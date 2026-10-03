"""Import the user's multi-tab compliance master workbook safely."""
from collections import Counter, defaultdict
from datetime import date, datetime
from hashlib import sha256
from io import BytesIO
from pathlib import Path
import json
import re
import zipfile

from django.conf import settings
from django.db import transaction
from django.utils import timezone
from openpyxl import load_workbook

from website.models import AuditLog, ImportBatch, ImportRecord, MasterWorkbook, Parameter
from website.services.errors import ImportValidationError
from website.services.snapshot_persistence import persist_snapshot


SHEET_PARAMETER_NAMES = {
    "CS": "CS (Craft)",
    "TGIM": "TGIM",
    "Splunk": "Splunk",
    "Logger": "Logger",
    "RSA": "RSA",
    "SNOW": "SNOW",
}
MAX_PACKAGE_BYTES = 250 * 1024 * 1024
MAX_ARCHIVE_MEMBERS = 2000
MAX_WORKSHEET_ROWS = 100_000
MAX_WORKSHEET_COLUMNS = 200
MAX_WORKSHEET_CELLS = 2_000_000
BLOCKED_PACKAGE_PARTS = (
    "xl/vbaproject.bin",
    "xl/externallinks/",
    "xl/embeddings/",
    "xl/activex/",
    "xl/connections.xml",
)
MONTHS = {
    name.casefold(): number
    for number, names in enumerate((
        ("january", "jan"), ("february", "feb"), ("march", "mar"),
        ("april", "apr"), ("may",), ("june", "jun"), ("july", "jul"),
        ("august", "aug"), ("september", "sep"), ("october", "oct"),
        ("november", "nov"), ("december", "dec"),
    ), start=1) for name in names
}
STATUS_MAP = {
    "compliant": "compliant",
    "non-compliant": "non_compliant",
    "not applicable": "not_applicable",
    "n/a": "not_applicable",
}
MISSING_IP_VALUES = {"", "nan", "not found", "n/a", "na", "none"}


def _as_json(value):
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def _text(value):
    return "" if value is None else str(value).strip()


def _month_from_heading(heading):
    match = re.search(r"\b([A-Za-z]+)[\s,_-]+(20\d{2})\b", heading)
    if not match:
        return None
    month = MONTHS.get(match.group(1).casefold())
    return date(int(match.group(2)), month, 1) if month else None


def _inspect_package(content):
    if len(content) > 25 * 1024 * 1024:
        raise ImportValidationError("The master workbook is larger than the 25 MB upload limit.")
    if not zipfile.is_zipfile(BytesIO(content)):
        raise ImportValidationError("This file is not a readable .xlsx workbook.")
    try:
        with zipfile.ZipFile(BytesIO(content), "r") as archive:
            members = archive.infolist()
            if len(members) > MAX_ARCHIVE_MEMBERS:
                raise ImportValidationError("The workbook contains too many internal files to process safely.")
            if sum(item.file_size for item in members) > MAX_PACKAGE_BYTES:
                raise ImportValidationError("The expanded workbook is too large to process safely.")
            if any(item.flag_bits & 0x1 for item in members):
                raise ImportValidationError("Encrypted workbooks are not supported.")
            names = [item.filename.casefold() for item in members]
            if "xl/workbook.xml" not in names:
                raise ImportValidationError("This ZIP file does not contain a valid Excel workbook.")
            if any(name.endswith(part) or name.startswith(part) for name in names for part in BLOCKED_PACKAGE_PARTS):
                raise ImportValidationError("Workbooks with macros, external links, embedded files, or active controls are not supported.")
    except ImportValidationError:
        raise
    except (OSError, zipfile.BadZipFile, RuntimeError) as exc:
        raise ImportValidationError("The workbook package could not be inspected safely.") from exc


def _headers(sheet):
    values = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if values is None:
        return []
    headings = ["" if value is None else str(value).strip() for value in values]
    while headings and not headings[-1]:
        headings.pop()
    if not headings or any(not heading for heading in headings):
        raise ImportValidationError(f"Worksheet '{sheet.title}' has a blank heading.")
    if len({heading.casefold() for heading in headings}) != len(headings):
        raise ImportValidationError(f"Worksheet '{sheet.title}' has repeated headings.")
    return headings


def _preview_for_month(sheet_name, parameter_name, headers, main_rows, status_rows, status_index, status_heading):
    rows = []
    groups = defaultdict(list)
    status_counts = Counter()
    missing_hostname_rows = missing_ip_rows = duplicate_rows = unmapped_rows = 0
    seen_source_rows = set()

    for main_row, sheet_row in zip(main_rows, status_rows):
        source_row = main_row[0]
        host = _text(main_row[1])
        ip_raw = _text(main_row[2])
        ip_address = "" if ip_raw.casefold() in MISSING_IP_VALUES else ip_raw
        raw_cell = sheet_row[status_index]
        if getattr(raw_cell, "data_type", "") == "f":
            raise ImportValidationError(
                f"'{status_heading}' contains a formula at Excel row {source_row}. Status cells must contain saved values."
            )
        raw_status = _text(raw_cell.value)
        if not host and not ip_raw and not raw_status:
            continue
        status_counts[raw_status or "(blank)"] += 1
        normalized = host.casefold() if host else ""
        mapped = STATUS_MAP.get(raw_status.casefold(), "unmapped")
        if mapped == "unmapped":
            unmapped_rows += 1
        source_data = {}
        for index, heading in enumerate(headers):
            cell = sheet_row[index]
            if heading.casefold() == "hostname":
                value = host
            elif heading.casefold() == "ip address":
                value = ip_raw
            elif getattr(cell, "data_type", "") == "f":
                value = None
            else:
                value = _as_json(cell.value)
            source_data[heading] = value
        if len(main_row) > 3:
            for heading, value in main_row[3].items():
                source_data.setdefault(heading, value)

        signature = json.dumps(source_data, ensure_ascii=False, sort_keys=True, default=str)
        is_duplicate = signature in seen_source_rows
        if is_duplicate:
            duplicate_rows += 1
        seen_source_rows.add(signature)
        issue_parts = []
        if not host:
            missing_hostname_rows += 1
            issue_parts.append("Missing Host Name; this row cannot be linked to a server.")
        if not ip_address:
            missing_ip_rows += 1
            issue_parts.append("Missing IP Address.")
        if mapped == "unmapped":
            issue_parts.append("Unmapped source status label; current state remains Unmapped.")
        issue = " ".join(issue_parts)
        row = {
            "source_row": source_row,
            "hostname": host,
            "normalized_hostname": normalized,
            "ip_address": ip_address,
            "operating_system": "",
            "environment": "",
            "raw_status": raw_status,
            "mapped_status": mapped,
            "source_data": source_data,
            "duplicate_row": is_duplicate,
            "issue": issue,
        }
        rows.append(row)
        if normalized:
            groups[normalized].append(row)

    conflicts = {
        hostname for hostname, group in groups.items()
        if len({(item["raw_status"], item["mapped_status"]) for item in group}) > 1
    }
    if conflicts:
        for row in rows:
            if row["normalized_hostname"] in conflicts:
                row["issue"] = (row["issue"] + " Conflicting statuses for this Host Name in one snapshot.").strip()
    total_servers = len(groups)
    return {
        "source_sheet": sheet_name,
        "parameter_name": parameter_name,
        "status_column": status_heading,
        "headers": headers,
        "rows": rows,
        "total_rows": len(rows),
        "unique_servers": total_servers,
        "missing_hostname_rows": missing_hostname_rows,
        "missing_ip_rows": missing_ip_rows,
        "duplicate_rows": duplicate_rows,
        "conflicting_server_groups": len(conflicts),
        "conflicting_hostnames": conflicts,
        "unmapped_status_rows": unmapped_rows,
        "multiple_ip_groups": 0,
        "status_counts": dict(status_counts),
    }


def _date_from_cell(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = _text(value)
    if not text:
        return None
    for parser in (date.fromisoformat,):
        try:
            return parser(text[:10])
        except ValueError:
            pass
    for pattern in ("%d/%m/%Y", "%m/%d/%Y", "%d-%b-%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, pattern).date()
        except ValueError:
            continue
    return None


def _rows_aligned_to_inventory(sheet, headers, main_rows):
    rows = []
    iterator = iter(sheet.iter_rows(min_row=2, values_only=False))
    for main_row in main_rows:
        row = next(iterator, None)
        if row is None:
            raise ImportValidationError(f"The '{sheet.title}' worksheet has fewer rows than Main Data.")
        row = list(row[:len(headers)])
        if len(row) < len(headers):
            row.extend([None] * (len(headers) - len(row)))
        rows.append(row)
    return rows


def preview_master_workbook(content):
    """Return one normalized preview per dated Splunk status column."""
    _inspect_package(content)
    try:
        workbook = load_workbook(BytesIO(content), read_only=True, data_only=False)
    except Exception as exc:
        raise ImportValidationError("The master .xlsx workbook could not be opened.") from exc
    try:
        total_cells = 0
        for worksheet in workbook.worksheets:
            rows = worksheet.max_row or 0
            columns = worksheet.max_column or 0
            total_cells += rows * columns
            if rows > MAX_WORKSHEET_ROWS or columns > MAX_WORKSHEET_COLUMNS or total_cells > MAX_WORKSHEET_CELLS:
                raise ImportValidationError(
                    f"The workbook exceeds the safe limit of {MAX_WORKSHEET_ROWS:,} rows, "
                    f"{MAX_WORKSHEET_COLUMNS} columns, or {MAX_WORKSHEET_CELLS:,} total cells."
                )
        required_sheets = {"Main Data", "Splunk"}
        if not required_sheets.issubset(workbook.sheetnames):
            raise ImportValidationError("The workbook must contain 'Main Data' and 'Splunk' worksheets.")
        inventory = workbook["Main Data"]
        splunk = workbook["Splunk"]
        inventory_headers = _headers(inventory)
        if "Hostname" not in inventory_headers or "IP Address" not in inventory_headers:
            raise ImportValidationError("'Main Data' needs 'Hostname' and 'IP Address' columns.")
        host_index = inventory_headers.index("Hostname")
        ip_index = inventory_headers.index("IP Address")
        main_rows = []
        inventory_iter = inventory.iter_rows(min_row=2, values_only=True)
        for source_row, values in enumerate(inventory_iter, start=2):
            if source_row > inventory.max_row:
                break
            values = list(values)
            host = values[host_index] if host_index < len(values) else None
            ip = values[ip_index] if ip_index < len(values) else None
            master_data = {
                heading: _as_json(values[index] if index < len(values) else None)
                for index, heading in enumerate(inventory_headers)
            }
            main_rows.append((source_row, host, ip, master_data))
        inventory_details = {}
        for _, host_value, ip_value, _ in main_rows:
            hostname = _text(host_value)
            normalized_hostname = hostname.casefold()
            if not normalized_hostname:
                continue
            entry = inventory_details.setdefault(normalized_hostname, {
                "normalized_hostname": normalized_hostname,
                "hostname": hostname,
                "ip_addresses": [],
            })
            ip_address = _text(ip_value)
            if ip_address.casefold() not in MISSING_IP_VALUES and ip_address not in entry["ip_addresses"]:
                entry["ip_addresses"].append(ip_address)
        parameter_sheets = []
        parameter_names = []
        reserved_sheets = {"Main Data", "Dashboard", "README", "Compliance Summary"}
        for sheet in workbook.worksheets:
            if sheet.title in reserved_sheets:
                continue
            parameter_name = SHEET_PARAMETER_NAMES.get(sheet.title)
            if parameter_name is None:
                raw_headers = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
                visible = {str(item).strip().casefold() for item in raw_headers if item is not None}
                if not {"hostname", "ip address"}.issubset(visible) or not any("status" in item for item in visible):
                    continue
                parameter_name = sheet.title.strip()
            if not parameter_name or len(parameter_name) > 120:
                continue
            headers = _headers(sheet)
            if not {"Hostname", "IP Address"}.issubset(set(headers)):
                continue
            parameter_sheets.append((sheet.title, parameter_name, sheet, headers))
            if parameter_name not in parameter_names:
                parameter_names.append(parameter_name)

        previews = []
        for sheet_name, parameter_name, sheet, headers in parameter_sheets:
            sheet_rows = _rows_aligned_to_inventory(sheet, headers, main_rows)
            status_columns = [
                (index, heading, _month_from_heading(heading))
                for index, heading in enumerate(headers)
                if "status" in heading.casefold()
            ]
            if not status_columns:
                continue
            check_date_index = next((index for index, heading in enumerate(headers) if heading.casefold() == "check date"), None)
            latest_check_date = None
            if check_date_index is not None:
                for row in sheet_rows:
                    cell = row[check_date_index]
                    value = None if cell is None or getattr(cell, "data_type", "") == "f" else cell.value
                    parsed = _date_from_cell(value)
                    if parsed and (latest_check_date is None or parsed > latest_check_date):
                        latest_check_date = parsed
            for status_index, status_heading, snapshot_date in status_columns:
                if snapshot_date is None:
                    snapshot_date = latest_check_date or timezone.localdate()
                preview = _preview_for_month(
                    sheet_name,
                    parameter_name,
                    headers,
                    iter(main_rows),
                    iter(sheet_rows),
                    status_index,
                    status_heading,
                )
                preview["snapshot_date"] = snapshot_date
                if any(raw != "(blank)" and count for raw, count in preview["status_counts"].items()):
                    previews.append(preview)
        previews.sort(key=lambda item: item["snapshot_date"])
        return {
            "previews": previews,
            "parameter_names": parameter_names,
            "hostnames": {
                str(item[1]).strip().casefold()
                for item in main_rows
                if item[1] is not None and str(item[1]).strip()
            },
            "host_details": [
                {
                    **entry,
                    "ip_address": ", ".join(entry["ip_addresses"]),
                }
                for entry in inventory_details.values()
            ],
            "source_sheets": list(workbook.sheetnames),
        }
    finally:
        workbook.close()


def validate_master_target_path(path, content, source_sheets):
    path = (path or "").strip()
    if not path:
        return ""
    target = Path(path).expanduser()
    if not target.is_absolute():
        raise ImportValidationError("Enter the full path to the original .xlsx master workbook.")
    if target.suffix.casefold() != ".xlsx" or not target.is_file():
        raise ImportValidationError("The master workbook path must point to an existing .xlsx file.")
    media_root = Path(settings.MEDIA_ROOT).resolve()
    target = target.resolve()
    if target == media_root or media_root in target.parents:
        raise ImportValidationError("The master workbook must be outside the app's uploads folder.")
    try:
        if sha256(target.read_bytes()).digest() != sha256(content).digest():
            raise ImportValidationError("The selected master workbook does not match the file you uploaded.")
        workbook = load_workbook(target, read_only=True, data_only=False)
    except ImportValidationError:
        raise
    except Exception as exc:
        raise ImportValidationError("The selected master workbook could not be opened.") from exc
    try:
        if not set(source_sheets).issubset(workbook.sheetnames):
            raise ImportValidationError("The selected workbook's worksheet names do not match the uploaded master.")
    finally:
        workbook.close()
    return str(target)


def parameter_sheet_name(parameter_name):
    reverse = {value.casefold(): key for key, value in SHEET_PARAMETER_NAMES.items()}
    return reverse.get(parameter_name.casefold(), parameter_name)


def read_master_parameter_sheet(parameter_name):
    """Return the real worksheet headings and formula-safe row values."""
    config = MasterWorkbook.objects.first()
    if not config or not config.path:
        raise ImportValidationError("The owner has not configured a master workbook path yet.")
    target = Path(config.path)
    try:
        content = target.read_bytes()
    except OSError as exc:
        raise ImportValidationError("The configured master workbook is not available on this computer.") from exc
    if sha256(content).hexdigest() != config.content_sha256:
        raise ImportValidationError("The master workbook changed outside the web app. The owner must re-import it first.")
    _inspect_package(content)
    try:
        workbook = load_workbook(BytesIO(content), read_only=False, data_only=False)
    except Exception as exc:
        raise ImportValidationError("The configured master workbook could not be opened safely.") from exc
    try:
        sheet_name = parameter_sheet_name(parameter_name)
        if sheet_name not in workbook.sheetnames:
            raise ImportValidationError(f"The master workbook has no worksheet for {parameter_name}.")
        sheet = workbook[sheet_name]
        headers = _headers(sheet)
        if "Main Data" not in workbook.sheetnames:
            raise ImportValidationError("The master workbook is missing its Main Data worksheet.")
        inventory = workbook["Main Data"]
        inventory_headers = _headers(inventory)
        index_by_heading = {heading: index + 1 for index, heading in enumerate(inventory_headers)}
        rows = []
        for source_row in range(2, max(sheet.max_row, inventory.max_row) + 1):
            values = {}
            formula_columns = []
            for column, heading in enumerate(headers, start=1):
                cell = sheet.cell(row=source_row, column=column)
                if cell.data_type == "f" and heading not in {"Hostname", "IP Address"}:
                    formula_columns.append(heading)
                if heading in {"Hostname", "IP Address"} and heading in index_by_heading:
                    inventory_cell = inventory.cell(row=source_row, column=index_by_heading[heading])
                    value = None if inventory_cell.data_type == "f" else inventory_cell.value
                else:
                    value = None if cell.data_type == "f" else cell.value
                values[heading] = _as_json(value)
            if not any(value not in (None, "") for value in values.values()):
                continue
            rows.append({"source_row": source_row, "values": values, "formula_columns": formula_columns})
        return {"sheet_name": sheet_name, "headers": headers, "rows": rows}
    finally:
        workbook.close()


def _preview_matches_batch(preview, batch):
    if list(batch.source_columns) != list(preview["headers"]):
        return False
    prior = {
        record.source_row: (
            record.hostname.casefold(), record.ip_address, record.raw_status,
            record.source_data or {},
        )
        for record in ImportRecord.objects.filter(batch=batch).only(
            "hostname", "ip_address", "source_row", "raw_status", "source_data"
        )
    }
    incoming = {
        row["source_row"]: (
            row["hostname"].casefold(), row["ip_address"], row["raw_status"],
            row["source_data"],
        )
        for row in preview["rows"]
    }
    return prior == incoming


@transaction.atomic
def import_master_workbook(content, *, original_filename, target_path=""):
    """Import new or changed monthly columns and activate this workbook as the source."""
    inspected = preview_master_workbook(content)
    source_hash = sha256(content).hexdigest()
    config = MasterWorkbook.objects.select_for_update().first()
    effective_target = target_path.strip() or (config.path if config else "")
    validated_path = validate_master_target_path(
        effective_target, content, inspected["source_sheets"]
    ) if effective_target else ""
    previous_master_batch = ImportBatch.objects.filter(
        is_master_source=True, parameter__name="Splunk"
    ).order_by(
        "-snapshot_date", "-imported_at"
    ).first()
    previous_hostnames = set()
    previous_roster_count = None
    if config and previous_master_batch:
        previous_hostnames = {
            str(entry.get("normalized_hostname") or entry.get("hostname", "")).strip().casefold()
            for entry in config.roster_details
            if entry.get("normalized_hostname") or entry.get("hostname")
        }
        if not previous_hostnames:
            previous_hostnames = {
                value.strip().casefold()
                for value in previous_master_batch.records.exclude(hostname="").values_list("hostname", flat=True)
            }
        previous_roster_count = len(previous_hostnames)

    prepared = []
    for preview in inspected["previews"]:
        existing = ImportBatch.objects.filter(
            is_master_source=True,
            parameter__name=preview["parameter_name"],
            snapshot_date=preview["snapshot_date"],
        ).order_by("-imported_at").first()
        if existing and _preview_matches_batch(preview, existing):
            continue
        prepared.append(preview)

    if not inspected["hostnames"]:
        raise ImportValidationError("The Main Data inventory does not contain any server hostnames.")

    ImportBatch.objects.filter(is_master_source=False).update(is_current_source=False)
    imported = []
    for preview in prepared:
        batch = persist_snapshot(
            preview,
            content,
            original_filename=original_filename,
            parameter_name=preview["parameter_name"],
            snapshot_date=preview["snapshot_date"],
            sync_target_path=validated_path,
            is_current_source=True,
            is_master_source=True,
            allow_same_date=True,
        )
        imported.append(batch)

    for parameter_name in inspected["parameter_names"]:
        Parameter.objects.get_or_create(name=parameter_name)

    all_master_batches = ImportBatch.objects.filter(is_master_source=True).order_by(
        "-snapshot_date", "-imported_at"
    )
    latest_batch = all_master_batches.first()
    baseline_batch = ImportBatch.objects.filter(
        is_master_source=True, parameter__name="Splunk"
    ).order_by("-snapshot_date", "-imported_at").first()
    if latest_batch is None:
        raise ImportValidationError("No monthly Splunk status data was found to set as the dashboard source.")
    new_server_count = (
        len(inspected["hostnames"] - previous_hostnames)
        if config is not None and previous_master_batch is not None
        else None
    )
    new_server_details = [
        entry for entry in inspected["host_details"]
        if entry["normalized_hostname"] not in previous_hostnames
    ] if new_server_count is not None else []
    if config is None:
        config = MasterWorkbook.objects.create(
            path=validated_path,
            original_filename=original_filename[:255],
            content_sha256=source_hash,
            parameter_names=inspected["parameter_names"],
            roster_details=inspected["host_details"],
            roster_count=len(inspected["hostnames"]),
            previous_roster_count=previous_roster_count,
            new_server_count=new_server_count,
            new_server_details=new_server_details,
            baseline_batch=baseline_batch,
        )
    else:
        config.path = validated_path or config.path
        config.original_filename = original_filename[:255]
        config.content_sha256 = source_hash
        config.parameter_names = inspected["parameter_names"]
        config.roster_details = inspected["host_details"]
        config.roster_count = len(inspected["hostnames"])
        config.previous_roster_count = previous_roster_count
        config.new_server_count = new_server_count
        config.new_server_details = new_server_details
        config.baseline_batch = baseline_batch
        config.save()

    AuditLog.objects.create(
        action="master_workbook_imported",
        detail=(
            f"Master workbook {original_filename[:255]} set as the current source; "
            f"{len(imported)} dated status column(s) imported and "
            f"{len(inspected['hostnames']):,} inventory host(s) recorded."
        ),
        import_batch=latest_batch,
    )
    return {
        "config": config,
        "imported": imported,
        "latest_batch": latest_batch,
        "new_server_count": new_server_count,
        "roster_count": len(inspected["hostnames"]),
    }
