"""Apply normalized snapshot rows to the relational compliance history."""
from collections import defaultdict
from hashlib import sha256
import uuid

from django.core.files.base import ContentFile
from django.db import transaction

from website.models import (
    AuditLog,
    ComplianceState,
    ComplianceStatus,
    ImportBatch,
    ImportRecord,
    Parameter,
    Server,
    StatusEvent,
)
from website.services.errors import ImportValidationError


def _chunks(values, size=500):
    values = list(values)
    for index in range(0, len(values), size):
        yield values[index:index + size]


@transaction.atomic
def persist_snapshot(
    preview, content, *, original_filename, parameter_name, snapshot_date,
    sync_target_path="", is_current_source=True, is_master_source=False,
    allow_same_date=False,
):
    """Persist normalized source rows and apply current-state transitions atomically."""
    file_hash = sha256(content).hexdigest()
    parameter, _ = Parameter.objects.get_or_create(name=parameter_name.strip())
    duplicate = ImportBatch.objects.filter(
        parameter=parameter,
        snapshot_date=snapshot_date,
        file_sha256=file_hash,
        result__in=[ImportBatch.Result.COMPLETED, ImportBatch.Result.PARTIAL],
    ).first()
    if duplicate:
        raise ImportValidationError(
            f"This same workbook was already imported for {parameter.name} on {snapshot_date:%d %b %Y}."
        )

    batch = ImportBatch.objects.create(
        parameter=parameter,
        original_filename=original_filename[:255],
        file_sha256=file_hash,
        source_sheet=preview["source_sheet"],
        status_column=preview["status_column"],
        source_columns=preview["headers"],
        sync_target_path=sync_target_path,
        is_current_source=is_current_source,
        is_master_source=is_master_source,
        snapshot_date=snapshot_date,
        total_rows=preview["total_rows"],
        unique_servers=preview["unique_servers"],
        missing_hostname_rows=preview["missing_hostname_rows"],
        missing_ip_rows=preview["missing_ip_rows"],
        duplicate_rows=preview["duplicate_rows"],
        conflicting_server_groups=preview["conflicting_server_groups"],
        unmapped_status_rows=preview["unmapped_status_rows"],
        result=(ImportBatch.Result.PARTIAL if (
            preview["missing_hostname_rows"] or preview["missing_ip_rows"]
            or preview["conflicting_server_groups"] or preview["unmapped_status_rows"]
        ) else ImportBatch.Result.COMPLETED),
    )
    safe_name = f"{uuid.uuid4().hex}.xlsx"
    batch.source_file.save(safe_name, ContentFile(content), save=True)
    batch.working_file.name = batch.source_file.name
    batch.working_sha256 = file_hash
    batch.save(update_fields=["working_file", "working_sha256"])

    normalized_names = {
        row["normalized_hostname"] for row in preview["rows"] if row["normalized_hostname"]
    }
    existing = {}
    for name_batch in _chunks(normalized_names):
        existing.update({
            server.normalized_hostname: server
            for server in Server.objects.filter(normalized_hostname__in=name_batch)
        })
    new_servers = [
        Server(hostname=name, normalized_hostname=name)
        for name in normalized_names if name not in existing
    ]
    if new_servers:
        Server.objects.bulk_create(new_servers, ignore_conflicts=True, batch_size=500)
    servers = {}
    for name_batch in _chunks(normalized_names):
        servers.update({
            server.normalized_hostname: server
            for server in Server.objects.filter(normalized_hostname__in=name_batch)
        })

    import_records = []
    for row in preview["rows"]:
        server = servers.get(row["normalized_hostname"])
        import_records.append(ImportRecord(
            batch=batch,
            source_row=row["source_row"],
            server=server,
            hostname=row["hostname"],
            ip_address=row["ip_address"],
            operating_system=row["operating_system"],
            environment=row["environment"],
            raw_status=row["raw_status"],
            mapped_status=row["mapped_status"],
            source_data=row["source_data"],
            duplicate_row=row["duplicate_row"],
            issue=row["issue"],
        ))
    ImportRecord.objects.bulk_create(import_records, batch_size=50)

    rows_by_host = defaultdict(list)
    for row in preview["rows"]:
        if row["normalized_hostname"]:
            rows_by_host[row["normalized_hostname"]].append(row)
    server_ids = [servers[name].id for name in rows_by_host]
    current_states = {}
    prior_non_compliant_server_ids = set()
    for id_batch in _chunks(server_ids):
        current_states.update({
            state.server.normalized_hostname: state
            for state in ComplianceState.objects.filter(
                parameter=parameter, server_id__in=id_batch
            ).select_related("server")
        })
        prior_non_compliant_server_ids.update(
            StatusEvent.objects.filter(
                parameter=parameter,
                server_id__in=id_batch,
                new_status=ComplianceStatus.NON_COMPLIANT,
            ).values_list("server_id", flat=True).distinct()
        )
    server_updates = []
    state_updates = []
    new_states = []
    new_events = []
    conflicting_hostnames = preview["conflicting_hostnames"]
    non_advancing_server_groups = 0
    for normalized_name, group in rows_by_host.items():
        server = servers[normalized_name]
        if server.first_seen_on is None or snapshot_date < server.first_seen_on:
            server.first_seen_on = snapshot_date
        if server.last_seen_on is None or snapshot_date > server.last_seen_on:
            server.last_seen_on = snapshot_date
            ips = {row["ip_address"] for row in group if row["ip_address"]}
            operating_systems = {row["operating_system"] for row in group if row["operating_system"]}
            environments = {row["environment"] for row in group if row["environment"]}
            server.hostname = group[0]["hostname"]
            server.current_ip = next(iter(ips)) if len(ips) == 1 else ""
            server.ip_ambiguous = len(ips) > 1
            server.operating_system = next(iter(operating_systems)) if len(operating_systems) == 1 else ""
            server.environment = next(iter(environments)) if len(environments) == 1 else ""
            server_updates.append(server)

        if normalized_name in conflicting_hostnames:
            continue

        latest_row = group[0]
        new_status = latest_row["mapped_status"]
        state = current_states.get(normalized_name)
        if state and (
            snapshot_date < state.observed_on
            or (snapshot_date == state.observed_on and not allow_same_date)
        ):
            non_advancing_server_groups += 1
            continue
        if state is None:
            new_states.append(ComplianceState(
                server=server,
                parameter=parameter,
                status=new_status,
                raw_status=latest_row["raw_status"],
                latest_import=batch,
                observed_on=snapshot_date,
            ))
            new_events.append(StatusEvent(
                server=server,
                parameter=parameter,
                import_batch=batch,
                previous_status="",
                new_status=new_status,
                kind=StatusEvent.Kind.INITIAL,
                event_date=snapshot_date,
                reason=(f"Source label: {latest_row['raw_status']}" if new_status == ComplianceStatus.UNMAPPED else ""),
            ))
            continue

        old_status = state.status
        recurrent = (
            old_status == ComplianceStatus.COMPLIANT
            and new_status == ComplianceStatus.NON_COMPLIANT
            and server.id in prior_non_compliant_server_ids
        )
        if old_status != new_status:
            new_events.append(StatusEvent(
                server=server,
                parameter=parameter,
                import_batch=batch,
                previous_status=old_status,
                new_status=new_status,
                kind=StatusEvent.Kind.RECURRENCE if recurrent else StatusEvent.Kind.CHANGE,
                event_date=snapshot_date,
                reason=(f"Source label: {latest_row['raw_status']}" if new_status == ComplianceStatus.UNMAPPED else ""),
            ))
        state.status = new_status
        state.raw_status = latest_row["raw_status"]
        state.latest_import = batch
        state.observed_on = snapshot_date
        if recurrent:
            state.recurrence_count += 1
        state_updates.append(state)

    if server_updates:
        Server.objects.bulk_update(
            server_updates,
            ["hostname", "current_ip", "ip_ambiguous", "operating_system", "environment", "first_seen_on", "last_seen_on"],
            batch_size=500,
        )
    if new_states:
        ComplianceState.objects.bulk_create(new_states, batch_size=500)
    if state_updates:
        ComplianceState.objects.bulk_update(
            state_updates,
            ["status", "raw_status", "latest_import", "observed_on", "recurrence_count"],
            batch_size=500,
        )
    if new_events:
        StatusEvent.objects.bulk_create(new_events, batch_size=1000)

    notes = []
    if preview["missing_hostname_rows"]:
        notes.append(f"{preview['missing_hostname_rows']} row(s) have no Host Name and were kept without a server link")
    if preview["missing_ip_rows"]:
        notes.append(f"{preview['missing_ip_rows']} row(s) have no IP Address")
    if preview["duplicate_rows"]:
        notes.append(f"{preview['duplicate_rows']} exact duplicate row(s) were retained")
    if preview["conflicting_server_groups"]:
        notes.append(f"{preview['conflicting_server_groups']} Host Name group(s) have conflicting statuses and did not update current state")
    if preview["unmapped_status_rows"]:
        notes.append(f"{preview['unmapped_status_rows']} row(s) have unmapped status labels and remain Unmapped")
    if preview["multiple_ip_groups"]:
        notes.append(f"{preview['multiple_ip_groups']} Host Name group(s) have multiple IP addresses; IP was left blank in the server summary")
    if non_advancing_server_groups:
        notes.append(f"{non_advancing_server_groups} server state(s) already had an equal or newer date; those current states were left unchanged")
        batch.non_advancing_server_groups = non_advancing_server_groups
        batch.result = ImportBatch.Result.PARTIAL
    batch.warning = "; ".join(notes)
    batch.save(update_fields=["warning", "non_advancing_server_groups", "result"])
    AuditLog.objects.create(
        action="excel_import_completed",
        detail=(f"Imported {batch.total_rows} rows for {parameter.name}; "
                f"snapshot {snapshot_date:%Y-%m-%d}; {batch.unmapped_status_rows} status row(s) left unmapped."),
        import_batch=batch,
    )
    return batch
