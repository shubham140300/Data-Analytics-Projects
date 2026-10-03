"""Owner-only removal of imported history and manual status events."""
from collections import defaultdict

from django.db import transaction
from django.db.models import Min, Max, Q

from website.models import (
    AuditLog,
    ComplianceState,
    ComplianceStatus,
    ImportBatch,
    ImportRecord,
    MasterWorkbook,
    Server,
    StatusEvent,
)


def _rebuild_parameter_history(parameter_id):
    """Recreate imported transitions and current states from remaining active snapshots."""
    batches = list(
        ImportBatch.objects.filter(
            parameter_id=parameter_id,
            is_current_source=True,
            result__in=[ImportBatch.Result.COMPLETED, ImportBatch.Result.PARTIAL],
        ).order_by("snapshot_date", "imported_at", "pk")
    )
    batch_ids = [batch.pk for batch in batches]
    manual_events = list(
        StatusEvent.objects.filter(parameter_id=parameter_id, import_batch__isnull=True)
        .select_related("server")
        .order_by("event_date", "created_at", "pk")
    )
    if batch_ids:
        StatusEvent.objects.filter(import_batch_id__in=batch_ids).delete()

    timelines = defaultdict(list)
    for batch in batches:
        grouped = defaultdict(list)
        for record in batch.records.filter(server__isnull=False).order_by("source_row").only(
            "server_id", "source_row", "raw_status", "mapped_status"
        ):
            grouped[record.server_id].append(record)
        for server_id, records in grouped.items():
            signatures = {(record.raw_status, record.mapped_status) for record in records}
            if len(signatures) > 1:
                continue
            record = records[0]
            timelines[server_id].append({
                "date": batch.snapshot_date,
                "created_at": batch.imported_at,
                "kind": "import",
                "batch": batch,
                "status": record.mapped_status,
                "raw_status": record.raw_status,
                "reason": (
                    f"Source label: {record.raw_status}"
                    if record.mapped_status == ComplianceStatus.UNMAPPED else ""
                ),
            })
    for event in manual_events:
        timelines[event.server_id].append({
            "date": event.event_date,
            "created_at": event.created_at,
            "kind": "manual",
            "event": event,
            "status": event.new_status,
            "raw_status": "Manual update",
            "reason": event.reason,
        })

    states = []
    generated_events = []
    changed_manual_events = []
    for server_id, entries in timelines.items():
        entries.sort(key=lambda item: (item["date"], item["created_at"], item["kind"] == "manual"))
        previous = ""
        ever_non_compliant = False
        recurrence_count = 0
        last_import = None
        observed_on = None
        raw_status = ""
        latest_status = ""
        for item in entries:
            new_status = item["status"]
            recurrent = (
                previous == ComplianceStatus.COMPLIANT
                and new_status == ComplianceStatus.NON_COMPLIANT
                and ever_non_compliant
            )
            if recurrent:
                recurrence_count += 1
            if item["kind"] == "manual":
                event = item["event"]
                event.previous_status = previous
                event.kind = StatusEvent.Kind.RECURRENCE if recurrent else StatusEvent.Kind.MANUAL
                changed_manual_events.append(event)
            elif not previous:
                generated_events.append(StatusEvent(
                    server_id=server_id,
                    parameter_id=parameter_id,
                    import_batch=item["batch"],
                    previous_status="",
                    new_status=new_status,
                    kind=StatusEvent.Kind.INITIAL,
                    event_date=item["date"],
                    reason=item["reason"],
                ))
            elif previous != new_status:
                generated_events.append(StatusEvent(
                    server_id=server_id,
                    parameter_id=parameter_id,
                    import_batch=item["batch"],
                    previous_status=previous,
                    new_status=new_status,
                    kind=StatusEvent.Kind.RECURRENCE if recurrent else StatusEvent.Kind.CHANGE,
                    event_date=item["date"],
                    reason=item["reason"],
                ))

            if new_status == ComplianceStatus.NON_COMPLIANT:
                ever_non_compliant = True
            previous = latest_status = new_status
            observed_on = item["date"]
            raw_status = item["raw_status"]
            if item["kind"] == "import":
                last_import = item["batch"]

        if latest_status:
            states.append(ComplianceState(
                server_id=server_id,
                parameter_id=parameter_id,
                status=latest_status,
                raw_status=raw_status,
                latest_import=last_import,
                observed_on=observed_on,
                recurrence_count=recurrence_count,
            ))

    ComplianceState.objects.filter(parameter_id=parameter_id).delete()
    if states:
        ComplianceState.objects.bulk_create(states, batch_size=500)
    if generated_events:
        StatusEvent.objects.bulk_create(generated_events, batch_size=1000)
    if changed_manual_events:
        StatusEvent.objects.bulk_update(changed_manual_events, ["previous_status", "kind"], batch_size=500)


def _refresh_server_metadata(server_ids):
    if not server_ids:
        return
    summaries = {
        item["server_id"]: item
        for item in ImportRecord.objects.filter(
            server_id__in=server_ids, batch__is_current_source=True
        ).values("server_id").annotate(
            first_seen=Min("batch__snapshot_date"), last_seen=Max("batch__snapshot_date")
        )
    }
    recent_records = ImportRecord.objects.filter(
        server_id__in=server_ids, batch__is_current_source=True
    ).select_related("batch").order_by(
        "server_id", "-batch__snapshot_date", "-batch__imported_at", "source_row"
    ).only("server_id", "ip_address", "batch__snapshot_date", "batch__imported_at")
    current_ips = {}
    latest_key = {}
    ip_sets = defaultdict(set)
    for record in recent_records:
        key = (record.batch.snapshot_date, record.batch.imported_at)
        if record.server_id not in latest_key:
            latest_key[record.server_id] = key
        if latest_key[record.server_id] != key:
            continue
        if record.ip_address:
            ip_sets[record.server_id].add(record.ip_address)
    for server_id, values in ip_sets.items():
        current_ips[server_id] = values

    updates = []
    for server in Server.objects.filter(pk__in=server_ids):
        summary = summaries.get(server.pk)
        ips = current_ips.get(server.pk, set())
        server.first_seen_on = summary["first_seen"] if summary else None
        server.last_seen_on = summary["last_seen"] if summary else None
        server.current_ip = next(iter(ips)) if len(ips) == 1 else ""
        server.ip_ambiguous = len(ips) > 1
        updates.append(server)
    if updates:
        Server.objects.bulk_update(
            updates,
            ["first_seen_on", "last_seen_on", "current_ip", "ip_ambiguous"],
            batch_size=500,
        )


@transaction.atomic
def delete_import_batch(batch_id, *, actor_id):
    batch = ImportBatch.objects.select_for_update().select_related("parameter").get(pk=batch_id)
    is_current = batch.is_current_source
    parameter_id = batch.parameter_id
    affected_server_ids = set(batch.records.exclude(server_id=None).values_list("server_id", flat=True))
    original_filename = batch.original_filename
    parameter_name = batch.parameter.name
    snapshot_date = batch.snapshot_date
    imported_at = batch.imported_at
    file_handles = [field for field in (batch.source_file, batch.working_file) if field and field.name]
    files = {field.name: field.storage for field in file_handles}

    if is_current:
        active_ids = list(ImportBatch.objects.filter(
            parameter_id=parameter_id, is_current_source=True
        ).values_list("pk", flat=True))
        StatusEvent.objects.filter(import_batch_id__in=active_ids).delete()
    else:
        StatusEvent.objects.filter(import_batch=batch).delete()

    batch.delete()

    if is_current:
        _rebuild_parameter_history(parameter_id)
        _refresh_server_metadata(affected_server_ids)

    master = MasterWorkbook.objects.select_for_update().first()
    if master and master.baseline_batch_id is None:
        master.baseline_batch = ImportBatch.objects.filter(
            is_master_source=True, parameter__name="Splunk", is_current_source=True
        ).order_by("-snapshot_date", "-imported_at").first()
        master.save(update_fields=["baseline_batch", "updated_at"])

    AuditLog.objects.create(
        action="import_batch_deleted",
        detail=(
            f"Owner {actor_id} deleted {parameter_name} snapshot dated {snapshot_date:%Y-%m-%d} "
            f"(uploaded {imported_at:%Y-%m-%d %H:%M}, file {original_filename}). "
            "The original workbook file was not changed."
        ),
    )

    def remove_unused_uploads():
        for name, storage in files.items():
            if not ImportBatch.objects.filter(Q(source_file=name) | Q(working_file=name)).exists():
                try:
                    storage.delete(name)
                except OSError:
                    pass

    transaction.on_commit(remove_unused_uploads)
    return {
        "parameter": parameter_name,
        "snapshot_date": snapshot_date,
        "filename": original_filename,
        "was_current": is_current,
    }


@transaction.atomic
def delete_manual_status_event(event_id, *, actor_id):
    event = StatusEvent.objects.select_for_update().select_related("server", "parameter").get(pk=event_id)
    if event.import_batch_id:
        raise ValueError("Imported status events are removed by deleting their import snapshot.")
    server_id = event.server_id
    detail = (
        f"Owner {actor_id} deleted a manual history event for {event.server.hostname} / "
        f"{event.parameter.name} dated {event.event_date:%Y-%m-%d}: "
        f"{event.previous_status or 'No prior status'} → {event.new_status}."
    )
    event.delete()
    AuditLog.objects.create(action="manual_history_event_deleted", detail=detail, server_id=server_id)


@transaction.atomic
def clear_compliance_workspace_data(*, actor_id):
    """Remove imported compliance data and app upload copies while retaining accounts."""
    batches = ImportBatch.objects.all().only("source_file", "working_file")
    batch_count = batches.count()
    server_count = Server.objects.count()
    event_count = StatusEvent.objects.count()
    upload_files = {}
    for batch in batches.iterator(chunk_size=500):
        for field in (batch.source_file, batch.working_file):
            if field and field.name:
                upload_files[field.name] = field.storage

    StatusEvent.objects.all().delete()
    ComplianceState.objects.all().delete()
    ImportBatch.objects.all().delete()
    Server.objects.all().delete()
    MasterWorkbook.objects.all().delete()
    AuditLog.objects.all().delete()
    AuditLog.objects.create(
        action="compliance_data_cleared",
        detail=(
            f"Owner {actor_id} cleared application compliance data: {batch_count} import(s), "
            f"{server_count} server(s), and {event_count} status event(s). "
            "User accounts and permission grants were retained; original workbook files were not changed."
        ),
    )

    def remove_uploads():
        for name, storage in upload_files.items():
            try:
                storage.delete(name)
            except OSError:
                pass

    transaction.on_commit(remove_uploads)
    return {
        "imports": batch_count,
        "servers": server_count,
        "events": event_count,
        "upload_copies": len(upload_files),
    }
