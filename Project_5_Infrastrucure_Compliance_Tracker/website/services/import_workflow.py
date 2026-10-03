"""Coordinate Excel parsing and source-independent snapshot persistence."""
from website.services.excel_import import preview_workbook
from website.services.snapshot_persistence import persist_snapshot


def import_snapshot(
    content,
    *,
    original_filename,
    parameter_name,
    snapshot_date,
    requested_status_column="",
    preview=None,
    sync_target_path="",
):
    if preview is None:
        preview = preview_workbook(content, requested_status_column)
    return persist_snapshot(
        preview,
        content,
        original_filename=original_filename,
        parameter_name=parameter_name,
        snapshot_date=snapshot_date,
        sync_target_path=sync_target_path,
    )
