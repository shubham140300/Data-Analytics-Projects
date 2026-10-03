from django.urls import path

from . import views


urlpatterns = [
    path("login/", views.login, name="login"),
    path("logout/", views.logout, name="logout"),
    path("account/password/", views.password_change, name="password_change"),
    path("owner/access/", views.access_control, name="access_control"),
    path("owner/imports/<int:batch_id>/columns/", views.column_settings, name="column_settings"),
    path("", views.dashboard, name="dashboard"),
    path("new-servers/", views.new_servers, name="new_servers"),
    path("servers/", views.servers, name="servers"),
    path("servers/<int:server_id>/", views.server_detail, name="server_detail"),
    path("non-compliant/", views.non_compliant, name="non_compliant"),
    path("parameters/", views.parameters, name="parameters"),
    path("parameters/<int:parameter_id>/data/", views.parameter_data, name="parameter_data"),
    path("parameters/<int:parameter_id>/rows/<int:source_row>/edit/", views.master_sheet_row_edit, name="master_sheet_row_edit"),
    path("compliance-summary/", views.compliance_summary, name="compliance_summary"),
    path("imports/", views.imports, name="imports"),
    path("imports/<int:batch_id>/delete/", views.delete_import, name="delete_import"),
    path("imports/new/", views.imports_new, name="imports_new"),
    path("imports/master/", views.imports_master, name="imports_master"),
    path("imports/<int:batch_id>/", views.import_detail, name="import_detail"),
    path("imports/<int:batch_id>/workbook/", views.download_workbook, name="download_workbook"),
    path("imports/<int:batch_id>/sync-target/", views.set_import_sync_target, name="set_import_sync_target"),
    path("source-rows/<int:record_id>/edit/", views.source_record_edit, name="source_record_edit"),
    path("history/", views.history, name="history"),
    path("history/<int:event_id>/delete/", views.delete_history_event, name="delete_history_event"),
    path("history/clear-all/", views.clear_all_data, name="clear_all_data"),
    path("exports/non-compliant.xlsx", views.export_non_compliant, name="export_non_compliant"),
]
