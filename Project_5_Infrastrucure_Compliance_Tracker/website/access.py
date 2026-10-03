from urllib.parse import quote

from django.conf import settings
from django.http import HttpResponseForbidden
from django.shortcuts import redirect
from django.urls import reverse

from .models import WorkspaceParameterPermission, WorkspaceUser


def allowed_parameter_ids(request, capability="read"):
    if getattr(request, "workspace_is_owner", False):
        return None
    return {
        parameter_id
        for parameter_id, permissions in getattr(request, "workspace_parameter_permissions", {}).items()
        if permissions.get(capability, False)
    }


def has_parameter_permission(request, parameter_id, capability="read"):
    allowed = allowed_parameter_ids(request, capability)
    return allowed is None or parameter_id in allowed


class WorkspaceAccessMiddleware:
    """Reload user grants on each request so owner changes take effect immediately."""

    public_views = {"login", "logout"}
    owner_views = {
        "access_control", "column_settings", "set_import_sync_target",
        "delete_import", "delete_history_event", "clear_all_data", "imports_master",
    }
    execute_views = {"imports_new", "imports_master", "export_non_compliant"}
    write_views = {"source_record_edit"}
    write_on_post_views = {"server_detail", "set_import_sync_target", "source_record_edit"}

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        user_id = request.session.get("workspace_user_id", "")
        if user_id and request.session.get("workspace_auth_method") != "password":
            request.session.pop("workspace_user_id", None)
            request.session.pop("workspace_auth_method", None)
            request.session.pop("workspace_must_change_password", None)
            user_id = ""
        request.workspace_user_id = ""
        request.workspace_username = ""
        request.workspace_is_owner = False
        request.workspace_permissions = {"read": False, "write": False, "execute": False}
        request.workspace_must_change_password = False
        request.workspace_user = None
        request.workspace_parameter_permissions = {}

        if user_id:
            grant = WorkspaceUser.objects.filter(user_id=user_id, is_active=True).first()
            if grant:
                request.workspace_user = grant
                request.workspace_username = grant.username
                request.workspace_is_owner = user_id == settings.WORKSPACE_OWNER_ID
                if request.workspace_is_owner:
                    # Renew the owner's persistent local sign-in when the session is used.
                    request.session.set_expiry(settings.WORKSPACE_OWNER_SESSION_AGE)
                request.workspace_must_change_password = grant.must_change_password
                if request.workspace_must_change_password:
                    request.session["workspace_must_change_password"] = True
                else:
                    request.session.pop("workspace_must_change_password", None)
                request.workspace_permissions = {
                    "read": request.workspace_is_owner or grant.can_read,
                    "write": request.workspace_is_owner or grant.can_write,
                    "execute": request.workspace_is_owner or grant.can_execute,
                }
                if request.workspace_is_owner:
                    request.workspace_parameter_permissions = "all"
                else:
                    request.workspace_parameter_permissions = {
                        row.parameter_id: {
                            "read": row.can_read,
                            "write": row.can_write,
                            "execute": row.can_execute,
                        }
                        for row in WorkspaceParameterPermission.objects.filter(user=grant)
                    }
            else:
                request.session.pop("workspace_user_id", None)
                request.session.pop("workspace_auth_method", None)
                request.session.pop("workspace_must_change_password", None)
                user_id = ""

        if user_id and request.workspace_user:
            request.workspace_user_id = user_id
        return self.get_response(request)

    def process_view(self, request, view_func, view_args, view_kwargs):
        if request.META.get("REMOTE_ADDR") not in {"127.0.0.1", "::1"}:
            return HttpResponseForbidden("This local workspace accepts browser connections from this computer only.")
        url_name = getattr(getattr(request, "resolver_match", None), "url_name", None)
        if url_name in self.public_views:
            return None
        if not request.workspace_user_id:
            login_url = reverse("login")
            return redirect(f"{login_url}?next={quote(request.get_full_path())}")
        if request.workspace_must_change_password and url_name not in {"password_change", "logout"}:
            return redirect("password_change")
        if request.workspace_is_owner:
            return None
        if url_name in self.owner_views:
            return HttpResponseForbidden("Only the workspace owner can manage access or workbook columns.")
        if not request.workspace_permissions["read"]:
            return HttpResponseForbidden("Your user ID does not have Read access.")
        if url_name in self.execute_views and not request.workspace_permissions["execute"]:
            return HttpResponseForbidden("Your user ID does not have Execute access.")
        if (url_name in self.write_views or (request.method == "POST" and url_name in self.write_on_post_views)):
            if not request.workspace_permissions["write"]:
                return HttpResponseForbidden("Your user ID does not have Write access.")
        return None
