from django.conf import settings


def workspace_access(request):
    return {
        "workspace_user_id": getattr(request, "workspace_user_id", ""),
        "workspace_username": getattr(request, "workspace_username", ""),
        "workspace_is_owner": getattr(request, "workspace_is_owner", False),
        "workspace_permissions": getattr(request, "workspace_permissions", {}),
        "workspace_owner_id": settings.WORKSPACE_OWNER_ID,
    }
