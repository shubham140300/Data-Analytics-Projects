from django.conf import settings
from django.contrib.auth.hashers import make_password
from django.db import migrations


def set_temporary_passwords(apps, schema_editor):
    WorkspaceUser = apps.get_model("website", "WorkspaceUser")
    database = schema_editor.connection.alias
    owner_id = settings.WORKSPACE_OWNER_ID
    temporary_password = settings.WORKSPACE_DEFAULT_PASSWORD
    if len(temporary_password) < 10:
        raise RuntimeError("WORKSPACE_DEFAULT_PASSWORD must be at least 10 characters long.")
    for user in WorkspaceUser.objects.using(database).exclude(user_id=owner_id).filter(password_hash="").iterator():
        user.password_hash = make_password(temporary_password)
        user.must_change_password = True
        user.save(update_fields=["password_hash", "must_change_password"])


class Migration(migrations.Migration):
    dependencies = [("website", "0009_password_login_and_server_roster")]

    operations = [migrations.RunPython(set_temporary_passwords, migrations.RunPython.noop)]
