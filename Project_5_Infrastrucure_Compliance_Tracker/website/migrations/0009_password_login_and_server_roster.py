from django.db import migrations, models


def backfill_usernames(apps, schema_editor):
    WorkspaceUser = apps.get_model("website", "WorkspaceUser")
    database = schema_editor.connection.alias
    for user in WorkspaceUser.objects.using(database).all().iterator():
        user.username = user.user_id.casefold()
        user.save(update_fields=["username"])


class Migration(migrations.Migration):
    dependencies = [("website", "0008_workspace_parameter_permissions")]

    operations = [
        migrations.AddField(
            model_name="masterworkbook",
            name="roster_details",
            field=models.JSONField(default=list),
        ),
        migrations.AddField(
            model_name="masterworkbook",
            name="new_server_details",
            field=models.JSONField(default=list),
        ),
        migrations.AddField(
            model_name="workspaceuser",
            name="username",
            field=models.CharField(blank=True, max_length=150),
        ),
        migrations.AddField(
            model_name="workspaceuser",
            name="password_hash",
            field=models.CharField(blank=True, max_length=128),
        ),
        migrations.AddField(
            model_name="workspaceuser",
            name="must_change_password",
            field=models.BooleanField(default=False),
        ),
        migrations.RunPython(backfill_usernames, migrations.RunPython.noop),
        migrations.AlterField(
            model_name="workspaceuser",
            name="username",
            field=models.CharField(max_length=150, unique=True),
        ),
    ]
