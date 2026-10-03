from django.db import migrations, models
import django.db.models.deletion


def copy_existing_grants(apps, schema_editor):
    WorkspaceUser = apps.get_model("website", "WorkspaceUser")
    Parameter = apps.get_model("website", "Parameter")
    Permission = apps.get_model("website", "WorkspaceParameterPermission")
    database = schema_editor.connection.alias
    parameters = list(Parameter.objects.using(database).all())
    for user in WorkspaceUser.objects.using(database).all():
        if not (user.can_read or user.can_write or user.can_execute):
            continue
        Permission.objects.using(database).bulk_create([
            Permission(
                user_id=user.pk,
                parameter_id=parameter.pk,
                can_read=user.can_read or user.can_write or user.can_execute,
                can_write=user.can_write,
                can_execute=user.can_execute,
            )
            for parameter in parameters
        ])


class Migration(migrations.Migration):
    dependencies = [("website", "0007_master_workbook_source")]

    operations = [
        migrations.CreateModel(
            name="WorkspaceParameterPermission",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("can_read", models.BooleanField(default=False)),
                ("can_write", models.BooleanField(default=False)),
                ("can_execute", models.BooleanField(default=False)),
                ("parameter", models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name="workspace_permissions", to="website.parameter")),
                ("user", models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name="parameter_permissions", to="website.workspaceuser")),
            ],
            options={"ordering": ["user__user_id", "parameter__name"]},
        ),
        migrations.AddConstraint(
            model_name="workspaceparameterpermission",
            constraint=models.UniqueConstraint(fields=("user", "parameter"), name="unique_user_parameter_permission"),
        ),
        migrations.RunPython(copy_existing_grants, migrations.RunPython.noop),
    ]
