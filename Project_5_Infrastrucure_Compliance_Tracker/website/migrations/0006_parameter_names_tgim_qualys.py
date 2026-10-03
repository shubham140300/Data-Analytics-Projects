from django.db import migrations


def rename_parameters(apps, schema_editor):
    Parameter = apps.get_model("website", "Parameter")
    database = schema_editor.connection.alias
    for old_name, new_name in (("EGM", "TGIM"), ("Colis", "Qualys")):
        old_parameter = Parameter.objects.using(database).filter(name=old_name).first()
        new_exists = Parameter.objects.using(database).filter(name=new_name).exists()
        if old_parameter and new_exists:
            raise RuntimeError(
                f"Both '{old_name}' and '{new_name}' parameters exist. Merge their records before renaming."
            )
        if old_parameter:
            old_parameter.name = new_name
            old_parameter.save(update_fields=["name"])


def reverse_parameter_names(apps, schema_editor):
    Parameter = apps.get_model("website", "Parameter")
    database = schema_editor.connection.alias
    for new_name, old_name in (("TGIM", "EGM"), ("Qualys", "Colis")):
        new_parameter = Parameter.objects.using(database).filter(name=new_name).first()
        old_exists = Parameter.objects.using(database).filter(name=old_name).exists()
        if new_parameter and old_exists:
            raise RuntimeError(
                f"Both '{new_name}' and '{old_name}' parameters exist. Merge their records before reverting."
            )
        if new_parameter:
            new_parameter.name = old_name
            new_parameter.save(update_fields=["name"])


class Migration(migrations.Migration):
    dependencies = [("website", "0005_workspaceuser")]

    operations = [migrations.RunPython(rename_parameters, reverse_parameter_names)]
