from django.db import migrations


PARAMETERS = ["Splunk", "CrowdStrike", "RSA", "Snow", "EGM", "Logger", "Colis"]


def seed_parameters(apps, schema_editor):
    Parameter = apps.get_model("website", "Parameter")
    for name in PARAMETERS:
        Parameter.objects.get_or_create(name=name)


class Migration(migrations.Migration):
    dependencies = [("website", "0001_initial")]

    operations = [migrations.RunPython(seed_parameters, migrations.RunPython.noop)]
