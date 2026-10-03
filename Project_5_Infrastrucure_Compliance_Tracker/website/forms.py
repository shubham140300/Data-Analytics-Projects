from django import forms
from django.conf import settings
from django.contrib.auth.hashers import make_password

from .models import ComplianceStatus, ImportRecord, Parameter, WorkspaceParameterPermission, WorkspaceUser


class ExcelImportForm(forms.Form):
    parameter_name = forms.CharField(
        max_length=120,
        label="Parameter name",
        help_text="For example, Splunk. This is stored separately from the workbook's column heading.",
    )
    snapshot_date = forms.DateField(
        required=False,
        widget=forms.DateInput(attrs={"type": "date"}),
        help_text="Optional if a date can be read from the status column heading.",
    )
    excel_file = forms.FileField(label="Excel workbook (.xlsx)")
    status_column = forms.CharField(
        max_length=255,
        required=False,
        label="Status column (optional)",
        help_text="Leave blank when the workbook has one clear status column. Otherwise type its exact heading.",
    )
    sync_target_path = forms.CharField(
        max_length=1024,
        required=False,
        label="Original workbook path for write-back (optional)",
        help_text="For a local app, enter the full path to the same .xlsx file you upload. Web edits will update that file too.",
    )

    def clean_parameter_name(self):
        return self.cleaned_data["parameter_name"].strip()

    def clean_status_column(self):
        value = self.cleaned_data.get("status_column", "")
        return value.strip()

    def clean_sync_target_path(self):
        return self.cleaned_data.get("sync_target_path", "").strip()

    def clean_excel_file(self):
        uploaded = self.cleaned_data["excel_file"]
        if not uploaded.name.lower().endswith(".xlsx"):
            raise forms.ValidationError("Choose an .xlsx Excel workbook.")
        if uploaded.size > 25 * 1024 * 1024:
            raise forms.ValidationError("The workbook is larger than the 25 MB limit.")
        return uploaded


class MasterWorkbookImportForm(forms.Form):
    master_file = forms.FileField(label="Master workbook (.xlsx)")
    target_path = forms.CharField(
        max_length=1024,
        required=True,
        label="Original workbook path for write-back",
        help_text="Enter the full path to the same workbook so edits in the web app update this file.",
    )

    def clean_target_path(self):
        return self.cleaned_data.get("target_path", "").strip()

    def clean_master_file(self):
        uploaded = self.cleaned_data["master_file"]
        if not uploaded.name.lower().endswith(".xlsx"):
            raise forms.ValidationError("Choose an .xlsx Excel master workbook.")
        if uploaded.size > 25 * 1024 * 1024:
            raise forms.ValidationError("The workbook is larger than the 25 MB upload limit.")
        return uploaded


class ManualStateChangeForm(forms.Form):
    parameter = forms.ModelChoiceField(queryset=Parameter.objects.none(), label="Parameter")
    status = forms.ChoiceField(choices=ComplianceStatus.choices, label="New status")
    reason = forms.CharField(
        max_length=500,
        widget=forms.Textarea(attrs={"rows": 3}),
        help_text="This note is stored with the history event.",
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["parameter"].queryset = Parameter.objects.all()


def _column_key(heading):
    return "".join(character for character in heading.casefold() if character.isalnum())


def _display_source_value(value):
    return "" if value is None else str(value)


class SourceRecordEditForm(forms.Form):
    changed_by = forms.CharField(
        max_length=120,
        label="Changes made by",
        disabled=True,
        help_text="Set automatically from the signed-in workspace user ID.",
    )

    def __init__(self, record: ImportRecord, *args, actor="", **kwargs):
        super().__init__(*args, **kwargs)
        self.record = record
        self.fields["changed_by"].initial = actor
        self.source_fields = []
        self.managed_columns = []
        self.column_headings = {}
        for index, heading in enumerate(record.batch.source_columns):
            key = _column_key(heading)
            current_value = record.source_data.get(heading)
            if key in {"lastupdated", "lastupdatedon", "lastmodified", "changesmadeby", "changedby", "updatedby", "checkdate"}:
                self.managed_columns.append((heading, _display_source_value(current_value)))
                continue
            field_name = f"column_{index}"
            widget = forms.Textarea(attrs={"rows": 3}) if any(word in key for word in ("comment", "note", "description")) else forms.TextInput()
            self.fields[field_name] = forms.CharField(
                label=heading,
                required=(key == "hostname"),
                max_length=32767,
                initial=_display_source_value(current_value),
                widget=widget,
                strip=(key == "hostname"),
            )
            self.column_headings[field_name] = heading
            self.source_fields.append(self[field_name])

        if record.batch.is_master_source:
            source_headings = set(record.batch.source_columns)
            extra_headings = [heading for heading in record.source_data if heading not in source_headings]
            for index, heading in enumerate(extra_headings):
                key = _column_key(heading)
                if key in {"lastupdated", "lastupdatedon", "lastmodified", "changesmadeby", "changedby", "updatedby", "checkdate"}:
                    self.managed_columns.append((heading, _display_source_value(record.source_data.get(heading))))
                    continue
                field_name = f"master_column_{index}"
                self.fields[field_name] = forms.CharField(
                    label=heading,
                    required=False,
                    max_length=32767,
                    initial=_display_source_value(record.source_data.get(heading)),
                    widget=forms.Textarea(attrs={"rows": 3}) if any(word in key for word in ("comment", "note", "description")) else forms.TextInput(),
                )
                self.column_headings[field_name] = heading
                self.source_fields.append(self[field_name])

    def changed_columns(self):
        changed = {}
        for field_name, heading in self.column_headings.items():
            value = self.cleaned_data.get(field_name, "")
            old_value = _display_source_value(self.record.source_data.get(heading))
            if value != old_value:
                changed[heading] = value
        return changed


class MasterSheetRowEditForm(forms.Form):
    def __init__(self, row, headers, formula_columns=(), *args, actor="", **kwargs):
        super().__init__(*args, **kwargs)
        self.row = row
        self.headers = list(headers)
        self.column_headings = {}
        self.managed_columns = []
        self.readonly_columns = []
        self.edit_field_names = []
        self.fields["changed_by"] = forms.CharField(
            label="Changes made by", disabled=True, initial=actor,
            help_text="Set automatically from the signed-in workspace user ID.",
        )
        values = row["values"]
        formulas = set(formula_columns)
        for index, heading in enumerate(self.headers):
            key = _column_key(heading)
            value = _display_source_value(values.get(heading))
            if key in {"lastupdated", "lastupdatedon", "lastmodified", "changesmadeby", "changedby", "updatedby", "checkdate"}:
                self.managed_columns.append((heading, value))
                continue
            if heading in formulas:
                self.readonly_columns.append((heading, value))
                continue
            field_name = f"column_{index}"
            self.fields[field_name] = forms.CharField(
                label=heading,
                required=(key == "hostname"),
                max_length=32767,
                initial=value,
                strip=(key == "hostname"),
                widget=forms.Textarea(attrs={"rows": 3}) if any(word in key for word in ("comment", "note", "description")) else forms.TextInput(),
            )
            self.column_headings[field_name] = heading
            self.edit_field_names.append(field_name)
        self.source_fields = [self[name] for name in self.edit_field_names]

    def changed_columns(self):
        changed = {}
        for field_name, heading in self.column_headings.items():
            value = self.cleaned_data.get(field_name, "")
            old_value = _display_source_value(self.row["values"].get(heading))
            if value != old_value:
                changed[heading] = value
        return changed


class SyncTargetForm(forms.Form):
    sync_target_path = forms.CharField(
        max_length=1024,
        required=False,
        label="Original workbook path",
        help_text="Paste the full path to the same .xlsx file. Leave blank to update only the copy stored by this app.",
    )

    def clean_sync_target_path(self):
        return self.cleaned_data.get("sync_target_path", "").strip()


class WorkspaceUserGrantForm(forms.Form):
    user_id = forms.CharField(
        max_length=32, label="User ID", widget=forms.TextInput(attrs={"inputmode": "numeric", "autocomplete": "off"}),
        help_text="The approved user's ID, used to identify their account.",
    )
    username = forms.CharField(
        max_length=150, required=False, label="Username", widget=forms.TextInput(attrs={"autocomplete": "off"}),
        help_text="Leave blank to use the user's ID as their username.",
    )
    password = forms.CharField(
        max_length=128, required=False, label="Set or reset password",
        widget=forms.PasswordInput(attrs={"autocomplete": "new-password"}, render_value=False),
        help_text="New accounts use the configured temporary password if blank. Editing with a blank field keeps the existing password.",
    )
    is_active = forms.BooleanField(required=False, initial=True, label="Access is active")

    def __init__(self, *args, instance=None, parameters=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.instance = instance
        self.fields["password"].required = instance is None
        self.parameters = list(parameters if parameters is not None else Parameter.objects.all())
        self.permission_groups = []
        current = {}
        if instance:
            current = {
                permission.parameter_id: permission
                for permission in WorkspaceParameterPermission.objects.filter(user=instance)
            }
            self.fields["user_id"].initial = instance.user_id
            self.fields["username"].initial = instance.username
            self.fields["is_active"].initial = instance.is_active
        for parameter in self.parameters:
            permission = current.get(parameter.pk)
            group = {"parameter": parameter}
            for capability, title in (("read", "Read"), ("write", "Write"), ("execute", "Execute")):
                field_name = f"perm_{parameter.pk}_{capability}"
                self.fields[field_name] = forms.BooleanField(
                    required=False,
                    label=f"{parameter.name} · {title}",
                    initial=bool(getattr(permission, f"can_{capability}", False)) if permission else False,
                )
                group[capability] = self[field_name]
            self.permission_groups.append(group)

    def clean_user_id(self):
        user_id = self.cleaned_data["user_id"].strip()
        if not user_id.isdigit():
            raise forms.ValidationError("User IDs must contain digits only.")
        if user_id == settings.WORKSPACE_OWNER_ID:
            raise forms.ValidationError("The owner ID is reserved and cannot be granted to another user.")
        duplicate = WorkspaceUser.objects.filter(user_id=user_id)
        if self.instance:
            duplicate = duplicate.exclude(pk=self.instance.pk)
        if duplicate.exists():
            raise forms.ValidationError("This user ID already has an access grant. Edit that grant instead.")
        return user_id

    def clean_username(self):
        username = self.cleaned_data.get("username", "").strip().casefold()
        if not username:
            username = self.cleaned_data.get("user_id", "").strip().casefold()
        if not username or not username[0].isalnum() or any(
            not (character.isalnum() or character in "_.@+-") for character in username
        ):
            raise forms.ValidationError("Use letters, numbers, dots, underscores, @, +, or hyphens in the username.")
        if username == settings.WORKSPACE_OWNER_USERNAME.casefold():
            raise forms.ValidationError("The owner username is reserved.")
        duplicate = WorkspaceUser.objects.filter(username__iexact=username)
        if self.instance:
            duplicate = duplicate.exclude(pk=self.instance.pk)
        if duplicate.exists():
            raise forms.ValidationError("This username is already in use.")
        return username

    def clean_password(self):
        password = self.cleaned_data.get("password", "")
        if password and len(password) < 10:
            raise forms.ValidationError("Passwords must be at least 10 characters long.")
        return password

    def clean(self):
        cleaned = super().clean()
        for parameter in self.parameters:
            read = cleaned.get(f"perm_{parameter.pk}_read", False)
            if (cleaned.get(f"perm_{parameter.pk}_write") or cleaned.get(f"perm_{parameter.pk}_execute")) and not read:
                raise forms.ValidationError(f"{parameter.name}: grant Read with Write or Execute.")
        return cleaned

    def save(self, *, granted_by=""):
        user_id = self.cleaned_data["user_id"].strip()
        grant = self.instance or WorkspaceUser(user_id=user_id)
        grant.user_id = user_id
        grant.username = self.cleaned_data["username"]
        password = self.cleaned_data.get("password")
        if password or not grant.password_hash:
            password = password or settings.WORKSPACE_DEFAULT_PASSWORD
            grant.password_hash = make_password(password)
            grant.must_change_password = True
        grant.is_active = self.cleaned_data.get("is_active", False)
        grant.granted_by = granted_by
        selected = {}
        for parameter in self.parameters:
            values = {
                capability: bool(self.cleaned_data.get(f"perm_{parameter.pk}_{capability}"))
                for capability in ("read", "write", "execute")
            }
            selected[parameter.pk] = values
        grant.can_read = any(value["read"] for value in selected.values())
        grant.can_write = any(value["write"] for value in selected.values())
        grant.can_execute = any(value["execute"] for value in selected.values())
        grant.save()
        for parameter in self.parameters:
            values = selected[parameter.pk]
            WorkspaceParameterPermission.objects.update_or_create(
                user=grant,
                parameter=parameter,
                defaults={
                    "can_read": values["read"],
                    "can_write": values["write"],
                    "can_execute": values["execute"],
                },
            )
        WorkspaceParameterPermission.objects.filter(user=grant).exclude(
            parameter__in=self.parameters
        ).delete()
        return grant


class WorkspacePasswordChangeForm(forms.Form):
    current_password = forms.CharField(
        label="Current password", widget=forms.PasswordInput(attrs={"autocomplete": "current-password"}),
    )
    new_password = forms.CharField(
        label="New password", min_length=10,
        widget=forms.PasswordInput(attrs={"autocomplete": "new-password"}),
        help_text="Use at least 10 characters.",
    )
    confirm_password = forms.CharField(
        label="Confirm new password", widget=forms.PasswordInput(attrs={"autocomplete": "new-password"}),
    )

    def clean(self):
        cleaned = super().clean()
        if cleaned.get("new_password") and cleaned.get("confirm_password") and cleaned["new_password"] != cleaned["confirm_password"]:
            self.add_error("confirm_password", "The new passwords do not match.")
        return cleaned


class ColumnChangeForm(forms.Form):
    operation = forms.ChoiceField(choices=(
        ("add", "Add a blank column"),
        ("rename", "Rename an optional column"),
        ("delete", "Remove an optional column"),
    ))
    existing_heading = forms.ChoiceField(choices=(("", "Choose a column"),), required=False, label="Existing optional column")
    new_heading = forms.CharField(max_length=255, required=False, label="New column name")
    confirm_delete = forms.BooleanField(required=False, label="I understand this removes the column and its values from the workbook")

    def clean(self):
        cleaned = super().clean()
        operation = cleaned.get("operation")
        existing = (cleaned.get("existing_heading") or "").strip()
        new = (cleaned.get("new_heading") or "").strip()
        if operation in {"rename", "delete"} and not existing:
            self.add_error("existing_heading", "Choose an existing optional column.")
        if operation in {"add", "rename"} and not new:
            self.add_error("new_heading", "Enter a column name.")
        if operation == "delete" and not cleaned.get("confirm_delete"):
            self.add_error("confirm_delete", "Confirm column removal before continuing.")
        cleaned["existing_heading"] = existing
        cleaned["new_heading"] = new
        return cleaned
