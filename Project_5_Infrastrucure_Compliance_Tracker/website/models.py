from django.db import models


class ComplianceStatus(models.TextChoices):
    COMPLIANT = "compliant", "Compliant"
    NON_COMPLIANT = "non_compliant", "Non-Compliant"
    NOT_APPLICABLE = "not_applicable", "Not Applicable"
    UNMAPPED = "unmapped", "Unmapped"


class Parameter(models.Model):
    """A configurable compliance check, such as the first Splunk dataset."""

    name = models.CharField(max_length=120, unique=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["name"]

    def __str__(self):
        return self.name


class Server(models.Model):
    """A server is keyed by the normalized source Host Name."""

    hostname = models.CharField(max_length=255)
    normalized_hostname = models.CharField(max_length=255, unique=True)
    current_ip = models.CharField(max_length=100, blank=True)
    ip_ambiguous = models.BooleanField(default=False)
    operating_system = models.CharField(max_length=150, blank=True)
    environment = models.CharField(max_length=100, blank=True)
    first_seen_on = models.DateField(null=True, blank=True)
    last_seen_on = models.DateField(null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["hostname"]
        indexes = [models.Index(fields=["hostname"])]

    def __str__(self):
        return self.hostname


class ImportBatch(models.Model):
    class Result(models.TextChoices):
        COMPLETED = "completed", "Completed"
        PARTIAL = "partial", "Completed with unresolved rows"
        FAILED = "failed", "Failed"

    parameter = models.ForeignKey(Parameter, on_delete=models.PROTECT, related_name="imports")
    source_file = models.FileField(upload_to="imports/%Y/%m/", blank=True)
    working_file = models.FileField(upload_to="imports/edited/", blank=True)
    original_filename = models.CharField(max_length=255)
    file_sha256 = models.CharField(max_length=64, db_index=True)
    working_sha256 = models.CharField(max_length=64, blank=True)
    source_sheet = models.CharField(max_length=120, blank=True)
    status_column = models.CharField(max_length=255, blank=True)
    source_columns = models.JSONField(default=list)
    sync_target_path = models.CharField(
        max_length=1024,
        blank=True,
        help_text="Optional local workbook updated when a source row is edited in the web app.",
    )
    is_current_source = models.BooleanField(
        default=True,
        help_text="Whether this import belongs to the currently selected workbook data source.",
    )
    is_master_source = models.BooleanField(default=False)
    snapshot_date = models.DateField()
    result = models.CharField(max_length=16, choices=Result.choices, default=Result.COMPLETED)
    total_rows = models.PositiveIntegerField(default=0)
    unique_servers = models.PositiveIntegerField(default=0)
    missing_hostname_rows = models.PositiveIntegerField(default=0)
    missing_ip_rows = models.PositiveIntegerField(default=0)
    duplicate_rows = models.PositiveIntegerField(default=0)
    conflicting_server_groups = models.PositiveIntegerField(default=0)
    non_advancing_server_groups = models.PositiveIntegerField(default=0)
    unmapped_status_rows = models.PositiveIntegerField(default=0)
    warning = models.TextField(blank=True)
    error_message = models.TextField(blank=True)
    imported_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-imported_at"]
        indexes = [models.Index(fields=["parameter", "snapshot_date"])]

    def __str__(self):
        return f"{self.parameter} · {self.snapshot_date}"


class MasterWorkbook(models.Model):
    """The owner-selected workbook that supplies the current dashboard data."""

    path = models.CharField(max_length=1024, blank=True)
    original_filename = models.CharField(max_length=255, blank=True)
    content_sha256 = models.CharField(max_length=64, blank=True)
    parameter_names = models.JSONField(default=list)
    roster_details = models.JSONField(default=list)
    roster_count = models.PositiveIntegerField(default=0)
    previous_roster_count = models.PositiveIntegerField(null=True, blank=True)
    new_server_count = models.PositiveIntegerField(null=True, blank=True)
    new_server_details = models.JSONField(default=list)
    baseline_batch = models.ForeignKey(
        ImportBatch,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="master_baselines",
    )
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "master workbook"

    def __str__(self):
        return self.original_filename or "Master workbook"


class ImportRecord(models.Model):
    """One source row, retained even when it cannot update a current state."""

    batch = models.ForeignKey(ImportBatch, on_delete=models.CASCADE, related_name="records")
    source_row = models.PositiveIntegerField()
    server = models.ForeignKey(
        Server, on_delete=models.SET_NULL, null=True, blank=True, related_name="import_records"
    )
    hostname = models.CharField(max_length=255, blank=True, db_index=True)
    ip_address = models.CharField(max_length=100, blank=True, db_index=True)
    operating_system = models.CharField(max_length=150, blank=True)
    environment = models.CharField(max_length=100, blank=True)
    raw_status = models.CharField(max_length=255, blank=True)
    mapped_status = models.CharField(
        max_length=20, choices=ComplianceStatus.choices, default=ComplianceStatus.UNMAPPED
    )
    source_data = models.JSONField(default=dict)
    duplicate_row = models.BooleanField(default=False)
    issue = models.CharField(max_length=255, blank=True)

    class Meta:
        ordering = ["source_row"]
        indexes = [models.Index(fields=["batch", "server"])]


class ComplianceState(models.Model):
    """The latest accepted status for one server and one parameter."""

    server = models.ForeignKey(Server, on_delete=models.CASCADE, related_name="states")
    parameter = models.ForeignKey(Parameter, on_delete=models.PROTECT, related_name="states")
    status = models.CharField(max_length=20, choices=ComplianceStatus.choices)
    raw_status = models.CharField(max_length=255, blank=True)
    latest_import = models.ForeignKey(
        ImportBatch, on_delete=models.SET_NULL, null=True, blank=True, related_name="current_states"
    )
    observed_on = models.DateField()
    recurrence_count = models.PositiveIntegerField(default=0)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["parameter__name", "server__hostname"]
        constraints = [
            models.UniqueConstraint(fields=["server", "parameter"], name="unique_server_parameter_state")
        ]
        indexes = [models.Index(fields=["status", "parameter"])]

    def __str__(self):
        return f"{self.server} · {self.parameter} · {self.get_status_display()}"


class StatusEvent(models.Model):
    class Kind(models.TextChoices):
        INITIAL = "initial", "First recorded status"
        CHANGE = "change", "Status change"
        RECURRENCE = "recurrence", "Recurring non-compliance"
        MANUAL = "manual", "Manual update"

    server = models.ForeignKey(Server, on_delete=models.CASCADE, related_name="status_events")
    parameter = models.ForeignKey(Parameter, on_delete=models.PROTECT, related_name="status_events")
    import_batch = models.ForeignKey(
        ImportBatch, on_delete=models.SET_NULL, null=True, blank=True, related_name="status_events"
    )
    previous_status = models.CharField(max_length=20, choices=ComplianceStatus.choices, blank=True)
    new_status = models.CharField(max_length=20, choices=ComplianceStatus.choices)
    kind = models.CharField(max_length=16, choices=Kind.choices)
    event_date = models.DateField()
    reason = models.CharField(max_length=500, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-event_date", "-id"]
        indexes = [models.Index(fields=["server", "parameter", "event_date"])]

    def __str__(self):
        return f"{self.server} · {self.get_new_status_display()} · {self.event_date}"


class AuditLog(models.Model):
    action = models.CharField(max_length=60)
    detail = models.TextField(blank=True)
    import_batch = models.ForeignKey(
        ImportBatch, on_delete=models.SET_NULL, null=True, blank=True, related_name="audit_entries"
    )
    server = models.ForeignKey(
        Server, on_delete=models.SET_NULL, null=True, blank=True, related_name="audit_entries"
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-created_at"]


class WorkspaceUser(models.Model):
    """An owner-approved username and password for this local workspace."""

    user_id = models.CharField(max_length=32, unique=True)
    username = models.CharField(max_length=150, unique=True)
    password_hash = models.CharField(max_length=128, blank=True)
    must_change_password = models.BooleanField(default=False)
    can_read = models.BooleanField(default=False)
    can_write = models.BooleanField(default=False)
    can_execute = models.BooleanField(default=False)
    is_active = models.BooleanField(default=True)
    granted_by = models.CharField(max_length=32, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["user_id"]

    def __str__(self):
        return self.user_id


class WorkspaceParameterPermission(models.Model):
    """A user's capabilities for one compliance parameter."""

    user = models.ForeignKey(WorkspaceUser, on_delete=models.CASCADE, related_name="parameter_permissions")
    parameter = models.ForeignKey(Parameter, on_delete=models.CASCADE, related_name="workspace_permissions")
    can_read = models.BooleanField(default=False)
    can_write = models.BooleanField(default=False)
    can_execute = models.BooleanField(default=False)

    class Meta:
        ordering = ["user__user_id", "parameter__name"]
        constraints = [
            models.UniqueConstraint(fields=["user", "parameter"], name="unique_user_parameter_permission")
        ]

    def __str__(self):
        return f"{self.user.user_id} · {self.parameter.name}"
