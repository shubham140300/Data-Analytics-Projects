# Application design

## Technology

- Python and Django for the web application and database access.
- SQLite for the first local version; it is included with Python and needs no separate database service.
- openpyxl for `.xlsx` reading and Excel exports.
- Django templates and local CSS for the browser UI. The pages do not load external JavaScript, fonts, or paid services.

## Database structure

| Model | Purpose | Important relationship |
| --- | --- | --- |
| `Parameter` | Stores a named compliance check. The name is entered during import and can be extended without adding code. | One parameter has many imports, current states, and events. |
| `Server` | Stores one server keyed by a case-insensitive, trimmed `Host Name`. | One server can have a current state for each parameter. |
| `ImportBatch` | Records the original upload, editable workbook copy, optional local write-back path, snapshot metadata, row counts, warnings, and import time. | Belongs to one parameter. |
| `ImportRecord` | Retains one row from the source, including the original column/value dictionary, exact status label, row number, and any issue. | Belongs to an import; links to a server when `Host Name` is available. |
| `ComplianceState` | Stores the latest accepted status for a server-parameter pair. | Unique on `(server, parameter)`. |
| `StatusEvent` | Append-only event for the first known status and later changes, including manual updates and recurrence markers. | Belongs to a server and parameter; import link is optional for manual events. |
| `AuditLog` | Records upload failures, completed imports, manual updates, and source-row edits with before/after values and workspace actor ID. | Can link to an import and/or server. |
| `WorkspaceUser` | Stores owner-approved numeric IDs and Read, Write, Execute, and active flags. | The configured owner ID bypasses grants and is the only identity allowed to manage grants or workbook columns. |

Every imported row remains in `ImportRecord`, even if it has no host name, repeats another row, or has an unmapped status. The original upload is retained under `uploads/` with a generated storage filename. A working workbook copy starts as the same file and receives web edits, preserving the uploaded original for audit.

## Application pages

- **Overview** — current compliance distribution and rate, a previous/current-month compliant-check chart, per-parameter compliant and non-compliant comparisons, compliant-to-non-compliant transitions, recent imports, and recent events.
- **Servers** — search by host name or current/imported IP; filter by state and environment; paginated list.
- **Non-compliant** — current non-compliant server-parameter pairs, searchable and exportable to Excel.
- **Parameters** — current counts and latest snapshot dates for each configured check.
- **Imports** — dated import list, upload form, and source-row details for each import.
- **Server detail** — current states, row-level observations, history events, audited manual status updates, and links to edit the latest Excel rows.
- **History** — status events, with parameter, server search, and recurrence filtering.
- **Owner console** — user grants, access revocation, and editable workbook columns.

## Import workflow

1. Accept only `.xlsx` uploads up to 25 MB; use defused XML parsing, package expansion limits, row/column bounds, and reject macros, formulas, external links, embedded files, and active controls.
2. Require the exact current server columns from the source: `Host Name`, `IP Address`, `OPERATING SYSTEM`, and `Environment`.
3. Detect one status column by its values, or ask the user to provide that column's exact heading when detection is ambiguous.
4. Read a snapshot date from the status-column heading when possible; otherwise require the user to enter one.
5. Keep original column headings and source cell values in each row's `source_data`. Standard statuses are mapped only by exact known label. Unknown labels are stored as Unmapped.
6. Group rows by normalized `Host Name` only for current-state comparison. Preserve duplicate rows and repeated IP relationships in the source record table.
7. Do not update a host's current status when that host has conflicting statuses in the same snapshot, or when the snapshot date is not newer than its accepted status date.
8. Save the import metadata and source rows, then update states and create events in a database transaction.

## Web edits and workbook write-back

- The row editor displays every heading found in the selected workbook. It does not invent columns that are absent from that workbook.
- The editor writes only changed cells in the current working workbook copy and preserves the original uploaded workbook separately.
- A configured local write-back path must point to the same workbook bytes and headings as the working copy. Before each write, the app checks the file hash to avoid overwriting manual changes made outside the app.
- The workbook replacement uses a temporary file and an atomic replace. If Excel has the original open and Windows locks it, the edit is rejected with a request to close the workbook and retry.
- `Last Updated` and `Changes Made By` columns are refreshed when their headings are present. The actor is taken from the signed-in workspace ID rather than typed by the editor.
- Status edits update the current `ComplianceState` and add a `StatusEvent`. Other column edits update the source row and audit log.
- Only rows in the latest snapshot can be edited, so historical imports remain read-only.
- Only the owner can add, rename, or delete optional columns. Core identity, status, and audit columns are locked. Non-owner imports must follow an owner-established schema, except for the dated status heading.

## Current state, events, and recurrence

- Current state answers what the latest accepted snapshot says for a server and parameter.
- Status history records the first known status and each later mapped status transition. Repeated identical statuses refresh the observation date without producing duplicate transitions.
- The snapshot date is the date the status was observed in a file. The application does not infer the exact time a change happened between snapshots.
- A recurrence requires a transition from Compliant to Non-Compliant and an earlier Non-Compliant event for the same server and parameter. A missing host from a snapshot is not interpreted as a fix.
- `Compliant to Non-Compliant` remains Unmapped until the user chooses its meaning. It is retained as a source label and excluded from compliant/non-compliant totals.

## Dashboard comparisons

- Each parameter chart selects its latest completed or partial import snapshot within the previous and current calendar months. Conflicting host groups are excluded from the chart because the snapshot did not establish one status for those servers.
- Each snapshot counts a server at most once for that parameter. The overall monthly compliant total sums parameter checks, not globally unique servers, and shows snapshot coverage so missing parameter data is visible.
- A compliant-to-non-compliant conversion is counted only for a server present in both selected snapshots with Compliant in the previous snapshot and Non-Compliant in the current snapshot. If either month is missing, the dashboard asks for both snapshots instead of reporting zero.
- Current compliance rate is compliant current states divided by compliant plus non-compliant current states. Not-applicable and unmapped records are excluded.

## Data-source boundary

Workbook parsing and validation are in `website/services/excel_import.py`. `website/services/import_workflow.py` coordinates the reader and `website/services/snapshot_persistence.py` writes normalized observations and transitions. A future SQL Server adapter or Sentinel connector should be implemented as a separate reader that emits the same normalized server observations and reuses persistence, history, permissions, and dashboard rules. SQL credentials and Sentinel API tokens should be stored outside source code and granted read-only scope wherever possible. The connector and source-specific mapping are not implemented yet.

## Local security and limits

The app is intended for local use and `start.bat` binds Django's development server to `127.0.0.1`; a request guard also rejects non-loopback clients. Debug mode defaults off, the local secret is randomly generated and git-ignored, CSRF and Content Security Policy are enabled, and private workbook downloads require a workspace session. Do not expose the development server to a public network.

Sign-in requires an active, owner-approved username and password. Passwords are stored as Django password hashes; the first local start prompts for the owner's password without echoing it. The owner creates or resets passwords for other approved users through the access console. The app also does not run an antivirus scanner; uploaded workbooks are parsed as data and are not executed, but they should still be trusted and scanned by the computer's antivirus. Public or shared-network deployment requires HTTPS, secure cookies, a production server, and a deployment review.

## Development phases

1. Workbook profile and unresolved label identification — complete.
2. SQLite schema and import/history design — complete.
3. Browser pages and current-state workflows — implemented in the project files.
4. Runtime setup and database initialization — automated by `start.bat`.
5. Runtime validation — prior dashboard checks were completed; re-run the Django system check after local changes.
