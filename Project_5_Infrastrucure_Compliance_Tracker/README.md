# TechCompliance

A local Django application for importing Excel compliance snapshots, viewing current server status, and recording status history. The first version uses free/open-source software and SQLite. It does not use paid APIs, hosted databases, external fonts, or other paid services.

## Start the application

1. Make sure Python is installed and available as `python` in Windows.
2. Double-click `start.bat` in this folder.
3. The first start creates a private `.venv`, downloads Django, openpyxl, and defusedxml, prepares the SQLite database, and asks you to set the owner password in the terminal before opening the site at <http://127.0.0.1:8001/>.
4. Leave the black terminal window open while using the site. Press `Ctrl+C` there to stop it.

The first dependency install needs an internet connection. Django and openpyxl are free Python packages. Future starts use the local installation.

### Run manually

Open PowerShell in this folder and run:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python manage.py migrate
python manage.py setup_workspace_owner
python manage.py runserver --insecure 127.0.0.1:8001
```

Then visit <http://127.0.0.1:8001/> in a browser on this computer.

Sign in with the owner username (defaults to `2798869`) and the password set during first start. The owner can approve user IDs and assign Read, Write, and Execute access from **Access & columns**. A blank username uses the user ID; the temporary password is `Password@123`, and users must change it at first sign-in. Owner sign-in persists and renews while used; other user sessions expire after eight hours. Passwords do not expire automatically, and users can change their own password from the account menu.

## What the application does

- Shows the latest known status for each server and parameter.
- Highlights current non-compliant and unmapped statuses.
- Imports `.xlsx` snapshots while preserving the uploaded source workbook and each nonblank source row.
- Reads the multi-tab SOE master workbook, combines the `Main Data` roster with dated Splunk status columns, and keeps older unrelated imports archived rather than mixing them into the active dashboard.
- Shows the number of servers newly present in a later master workbook compared with the previous master roster. The first master upload establishes the baseline.
- Opens the new-server count to a searchable hostname and IP address list.
- Shows overall compliant and non-compliant monthly trends as a line chart.
- Records transitions when a later dated snapshot changes a known status.
- Marks a return to non-compliant as recurring after the same server and parameter were previously non-compliant and then compliant.
- Searches and filters server lists with 25, 50, or 100 rows per page.
- Allows a manual status update with a reason; it creates a history event and an audit entry.
- Allows editing any imported source row by its Excel headings, updates the saved workbook, and optionally writes back to the matching original `.xlsx` file.
- Lets the owner add, rename, or remove optional columns in the latest workbook snapshot. Core identity, status, and audit columns stay locked.
- Shows per-parameter compliant-to-non-compliant and recovery counts, with expandable lists of servers whose statuses changed.
- Exports current non-compliant server and parameter pairs to `.xlsx`.

## Import a workbook

For the SOE master workbook, open **Imports** and choose **Update master workbook**. Upload the complete workbook after adding the new month column, and enter the full path to that same local file. The app reads the server roster from `Main Data` and dated status values from `Splunk`. It recognizes the `CS` tab as **CS (Craft)**, along with `TGIM`, `Splunk`, `Logger`, `RSA`, and `SNOW`. Tabs without status values remain available but show no compliance snapshot.

The initial master workbook currently contains August, September, and October 2026 Splunk history. Those month headings become snapshots. A server count is not shown as “new” until the next master upload because the previous app dataset used anonymized hostnames and cannot be matched to this workbook. Each later upload compares the new roster to the most recently imported master roster.

For a separate one-parameter workbook, use **Import snapshot**:

1. Open **Imports** and choose **Import workbook**.
2. Enter a parameter name, such as `Splunk`. This is a separate label; the original status-column heading is retained exactly.
3. Choose the `.xlsx` workbook. The file is limited to 25 MB.
4. Enter the snapshot date if the app cannot read one from the status-column heading.
5. If the sheet has more than one possible status column, enter the exact heading for the correct status column.
6. Only the owner can configure write-back to the original workbook. The path must point to the same workbook you upload.
7. Select **Validate and import**. The import details page shows the source labels and any rows needing review.

The active workbook profile is in [docs/master-workbook-profile.md](docs/master-workbook-profile.md); the earlier anonymized workbook profile remains in [docs/data-profile.md](docs/data-profile.md). The `Compliant to Non-Compliant` label remains **Unmapped** in this master import too, as previously selected. Those records stay visible and are excluded from compliant and non-compliant totals.

Only the owner can establish a parameter's first workbook schema or change its columns. A user with Execute can import later snapshots only when the non-status headings, count, and order match the approved schema; the dated status heading can change from month to month.

The importer requires these source headings exactly:

- `Host Name`
- `IP Address`
- `OPERATING SYSTEM`
- `Environment`

The status heading is detected from known status values or can be entered on the form. Other workbook columns and cell values are retained in the imported row data.

## Edit a workbook row

1. Open a server detail page and choose **Edit this row** under **Edit source data**. Rows from older snapshots stay read-only.
2. Edit source columns shown from the workbook's exact headings. The Host Name remains required because it identifies the server.
3. Save. **Changes made by** is populated from the signed-in workspace user ID; users cannot type another identity into the audit field. If the workbook has a matching `Changes Made By` column, the app fills it automatically. A matching `Last Updated` column is also refreshed automatically.
4. In master-workbook mode, the app updates the original workbook and all saved master copies. The first edit creates a one-time backup beside the original file. Status edits also update current status and add a history event. Other edited values are recorded in the audit log.

The owner configures the write-back path on the master import page. Before a write, the app checks the master file's saved fingerprint and confirms the row still matches the server. If the workbook changed outside the app, the edit is stopped to prevent overwriting newer spreadsheet changes. Upload the latest workbook again before editing.

## Owner-only column settings

The owner can open the latest import and choose **Manage columns** to add, rename, or remove optional headings. For a master workbook, the change is written to the original file and applied across all saved snapshots for that parameter tab. Core server identity, dated status, and automatically managed audit columns stay locked. New monthly status columns are detected when the owner uploads the updated master workbook.

## How current status and history work

- **Current state** is one latest status for each server and parameter.
- **Data** stores the first recorded status and each later status change as a dated event, with an owner-only bulk clear option.
- An unchanged status in a later snapshot refreshes the current observation date without adding a duplicate transition.
- A transition to non-compliant is marked recurring only if that server and parameter had an earlier non-compliant event and were compliant immediately before the new non-compliant status.
- The snapshot date is used as the observed date. It does not claim the exact time a change happened between two imports.
- Servers missing from a new file are not automatically marked compliant or removed.
- Rows without `Host Name` are retained but cannot be linked to a server. Exact duplicate rows remain in import history and are flagged.
- Conflicting statuses for one host in a snapshot are kept in the raw import and do not update that host's current state.
- An older snapshot is stored for traceability but does not roll current state backward. When the owner re-uploads the master workbook with a corrected value for the same month, that same-date master snapshot updates current state.

## Dashboard definitions

- **Known servers**: distinct named hosts in the latest master inventory snapshot when master-workbook mode is active.
- **Non-compliant**: current server and parameter pairs with a mapped Non-Compliant status.
- **Compliant**: current server and parameter pairs with a mapped Compliant status.
- **Recurring**: current server and parameter pairs with at least one recorded return to Non-Compliant after a compliant state.
- **Unmapped**: current records whose exact source labels have not been mapped to a business status.
- **Compliance rate**: compliant current checks divided by compliant plus non-compliant current checks. Not-applicable and unmapped checks are excluded.
- **Monthly comparison**: the latest completed snapshot for each parameter within the previous and current calendar months. The combined chart sums server-parameter checks, so a server checked by two parameters counts once in each parameter.
- **Compliant to non-compliant**: for a parameter, servers present as Compliant in its selected previous-month snapshot and Non-Compliant in its selected current-month snapshot. Both months must have snapshots; missing months display as “No snapshot.”
- The dashboard's month-to-month conversion details show each server changed from Compliant to Non-Compliant, with Hostname, IP Address, and both statuses.
- The Servers search matches hostnames, IP addresses, and active parameter names. For example, searching `Nmap` lists servers with an active Nmap snapshot; importing Nmap data first is required if the current workbook does not contain it.

The active master dashboard includes cards for Splunk, CS (Craft), RSA, SNOW, TGIM, and Logger. A parameter without a snapshot shows “No snapshot” rather than zero. Monthly charts show the latest three months and compare status changes between adjacent snapshots. Previous anonymized imports remain stored for traceability but are excluded from the active master dashboard.

## Project structure

```text
config/                  Django settings and entry points
website/models.py        Servers, parameters, imports, current state, history, audit
website/services/        Excel parsing, snapshot workflow, and database persistence
website/views.py          Dashboard, search, history, import and export pages
templates/               HTML pages
static/                  Local CSS; no external asset service
uploads/                 Uploaded source workbooks (created at runtime)
db.sqlite3               Local database (created at runtime)
```

## Data storage and backup

The SQLite database is `db.sqlite3`. Uploaded and edited workbook copies are stored under `uploads/`. Stop the server before making a backup, then copy both `db.sqlite3` and the `uploads` folder to a safe location. When original-file write-back is configured, web edits also change that local `.xlsx` file.

## Remove import or clear application data

The owner can filter **Imports** by parameter, snapshot date, or upload date, then remove a selected import. The **Data** page also has a bulk clear action. It requires typing a confirmation phrase and removes all app snapshots, servers, status events, audit entries, master-workbook selection, and stored upload copies. User accounts and permissions remain. The original Excel workbook is not modified; reimport it to repopulate the app. Individual deletion and bulk-clear actions are recorded in the audit log.

## Adding future parameters and data sources

An import creates a parameter by the name entered in the form; existing parameters can be reused by entering the same name. New parameter names do not require code changes. Excel parsing is isolated in `website/services/excel_import.py` from the database and views, so another data source can later provide the same normalized server observations without replacing the history workflow.

The initial importer uses the workbook's exact `Host Name` value as the server key, compared without case differences. IP addresses are retained as server details rather than used as a unique key because the source contains repeated host and IP values.

## Security and local access

The app is intended to run on this computer and `start.bat` binds Django's development server to `127.0.0.1`. A request guard also rejects non-loopback clients. It is not deployed to the public internet. Debug mode is off by default; the app generates a local secret key, uses CSRF protection, short-lived server-side sessions, restrictive browser headers, and session-protected workbook downloads. The key file is git-ignored.

This application does not scan uploads with antivirus software, and no web app can guarantee that every attack or virus will be blocked. Uploads are size-limited and macro-enabled files, external links, embedded files, and active controls are rejected. The master importer reads formula-linked server details from the `Main Data` values and does not evaluate workbook formulas; dated status cells must contain saved values. Generic single-parameter imports continue to reject formulas. Only use trusted `.xlsx` files and keep Windows Defender or another antivirus product enabled on the computer.

Sign-in requires an active, owner-approved username and password. Passwords are stored as Django password hashes. The owner account uses ID `2798869` and a configurable username (default `2798869`); the first start prompts for its password without echoing it. Existing user grants need an initial password set by the owner from **Access & columns**. Keep the app local, do not bind it to `0.0.0.0`, open a firewall port, or expose it to an untrusted network. Cookie Secure/HSTS settings are intentionally off for local HTTP; they must be enabled behind HTTPS in production.

For deployment, use a production WSGI/ASGI server, a protected `DJANGO_SECRET_KEY`, HTTPS, restrictive allowed hosts, secure cookies, and a deployment security review. Django's official [deployment checklist](https://docs.djangoproject.com/en/6.0/howto/deployment/) describes the production steps. Free hosting terms and limits can change. No hosting provider is required for local use.

## Troubleshooting

- **`python` is not recognized**: install Python, then reopen the terminal. Django 6.0 supports Python 3.13 according to the [official Django installation FAQ](https://docs.djangoproject.com/en/6.0/faq/install/).
- **First install fails**: check the internet connection and try `start.bat` again.
- **Port 8001 is busy**: stop the other local server or change the run command to a free local port such as `python manage.py runserver 127.0.0.1:8010`, then open that port in the browser.
- **Workbook headings error**: check the exact required headings above; do not rename source values to force an import.
- **A status is Unmapped**: the exact label is retained. Decide its meaning before treating it as compliant or non-compliant.
- **Missing database tables**: close the app and run `start.bat` again so Django can prepare the database.
