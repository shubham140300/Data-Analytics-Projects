# Workbook profile

## Source structure

- Workbook: `Anonymized_Data.xlsx`
- Worksheet: `Sheet1`
- Snapshot date indicated in the status heading: 2 October 2026
- Nonblank data rows: 6,077
- Columns, exactly as written in the workbook:
  - `Host Name`
  - `IP Address`
  - `OPERATING SYSTEM`
  - `Environment`
  - `SplunkAgent 2 october 2026`
- There is no separate date column. The date is part of the status-column heading.

## Status values

| Exact source value | Rows | Initial application treatment |
| --- | ---: | --- |
| `Compliant` | 4,350 | Compliant |
| `Non-Compliant` | 1,086 | Non-Compliant |
| `Not Applicable` | 486 | Not Applicable |
| `Compliant to Non-Compliant` | 155 | Unmapped until its meaning is chosen |

The source wording is retained on every imported row. The final category is not counted as Compliant or Non-Compliant until the user chooses how to interpret it.

## Other column values

`Environment` has two values: `Development` (3,055 rows) and `Production` (3,022 rows).

`OPERATING SYSTEM` has nine values:

| Exact source value | Rows |
| --- | ---: |
| `Linux` | 4,520 |
| `Windows` | 908 |
| `CentOS` | 377 |
| `VMWare ESXi` | 157 |
| `Custom OS` | 55 |
| `Storage` | 42 |
| `Other Operating System(OS)` | 11 |
| `Unknown` | 5 |
| `Backup Device` | 2 |

## Record quality and duplicate patterns

- `Host Name` and `IP Address` each have 6,072 nonblank cells and 5 blank cells.
- There are 6,045 distinct nonblank host names and 6,057 distinct nonblank IP strings.
- All 6,072 nonblank host-name values are 11 characters long, contain a hyphen, and contain no dot or whitespace. They appear to be short host identifiers rather than fully qualified domain names.
- All 6,072 nonblank IP values parse as IPv4 addresses; there are no IPv6 values or other IP formats in this file.
- The 5 rows with blank host names are the same 5 rows with blank IP addresses. Each of those rows still has other source data.
- 12 host-name values repeat, with 27 extra appearances. Two host names are paired with more than one IP address in this snapshot.
- 15 IP values repeat. Two IP values are paired with more than one host name.
- There are 16 exact duplicate rows beyond their first occurrence, across 14 duplicate groups.
- Repeated rows for a host name do not have differing status values in this snapshot. Repeated IP/host relationships are retained rather than deduplicated.

Rows without a host name cannot be matched to a server; they should remain visible as import issues. The application keeps every source row, and uses normalized `Host Name` as the current server key. IP is not used as a unique key.

## Initial system design

- **Django** renders browser pages and handles forms.
- **SQLite** stores parameter definitions, servers, current server-parameter state, import batches, row-level source records, status events, and audit entries.
- **Excel service layer** validates `.xlsx` input and converts each source row into a normalized observation while preserving the exact original column/value pairs.
- **Pages**: Overview, Servers, Non-compliant, Parameters, Imports, Import details, Server details, and History.
- **Import flow**: validate headings and workbook; detect or request the status column; derive or request snapshot date; preserve uploaded file and rows; update states only when a snapshot is newer; retain issues for review.
- **History flow**: record first-known status and actual mapped status changes; do not make a repeated status into a new transition.
- **Recurrence flow**: mark a return as recurring only for a Compliant → Non-Compliant transition when a prior Non-Compliant event exists for the same host and parameter.

## Open business rule

The label `Compliant to Non-Compliant` sounds transition-like, but this workbook does not say whether it is a current-state value, an event marker, or both. It remains Unmapped until the user chooses. This is the only status category with an unresolved meaning in the current sheet.
