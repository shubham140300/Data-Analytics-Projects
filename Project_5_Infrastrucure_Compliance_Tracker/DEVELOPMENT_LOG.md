# Development log

- Phase 1 - Workbook inspection: complete. Read-only profile saved in `docs/data-profile.md`.
- Phase 2 - Architecture and data flow: complete. Django and SQLite store source rows, current state, events, and audit history.
- Phase 3 - Tracker pages and models: implemented.
- Phase 4 - Excel import, search, filtering, and export: implemented.
- Phase 5 - Dashboard monthly comparison charts: implemented for the seven requested parameters; missing snapshots remain visibly missing.
- Phase 6 - Web row editing and Excel write-back: implemented. The original upload is preserved, a separate working copy is maintained, and the configured original workbook path is checked before each write.
- Phase 7 - Runtime setup: complete. Django system check passed; import details and row editor returned HTTP 200 on port 8001.
- Phase 8 - Owner permissions and workbook schema controls: implemented. Owner ID 2798869 manages grants; delegated imports must match the owner-established non-status schema.
- Phase 9 - Local security hardening: implemented. Loopback-only requests, DEBUG off by default, generated ignored secret key, CSRF/CSP/security headers, expiring sessions, protected workbook downloads, and hardened/bounded .xlsx parsing.
- Security verification: `manage.py check`, migration drift check, and template compilation passed. `check --deploy` reports the expected HTTPS-only HSTS/secure-cookie warnings because this local app uses HTTP on loopback. Local login redirect and security headers were confirmed over HTTP.
- Virus scanning: no antivirus engine is built into the app; workbook active content is rejected and files are handled as data. Keep Windows Defender or another antivirus enabled.
- Initial snapshot: imported from the attached workbook as `Splunk` for 2 October 2026. The original `Compliant to Non-Compliant` label remains Unmapped.
- Automated tests: not added or run.
