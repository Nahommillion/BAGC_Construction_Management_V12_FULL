# BAGC Construction Management V28

V28 adds strict separation between Head Office personnel and Project Team personnel, cross-project Head Office responsibility, filtered department reports, and multi-person responsibility/contact selection.

## V28 highlights
- Head Office accounts and Project Personnel accounts are separate using `personnel_scope`.
- Head Office personnel are never added to the Project Team register.
- A Head Office employee can be assigned functional responsibility for any number of projects without becoming project staff.
- Project Team responsibility and Head Office responsibility have separate multi-select boxes with no 5/10-person limit.
- Head Office responsibility gives project access for the assigned project/function while preserving Head Office identity.
- Workflow contacts include relevant project responsibility contacts.
- Machinery/Fuel personnel can only access Machinery and Fuel reports.
- Store personnel can only access Store reports.
- HR personnel can only access Manpower reports.
- Finance personnel can only access Finance reports.
- Design personnel can only access Design reports.
- Project Manager has full reporting access within assigned projects.
- Saved reports are also filtered and protected by the same report permissions.
- Project Manager remains the project administrator for assigned projects.
- Existing V27 request approval, Head Office routing, forwarding, resource registration, Excel imports and other features are retained.

## Deployment
- Start command: `gunicorn app:app`
- Requirements: Flask, openpyxl, Werkzeug, gunicorn.
- For persistent Render data, configure `BAGC_DATA_DIR` on a persistent disk or use PostgreSQL.
