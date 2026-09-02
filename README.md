# BAGC Construction Management V18

Multi-project construction management system for Birhanu Abebe General Contractor.

## V18 changes
- Robust BOQ Excel import: scans sheets for the real header row and accepts common BOQ labels such as No./Item No., Description of Works, Unit, Quantity, Contract Quantity, Rate, Unit Rate and Unit Price.
- BOQ import supports `.xlsx` and `.xlsm` and handles numeric values stored as text with commas/currency labels.
- Imported BOQ items become the project master and are used by daily work and progress reporting.
- Mandatory professional staff photo during user registration.
- Permanent BAGC Staff ID, department, year and sequence.
- Printable staff ID card with photo, name, department, position and status.
- Existing staff records are never deleted. Disable only blocks access; the record and Staff ID remain.
- Existing users can have their photo updated.
- RFI workflow: Site Engineer → Office Engineer → Project Manager.
- Daily data feeds weekly/monthly/semi-annual/annual reporting with custom From/End dates and saved report history.
- Machinery assignment lifecycle requires signed start/end hour records and a new assignment after an assignment is ended.

## Render
Build: `pip install -r requirements.txt`
Start: `gunicorn app:app`

Environment variables:
- `ADMIN_USERNAME`
- `ADMIN_PASSWORD`
- `SECRET_KEY`

## Deployment note
The application currently uses SQLite and runtime uploads. For production permanence on Render, use persistent storage or preferably PostgreSQL/object storage so database records and staff photos survive service replacement/redeployments.

## V20 persistence requirement
The application never deletes users, projects, BOQ, daily reports or saved reports during startup. SQLite is stored in `BAGC_DATA_DIR` when configured; if `/data` is writable it is preferred automatically. On Render, the normal ephemeral filesystem can still be reset during redeploy/restart. For permanent production data, configure a Render persistent disk mounted at `/data` (paid service) or migrate the SQLite database to PostgreSQL. Do not rely on the default ephemeral filesystem for company records.

V20 also fixes the machinery/fuel save errors, adds linked daily work packages, crew evaluation/capacity, BOQ variation alerts, staff ID bilingual front/back and unique Code 128 barcodes, and automatic machine assignment completion based on signed total hours.


## V22 Daily Report and Machinery workflow
- Daily Report is a single-save workflow: add multiple BOQ work packages, then select only the machinery, manpower, crews, store, fuel and finance records used for each BOQ item.
- Each selected resource has an allocation field so one resource record can be split across several BOQ work items on the same day. Store allocation cannot exceed the issued quantity; machinery cannot exceed logged hours; manpower cannot exceed attendance/hours; fuel and finance cannot exceed their source records.
- Daily Report includes a Unit selector populated from common construction units plus units found in the project's BOQ.
- Machinery has a signed assignment start date, start meter/hour and total signed hours. The assignment automatically ends when cumulative work + idle + down hours reach the signed total. A new signed assignment is required for reuse.
- Machinery assignment endpoints are included and the machine add/log workflow is committed before the page is rendered.
- SQLite WAL mode is configured during initialization rather than on every request, reducing lock contention. Login no longer reruns database migrations on every authentication request.
