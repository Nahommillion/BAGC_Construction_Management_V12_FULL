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
