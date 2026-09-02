# BAGC Construction Management V17

Birhanu Abebe General Contractor multi-project construction management system.

V17 fixes and adds:
- Manpower overtime by normal day, night, Sunday and holiday, each with hours and rate.
- Machinery idle classification: payable vs non-payable; payable idle is included in machinery cost.
- Fuel register error fix; fuel reports calculate received, actual consumption, expected consumption, variance and cost.
- Persistent selectable/manual crew groups and positions; new values are saved for future selection.
- User IDs (`BAGC-00001` format) shown in administration.
- More robust BOQ Excel synchronization: detects headers below title rows, processes multiple BOQ worksheets, updates existing item numbers, and records import history.

Render:
Build: `pip install -r requirements.txt`
Start: `gunicorn app:app`

Environment variables:
`ADMIN_USERNAME`, `ADMIN_PASSWORD`, `SECRET_KEY`

## V17 additions
- Fixed BOQ upload Internal Server Error caused by closing the request DB connection before exception handling.
- BOQ importer safely reports import errors instead of returning a server error.
- RFI hierarchy: Site Engineer → Office Engineer → Project Manager, sequential sign-off.
- Daily entries automatically create/update a saved daily report snapshot.
- Reports support user-selected From Date and End Date for Daily, Weekly, Monthly, Semi-Annual and Annual periods.
- Saved report archive with lineage links from higher reports to lower reports.
- Machinery, manpower, store and fuel reports can be generated for any date range.
- Machinery assignment lifecycle: signed start date/hour, signed end date/hour, ended status and mandatory new signature before reuse.
- Daily machinery fuel/gauge data is linked to the Fuel Report.
- Professional staff IDs: BAGC-DEPARTMENT-YEAR-SEQUENCE, with staff position stored for workflow assignment.
