# BAGC Construction Management V16

Birhanu Abebe General Contractor multi-project construction management system.

V16 fixes and adds:
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
