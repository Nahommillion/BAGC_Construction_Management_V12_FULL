# BAGC Construction Management V26

Multi-project construction management system with strict department/project visibility and hierarchy-driven workflows.

## V26 additions
- Department/team/project visibility filtering.
- Head Office and Project responsibility register.
- Super Admin assignment of department/team, position, project position and reporting lines.
- Controlled recipient filtering for company workflow correspondence.
- Resource Requests for Material, Fuel, Machinery, Manpower, Expense, Design, Procurement and Other.
- Hierarchy-based approval: Project requests route to the project reporting chain; Head Office requests route to Head Office reporting chain.
- Approved requests automatically register into Store, Fuel, Machinery, Manpower, Finance or Design where applicable.
- Request attachments and permanent registration references.
- Department-filtered dashboard KPIs and request status chart.
- Inter-project material transfer deducts source stock and receiving confirmation registers the material at the destination.
- Existing consultant/company-paid expense, workflow files, fuel, machinery and project reporting features retained.

## Run locally
`pip install -r requirements.txt`
`python app.py`

## Render
Start command: `gunicorn app:app`
Configure `BAGC_DATA_DIR` to a persistent disk or use an external database for permanent production data.

## V27 additions
- Project Manager can administer all modules inside assigned projects, including BOQ, daily reporting, machinery, fuel, manpower, store, finance, design, reports and project requests.
- Project resource requests now use a project approval chain: Office Engineer -> Office Head -> Project Manager.
- After project-side approval, the Project Manager explicitly routes the request to Head Office by selecting Department -> Team -> Responsible Person.
- Head Office recipients can approve, reject, or forward requests to another Head Office person; the full request route is preserved in `request_steps`.
- Final approval automatically registers the approved resource in the appropriate project register (Store, Fuel, Machinery, Manpower, Finance or Design).
- Projects support unlimited responsible personnel by responsibility area, including Head Office personnel who can be responsible for multiple projects.
- Optional Excel import added for project manpower and machinery registers. Manual entry remains available.
