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
