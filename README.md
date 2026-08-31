# BAGC Construction Management V9

Multi-project construction management web application for Birhanu Abebe General Contractor.

## Main capabilities
- Super Admin sees all projects and all-company income/expenses.
- Staff see only assigned projects.
- Department permissions: Design, Machinery, Finance, HR, Store, Project.
- Daily project reporting: BOQ work, machinery time, manpower, store movements and other finance.
- BOQ income = executed quantity × BOQ rate.
- Weekly/monthly BOQ report: Previous Qty, This Period Qty, To-Date Qty; Previous Amount, This Period Amount, To-Date Amount.
- Machinery fleet: add/remove Dozer, Excavator, Dump Truck, Crusher, Asphalt Plant, etc.; owned/rental; work/idle/down; idle/down reasons; gauge and fuel consumption; hourly expense.
- Store: common/concrete/rebar/formwork/sanitary/aluminium/electrical/fuel/etc.; received/issued/balance; physical ending balance.
- Manpower: permanent/temporary, attendance, rates, overtime.
- Design approval workflow.
- Excel BOQ import using Item No, Description, Unit, Contract Quantity and Contract Rate headers.
- Animated construction dashboard and charts.

## Render
Build: `pip install -r requirements.txt`
Start: `gunicorn app:app`

Environment variables:
- `ADMIN_USERNAME`
- `ADMIN_PASSWORD`
- `SECRET_KEY`

V8 synchronizes the configured Super Admin credentials on application startup, including an existing SQLite database.


## V9 login fix
The authenticated dashboard is rendered directly at `/` after sign-in, while `/dashboard` remains available as an alias. This avoids a post-login redirect 404 on hosted deployments.


## V11 login fix
Login now authenticates the user and redirects to the canonical `/dashboard` route. The Super Admin is synchronized from `ADMIN_USERNAME` and `ADMIN_PASSWORD` before authentication, including on an existing SQLite database.

## V14 additions
- Project contract baseline: signed date, commencement, end date, contract days, contract value
- Planned vs actual income and physical progress
- Ahead/behind schedule and time discrepancy dashboard
- Project and company expense percentage charts
- Daily machinery/material/fuel lists on dashboard
- Fuel and gauge reporting with cost and consumption
- Site crew hierarchy and grouped positions
- Expanded construction material catalogue
- Print-friendly reporting foundation


## V14.1 additions
- Editable project contract baseline
- Planned vs actual income controls
- Physical vs schedule variance and ahead/behind indicators
- Expanded construction material categories/catalogue
- Fuel register and printable controls
- Site crew hierarchy groups and position catalogue
