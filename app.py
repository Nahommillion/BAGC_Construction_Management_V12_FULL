import os, sqlite3, calendar, datetime as dt, json
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, send_from_directory
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from openpyxl import load_workbook

BASE=os.path.dirname(os.path.abspath(__file__))
# Keep application data outside the code directory when BAGC_DATA_DIR is configured.
# This prevents redeploying/replacing the source from replacing the database.
DATA_DIR=os.environ.get("BAGC_DATA_DIR", "/data" if os.path.isdir("/data") and os.access("/data", os.W_OK) else BASE).strip() or BASE
os.makedirs(DATA_DIR, exist_ok=True)
DB=os.path.join(DATA_DIR,"bagc.db")
UPLOADS=os.path.join(DATA_DIR,"uploads")
USER_PHOTOS=os.path.join(UPLOADS,"user_photos")
os.makedirs(UPLOADS,exist_ok=True)
os.makedirs(USER_PHOTOS,exist_ok=True)
ALLOWED_PHOTO_EXT={"jpg","jpeg","png","webp"}
app=Flask(__name__)
app.secret_key=os.environ.get("SECRET_KEY","bagc-change-this-secret")

DEPARTMENTS=["Administration","Design","Machinery","Finance","HR","Store","Project"]
MACHINE_TYPES=["Dozer","Excavator","Wheel Loader","Backhoe Loader","Motor Grader","Roller","Dump Truck","Water Truck","Crane","Forklift","Concrete Mixer","Concrete Pump","Batching Plant","Crusher","Asphalt Plant","Asphalt Paver","Bitumen Distributor","Road Sweeper","Generator","Welding Machine","Vibrator","Air Compressor","Pickup","Other"]
MATERIAL_CATEGORIES=["Common Construction","Earthworks","Concrete","Rebar","Structural Steel / RHS","Formwork","Masonry","Roofing","Waterproofing","Finishing","Tiles","Natural Stone","Sanitary","Plumbing","Drainage","Electrical","Low Voltage / ICT","Aluminium","Glass","Road Works","Culverts","Landscaping","Fuel & Oil","Lubricants","Spare Parts","Welding / Cutting","PPE / Safety","Stationery & Cleaning","Tools","Other"]
MATERIAL_CATALOG=[
"Cement OPC","Cement PPC","Cement Rapid Hardening","Sand Fine","Sand Coarse","Quarry Dust","Selected Fill","Subbase","Base Course",
"Aggregate 5mm","Aggregate 10mm","Aggregate 14mm","Aggregate 20mm","Aggregate 25mm","Aggregate 40mm","Aggregate 50mm","Water",
"Rebar Ø6","Rebar Ø8","Rebar Ø10","Rebar Ø12","Rebar Ø14","Rebar Ø16","Rebar Ø20","Rebar Ø25","Rebar Ø32","Binding Wire 0.9mm","Binding Wire 1.2mm","Black Wire Roll",
"Welded Wire Mesh","Tie Wire","Anchor Bolt","Steel Plate","Angle Bar","Flat Bar","Round Bar","Channel","I-Beam","H-Beam","RHS 40x40","RHS 50x50","RHS 60x60","RHS 80x80","RHS 100x50","RHS 100x100","SHS","GI Sheet","Corrugated Sheet","Roofing Screw","Self Tapping Screw",
"Plywood 4x8","Plywood 12mm","Plywood 15mm","Plywood 18mm","Marine Plywood","Eucalyptus Pole","Eucalyptus Plank","Timber 2x4","Timber 2x6","Timber 2x8","Timber 4x4","Scaffold Tube","Scaffold Coupler","Form Tie","Form Oil","Release Agent","Nails 1in","Nails 2in","Nails 3in","Nails 4in","Nails 5in","Nails 6in",
"Concrete Block","Hollow Concrete Block 10cm","Hollow Concrete Block 15cm","Hollow Concrete Block 20cm","Solid Block","Brick","Kerbstone","Interlock Paver","Concrete Pipe","U-Ditch","Manhole Cover",
"PVC Pipe 25mm","PVC Pipe 32mm","PVC Pipe 50mm","PVC Pipe 75mm","PVC Pipe 110mm","PVC Pipe 160mm","PVC Pipe 200mm","PVC Pipe 250mm","HDPE Pipe","PPR Pipe","UPVC Fitting","Elbow","Tee","Reducer","Valve","Manhole Ring","Geotextile","Waterproofing Membrane",
"Electrical Cable 1.5mm2","Electrical Cable 2.5mm2","Electrical Cable 4mm2","Electrical Cable 6mm2","Electrical Cable 10mm2","Electrical Cable 16mm2","Armoured Cable","Control Cable","Conduit PVC","Flexible Conduit","Junction Box","Distribution Board","MCB","MCCB","RCD","Isolator","Switch","Socket","Industrial Socket","Light Fixture","LED Light","Street Light","Earthing Rod","Earthing Cable","Cable Lug","Cable Tray","Cable Tie","Electrical Tape",
"Aluminium Frame","Aluminium Sheet","Aluminium Composite Panel","Glass 5mm","Glass 6mm","Tempered Glass","Laminated Glass","Silicone","Rubber Gasket","Door Handle","Lockset","Hinge","Door Closer",
"Ceramic Tile","Porcelain Tile","Granite Tile","Marble","Natural Stone Cladding","Stone Veneer","Terrazzo","Skirting Tile","Tile Adhesive","Grout","Cement Mortar","Waterproofing Additive","Paint Primer","Emulsion Paint","Enamel Paint","Exterior Paint","Thinner","Putty","Gypsum Board","Gypsum Powder","Ceiling Tile","Acoustic Panel","Insulation","Bituminous Paint",
"Sanitary WC","Wash Basin","Urinal","Shower Mixer","Tap","Bib Tap","Floor Drain","Bottle Trap","P-Trap","S-Trap","Flush Valve","Flush Tank","Flexible Hose","Mirror","Soap Holder","Towel Rail","Sanitary Silicone",
"Diesel","Petrol","Engine Oil","Hydraulic Oil","Gear Oil","Brake Fluid","Coolant","Grease","AdBlue","Bitumen 60/70","Emulsion","Prime Coat","Tack Coat","Welding Rod 2.5mm","Welding Rod 3.2mm","Welding Rod 4mm","Oxygen","Acetylene","Grinding Disc","Cutting Disc","Drill Bit","Saw Blade",
"Safety Helmet","Safety Vest","Safety Boot","Safety Glove","Safety Goggles","Ear Plug","Dust Mask","Harness","Reflective Tape","Barricade","Warning Sign","Cones","Fire Extinguisher","First Aid Kit",
"Other"]
CREW_GROUPS=["Project Management","Key Staff","Earthwork Crew","Structure Crew","Road/Culvert Crew","Equipment / Machinery","Store","DL","Data Collector","Survey Team","Supporting Staff","Office Staff","Skilled Labour","Semi-Skilled Labour","Non-Skilled Labour","Security / General Support"]
POSITION_CATALOG=["Project Manager","Deputy Project Manager","Construction Manager","Site Engineer","Office Engineer","Quantity Surveyor","Planning Engineer","Design Engineer","QA/QC Engineer","Materials Engineer","Surveyor","Survey Assistant","HSE Officer","Foreman","Earthwork Foreman","Structure Foreman","Road Foreman","DL","Data Collector","Store Keeper","Store Assistant","Mechanic","Electrician","Plumber","Mason","Carpenter","Steel Fixer","Welder","Painter","Aluminium Worker","Equipment Operator","Driver","Labourer","Security Guard","Cleaner","Office Assistant","Document Controller","Accountant","Procurement Officer","Other"]
DESIGN_STATUSES=["Draft","Submitted","Under Review","Approved","Approved with Comments","Revise & Resubmit","Rejected","As-Built","Handed Over"]


@app.context_processor
def template_helpers():
    return {"dt": dt}


def db():
    c=sqlite3.connect(DB)
    c.row_factory=sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    return c


DEPT_CODES={'Administration':'ADM','Design':'DSN','Machinery':'MCH','Finance':'FIN','HR':'HR','Store':'STR','Project':'PRJ'}
def make_staff_id(department, seq):
    code=DEPT_CODES.get(department,'STF')
    return f"BAGC-{code}-{dt.date.today().year}-{seq:04d}"


def init_db():
    if os.environ.get("RENDER") and not os.environ.get("BAGC_DATA_DIR"):
        app.logger.warning("BAGC_DATA_DIR is not set on Render. SQLite will be ephemeral; configure a persistent disk or external database for permanent users/reports/BOQ.")
    c=db()
    c.executescript('''
    CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY,full_name TEXT,username TEXT UNIQUE,password_hash TEXT,department TEXT,position TEXT,location TEXT,role TEXT,active INTEGER DEFAULT 1,staff_id TEXT UNIQUE,photo_filename TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS projects(id INTEGER PRIMARY KEY,name TEXT UNIQUE,code TEXT,location TEXT,client TEXT,consultant TEXT,status TEXT DEFAULT 'Active',start_date TEXT,end_date TEXT);
    CREATE TABLE IF NOT EXISTS user_projects(user_id INTEGER,project_id INTEGER,UNIQUE(user_id,project_id));
    CREATE TABLE IF NOT EXISTS boq(id INTEGER PRIMARY KEY,project_id INTEGER,item_no TEXT,description TEXT,unit TEXT,rate REAL DEFAULT 0,contract_qty REAL DEFAULT 0,source_sheet TEXT,series TEXT DEFAULT '',title TEXT DEFAULT '',UNIQUE(project_id,item_no));
    CREATE TABLE IF NOT EXISTS boq_settings(id INTEGER PRIMARY KEY,project_id INTEGER UNIQUE,title TEXT DEFAULT '',revision TEXT DEFAULT '',effective_date TEXT);
    CREATE TABLE IF NOT EXISTS daily_work(id INTEGER PRIMARY KEY,project_id INTEGER,date TEXT,boq_id INTEGER,quantity REAL,station_from TEXT,station_to TEXT,notes TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS machines(id INTEGER PRIMARY KEY,project_id INTEGER,machine_type TEXT,code TEXT,ownership TEXT,hourly_rate REAL DEFAULT 0,expected_fuel REAL DEFAULT 0,active INTEGER DEFAULT 1);
    CREATE TABLE IF NOT EXISTS machine_logs(id INTEGER PRIMARY KEY,project_id INTEGER,machine_id INTEGER,date TEXT,work_hours REAL DEFAULT 0,idle_hours REAL DEFAULT 0,idle_reason TEXT,idle_payable INTEGER DEFAULT 0,down_hours REAL DEFAULT 0,down_reason TEXT,opening_gauge REAL DEFAULT 0,fuel_received REAL DEFAULT 0,closing_gauge REAL DEFAULT 0,notes TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS materials(id INTEGER PRIMARY KEY,project_id INTEGER,category TEXT,name TEXT,unit TEXT,min_stock REAL DEFAULT 0,active INTEGER DEFAULT 1,UNIQUE(project_id,name));
    CREATE TABLE IF NOT EXISTS store_logs(id INTEGER PRIMARY KEY,project_id INTEGER,material_id INTEGER,date TEXT,received REAL DEFAULT 0,issued REAL DEFAULT 0,unit_cost REAL DEFAULT 0,physical_balance REAL,reference TEXT,notes TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS manpower(id INTEGER PRIMARY KEY,project_id INTEGER,date TEXT,name TEXT,employment TEXT,position TEXT,present REAL DEFAULT 1,working_hours REAL DEFAULT 8,hourly_rate REAL DEFAULT 0,daily_rate REAL DEFAULT 0,normal_ot_hours REAL DEFAULT 0,normal_ot_rate REAL DEFAULT 0,night_ot_hours REAL DEFAULT 0,night_ot_rate REAL DEFAULT 0,sunday_ot_hours REAL DEFAULT 0,sunday_ot_rate REAL DEFAULT 0,holiday_ot_hours REAL DEFAULT 0,holiday_ot_rate REAL DEFAULT 0,overtime_hours REAL DEFAULT 0,overtime_rate REAL DEFAULT 0,notes TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS design_items(id INTEGER PRIMARY KEY,project_id INTEGER,drawing_no TEXT,title TEXT,discipline TEXT,revision TEXT,status TEXT,submitted TEXT,approved TEXT,comments TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS finance_logs(id INTEGER PRIMARY KEY,project_id INTEGER,date TEXT,category TEXT,kind TEXT,description TEXT,amount REAL DEFAULT 0,reference TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS boq_uploads(id INTEGER PRIMARY KEY,project_id INTEGER,filename TEXT,uploaded_at TEXT,user_id INTEGER,rows_imported INTEGER DEFAULT 0);
    CREATE TABLE IF NOT EXISTS performance_rates(id INTEGER PRIMARY KEY,project_id INTEGER,work_type TEXT,worker_type TEXT,unit TEXT,qty_per_hour REAL DEFAULT 0,notes TEXT,UNIQUE(project_id,work_type,worker_type));
    CREATE TABLE IF NOT EXISTS daily_activities(id INTEGER PRIMARY KEY,project_id INTEGER,date TEXT,boq_id INTEGER,work_type TEXT,executed_qty REAL DEFAULT 0,machine_id INTEGER,machine_hours REAL DEFAULT 0,manpower_position TEXT,manpower_qty REAL DEFAULT 0,manpower_hours REAL DEFAULT 0,material_id INTEGER,material_qty REAL DEFAULT 0,remarks TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS problems(id INTEGER PRIMARY KEY,project_id INTEGER,date TEXT,problem TEXT,remark TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS fuel_logs(id INTEGER PRIMARY KEY,project_id INTEGER,machine_id INTEGER,date TEXT,opening_gauge REAL DEFAULT 0,fuel_received REAL DEFAULT 0,closing_gauge REAL DEFAULT 0,fuel_price REAL DEFAULT 0,reference TEXT,notes TEXT,user_id INTEGER,source TEXT DEFAULT 'Fuel Register');
    CREATE TABLE IF NOT EXISTS project_crews(id INTEGER PRIMARY KEY,project_id INTEGER,date TEXT,group_name TEXT,position TEXT,name TEXT,employment TEXT,skill_level TEXT,working_hours REAL DEFAULT 0,hourly_rate REAL DEFAULT 0,notes TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS crew_group_capacity(id INTEGER PRIMARY KEY,group_name TEXT UNIQUE,foreman_qty REAL DEFAULT 0,dl_qty REAL DEFAULT 0,surveyor_qty REAL DEFAULT 0,data_collector_qty REAL DEFAULT 0,time_keeper_qty REAL DEFAULT 0,other_qty REAL DEFAULT 0,total_qty REAL DEFAULT 0,active INTEGER DEFAULT 1);
    CREATE TABLE IF NOT EXISTS activity_machines(id INTEGER PRIMARY KEY,activity_id INTEGER,machine_log_id INTEGER,machine_id INTEGER,hours REAL DEFAULT 0);
    CREATE TABLE IF NOT EXISTS activity_manpower(id INTEGER PRIMARY KEY,activity_id INTEGER,manpower_id INTEGER,crew_id INTEGER,qty REAL DEFAULT 0,hours REAL DEFAULT 0);
    CREATE TABLE IF NOT EXISTS activity_store(id INTEGER PRIMARY KEY,activity_id INTEGER,store_log_id INTEGER,material_id INTEGER,qty REAL DEFAULT 0);
    CREATE TABLE IF NOT EXISTS activity_fuel(id INTEGER PRIMARY KEY,activity_id INTEGER,fuel_log_id INTEGER,litres REAL DEFAULT 0);
    CREATE TABLE IF NOT EXISTS activity_finance(id INTEGER PRIMARY KEY,activity_id INTEGER,finance_log_id INTEGER,amount REAL DEFAULT 0);
    CREATE TABLE IF NOT EXISTS crew_evaluations(id INTEGER PRIMARY KEY,activity_id INTEGER,crew_id INTEGER,evaluation TEXT,remarks TEXT,score REAL);
    CREATE TABLE IF NOT EXISTS variation_alerts(id INTEGER PRIMARY KEY,project_id INTEGER,boq_id INTEGER,date TEXT,contract_qty REAL,previous_qty REAL,period_qty REAL,to_date_qty REAL,excess_qty REAL,status TEXT DEFAULT 'OPEN',message TEXT,created_by INTEGER,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS crew_groups(id INTEGER PRIMARY KEY,name TEXT UNIQUE,active INTEGER DEFAULT 1);
    CREATE TABLE IF NOT EXISTS crew_positions(id INTEGER PRIMARY KEY,name TEXT UNIQUE,active INTEGER DEFAULT 1);
    CREATE TABLE IF NOT EXISTS report_settings(id INTEGER PRIMARY KEY,project_id INTEGER UNIQUE,contractor_role TEXT DEFAULT 'Main Contractor',phone TEXT,email TEXT,website TEXT,fax TEXT,address TEXT,logo_text TEXT);
    CREATE TABLE IF NOT EXISTS rfis(id INTEGER PRIMARY KEY,project_id INTEGER,rfi_no TEXT,date_requested TEXT,inspection_date TEXT,location TEXT,boq_id INTEGER,work_description TEXT,drawing_no TEXT,drawing_revision TEXT,specification TEXT,work_stage TEXT,submitted_by INTEGER,status TEXT DEFAULT 'PENDING INSPECTION',overall_comment TEXT,corrective_action TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS rfi_inspections(id INTEGER PRIMARY KEY,rfi_id INTEGER,inspector_user_id INTEGER,inspector_role TEXT,decision TEXT DEFAULT 'PENDING',comments TEXT,inspection_date TEXT,signed_at TEXT,UNIQUE(rfi_id,inspector_user_id));
    CREATE TABLE IF NOT EXISTS rfi_steps(id INTEGER PRIMARY KEY,rfi_id INTEGER,step_order INTEGER,stage TEXT,assigned_user_id INTEGER,decision TEXT DEFAULT 'PENDING',comments TEXT,inspection_date TEXT,signed_at TEXT,UNIQUE(rfi_id,step_order));
    CREATE TABLE IF NOT EXISTS saved_reports(id INTEGER PRIMARY KEY,project_id INTEGER,report_no TEXT,report_type TEXT,scope TEXT DEFAULT 'ALL',start_date TEXT,end_date TEXT,generated_by INTEGER,generated_at TEXT DEFAULT CURRENT_TIMESTAMP,snapshot_json TEXT,source_report_ids TEXT DEFAULT '[]',UNIQUE(project_id,report_type,scope,start_date,end_date));
    CREATE TABLE IF NOT EXISTS machine_assignments(id INTEGER PRIMARY KEY,machine_id INTEGER,project_id INTEGER,start_date TEXT,start_hour REAL DEFAULT 0,end_date TEXT,end_hour REAL,status TEXT DEFAULT 'ACTIVE',assigned_by INTEGER,ended_by INTEGER,ended_at TEXT,notes TEXT);
    ''')
    # Safe migrations for databases created by earlier BAGC versions.
    existing_bq=[r['name'] for r in c.execute("PRAGMA table_info(boq)").fetchall()]
    if 'series' not in existing_bq: c.execute("ALTER TABLE boq ADD COLUMN series TEXT DEFAULT ''")
    if 'title' not in existing_bq: c.execute("ALTER TABLE boq ADD COLUMN title TEXT DEFAULT ''")
    existing_sr=[r['name'] for r in c.execute("PRAGMA table_info(saved_reports)").fetchall()]
    if 'source_report_ids' not in existing_sr: c.execute("ALTER TABLE saved_reports ADD COLUMN source_report_ids TEXT DEFAULT '[]'")
    existing_u=[r['name'] for r in c.execute("PRAGMA table_info(users)").fetchall()]
    if 'position' not in existing_u: c.execute("ALTER TABLE users ADD COLUMN position TEXT")
    if 'staff_id' not in existing_u: c.execute("ALTER TABLE users ADD COLUMN staff_id TEXT")
    if 'photo_filename' not in existing_u: c.execute("ALTER TABLE users ADD COLUMN photo_filename TEXT")
    for ur in c.execute("SELECT id,department,staff_id FROM users").fetchall():
        if not ur['staff_id'] or str(ur['staff_id']).startswith('BAGC-') and str(ur['staff_id'])[5:].isdigit(): c.execute("UPDATE users SET staff_id=? WHERE id=?",(make_staff_id(ur['department'],ur['id']),ur['id']))
    existing=[r['name'] for r in c.execute("PRAGMA table_info(machines)").fetchall()]
    for col,typ in [('plate_no','TEXT'),('engine_no','TEXT'),('fuel_price','REAL DEFAULT 0')]:
        if col not in existing: c.execute(f"ALTER TABLE machines ADD COLUMN {col} {typ}")
    existing_mp=[r['name'] for r in c.execute("PRAGMA table_info(manpower)").fetchall()]
    if 'crew_id' not in existing_mp: c.execute("ALTER TABLE manpower ADD COLUMN crew_id INTEGER")
    for col,typ in [('working_hours','REAL DEFAULT 8'),('hourly_rate','REAL DEFAULT 0'),('normal_ot_hours','REAL DEFAULT 0'),('normal_ot_rate','REAL DEFAULT 0'),('night_ot_hours','REAL DEFAULT 0'),('night_ot_rate','REAL DEFAULT 0'),('sunday_ot_hours','REAL DEFAULT 0'),('sunday_ot_rate','REAL DEFAULT 0'),('holiday_ot_hours','REAL DEFAULT 0'),('holiday_ot_rate','REAL DEFAULT 0'),('overtime_hours','REAL DEFAULT 0'),('overtime_rate','REAL DEFAULT 0')]:
        if col not in existing_mp: c.execute(f"ALTER TABLE manpower ADD COLUMN {col} {typ}")
    existing_ml=[r['name'] for r in c.execute("PRAGMA table_info(machine_logs)").fetchall()]
    if 'idle_payable' not in existing_ml: c.execute("ALTER TABLE machine_logs ADD COLUMN idle_payable INTEGER DEFAULT 0")
    existing_f=[r['name'] for r in c.execute("PRAGMA table_info(fuel_logs)").fetchall()]
    if 'source' not in existing_f: c.execute("ALTER TABLE fuel_logs ADD COLUMN source TEXT DEFAULT 'Fuel Register'")
    existing_m=[r['name'] for r in c.execute("PRAGMA table_info(machines)").fetchall()]
    for col,typ in [('assignment_start_date','TEXT'),('assignment_start_hour','REAL DEFAULT 0'),('assignment_end_date','TEXT'),('assignment_end_hour','REAL'),('total_signed_hours','REAL DEFAULT 0'),('hours_used','REAL DEFAULT 0'),('lifecycle_status',"TEXT DEFAULT 'ACTIVE'"),('assignment_signed_by','INTEGER'),('assignment_ended_by','INTEGER'),('assignment_ended_at','TEXT')]:
        if col not in existing_m: c.execute(f"ALTER TABLE machines ADD COLUMN {col} {typ}")
    for g in CREW_GROUPS: c.execute("INSERT OR IGNORE INTO crew_groups(name) VALUES(?)",(g,))
    for pos in POSITION_CATALOG: c.execute("INSERT OR IGNORE INTO crew_positions(name) VALUES(?)",(pos,))
    for g in CREW_GROUPS: c.execute("INSERT OR IGNORE INTO crew_group_capacity(group_name) VALUES(?)",(g,))
    existing_p=[r['name'] for r in c.execute("PRAGMA table_info(projects)").fetchall()]
    for col,typ in [('contractor_role',"TEXT DEFAULT 'Main Contractor'"),('contract_sign_date','TEXT'),('commencement_date','TEXT'),('contract_end_date','TEXT'),('contract_days','INTEGER DEFAULT 0'),('planned_income','REAL DEFAULT 0'),('planned_physical_pct','REAL DEFAULT 0'),('contract_value','REAL DEFAULT 0')]:
        if col not in existing_p: c.execute(f"ALTER TABLE projects ADD COLUMN {col} {typ}")
    # ENV-controlled Super Admin synchronization fixes an already-created SQLite DB.
    u=os.environ.get("ADMIN_USERNAME","admin").strip() or "admin"
    p=os.environ.get("ADMIN_PASSWORD","admin123")
    admin=c.execute("SELECT id FROM users WHERE role='SUPER_ADMIN' ORDER BY id LIMIT 1").fetchone()
    if not admin:
        c.execute("INSERT INTO users(full_name,username,password_hash,department,position,location,role) VALUES(?,?,?,?,?,?,?)",("System Administrator",u,generate_password_hash(p),"Administration","Super Admin","Head Office","SUPER_ADMIN"))
    else:
        c.execute("UPDATE users SET username=?,password_hash=?,active=1,department='Administration',position=COALESCE(position,'Super Admin'),location='Head Office' WHERE id=?",(u,generate_password_hash(p),admin["id"]))
    if not c.execute("SELECT id FROM projects").fetchone():
        c.execute("INSERT INTO projects(name,code,location,status) VALUES(?,?,?,?)",("Koye Feche","KOYE","Koye Feche","Active"))
    for ur in c.execute("SELECT id,department FROM users WHERE staff_id IS NULL OR staff_id=''").fetchall(): c.execute("UPDATE users SET staff_id=? WHERE id=?",(make_staff_id(ur['department'],ur['id']),ur['id']))
    c.commit();c.close()


def current_user():
    if not session.get("user_id"): return None
    c=db(); u=c.execute("SELECT * FROM users WHERE id=? AND active=1",(session["user_id"],)).fetchone(); c.close(); return u

@app.context_processor
def inject():
    return {"me":current_user(),"machine_types":MACHINE_TYPES,"material_categories":MATERIAL_CATEGORIES,"material_catalog":MATERIAL_CATALOG,"design_statuses":DESIGN_STATUSES,"today":dt.date.today().isoformat(),"crew_groups":CREW_GROUPS,"position_catalog":POSITION_CATALOG}


def login_required(f):
    @wraps(f)
    def w(*a,**k):
        if not current_user(): return redirect(url_for("login"))
        return f(*a,**k)
    return w


def admin_required(f):
    @wraps(f)
    def w(*a,**k):
        u=current_user()
        if not u or u["role"]!="SUPER_ADMIN":
            flash("👑 Super Admin access required.","error"); return redirect(url_for("dashboard"))
        return f(*a,**k)
    return w


def allowed_project(pid):
    u=current_user()
    if not u:return False
    if u["role"]=="SUPER_ADMIN":return True
    c=db();ok=c.execute("SELECT 1 FROM user_projects WHERE user_id=? AND project_id=?",(u["id"],pid)).fetchone();c.close();return bool(ok)


def can_module(module):
    u=current_user()
    if not u:return False
    if u["role"]=="SUPER_ADMIN":return True
    if u["department"]=="Project":return True
    return u["department"]==module


def parse_float(v):
    try:return float(v or 0)
    except:return 0.0


def period_bounds(period, anchor):
    d=dt.date.fromisoformat(anchor)
    if period=="day": return d,d
    if period=="week":
        start=d-dt.timedelta(days=d.weekday()); return start,start+dt.timedelta(days=6)
    start=d.replace(day=1); end=d.replace(day=calendar.monthrange(d.year,d.month)[1]); return start,end


def money(v):return round(float(v or 0),2)

def snapshot_json(obj):
    return json.dumps(obj, default=lambda x: dict(x) if isinstance(x, sqlite3.Row) else str(x))

def report_dates(report_type, start, end):
    if start and end: return dt.date.fromisoformat(start), dt.date.fromisoformat(end)
    today=dt.date.today()
    if report_type=='DAILY': return today,today
    if report_type=='WEEKLY': return today-dt.timedelta(days=today.weekday()),today
    if report_type=='MONTHLY': return today.replace(day=1),today
    if report_type=='SEMI_ANNUAL': return (dt.date(today.year,1,1),dt.date(today.year,6,30)) if today.month<=6 else (dt.date(today.year,7,1),dt.date(today.year,12,31))
    if report_type=='ANNUAL': return dt.date(today.year,1,1),dt.date(today.year,12,31)
    return today,today

def build_report_snapshot(pid,start,end,scope='ALL'):
    c=db(); out={'project_id':pid,'start_date':start.isoformat(),'end_date':end.isoformat(),'scope':scope}
    if scope in ('ALL','BOQ'):
        rows=c.execute("SELECT b.*,COALESCE(SUM(CASE WHEN dw.date<? THEN dw.quantity ELSE 0 END),0) previous_qty,COALESCE(SUM(CASE WHEN dw.date BETWEEN ? AND ? THEN dw.quantity ELSE 0 END),0) period_qty,COALESCE(SUM(dw.quantity),0) todate_qty FROM boq b LEFT JOIN daily_work dw ON dw.boq_id=b.id WHERE b.project_id=? GROUP BY b.id ORDER BY b.item_no",(start.isoformat(),start.isoformat(),end.isoformat(),pid)).fetchall()
        out['boq']=[dict(r,previous_amount=r['previous_qty']*r['rate'],period_amount=r['period_qty']*r['rate'],todate_amount=r['todate_qty']*r['rate']) for r in rows]
    if scope in ('ALL','MACHINERY'):
        out['machinery']=[dict(r) for r in c.execute("SELECT ml.*,m.machine_type,m.code,m.plate_no,m.engine_no,m.ownership,m.hourly_rate,m.expected_fuel,((ml.work_hours+CASE WHEN ml.idle_payable=1 THEN ml.idle_hours ELSE 0 END)*m.hourly_rate) expense,(ml.opening_gauge+ml.fuel_received-ml.closing_gauge) actual_fuel,(ml.work_hours*m.expected_fuel) expected_fuel FROM machine_logs ml JOIN machines m ON m.id=ml.machine_id WHERE ml.project_id=? AND ml.date BETWEEN ? AND ? ORDER BY ml.date,ml.id",(pid,start.isoformat(),end.isoformat())).fetchall()]
    if scope in ('ALL','MANPOWER'):
        out['manpower']=[dict(r) for r in c.execute("SELECT * FROM manpower WHERE project_id=? AND date BETWEEN ? AND ? ORDER BY date,id",(pid,start.isoformat(),end.isoformat())).fetchall()]
    if scope in ('ALL','STORE'):
        out['store']=[dict(r) for r in c.execute("SELECT sl.*,m.name,m.category,m.unit FROM store_logs sl JOIN materials m ON m.id=sl.material_id WHERE sl.project_id=? AND sl.date BETWEEN ? AND ? ORDER BY sl.date,sl.id",(pid,start.isoformat(),end.isoformat())).fetchall()]
    if scope in ('ALL','FUEL'):
        out['fuel']=[dict(r) for r in c.execute("SELECT f.*,m.machine_type,m.code,m.plate_no,m.engine_no,m.ownership,m.expected_fuel,COALESCE((SELECT SUM(ml.work_hours) FROM machine_logs ml WHERE ml.machine_id=f.machine_id AND ml.date=f.date),0) work_hours,(f.opening_gauge+f.fuel_received-f.closing_gauge) consumption,(f.fuel_received*f.fuel_price) cost FROM fuel_logs f JOIN machines m ON m.id=f.machine_id WHERE f.project_id=? AND f.date BETWEEN ? AND ? ORDER BY f.date,f.id",(pid,start.isoformat(),end.isoformat())).fetchall()]
    if scope in ('ALL','FINANCE'):
        out['finance']=[dict(r) for r in c.execute("SELECT * FROM finance_logs WHERE project_id=? AND date BETWEEN ? AND ? ORDER BY date,id",(pid,start.isoformat(),end.isoformat())).fetchall()]
    if scope in ('ALL','PROBLEMS'):
        out['variations']=[dict(r) for r in c.execute("SELECT va.*,b.item_no,b.description,b.unit FROM variation_alerts va JOIN boq b ON b.id=va.boq_id WHERE va.project_id=? AND va.date BETWEEN ? AND ? ORDER BY va.date,va.id",(pid,start.isoformat(),end.isoformat())).fetchall()]
        out['problems']=[dict(r) for r in c.execute("SELECT * FROM problems WHERE project_id=? AND date BETWEEN ? AND ? ORDER BY date,id",(pid,start.isoformat(),end.isoformat())).fetchall()]
    # summary totals, always useful for every saved report
    out['summary']={
        'income': c.execute("SELECT COALESCE(SUM(dw.quantity*b.rate),0) FROM daily_work dw JOIN boq b ON b.id=dw.boq_id WHERE dw.project_id=? AND dw.date BETWEEN ? AND ?",(pid,start.isoformat(),end.isoformat())).fetchone()[0],
        'machinery_expense': c.execute("SELECT COALESCE(SUM((ml.work_hours+CASE WHEN ml.idle_payable=1 THEN ml.idle_hours ELSE 0 END)*m.hourly_rate),0) FROM machine_logs ml JOIN machines m ON m.id=ml.machine_id WHERE ml.project_id=? AND ml.date BETWEEN ? AND ?",(pid,start.isoformat(),end.isoformat())).fetchone()[0],
        'manpower_expense': c.execute("SELECT COALESCE(SUM(CASE WHEN hourly_rate>0 THEN present*working_hours*hourly_rate ELSE present*daily_rate END+normal_ot_hours*normal_ot_rate+night_ot_hours*night_ot_rate+sunday_ot_hours*sunday_ot_rate+holiday_ot_hours*holiday_ot_rate),0) FROM manpower WHERE project_id=? AND date BETWEEN ? AND ?",(pid,start.isoformat(),end.isoformat())).fetchone()[0],
        'store_expense': c.execute("SELECT COALESCE(SUM(issued*unit_cost),0) FROM store_logs WHERE project_id=? AND date BETWEEN ? AND ?",(pid,start.isoformat(),end.isoformat())).fetchone()[0],
        'other_expense': c.execute("SELECT COALESCE(SUM(amount),0) FROM finance_logs WHERE project_id=? AND kind='Expense' AND date BETWEEN ? AND ?",(pid,start.isoformat(),end.isoformat())).fetchone()[0],
        'fuel_cost': c.execute("SELECT COALESCE(SUM(fuel_received*fuel_price),0) FROM fuel_logs WHERE project_id=? AND date BETWEEN ? AND ?",(pid,start.isoformat(),end.isoformat())).fetchone()[0]
    }
    c.close(); return out

def save_report(pid, report_type, start, end, scope='ALL', user_id=None):
    snap=build_report_snapshot(pid,start,end,scope); c=db(); existing=c.execute("SELECT id FROM saved_reports WHERE project_id=? AND report_type=? AND scope=? AND start_date=? AND end_date=?",(pid,report_type,scope,start.isoformat(),end.isoformat())).fetchone()
    if existing:
        rid=existing['id']; c.execute("UPDATE saved_reports SET snapshot_json=?,generated_by=?,generated_at=CURRENT_TIMESTAMP WHERE id=?",(snapshot_json(snap),user_id or session.get('user_id'),rid))
    else:
        n=c.execute("SELECT COUNT(*) FROM saved_reports WHERE project_id=? AND report_type=?",(pid,report_type)).fetchone()[0]+1; no=f"{report_type[:3]}-{dt.date.today().year}-{n:04d}"
        c.execute("INSERT INTO saved_reports(project_id,report_no,report_type,scope,start_date,end_date,generated_by,snapshot_json) VALUES(?,?,?,?,?,?,?,?)",(pid,no,report_type,scope,start.isoformat(),end.isoformat(),user_id or session.get('user_id'),snapshot_json(snap))); rid=c.execute("SELECT last_insert_rowid()").fetchone()[0]
    # Explicit lineage: higher reports retain the saved lower-level reports they consolidate.
    source_types={'WEEKLY':['DAILY'],'MONTHLY':['DAILY','WEEKLY'],'SEMI_ANNUAL':['MONTHLY','WEEKLY'],'ANNUAL':['SEMI_ANNUAL','MONTHLY']} .get(report_type,[])
    src=[]
    for rt in source_types:
        src += [x['id'] for x in c.execute("SELECT id FROM saved_reports WHERE project_id=? AND report_type=? AND start_date<=? AND end_date>=? ORDER BY start_date",(pid,rt,end.isoformat(),start.isoformat())).fetchall()]
    c.execute("UPDATE saved_reports SET source_report_ids=? WHERE id=?",(json.dumps(src),rid)); c.commit(); c.close(); return rid



def dashboard_data(pid=None):
    c=db(); where="" if pid is None else " WHERE p.id=?"; args=() if pid is None else (pid,)
    projects=c.execute("SELECT p.* FROM projects p"+where+" ORDER BY p.name",args).fetchall()
    out=[]
    for p in projects:
        inc=c.execute("SELECT COALESCE(SUM(dw.quantity*b.rate),0) x FROM daily_work dw JOIN boq b ON b.id=dw.boq_id WHERE dw.project_id=?",(p["id"],)).fetchone()["x"]
        me=c.execute("SELECT COALESCE(SUM(ml.work_hours*m.hourly_rate),0) x FROM machine_logs ml JOIN machines m ON m.id=ml.machine_id WHERE ml.project_id=?",(p["id"],)).fetchone()["x"]
        pe=c.execute("SELECT COALESCE(SUM((CASE WHEN mp.hourly_rate>0 THEN mp.present*mp.working_hours*mp.hourly_rate ELSE mp.present*mp.daily_rate END + mp.normal_ot_hours*mp.normal_ot_rate + mp.night_ot_hours*mp.night_ot_rate + mp.sunday_ot_hours*mp.sunday_ot_rate + mp.holiday_ot_hours*mp.holiday_ot_rate)),0) x FROM manpower mp WHERE mp.project_id=?",(p["id"],)).fetchone()["x"]
        se=c.execute("SELECT COALESCE(SUM(sl.issued*sl.unit_cost),0) x FROM store_logs sl WHERE sl.project_id=?",(p["id"],)).fetchone()["x"]
        other=c.execute("SELECT COALESCE(SUM(amount),0) x FROM finance_logs WHERE project_id=? AND kind='Expense'",(p["id"],)).fetchone()["x"]
        workers=c.execute("SELECT COUNT(*) x FROM manpower WHERE project_id=? AND present>0",(p["id"],)).fetchone()["x"]
        machines=c.execute("SELECT COUNT(*) x FROM machines WHERE project_id=? AND active=1",(p["id"],)).fetchone()["x"]
        total_exp=money(me+pe+se+other)
        daily_m=c.execute("SELECT ml.*,m.machine_type,m.code,m.plate_no,m.engine_no,m.ownership,((ml.work_hours + CASE WHEN ml.idle_payable=1 THEN ml.idle_hours ELSE 0 END)*m.hourly_rate) expense,(ml.opening_gauge+ml.fuel_received-ml.closing_gauge) actual_fuel FROM machine_logs ml JOIN machines m ON m.id=ml.machine_id WHERE ml.project_id=? ORDER BY ml.date DESC,ml.id DESC LIMIT 8",(p["id"],)).fetchall()
        daily_mat=c.execute("SELECT sl.*,m.name,m.category,m.unit FROM store_logs sl JOIN materials m ON m.id=sl.material_id WHERE sl.project_id=? ORDER BY sl.date DESC,sl.id DESC LIMIT 8",(p["id"],)).fetchall()
        fuel_recent=c.execute("SELECT f.*,m.machine_type,m.code,m.plate_no,m.engine_no,m.ownership,(f.opening_gauge+f.fuel_received-f.closing_gauge) consumption,(f.fuel_received*f.fuel_price) cost FROM fuel_logs f JOIN machines m ON m.id=f.machine_id WHERE f.project_id=? ORDER BY f.date DESC,f.id DESC LIMIT 8",(p["id"],)).fetchall()
        contract_value=p["contract_value"] or c.execute("SELECT COALESCE(SUM(contract_qty*rate),0) FROM boq WHERE project_id=?",(p["id"],)).fetchone()[0]
        physical_pct=(inc/contract_value*100) if contract_value else 0
        start=p["commencement_date"] or p["start_date"]; end=p["contract_end_date"] or p["end_date"]; schedule_pct=0; time_var=0; days_remaining=None
        if start and end:
            sd=dt.date.fromisoformat(start); ed=dt.date.fromisoformat(end); total=max((ed-sd).days,0); elapsed=max(min((dt.date.today()-sd).days,total),0); schedule_pct=(elapsed/total*100) if total else 0; time_var=physical_pct-schedule_pct; days_remaining=(ed-dt.date.today()).days
        out.append({"p":p,"income":money(inc),"expense":total_exp,"expense_pct":money((total_exp/(inc+0.0001))*100) if inc else 0,"machine_expense":money(me),"manpower_expense":money(pe),"store_expense":money(se),"other_expense":money(other),"workers":workers,"machines":machines,"daily_machines":daily_m,"daily_materials":daily_mat,"fuel_recent":fuel_recent,"contract_value":money(contract_value),"physical_pct":money(physical_pct),"schedule_pct":money(schedule_pct),"time_variance_pct":money(time_var),"days_remaining":days_remaining,"planned_income":money(p["planned_income"] or 0)})
    company_exp=sum(x["expense"] for x in out) or 1
    out=[dict(x,company_expense_pct=money((x["expense"]/company_exp)*100)) for x in out]
    c.close();return out

@app.route("/")
def home():
    # Render the authenticated dashboard directly from / so the post-login
    # flow does not depend on a second redirect that can produce a 404.
    u=current_user()
    if not u:
        return redirect(url_for("login"))
    c=db()
    if u["role"]=="SUPER_ADMIN":
        projects=c.execute("SELECT p.* FROM projects p ORDER BY p.name").fetchall()
    else:
        projects=c.execute("SELECT p.* FROM projects p JOIN user_projects up ON up.project_id=p.id WHERE up.user_id=? ORDER BY p.name",(u["id"],)).fetchall()
    c.close()
    data=dashboard_data(); allowed_ids={p["id"] for p in projects}; data=[x for x in data if x["p"]["id"] in allowed_ids]
    totals={k:sum(x[k] for x in data) for k in ["income","expense","machine_expense","manpower_expense","store_expense","other_expense"]}
    return render_template("dashboard.html",data=data,totals=totals)

@app.route("/login", methods=["GET", "POST"])
def login():
    # Always initialize/synchronize the configured Super Admin before authentication.
    # This makes a fresh Render deployment and a changed Render ENV behave the same.
    init_db()
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        c = db()
        u = c.execute("SELECT * FROM users WHERE username=? AND active=1", (username,)).fetchone()
        c.close()
        if u and check_password_hash(u["password_hash"], password):
            session.clear()
            session["user_id"] = int(u["id"])
            session.permanent = True
            # Use a normal redirect to the canonical dashboard endpoint.
            # The dashboard route itself rebuilds the project scope from the session.
            return redirect(url_for("dashboard"))
        flash("🔐 Invalid username or password.", "error")
    return render_template("login.html")

@app.route("/logout")
def logout():session.clear();return redirect(url_for("login"))

@app.route("/dashboard")
@login_required
def dashboard():
    return home()

@app.route("/home")
@login_required
def home_alias():
    return home()

@app.route("/projects/<int:pid>")
@login_required
def project(pid):
    if not allowed_project(pid): flash("🚫 You do not have access to this project.","error"); return redirect(url_for("dashboard"))
    c=db(); p=c.execute("SELECT * FROM projects WHERE id=?",(pid,)).fetchone()
    if not p: c.close(); return redirect(url_for("dashboard"))
    boq_count=c.execute("SELECT COUNT(*) n FROM boq WHERE project_id=?",(pid,)).fetchone()["n"]
    machine_count=c.execute("SELECT COUNT(*) n FROM machines WHERE project_id=? AND active=1",(pid,)).fetchone()["n"]
    mat_count=c.execute("SELECT COUNT(*) n FROM materials WHERE project_id=? AND active=1",(pid,)).fetchone()["n"]
    actual_income=c.execute("SELECT COALESCE(SUM(dw.quantity*b.rate),0) x FROM daily_work dw JOIN boq b ON b.id=dw.boq_id WHERE dw.project_id=?",(pid,)).fetchone()["x"]
    actual_expense=c.execute("SELECT COALESCE(SUM((ml.work_hours + CASE WHEN ml.idle_payable=1 THEN ml.idle_hours ELSE 0 END)*m.hourly_rate),0) FROM machine_logs ml JOIN machines m ON m.id=ml.machine_id WHERE ml.project_id=?",(pid,)).fetchone()[0]
    actual_expense+=c.execute("SELECT COALESCE(SUM((CASE WHEN mp.hourly_rate>0 THEN mp.present*mp.working_hours*mp.hourly_rate ELSE mp.present*mp.daily_rate END + mp.normal_ot_hours*mp.normal_ot_rate + mp.night_ot_hours*mp.night_ot_rate + mp.sunday_ot_hours*mp.sunday_ot_rate + mp.holiday_ot_hours*mp.holiday_ot_rate)),0) FROM manpower mp WHERE mp.project_id=?",(pid,)).fetchone()[0]
    actual_expense+=c.execute("SELECT COALESCE(SUM(issued*unit_cost),0) FROM store_logs WHERE project_id=?",(pid,)).fetchone()[0]
    actual_expense+=c.execute("SELECT COALESCE(SUM(amount),0) FROM finance_logs WHERE project_id=? AND kind='Expense'",(pid,)).fetchone()[0]
    contract_value=p["contract_value"] or c.execute("SELECT COALESCE(SUM(contract_qty*rate),0) FROM boq WHERE project_id=?",(pid,)).fetchone()[0]
    physical_pct=(actual_income/contract_value*100) if contract_value else 0
    today=dt.date.today(); start=p["commencement_date"] or p["start_date"]; end=p["contract_end_date"] or p["end_date"]
    elapsed=0; schedule_pct=0; time_variance_pct=0; days_remaining=None
    if start and end:
        sd=dt.date.fromisoformat(start); ed=dt.date.fromisoformat(end); total=max((ed-sd).days,0); elapsed=max(min((today-sd).days,total),0); schedule_pct=(elapsed/total*100) if total else 0; time_variance_pct=physical_pct-schedule_pct; days_remaining=(ed-today).days
    planned_income=p["planned_income"] or 0; income_variance=actual_income-planned_income
    crew_count=c.execute("SELECT COUNT(*) n FROM project_crews WHERE project_id=?",(pid,)).fetchone()["n"]
    c.close()
    return render_template("project.html",p=p,boq_count=boq_count,machine_count=machine_count,mat_count=mat_count,actual_income=actual_income,actual_expense=actual_expense,contract_value=contract_value,physical_pct=physical_pct,schedule_pct=schedule_pct,time_variance_pct=time_variance_pct,days_remaining=days_remaining,planned_income=planned_income,income_variance=income_variance,crew_count=crew_count)


def variation_check(c,pid,boq_id,additional_qty,d):
    b=c.execute("SELECT * FROM boq WHERE id=? AND project_id=?",(boq_id,pid)).fetchone()
    if not b: raise ValueError("Selected BOQ item was not found in this project.")
    prev=c.execute("SELECT COALESCE(SUM(quantity),0) FROM daily_work WHERE boq_id=? AND date<?",(boq_id,d)).fetchone()[0]
    period=c.execute("SELECT COALESCE(SUM(quantity),0) FROM daily_work WHERE boq_id=? AND date BETWEEN ? AND ?",(boq_id,d,d)).fetchone()[0]
    tod=prev+period
    new_tod=tod+additional_qty
    excess=max(new_tod-(b['contract_qty'] or 0),0)
    if excess>0:
        msg=f"Variation Order required: BOQ {b['item_no']} will exceed contract quantity by {excess:g} {b['unit']}."
        c.execute("INSERT INTO variation_alerts(project_id,boq_id,date,contract_qty,previous_qty,period_qty,to_date_qty,excess_qty,status,message,created_by) VALUES(?,?,?,?,?,?,?,?,?,?,?)",(pid,boq_id,d,b['contract_qty'] or 0,prev,period, new_tod, excess,'OPEN',msg,current_user()['id']))
        return msg
    return None

@app.errorhandler(Exception)
def application_error(error):
    # Never expose a generic blank 500 to field staff; log the actual exception and return a useful message.
    import traceback
    app.logger.error("Unhandled BAGC application error: %s\n%s", error, traceback.format_exc())
    if request.path.startswith('/projects/'):
        flash("⚠️ Could not save this entry. Nothing was intentionally deleted. Please check the fields and try again. Error: "+str(error),"error")
        return redirect(request.referrer or url_for('dashboard'))
    return ("Internal Server Error: "+str(error),500)

@app.route("/projects/<int:pid>/daily",methods=["GET","POST"])
@login_required
def daily(pid):
    u=current_user()
    if not allowed_project(pid) or (u["role"]!="SUPER_ADMIN" and u["position"]!="Office Engineer"):
        flash("🚫 Only the assigned Office Engineer can prepare the Daily Report. Super Admin has override access.","error")
        return redirect(url_for("project",pid=pid))
    c=db(); default_date=request.args.get("date",dt.date.today().isoformat())
    try:
        if request.method=="POST":
            d=request.form.get("date") or default_date; section=request.form.get("section")
            if section=="boq":
                qty=parse_float(request.form.get("quantity")); bid=int(request.form["boq_id"]); msg=variation_check(c,pid,bid,qty,d)
                c.execute("INSERT INTO daily_work(project_id,date,boq_id,quantity,station_from,station_to,notes,user_id) VALUES(?,?,?,?,?,?,?,?)",(pid,d,bid,qty,request.form.get("station_from",""),request.form.get("station_to",""),request.form.get("notes",""),u["id"]))
                flash("📐 BOQ work registered — income = quantity × BOQ rate.","success")
                if msg: flash("🚨 "+msg,"error")
            elif section=="activity":
                bid=request.form.get("boq_id") or None; qty=parse_float(request.form.get("executed_qty"))
                if bid and qty>0:
                    msg=variation_check(c,pid,int(bid),qty,d)
                    c.execute("INSERT INTO daily_work(project_id,date,boq_id,quantity,station_from,station_to,notes,user_id) VALUES(?,?,?,?,?,?,?,?)",(pid,d,int(bid),qty,request.form.get("station_from",""),request.form.get("station_to",""),request.form.get("remarks",""),u["id"]))
                    if msg: flash("🚨 "+msg,"error")
                c.execute("INSERT INTO daily_activities(project_id,date,boq_id,work_type,executed_qty,machine_id,machine_hours,manpower_position,manpower_qty,manpower_hours,material_id,material_qty,remarks,user_id) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(pid,d,bid,request.form.get("work_type",""),qty,None,0,"",0,0,None,0,request.form.get("remarks",""),u["id"]))
                aid=c.execute("SELECT last_insert_rowid()").fetchone()[0]
                for lid in request.form.getlist("machine_log_ids"):
                    if lid: c.execute("INSERT INTO activity_machines(activity_id,machine_log_id,machine_id,hours) SELECT ?,id,machine_id,work_hours+idle_hours+down_hours FROM machine_logs WHERE id=? AND project_id=?",(aid,int(lid),pid))
                for mid in request.form.getlist("manpower_ids"):
                    if mid: c.execute("INSERT INTO activity_manpower(activity_id,manpower_id,crew_id,qty,hours) SELECT ?,id,crew_id,present,working_hours FROM manpower WHERE id=? AND project_id=?",(aid,int(mid),pid))
                for sid in request.form.getlist("store_log_ids"):
                    if sid: c.execute("INSERT INTO activity_store(activity_id,store_log_id,material_id,qty) SELECT ?,id,material_id,issued FROM store_logs WHERE id=? AND project_id=?",(aid,int(sid),pid))
                for fid in request.form.getlist("fuel_log_ids"):
                    if fid: c.execute("INSERT INTO activity_fuel(activity_id,fuel_log_id,litres) SELECT ?,id,opening_gauge+fuel_received-closing_gauge FROM fuel_logs WHERE id=? AND project_id=?",(aid,int(fid),pid))
                # Correct activity_fuel machine-independent insert after the insert above if schema mismatch is encountered.
                c.execute("DELETE FROM activity_fuel WHERE activity_id=? AND fuel_log_id NOT IN (SELECT id FROM fuel_logs WHERE project_id=?)",(aid,pid))
                for xid in request.form.getlist("finance_log_ids"):
                    if xid: c.execute("INSERT INTO activity_finance(activity_id,finance_log_id,amount) SELECT ?,id,amount FROM finance_logs WHERE id=? AND project_id=?",(aid,int(xid),pid))
                for crewid in request.form.getlist("crew_ids"):
                    if crewid: c.execute("INSERT INTO crew_evaluations(activity_id,crew_id,evaluation,remarks,score) VALUES(?,?,?,?,?)",(aid,int(crewid),request.form.get("evaluation",""),request.form.get("evaluation_remarks",""),parse_float(request.form.get("score"))))
                flash("🏗️ Daily work package saved with all selected machinery, manpower, fuel, store, finance and crew links.","success")
            elif section=="problem":
                c.execute("INSERT INTO problems(project_id,date,problem,remark,user_id) VALUES(?,?,?,?,?)",(pid,d,request.form.get("problem",""),request.form.get("remark",""),u["id"])); flash("⚠️ Problem and corrective action saved.","success")
            c.commit()
            try: save_report(pid,'DAILY',dt.date.fromisoformat(d),dt.date.fromisoformat(d),'ALL',u['id'])
            except Exception as e: app.logger.warning("Daily snapshot failed: %s",e)
            default_date=d
    except Exception as e:
        c.rollback(); flash("Daily report save failed: "+str(e),"error")
    boq=c.execute("SELECT * FROM boq WHERE project_id=? ORDER BY series,item_no",(pid,)).fetchall()
    machines=c.execute("SELECT * FROM machines WHERE project_id=? AND active=1 ORDER BY machine_type,code",(pid,)).fetchall()
    materials=c.execute("SELECT * FROM materials WHERE project_id=? AND active=1 ORDER BY category,name",(pid,)).fetchall()
    crews=c.execute("SELECT * FROM project_crews WHERE project_id=? ORDER BY group_name,position,name",(pid,)).fetchall()
    linked_machines=c.execute("SELECT ml.*,m.machine_type,m.code,m.plate_no,m.engine_no,m.ownership FROM machine_logs ml JOIN machines m ON m.id=ml.machine_id WHERE ml.project_id=? AND ml.date=? ORDER BY ml.id",(pid,default_date)).fetchall()
    linked_manpower=c.execute("SELECT mp.*,pc.group_name,pc.position AS crew_position,pc.name AS crew_name FROM manpower mp LEFT JOIN project_crews pc ON pc.id=mp.crew_id WHERE mp.project_id=? AND mp.date=? ORDER BY mp.id",(pid,default_date)).fetchall()
    linked_fuel=c.execute("SELECT f.*,m.machine_type,m.code,m.plate_no,m.engine_no FROM fuel_logs f JOIN machines m ON m.id=f.machine_id WHERE f.project_id=? AND f.date=? ORDER BY f.id",(pid,default_date)).fetchall()
    linked_store=c.execute("SELECT sl.*,m.name,m.unit FROM store_logs sl JOIN materials m ON m.id=sl.material_id WHERE sl.project_id=? AND sl.date=? ORDER BY sl.id",(pid,default_date)).fetchall()
    linked_finance=c.execute("SELECT * FROM finance_logs WHERE project_id=? AND date=? ORDER BY id",(pid,default_date)).fetchall()
    recent=c.execute("SELECT dw.*,b.item_no,b.description,b.unit,b.rate,dw.quantity*b.rate amount FROM daily_work dw JOIN boq b ON b.id=dw.boq_id WHERE dw.project_id=? ORDER BY dw.date DESC,dw.id DESC LIMIT 30",(pid,)).fetchall()
    activities=c.execute("SELECT a.*,b.item_no,b.description FROM daily_activities a LEFT JOIN boq b ON b.id=a.boq_id WHERE a.project_id=? AND a.date=? ORDER BY a.id DESC",(pid,default_date)).fetchall()
    alerts=c.execute("SELECT va.*,b.item_no,b.description FROM variation_alerts va JOIN boq b ON b.id=va.boq_id WHERE va.project_id=? ORDER BY va.id DESC LIMIT 20",(pid,)).fetchall()
    c.close()
    return render_template("daily.html",pid=pid,date=default_date,boq=boq,machines=machines,materials=materials,crews=crews,recent=recent,activities=activities,alerts=alerts,linked_machines=linked_machines,linked_manpower=linked_manpower,linked_fuel=linked_fuel,linked_store=linked_store,linked_finance=linked_finance)

@app.route("/projects/<int:pid>/machinery/assign",methods=['POST'])
@login_required
def assign_machine(pid):
    if not allowed_project(pid) or not can_module('Machinery'): return redirect(url_for('project',pid=pid))
    c=db()
    try:
        mid=int(request.form['machine_id']); m=c.execute('SELECT * FROM machines WHERE id=? AND project_id=?',(mid,pid)).fetchone()
        if not m: raise ValueError('Machine not found.')
        # Close any old active assignment before creating a new signed assignment.
        c.execute("UPDATE machine_assignments SET status='ENDED',end_date=COALESCE(end_date,start_date),ended_by=?,ended_at=CURRENT_TIMESTAMP WHERE machine_id=? AND project_id=? AND status='ACTIVE'",(current_user()['id'],mid,pid))
        total=parse_float(request.form.get('total_hours'))
        if total<=0: raise ValueError('Total signed hours must be greater than zero.')
        start_date=request.form['start_date']; start_hour=parse_float(request.form.get('start_hour'))
        end_meter=start_hour+total
        c.execute("UPDATE machines SET lifecycle_status='ACTIVE',assignment_start_date=?,assignment_start_hour=?,assignment_end_date=NULL,assignment_end_hour=?,total_signed_hours=?,hours_used=0,assignment_signed_by=?,assignment_ended_by=NULL,assignment_ended_at=NULL WHERE id=?",(start_date,start_hour,end_meter,total,current_user()['id'],mid))
        c.execute("INSERT INTO machine_assignments(machine_id,project_id,start_date,start_hour,end_date,end_hour,status,assigned_by,notes) VALUES(?,?,?,?,?,?,?,?,?)",(mid,pid,start_date,start_hour,None,end_meter,'ACTIVE',current_user()['id'],request.form.get('notes','')))
        c.commit(); flash(f'✍️ Assignment signed for {total:g} hours. Ending meter will be {end_meter:g}. The actual ending date is calculated automatically when logged hours reach the signed total.','success')
    except Exception as e:
        c.rollback(); flash('Machinery assignment failed: '+str(e),'error')
    c.close(); return redirect(url_for('machinery',pid=pid))

@app.route("/projects/<int:pid>/machinery/end",methods=['POST'])
@login_required
def end_machine_assignment(pid):
    if not allowed_project(pid) or not can_module('Machinery'): return redirect(url_for('project',pid=pid))
    c=db()
    try:
        mid=int(request.form['machine_id']); end_date=request.form.get('end_date') or dt.date.today().isoformat(); end_hour=parse_float(request.form.get('end_hour'))
        c.execute("UPDATE machines SET lifecycle_status='ENDED',assignment_end_date=?,assignment_end_hour=?,assignment_ended_by=?,assignment_ended_at=? WHERE id=? AND project_id=?",(end_date,end_hour,current_user()['id'],dt.datetime.now().isoformat(timespec='seconds'),mid,pid))
        c.execute("UPDATE machine_assignments SET end_date=?,end_hour=?,status='ENDED',ended_by=?,ended_at=? WHERE machine_id=? AND project_id=? AND status='ACTIVE'",(end_date,end_hour,current_user()['id'],dt.datetime.now().isoformat(timespec='seconds'),mid,pid))
        c.commit(); flash('🛑 Machine assignment ended. A new signed assignment is required before reuse.','success')
    except Exception as e: c.rollback(); flash('Could not end assignment: '+str(e),'error')
    c.close(); return redirect(url_for('machinery',pid=pid))

@app.route("/projects/<int:pid>/fuel",methods=["GET","POST"])
@login_required
def fuel(pid):
    if not allowed_project(pid) or not can_module("Machinery"): flash("🚫 Machinery/Fuel access is not assigned.","error"); return redirect(url_for("project",pid=pid))
    c=db()
    try:
        if request.method=="POST":
            mid=int(request.form["machine_id"]); d=request.form["date"]
            if not c.execute("SELECT 1 FROM machines WHERE id=? AND project_id=? AND active=1",(mid,pid)).fetchone(): raise ValueError('Selected machine is not active in this project.')
            vals=(pid,mid,d,parse_float(request.form.get("opening_gauge")),parse_float(request.form.get("fuel_received")),parse_float(request.form.get("closing_gauge")),parse_float(request.form.get("fuel_price")),request.form.get("reference",""),request.form.get("notes",""),current_user()["id"],"Fuel Register")
            c.execute("INSERT INTO fuel_logs(project_id,machine_id,date,opening_gauge,fuel_received,closing_gauge,fuel_price,reference,notes,user_id,source) VALUES(?,?,?,?,?,?,?,?,?,?,?)",vals)
            c.commit(); flash("⛽ Fuel log saved.","success")
    except Exception as e:
        c.rollback(); flash("Fuel log failed: "+str(e),"error")
    machines=c.execute("SELECT * FROM machines WHERE project_id=? AND active=1 ORDER BY machine_type,code",(pid,)).fetchall()
    logs=c.execute("SELECT f.*,m.machine_type,m.code,m.plate_no,m.engine_no,m.ownership,(f.opening_gauge+f.fuel_received-f.closing_gauge) consumption,(f.fuel_received*f.fuel_price) cost,COALESCE((SELECT SUM(ml.work_hours) FROM machine_logs ml WHERE ml.machine_id=f.machine_id AND ml.date=f.date),0) work_hours,COALESCE((SELECT SUM(ml.work_hours) FROM machine_logs ml WHERE ml.machine_id=f.machine_id AND ml.date=f.date),0)*m.expected_fuel expected_consumption FROM fuel_logs f JOIN machines m ON m.id=f.machine_id WHERE f.project_id=? ORDER BY f.date DESC,f.id DESC LIMIT 100",(pid,)).fetchall()
    total_l=c.execute("SELECT COALESCE(SUM(fuel_received),0) FROM fuel_logs WHERE project_id=?",(pid,)).fetchone()[0]; total_cost=c.execute("SELECT COALESCE(SUM(fuel_received*fuel_price),0) FROM fuel_logs WHERE project_id=?",(pid,)).fetchone()[0]
    c.close(); return render_template("fuel.html",pid=pid,machines=machines,logs=logs,total_l=total_l,total_cost=total_cost,today=dt.date.today().isoformat())

@app.route("/projects/<int:pid>/performance",methods=["GET","POST"])
@login_required
def performance(pid):
    if not allowed_project(pid): return redirect(url_for("dashboard"))
    c=db()
    if request.method=="POST":
        c.execute("INSERT INTO performance_rates(project_id,work_type,worker_type,unit,qty_per_hour,notes) VALUES(?,?,?,?,?,?) ON CONFLICT(project_id,work_type,worker_type) DO UPDATE SET unit=excluded.unit,qty_per_hour=excluded.qty_per_hour,notes=excluded.notes",(pid,request.form["work_type"],request.form["worker_type"],request.form["unit"],parse_float(request.form["qty_per_hour"]),request.form.get("notes",""))); c.commit(); flash("📈 Performance rate saved.","success")
    rows=c.execute("SELECT * FROM performance_rates WHERE project_id=? ORDER BY work_type,worker_type",(pid,)).fetchall(); c.close(); return render_template("performance.html",pid=pid,rows=rows)

@app.route("/projects/<int:pid>/print-report")
@login_required
def print_report(pid):
    if not allowed_project(pid): return redirect(url_for("dashboard"))
    date=request.args.get("date",dt.date.today().isoformat()); c=db()
    p=c.execute("SELECT * FROM projects WHERE id=?",(pid,)).fetchone(); settings=c.execute("SELECT * FROM report_settings WHERE project_id=?",(pid,)).fetchone()
    boq=c.execute("SELECT b.item_no,b.description,b.unit,b.rate,COALESCE(SUM(dw.quantity),0) qty FROM boq b LEFT JOIN daily_work dw ON dw.boq_id=b.id AND dw.date=? WHERE b.project_id=? GROUP BY b.id ORDER BY b.item_no",(date,pid)).fetchall()
    activities=c.execute("SELECT a.*,b.item_no,b.description FROM daily_activities a LEFT JOIN boq b ON b.id=a.boq_id WHERE a.project_id=? AND a.date=? ORDER BY a.id",(pid,date)).fetchall()
    activity_links={}
    for a in activities:
        aid=a['id']
        activity_links[aid]={
            'machines':c.execute("SELECT am.*,m.machine_type,m.code,m.plate_no FROM activity_machines am JOIN machines m ON m.id=am.machine_id WHERE am.activity_id=?",(aid,)).fetchall(),
            'manpower':c.execute("SELECT ap.*,mp.name,mp.position,pc.group_name,pc.name crew_name FROM activity_manpower ap JOIN manpower mp ON mp.id=ap.manpower_id LEFT JOIN project_crews pc ON pc.id=ap.crew_id WHERE ap.activity_id=?",(aid,)).fetchall(),
            'store':c.execute("SELECT ast.*,m.name,m.unit FROM activity_store ast JOIN materials m ON m.id=ast.material_id WHERE ast.activity_id=?",(aid,)).fetchall(),
            'fuel':c.execute("SELECT af.*,f.opening_gauge,f.fuel_received,f.closing_gauge,m.machine_type,m.code FROM activity_fuel af JOIN fuel_logs f ON f.id=af.fuel_log_id JOIN machines m ON m.id=f.machine_id WHERE af.activity_id=?",(aid,)).fetchall(),
            'finance':c.execute("SELECT ax.*,f.kind,f.description FROM activity_finance ax JOIN finance_logs f ON f.id=ax.finance_log_id WHERE ax.activity_id=?",(aid,)).fetchall(),
            'evaluations':c.execute("SELECT ce.*,pc.name,pc.position,pc.group_name FROM crew_evaluations ce JOIN project_crews pc ON pc.id=ce.crew_id WHERE ce.activity_id=?",(aid,)).fetchall()
        }
    machines=c.execute("SELECT ml.*,m.machine_type,m.code,m.plate_no,m.ownership FROM machine_logs ml JOIN machines m ON m.id=ml.machine_id WHERE ml.project_id=? AND ml.date=?",(pid,date)).fetchall()
    manpower=c.execute("SELECT * FROM manpower WHERE project_id=? AND date=?",(pid,date)).fetchall(); store=c.execute("SELECT sl.*,m.name,m.unit FROM store_logs sl JOIN materials m ON m.id=sl.material_id WHERE sl.project_id=? AND sl.date=?",(pid,date)).fetchall(); problems=c.execute("SELECT * FROM problems WHERE project_id=? AND date=?",(pid,date)).fetchall(); c.close()
    return render_template("print_report.html",p=p,settings=settings,date=date,boq=boq,activities=activities,activity_links=activity_links,machines=machines,manpower=manpower,store=store,problems=problems)

@app.route("/projects/<int:pid>/machinery",methods=["GET","POST"])
@login_required
def machinery(pid):
    if not allowed_project(pid) or not can_module("Machinery"):flash("🚫 Machinery access is not assigned.","error");return redirect(url_for("project",pid=pid))
    c=db()
    if request.method=="POST":
        action=request.form.get("action")
        if action=="add":
            c.execute("INSERT INTO machines(project_id,machine_type,code,plate_no,engine_no,ownership,hourly_rate,expected_fuel,fuel_price,lifecycle_status) VALUES(?,?,?,?,?,?,?,?,?,?)",(pid,request.form["machine_type"],request.form["code"],request.form.get("plate_no",request.form["code"]),request.form.get("engine_no",""),request.form["ownership"],parse_float(request.form["hourly_rate"]),parse_float(request.form["expected_fuel"]),parse_float(request.form.get("fuel_price")),"UNASSIGNED"))
            flash("🚜 Machine added to this project's fleet.","success")
        elif action=="remove":c.execute("UPDATE machines SET active=0 WHERE id=? AND project_id=?",(request.form["machine_id"],pid));flash("Machine removed from active fleet.","success")
        elif action=="log":
            active_assignment=c.execute("SELECT id FROM machine_assignments WHERE machine_id=? AND project_id=? AND status='ACTIVE' ORDER BY id DESC LIMIT 1",(request.form['machine_id'],pid)).fetchone()
            if not active_assignment: raise ValueError('Machine has no active signed assignment. Machinery Admin must sign a new start date/hour first.')
            vals=(pid,request.form["machine_id"],request.form["date"],parse_float(request.form.get("work_hours")),parse_float(request.form.get("idle_hours")),request.form.get("idle_reason",""),1 if request.form.get("idle_payable")=="1" else 0,parse_float(request.form.get("down_hours")),request.form.get("down_reason",""),parse_float(request.form.get("opening_gauge")),parse_float(request.form.get("fuel_received")),parse_float(request.form.get("closing_gauge")),request.form.get("notes",""),current_user()["id"])
            c.execute("INSERT INTO machine_logs(project_id,machine_id,date,work_hours,idle_hours,idle_reason,idle_payable,down_hours,down_reason,opening_gauge,fuel_received,closing_gauge,notes,user_id) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",vals)
            if any(parse_float(request.form.get(k)) for k in ("opening_gauge","fuel_received","closing_gauge")):
                mp=c.execute("SELECT fuel_price FROM machines WHERE id=?",(request.form['machine_id'],)).fetchone()
                c.execute("INSERT INTO fuel_logs(project_id,machine_id,date,opening_gauge,fuel_received,closing_gauge,fuel_price,reference,notes,user_id,source) VALUES(?,?,?,?,?,?,?,?,?,?,?)",(pid,request.form['machine_id'],request.form['date'],parse_float(request.form.get('opening_gauge')),parse_float(request.form.get('fuel_received')),parse_float(request.form.get('closing_gauge')),mp['fuel_price'] if mp else 0,'MCH-'+request.form['date'],request.form.get('notes',''),current_user()['id'],'Machinery Log'))
            used=parse_float(request.form.get("work_hours"))+parse_float(request.form.get("idle_hours"))+parse_float(request.form.get("down_hours"))
            ma=c.execute("SELECT id,total_signed_hours,hours_used FROM machine_assignments WHERE machine_id=? AND project_id=? AND status='ACTIVE' ORDER BY id DESC LIMIT 1",(request.form['machine_id'],pid)).fetchone()
            if ma:
                new_used=(ma['hours_used'] or 0)+used; c.execute("UPDATE machine_assignments SET status=CASE WHEN total_signed_hours>0 AND ?>=total_signed_hours THEN 'ENDED' ELSE status END,end_date=CASE WHEN total_signed_hours>0 AND ?>=total_signed_hours THEN ? ELSE end_date END,end_hour=CASE WHEN total_signed_hours>0 AND ?>=total_signed_hours THEN ? ELSE end_hour END WHERE id=?",(new_used,new_used,request.form['date'],parse_float(request.form.get('closing_gauge')),ma['id']))
                c.execute("UPDATE machines SET hours_used=?,lifecycle_status=CASE WHEN total_signed_hours>0 AND ?>=total_signed_hours THEN 'ENDED' ELSE lifecycle_status END,assignment_end_date=CASE WHEN total_signed_hours>0 AND ?>=total_signed_hours THEN ? ELSE assignment_end_date END WHERE id=?",(new_used,new_used,new_used,request.form['date'],request.form['machine_id']))
            flash("⏱️ Machine hours / idle / down / gauge saved successfully.","success")
        try: c.commit()
        except Exception as e: c.rollback(); flash("Machinery save failed: "+str(e),"error")
    machines=c.execute("SELECT * FROM machines WHERE project_id=? AND active=1 ORDER BY machine_type,code",(pid,)).fetchall();assignments=c.execute("SELECT ma.*,m.machine_type,m.code,m.plate_no FROM machine_assignments ma JOIN machines m ON m.id=ma.machine_id WHERE ma.project_id=? ORDER BY ma.id DESC LIMIT 100",(pid,)).fetchall();logs=c.execute("SELECT ml.*,m.machine_type,m.code,m.ownership,m.hourly_rate,m.expected_fuel,((ml.work_hours + CASE WHEN ml.idle_payable=1 THEN ml.idle_hours ELSE 0 END)*m.hourly_rate) expense,(ml.opening_gauge+ml.fuel_received-ml.closing_gauge) actual_fuel,(ml.work_hours*m.expected_fuel) expected_fuel_qty,CASE WHEN (ml.work_hours+ml.idle_hours+ml.down_hours)>0 THEN ml.work_hours*100.0/(ml.work_hours+ml.idle_hours+ml.down_hours) ELSE 0 END utilization,CASE WHEN (ml.work_hours+ml.idle_hours+ml.down_hours)>0 THEN (ml.work_hours+ml.idle_hours)*100.0/(ml.work_hours+ml.idle_hours+ml.down_hours) ELSE 0 END availability,(ml.opening_gauge+ml.fuel_received-ml.closing_gauge)-(ml.work_hours*m.expected_fuel) fuel_discrepancy FROM machine_logs ml JOIN machines m ON m.id=ml.machine_id WHERE ml.project_id=? ORDER BY ml.date DESC,ml.id DESC LIMIT 50",(pid,)).fetchall();c.close()
    return render_template("machinery.html",pid=pid,machines=machines,assignments=assignments,logs=logs)

@app.route("/projects/<int:pid>/manpower",methods=["GET","POST"])
@login_required
def manpower(pid):
    if not allowed_project(pid) or not can_module("HR"): flash("🚫 HR/manpower access is not assigned.","error"); return redirect(url_for("project",pid=pid))
    c=db()
    try:
        if request.method=="POST":
            vals=(pid,request.form["date"],request.form["name"],request.form["employment"],request.form["position"],request.form.get("crew_id") or None,parse_float(request.form.get("present")),parse_float(request.form.get("working_hours",8)),parse_float(request.form.get("hourly_rate")),parse_float(request.form.get("daily_rate")),parse_float(request.form.get("normal_ot_hours")),parse_float(request.form.get("normal_ot_rate")),parse_float(request.form.get("night_ot_hours")),parse_float(request.form.get("night_ot_rate")),parse_float(request.form.get("sunday_ot_hours")),parse_float(request.form.get("sunday_ot_rate")),parse_float(request.form.get("holiday_ot_hours")),parse_float(request.form.get("holiday_ot_rate")),0,0,request.form.get("notes",""),current_user()["id"])
            c.execute("INSERT INTO manpower(project_id,date,name,employment,position,crew_id,present,working_hours,hourly_rate,daily_rate,normal_ot_hours,normal_ot_rate,night_ot_hours,night_ot_rate,sunday_ot_hours,sunday_ot_rate,holiday_ot_hours,holiday_ot_rate,overtime_hours,overtime_rate,notes,user_id) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",vals); c.commit(); flash("👷 Manpower record saved and linked to the selected crew.","success")
    except Exception as e: c.rollback(); flash("Manpower save failed: "+str(e),"error")
    rows=c.execute("SELECT mp.*,pc.group_name,pc.position crew_position,pc.name crew_name FROM manpower mp LEFT JOIN project_crews pc ON pc.id=mp.crew_id WHERE mp.project_id=? ORDER BY mp.date DESC,mp.id DESC LIMIT 100",(pid,)).fetchall(); crews=c.execute("SELECT * FROM project_crews WHERE project_id=? ORDER BY group_name,position,name",(pid,)).fetchall(); c.close(); return render_template("manpower.html",pid=pid,rows=rows,crews=crews)

@app.route("/projects/<int:pid>/store",methods=["GET","POST"])
@login_required
def store(pid):
    if not allowed_project(pid) or not can_module("Store"):flash("🚫 Store access is not assigned.","error");return redirect(url_for("project",pid=pid))
    c=db()
    if request.method=="POST":
        if request.form.get("action")=="add":c.execute("INSERT INTO materials(project_id,category,name,unit,min_stock) VALUES(?,?,?,?,?)",(pid,request.form["category"],request.form["name"],request.form["unit"],parse_float(request.form["min_stock"])))
        else:
            physical=request.form.get("physical_balance");physical=parse_float(physical) if physical not in (None,"") else None
            c.execute("INSERT INTO store_logs(project_id,material_id,date,received,issued,unit_cost,physical_balance,reference,notes,user_id) VALUES(?,?,?,?,?,?,?,?,?,?)",(pid,request.form["material_id"],request.form["date"],parse_float(request.form["received"]),parse_float(request.form["issued"]),parse_float(request.form["unit_cost"]),physical,request.form.get("reference",""),request.form.get("notes",""),current_user()["id"]))
        c.commit();flash("📦 Store record saved.","success")
    mats=c.execute("SELECT m.*,COALESCE(SUM(sl.received),0) rec,COALESCE(SUM(sl.issued),0) iss,COALESCE(SUM(sl.received)-SUM(sl.issued),0) balance,(SELECT sl2.physical_balance FROM store_logs sl2 WHERE sl2.material_id=m.id AND sl2.physical_balance IS NOT NULL ORDER BY sl2.date DESC,sl2.id DESC LIMIT 1) physical FROM materials m LEFT JOIN store_logs sl ON sl.material_id=m.id WHERE m.project_id=? AND m.active=1 GROUP BY m.id ORDER BY m.category,m.name",(pid,)).fetchall();logs=c.execute("SELECT sl.*,m.name,m.unit FROM store_logs sl JOIN materials m ON m.id=sl.material_id WHERE sl.project_id=? ORDER BY sl.date DESC,sl.id DESC LIMIT 100",(pid,)).fetchall();c.close();return render_template("store.html",pid=pid,materials=mats,logs=logs)


@app.route("/projects/<int:pid>/crew",methods=["GET","POST"])
@login_required
def crew(pid):
    if not allowed_project(pid) or not can_module("HR"): flash("🚫 Crew access is not assigned.","error"); return redirect(url_for("project",pid=pid))
    c=db()
    try:
        if request.method=="POST":
            action=request.form.get("action","add")
            if action=="capacity":
                name=request.form.get("capacity_group","").strip()
                if not name: raise ValueError("Select or enter a crew group.")
                c.execute("INSERT OR IGNORE INTO crew_groups(name) VALUES(?)",(name,))
                c.execute("INSERT INTO crew_group_capacity(group_name,foreman_qty,dl_qty,surveyor_qty,data_collector_qty,time_keeper_qty,other_qty,total_qty) VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(group_name) DO UPDATE SET foreman_qty=excluded.foreman_qty,dl_qty=excluded.dl_qty,surveyor_qty=excluded.surveyor_qty,data_collector_qty=excluded.data_collector_qty,time_keeper_qty=excluded.time_keeper_qty,other_qty=excluded.other_qty,total_qty=excluded.total_qty",(name,parse_float(request.form.get('foreman_qty')),parse_float(request.form.get('dl_qty')),parse_float(request.form.get('surveyor_qty')),parse_float(request.form.get('data_collector_qty')),parse_float(request.form.get('time_keeper_qty')),parse_float(request.form.get('other_qty')),parse_float(request.form.get('total_qty'))))
                c.commit(); flash("👥 Crew capacity saved for future daily crew planning.","success")
            elif action=="add_group":
                name=request.form.get("new_group","").strip()
                if name: c.execute("INSERT OR IGNORE INTO crew_groups(name) VALUES(?)",(name,)); c.execute("INSERT OR IGNORE INTO crew_group_capacity(group_name) VALUES(?)",(name,)); c.commit(); flash("👥 Group added for future selection.","success")
            elif action=="add_position":
                name=request.form.get("new_position","").strip()
                if name: c.execute("INSERT OR IGNORE INTO crew_positions(name) VALUES(?)",(name,)); c.commit(); flash("👷 Position added for future selection.","success")
            else:
                group=request.form.get("group_name","").strip(); position=request.form.get("position","").strip()
                if group: c.execute("INSERT OR IGNORE INTO crew_groups(name) VALUES(?)",(group,)); c.execute("INSERT OR IGNORE INTO crew_group_capacity(group_name) VALUES(?)",(group,))
                if position: c.execute("INSERT OR IGNORE INTO crew_positions(name) VALUES(?)",(position,))
                c.execute("INSERT INTO project_crews(project_id,date,group_name,position,name,employment,skill_level,working_hours,hourly_rate,notes,user_id) VALUES(?,?,?,?,?,?,?,?,?,?,?)",(pid,request.form.get("date",dt.date.today().isoformat()),group,position,request.form["name"],request.form["employment"],request.form.get("skill_level","Skilled"),parse_float(request.form.get("working_hours")),parse_float(request.form.get("hourly_rate")),request.form.get("notes",""),current_user()["id"])); c.commit(); flash("👷 Crew member registered.","success")
    except Exception as e: c.rollback(); flash("Crew save failed: "+str(e),"error")
    groups=[r["name"] for r in c.execute("SELECT name FROM crew_groups WHERE active=1 ORDER BY name").fetchall()]; positions=[r["name"] for r in c.execute("SELECT name FROM crew_positions WHERE active=1 ORDER BY name").fetchall()]; rows=c.execute("SELECT * FROM project_crews WHERE project_id=? ORDER BY group_name,position,name",(pid,)).fetchall(); capacities=c.execute("SELECT * FROM crew_group_capacity WHERE active=1 ORDER BY group_name").fetchall(); c.close()
    return render_template("crew.html",pid=pid,rows=rows,crew_groups=groups,position_catalog=positions,capacities=capacities)

@app.route("/projects/<int:pid>/report-settings",methods=["GET","POST"])
@admin_required
def report_settings(pid):
    c=db(); s=c.execute("SELECT * FROM report_settings WHERE project_id=?",(pid,)).fetchone()
    if request.method=="POST":
        vals=(pid,request.form.get("contractor_role","Main Contractor"),request.form.get("phone",""),request.form.get("email",""),request.form.get("website",""),request.form.get("fax",""),request.form.get("address",""),request.form.get("logo_text","BAGC"))
        c.execute("INSERT INTO report_settings(project_id,contractor_role,phone,email,website,fax,address,logo_text) VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(project_id) DO UPDATE SET contractor_role=excluded.contractor_role,phone=excluded.phone,email=excluded.email,website=excluded.website,fax=excluded.fax,address=excluded.address,logo_text=excluded.logo_text",vals); c.commit(); flash("🖨️ Report header/footer saved.","success"); s=c.execute("SELECT * FROM report_settings WHERE project_id=?",(pid,)).fetchone()
    c.close(); return render_template("report_settings.html",pid=pid,s=s)

@app.route("/projects/<int:pid>/design",methods=["GET","POST"])
@login_required
def design(pid):
    if not allowed_project(pid) or not can_module("Design"):flash("🚫 Design access is not assigned.","error");return redirect(url_for("project",pid=pid))
    c=db()
    if request.method=="POST":
        c.execute("INSERT INTO design_items(project_id,drawing_no,title,discipline,revision,status,submitted,approved,comments,user_id) VALUES(?,?,?,?,?,?,?,?,?,?)",(pid,request.form["drawing_no"],request.form["title"],request.form["discipline"],request.form["revision"],request.form["status"],request.form.get("submitted",""),request.form.get("approved",""),request.form.get("comments",""),current_user()["id"]));c.commit();flash("🎨 Design record saved.","success")
    rows=c.execute("SELECT * FROM design_items WHERE project_id=? ORDER BY id DESC",(pid,)).fetchall();c.close();return render_template("design.html",pid=pid,rows=rows)

@app.route("/projects/<int:pid>/finance",methods=["GET","POST"])
@login_required
def finance(pid):
    if not allowed_project(pid) or not can_module("Finance"):flash("🚫 Finance access is not assigned.","error");return redirect(url_for("project",pid=pid))
    c=db()
    if request.method=="POST":c.execute("INSERT INTO finance_logs(project_id,date,category,kind,description,amount,reference,user_id) VALUES(?,?,?,?,?,?,?,?)",(pid,request.form["date"],request.form["category"],request.form["kind"],request.form["description"],parse_float(request.form["amount"]),request.form.get("reference",""),current_user()["id"]));c.commit();flash("💰 Finance record saved.","success")
    rows=c.execute("SELECT * FROM finance_logs WHERE project_id=? ORDER BY date DESC,id DESC LIMIT 100",(pid,)).fetchall();c.close();return render_template("finance.html",pid=pid,rows=rows)

@app.route("/projects/<int:pid>/reports")
@login_required
def reports(pid):
    if not allowed_project(pid): return redirect(url_for("dashboard"))
    report_type=request.args.get('report_type','MONTHLY').upper()
    scope=request.args.get('scope','ALL').upper()
    start_s=request.args.get('start',''); end_s=request.args.get('end','')
    try: start,end=report_dates(report_type,start_s,end_s)
    except Exception: start,end=report_dates(report_type,'','')
    c=db(); p=c.execute("SELECT * FROM projects WHERE id=?",(pid,)).fetchone()
    saved=c.execute("SELECT sr.*,u.full_name generated_name FROM saved_reports sr LEFT JOIN users u ON u.id=sr.generated_by WHERE sr.project_id=? ORDER BY sr.generated_at DESC,sr.id DESC LIMIT 50",(pid,)).fetchall(); c.close()
    snapshot=build_report_snapshot(pid,start,end,scope)
    if request.args.get('save')=='1':
        rid=save_report(pid,report_type,start,end,scope,current_user()['id']); flash(f'📚 {report_type} report saved as a permanent report record.','success'); return redirect(url_for('reports',pid=pid,report_type=report_type,scope=scope,start=start.isoformat(),end=end.isoformat()))
    return render_template('reports.html',pid=pid,p=p,report_type=report_type,scope=scope,start=start,end=end,snapshot=snapshot,saved=saved)

@app.route("/projects/<int:pid>/reports/save",methods=['POST'])
@login_required
def save_report_route(pid):
    if not allowed_project(pid): return redirect(url_for('dashboard'))
    rt=request.form.get('report_type','MONTHLY').upper(); scope=request.form.get('scope','ALL').upper()
    start=dt.date.fromisoformat(request.form['start']); end=dt.date.fromisoformat(request.form['end'])
    rid=save_report(pid,rt,start,end,scope,current_user()['id']); flash(f'📚 Report saved permanently (Record #{rid}).','success')
    return redirect(url_for('reports',pid=pid,report_type=rt,scope=scope,start=start.isoformat(),end=end.isoformat()))

@app.route("/projects/<int:pid>/reports/<int:rid>")
@login_required
def saved_report(pid,rid):
    if not allowed_project(pid): return redirect(url_for('dashboard'))
    c=db(); r=c.execute("SELECT sr.*,p.name project_name,p.client,p.consultant,p.contractor_role,u.full_name generated_name FROM saved_reports sr JOIN projects p ON p.id=sr.project_id LEFT JOIN users u ON u.id=sr.generated_by WHERE sr.id=? AND sr.project_id=?",(rid,pid)).fetchone()
    if not r: c.close(); return ('Report not found',404)
    source_ids=json.loads(r['source_report_ids'] or '[]'); sources=[]
    if source_ids:
        marks=','.join('?'*len(source_ids)); sources=c.execute(f"SELECT id,report_no,report_type,scope,start_date,end_date FROM saved_reports WHERE id IN ({marks}) ORDER BY start_date",source_ids).fetchall()
    c.close(); return render_template('saved_report.html',r=r,snapshot=json.loads(r['snapshot_json']),sources=sources)

@app.route("/projects/<int:pid>/rfi", methods=["GET","POST"])
@login_required
def rfi(pid):
    if not allowed_project(pid): flash("🚫 You do not have access to this project.","error"); return redirect(url_for("dashboard"))
    c=db(); project=c.execute("SELECT * FROM projects WHERE id=?",(pid,)).fetchone(); boqs=c.execute("SELECT * FROM boq WHERE project_id=? ORDER BY item_no",(pid,)).fetchall()
    users=c.execute("SELECT id,full_name,username,department,position,role FROM users WHERE active=1 AND (role='SUPER_ADMIN' OR id IN (SELECT user_id FROM user_projects WHERE project_id=?)) ORDER BY full_name",(pid,)).fetchall()
    if request.method=='POST':
        f=request.form
        roles=[('Site Engineer',f.get('site_engineer')),('Office Engineer',f.get('office_engineer')),('Project Manager',f.get('project_manager'))]
        if any(not uid for _,uid in roles):
            flash('👷 Select Site Engineer → Office Engineer → Project Manager in sequence.','error'); c.close(); return render_template('rfi.html',p=project,boqs=boqs,users=users,rfis=[],selected_inspectors=[])
        count=c.execute("SELECT COUNT(*) n FROM rfis WHERE project_id=?",(pid,)).fetchone()['n']+1; rfi_no=f"RFI-{dt.date.today().year}-{count:03d}"
        c.execute("INSERT INTO rfis(project_id,rfi_no,date_requested,inspection_date,location,boq_id,work_description,drawing_no,drawing_revision,specification,work_stage,submitted_by,status,overall_comment,corrective_action) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(pid,rfi_no,f.get('date_requested') or dt.date.today().isoformat(),f.get('inspection_date') or None,f.get('location',''),f.get('boq_id') or None,f.get('work_description',''),f.get('drawing_no',''),f.get('drawing_revision',''),f.get('specification',''),f.get('work_stage',''),session['user_id'],'PENDING SITE ENGINEER','', ''))
        rid=c.execute('SELECT last_insert_rowid()').fetchone()[0]
        for order,(stage,uid) in enumerate(roles,1):
            c.execute("INSERT INTO rfi_steps(rfi_id,step_order,stage,assigned_user_id) VALUES(?,?,?,?)",(rid,order,stage,int(uid)))
        # retain legacy participant table for compatibility/printing
        for stage,uid in roles: c.execute("INSERT OR IGNORE INTO rfi_inspections(rfi_id,inspector_user_id,inspector_role) VALUES(?,?,?)",(rid,int(uid),stage))
        c.commit(); c.close(); flash(f'📋 {rfi_no} submitted. Workflow: Site Engineer → Office Engineer → Project Manager.','success'); return redirect(url_for('rfi',pid=pid))
    rfis=c.execute("SELECT r.*,u.full_name AS submitter,(SELECT COUNT(*) FROM rfi_steps s WHERE s.rfi_id=r.id) step_count,(SELECT COUNT(*) FROM rfi_steps s WHERE s.rfi_id=r.id AND s.decision IN ('APPROVED','APPROVED WITH COMMENTS')) approved_count FROM rfis r LEFT JOIN users u ON u.id=r.submitted_by WHERE r.project_id=? ORDER BY r.id DESC",(pid,)).fetchall(); c.close()
    return render_template('rfi.html',p=project,boqs=boqs,users=users,rfis=rfis,selected_inspectors=[])

@app.route("/projects/<int:pid>/rfi/<int:rid>", methods=["GET","POST"])
@login_required
def rfi_detail(pid,rid):
    if not allowed_project(pid): return redirect(url_for('dashboard'))
    c=db(); r=c.execute("SELECT r.*,p.name project_name,p.client,p.consultant,u.full_name submitter FROM rfis r JOIN projects p ON p.id=r.project_id LEFT JOIN users u ON u.id=r.submitted_by WHERE r.id=? AND r.project_id=?",(rid,pid)).fetchone()
    if not r: c.close(); return ('RFI not found',404)
    if request.method=='POST':
        sid=int(request.form.get('step_id','0')); decision=request.form.get('decision','PENDING'); comment=request.form.get('comments',''); me=c.execute('SELECT * FROM users WHERE id=?',(session.get('user_id'),)).fetchone()
        step=c.execute('SELECT * FROM rfi_steps WHERE id=? AND rfi_id=?',(sid,rid)).fetchone()
        if not step or (me['role']!='SUPER_ADMIN' and step['assigned_user_id']!=me['id']): flash('🚫 This RFI step is not assigned to your account.','error')
        else:
            previous=c.execute("SELECT COUNT(*) n FROM rfi_steps WHERE rfi_id=? AND step_order<? AND decision NOT IN ('APPROVED','APPROVED WITH COMMENTS')",(rid,step['step_order'])).fetchone()['n']
            if previous and me['role']!='SUPER_ADMIN': flash('⏳ The previous approval stage must be completed first.','error')
            else:
                c.execute("UPDATE rfi_steps SET decision=?,comments=?,inspection_date=?,signed_at=CURRENT_TIMESTAMP WHERE id=?",(decision,comment,request.form.get('inspection_date') or dt.date.today().isoformat(),sid))
                c.execute("UPDATE rfi_inspections SET decision=?,comments=?,inspection_date=?,signed_at=CURRENT_TIMESTAMP WHERE rfi_id=? AND inspector_user_id=?",(decision,comment,request.form.get('inspection_date') or dt.date.today().isoformat(),rid,step['assigned_user_id']))
                if decision=='REJECTED': status='REJECTED'
                elif decision=='RESUBMISSION REQUIRED': status='RESUBMISSION REQUIRED'
                else:
                    remaining=c.execute("SELECT COUNT(*) n FROM rfi_steps WHERE rfi_id=? AND decision NOT IN ('APPROVED','APPROVED WITH COMMENTS')",(rid,)).fetchone()['n']
                    if remaining==0: status='APPROVED'
                    else:
                        nxt=c.execute("SELECT stage FROM rfi_steps WHERE rfi_id=? AND decision NOT IN ('APPROVED','APPROVED WITH COMMENTS') ORDER BY step_order LIMIT 1",(rid,)).fetchone(); status=f"PENDING {nxt['stage'].upper()}" if nxt else 'PENDING INSPECTION'
                c.execute('UPDATE rfis SET status=? WHERE id=?',(status,rid)); c.commit(); flash('✅ RFI approval stage signed and workflow advanced.','success')
    r=c.execute("SELECT r.*,p.name project_name,p.client,p.consultant,u.full_name submitter FROM rfis r JOIN projects p ON p.id=r.project_id LEFT JOIN users u ON u.id=r.submitted_by WHERE r.id=?",(rid,)).fetchone(); steps=c.execute("SELECT s.*,u.full_name,u.username,u.department FROM rfi_steps s LEFT JOIN users u ON u.id=s.assigned_user_id WHERE s.rfi_id=? ORDER BY s.step_order",(rid,)).fetchall(); c.close(); return render_template('rfi_detail.html',r=r,inspections=steps)

@app.route("/projects/<int:pid>/rfi/<int:rid>/print")
@login_required
def rfi_print(pid,rid):
    if not allowed_project(pid): return redirect(url_for("dashboard"))
    c=db(); r=c.execute("SELECT r.*,p.name project_name,p.client,p.consultant,p.contractor_role,u.full_name submitter FROM rfis r JOIN projects p ON p.id=r.project_id LEFT JOIN users u ON u.id=r.submitted_by WHERE r.id=? AND r.project_id=?",(rid,pid)).fetchone(); inspections=c.execute("SELECT s.*,u.full_name,u.department,u.username FROM rfi_steps s LEFT JOIN users u ON u.id=s.assigned_user_id WHERE s.rfi_id=? ORDER BY s.step_order",(rid,)).fetchall(); c.close()
    if not r: return ("RFI not found",404)
    return render_template("rfi_print.html",r=r,inspections=inspections)

@app.route("/uploads/user_photos/<path:filename>")
def user_photo(filename):
    return send_from_directory(USER_PHOTOS, filename)

@app.route("/admin/users")
@admin_required
def users():
    c=db();users=c.execute("SELECT * FROM users ORDER BY full_name").fetchall();projects=c.execute("SELECT * FROM projects ORDER BY name").fetchall();assign={u["id"]:[r["project_id"] for r in c.execute("SELECT project_id FROM user_projects WHERE user_id=?",(u["id"],)).fetchall()] for u in users};c.close();return render_template("users.html",users=users,projects=projects,assign=assign)

@app.route("/admin/users/add",methods=["POST"])
@admin_required
def add_user():
    c=db()
    try:
        photo=request.files.get("photo")
        if not photo or not photo.filename:
            raise ValueError("Staff photo is required. Upload a passport-style JPG, PNG or WEBP photo.")
        ext=secure_filename(photo.filename).rsplit('.',1)[-1].lower() if '.' in photo.filename else ''
        if ext not in ALLOWED_PHOTO_EXT: raise ValueError("Photo must be JPG, JPEG, PNG or WEBP.")
        role="SUPER_ADMIN" if request.form.get("role")=="SUPER_ADMIN" else "STAFF"
        c.execute("INSERT INTO users(full_name,username,password_hash,department,position,location,role,photo_filename) VALUES(?,?,?,?,?,?,?,?)",(request.form["full_name"].strip(),request.form["username"].strip(),generate_password_hash(request.form["password"]),request.form["department"],request.form.get("position","Other"),request.form.get("location","").strip(),role,None))
        uid=c.execute("SELECT id FROM users WHERE username=?",(request.form["username"].strip(),)).fetchone()["id"]
        staff_id=make_staff_id(request.form["department"],uid)
        filename=f"{staff_id}_{uid}.{ext}"
        photo.save(os.path.join(USER_PHOTOS,filename))
        c.execute("UPDATE users SET staff_id=?,photo_filename=? WHERE id=?",(staff_id,filename,uid))
        for pid in request.form.getlist("project_ids"): c.execute("INSERT OR IGNORE INTO user_projects(user_id,project_id) VALUES(?,?)",(uid,pid))
        c.commit(); flash(f"👤 Staff registered permanently: {staff_id}. The account can be disabled later, but it is never deleted.","success")
    except Exception as e:
        c.rollback(); flash("Could not create user: "+str(e),"error")
    c.close(); return redirect(url_for("users"))

@app.route("/admin/users/<int:uid>/photo",methods=["POST"])
@admin_required
def update_user_photo(uid):
    photo=request.files.get("photo")
    c=db(); u=c.execute("SELECT * FROM users WHERE id=?",(uid,)).fetchone()
    if not u: c.close(); flash("User not found.","error"); return redirect(url_for("users"))
    try:
        if not photo or not photo.filename: raise ValueError("Select a photo.")
        ext=secure_filename(photo.filename).rsplit('.',1)[-1].lower() if '.' in photo.filename else ''
        if ext not in ALLOWED_PHOTO_EXT: raise ValueError("Photo must be JPG, JPEG, PNG or WEBP.")
        if u['photo_filename']:
            old=os.path.join(USER_PHOTOS,u['photo_filename'])
            if os.path.isfile(old): os.remove(old)
        filename=f"{u['staff_id']}_{uid}.{ext}"
        photo.save(os.path.join(USER_PHOTOS,filename)); c.execute("UPDATE users SET photo_filename=? WHERE id=?",(filename,uid)); c.commit(); flash("📷 Staff photo updated.","success")
    except Exception as e: c.rollback(); flash("Photo update failed: "+str(e),"error")
    c.close(); return redirect(url_for("users"))


CODE128_B_PATTERNS=["212222","222122","222221","121223","121322","131222","122213","122312","132212","221213","221312","231212","112232","122132","122231","113222","123122","123221","223211","221132","221231","213212","223112","312131","311222","321122","321221","312212","322112","322211","212123","212321","232121","111323","131123","131321","112313","132113","132311","211313","231113","231311","112133","112331","132131","113123","113321","133121","313121","211331","231131","213113","213311","213131","311123","311321","331121","312113","312311","332111","314111","221411","431111","111224","111422","121124","121421","141122","141221","112214","112412","122114","122411","142112","142211","241211","221114","413111","241112","134111","111242","121142","121241","114212","124112","124211","411212","421112","421211","212141","214121","412121","111143","111341","131141","114113","114311","411113","411311","113141","114131","311141","411131","211412","211214","211232","2331112"]
def code128_svg(value):
    value=str(value or "STAFF"); vals=[104]+[ord(ch)-32 if 32<=ord(ch)<=127 else 0 for ch in value]; checksum=(104+sum((i+1)*v for i,v in enumerate(vals[1:],1)))%103; vals.append(checksum); vals.append(106)
    x=4; scale=1.35; h=34; parts=[]
    for code in vals:
        pat=CODE128_B_PATTERNS[code]
        black=True
        for n in map(int,pat):
            if black: parts.append(f'<rect x="{x:.2f}" y="0" width="{n*scale:.2f}" height="{h}"/>')
            x+=n*scale; black=not black
    w=x+4
    return f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w:.2f} 48" preserveAspectRatio="none"><rect width="100%" height="100%" fill="white"/>'+''.join(parts)+f'<text x="{w/2:.2f}" y="46" text-anchor="middle" font-family="Arial" font-size="8">{value}</text></svg>'

@app.route("/admin/users/<int:uid>/id-card")
@admin_required
def user_id_card(uid):
    c=db(); u=c.execute("SELECT * FROM users WHERE id=?",(uid,)).fetchone(); c.close()
    if not u: return ("User not found",404)
    return render_template("user_id_card.html",u=u,barcode_svg=code128_svg(u["staff_id"]))

@app.route("/admin/users/<int:uid>/edit",methods=["POST"])
@admin_required
def edit_user(uid):
    c=db(); u=c.execute("SELECT * FROM users WHERE id=?",(uid,)).fetchone()
    if not u: c.close(); flash("User not found.","error"); return redirect(url_for("users"))
    try:
        role="SUPER_ADMIN" if request.form.get("role")=="SUPER_ADMIN" else "STAFF"
        username=request.form.get("username","").strip()
        if not username: raise ValueError("Username is required.")
        clash=c.execute("SELECT id FROM users WHERE username=? AND id<>?",(username,uid)).fetchone()
        if clash: raise ValueError("Username already exists.")
        c.execute("UPDATE users SET full_name=?,username=?,department=?,position=?,location=?,role=? WHERE id=?",(request.form.get("full_name","").strip(),username,request.form.get("department","Project"),request.form.get("position","Other"),request.form.get("location","").strip(),role,uid))
        c.commit(); flash("✏️ User profile, department, position and role updated.","success")
    except Exception as e: c.rollback(); flash("User update failed: "+str(e),"error")
    c.close(); return redirect(url_for("users"))

@app.route("/admin/users/<int:uid>/reset",methods=["POST"])
@admin_required
def reset_user(uid):
    c=db();p=request.form.get("password","").strip()
    if not p:flash("Password cannot be empty.","error")
    else:c.execute("UPDATE users SET password_hash=?,active=1 WHERE id=?",(generate_password_hash(p),uid));c.commit();flash("🔑 User password reset.","success")
    c.close();return redirect(url_for("users"))

@app.route("/admin/users/<int:uid>/toggle",methods=["POST"])
@admin_required
def toggle_user(uid):
    if uid==current_user()["id"]:flash("You cannot disable your own Super Admin account.","error")
    else:
        c=db();c.execute("UPDATE users SET active=CASE WHEN active=1 THEN 0 ELSE 1 END WHERE id=?",(uid,));c.commit();c.close();flash("User status updated.","success")
    return redirect(url_for("users"))

@app.route("/admin/users/<int:uid>/projects",methods=["POST"])
@admin_required
def assign_projects(uid):
    c=db();c.execute("DELETE FROM user_projects WHERE user_id=?",(uid,))
    for pid in request.form.getlist("project_ids"):c.execute("INSERT OR IGNORE INTO user_projects(user_id,project_id) VALUES(?,?)",(uid,pid))
    c.commit();c.close();flash("🏗️ Project access updated.","success");return redirect(url_for("users"))

@app.route("/admin/projects",methods=["GET","POST"])
@admin_required
def projects_admin():
    c=db()
    if request.method=="POST":
        try:c.execute("INSERT INTO projects(name,code,location,client,consultant,status,start_date,end_date,contractor_role,contract_sign_date,commencement_date,contract_end_date,contract_days,planned_income,planned_physical_pct,contract_value) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(request.form["name"],request.form["code"],request.form["location"],request.form.get("client",""),request.form.get("consultant",""),request.form.get("status","Active"),request.form.get("start_date","") ,request.form.get("end_date",""),request.form.get("contractor_role","Main Contractor"),request.form.get("contract_sign_date",""),request.form.get("commencement_date",request.form.get("start_date","")),request.form.get("contract_end_date",request.form.get("end_date","")),int(parse_float(request.form.get("contract_days"))),parse_float(request.form.get("planned_income")),parse_float(request.form.get("planned_physical_pct")),parse_float(request.form.get("contract_value"))));c.commit();flash("🏗️ Project created.","success")
        except Exception as e:c.rollback();flash(str(e),"error")
    projects=c.execute("SELECT * FROM projects ORDER BY name").fetchall();c.close();return render_template("projects_admin.html",projects=projects)



@app.route("/admin/projects/<int:pid>/edit",methods=["GET","POST"])
@admin_required
def edit_project(pid):
    c=db(); p=c.execute("SELECT * FROM projects WHERE id=?",(pid,)).fetchone()
    if not p: c.close(); flash("Project not found.","error"); return redirect(url_for("projects_admin"))
    if request.method=="POST":
        c.execute("""UPDATE projects SET name=?,code=?,location=?,client=?,consultant=?,status=?,contractor_role=?,contract_sign_date=?,commencement_date=?,contract_end_date=?,contract_days=?,planned_income=?,planned_physical_pct=?,contract_value=?,start_date=?,end_date=? WHERE id=?""",(request.form["name"],request.form["code"],request.form.get("location",""),request.form.get("client",""),request.form.get("consultant",""),request.form.get("status","Active"),request.form.get("contractor_role","Main Contractor"),request.form.get("contract_sign_date",""),request.form.get("commencement_date",""),request.form.get("contract_end_date",""),int(parse_float(request.form.get("contract_days"))),parse_float(request.form.get("planned_income")),parse_float(request.form.get("planned_physical_pct")),parse_float(request.form.get("contract_value")),request.form.get("start_date",""),request.form.get("end_date",""),pid)); c.commit(); flash("🏗️ Project baseline updated.","success"); p=c.execute("SELECT * FROM projects WHERE id=?",(pid,)).fetchone()
    c.close(); return render_template("project_edit.html",p=p)

def _norm_header(v):
    if v is None: return ""
    x=str(v).strip().lower().replace("\n"," ")
    for ch in [".",":","/","\\","(",")","-","_"]:
        x=x.replace(ch," ")
    return " ".join(x.split())

def _excel_number(v):
    if v is None or v=="": return 0.0
    if isinstance(v,(int,float)): return float(v)
    x=str(v).strip().replace(",","").replace("ETB","").replace("Birr","").strip()
    if x.startswith("="): return 0.0
    try: return float(x)
    except Exception: return 0.0

def import_boq_xlsx(path,pid):
    """Import real-world BOQ spreadsheets without requiring exact column wording.
    Finds the header row anywhere in each sheet and accepts common Ethiopian/contract BOQ labels.
    """
    wb_values=load_workbook(path,data_only=True,read_only=True)
    wb_formulas=load_workbook(path,data_only=False,read_only=True)
    c=db(); count=0; sheets=[]; scanned=0
    try:
        for ws_val, ws_formula in zip(wb_values.worksheets, wb_formulas.worksheets):
            max_rows=min(ws_val.max_row or 1,10000)
            rows=list(ws_val.iter_rows(min_row=1,max_row=max_rows,values_only=True))
            header_idx=None; header=None
            for idx,row in enumerate(rows):
                vals=[_norm_header(v) for v in row]
                joined=" | ".join(vals)
                has_item=any(v in vals or v.startswith("item no") or v.startswith("item number") or v in ("no","s no","s n","bill item","bill no") for v in vals)
                has_desc=any(("description" in v) or ("activity" in v) or ("work description" in v) or ("particular" in v) or ("item description" in v) for v in vals)
                has_unit=any(v=="unit" or v.endswith(" unit") or v=="uom" for v in vals)
                has_qty=any(("quantity" in v) or v in ("qty","contract qty","contract quantity","total quantity") for v in vals)
                has_rate=any(("rate" in v) or ("unit price" in v) or ("unit cost" in v) or ("price"==v) for v in vals)
                if has_desc and has_unit and has_rate and (has_item or has_qty):
                    header_idx=idx; header=vals; break
            if header_idx is None: continue
            def find_col(kind):
                exact={
                  "item": ["item no","item number","item no.","bill item","bill no","no","s no","s n","item"],
                  "desc": ["description","description of work","description of works","work description","activity","particular","particulars","item description"],
                  "unit": ["unit","uom","unit of measurement"],
                  "qty": ["contract quantity","contract qty","total contract quantity","quantity","qty","total quantity","original quantity"],
                  "rate": ["contract rate","unit rate","unit price","unit cost","rate","price"],
                }[kind]
                for i,v in enumerate(header):
                    if v in exact: return i
                for i,v in enumerate(header):
                    if any(x in v for x in exact): return i
                return None
            ci,cd,cu,cq,cr=[find_col(k) for k in ("item","desc","unit","qty","rate")]
            cs,ct=find_col("series") if False else (None,None)
            for i,v in enumerate(header):
                if v in ("series","section","bill section","section no","series no","series number") or "series" in v: cs=i; break
            for i,v in enumerate(header):
                if v in ("title","item title","work title") or "title" in v: ct=i; break
            if cd is None or cu is None or cr is None or (ci is None and cq is None): continue
            sheet_count=0
            for row_idx,row in enumerate(rows[header_idx+1:], start=header_idx+2):
                scanned+=1
                def cell(i): return row[i] if i is not None and i<len(row) else None
                item=cell(ci); desc=cell(cd); unit=cell(cu)
                # Skip repeated headers, section totals and empty lines.
                item_s="" if item is None else str(item).strip()
                desc_s="" if desc is None else str(desc).strip()
                if not desc_s or not item_s: continue
                if _norm_header(item_s) in {"no","item","item no","item number","total","subtotal"}: continue
                lowdesc=_norm_header(desc_s)
                if lowdesc in {"description","description of works","description of work","activity"}: continue
                qty=_excel_number(cell(cq)); rate=_excel_number(cell(cr))
                unit_s="" if unit is None else str(unit).strip()
                # Some BOQs have formulas in the data-only workbook. Try the formula workbook's cached-looking numeric cells.
                if rate==0 and cr is not None and row_idx<=ws_formula.max_row:
                    try:
                        fv=ws_formula.cell(row=row_idx,column=cr+1).value
                        if isinstance(fv,(int,float)): rate=float(fv)
                    except Exception: pass
                if qty==0 and cq is not None and row_idx<=ws_formula.max_row:
                    try:
                        qv=ws_formula.cell(row=row_idx,column=cq+1).value
                        if isinstance(qv,(int,float)): qty=float(qv)
                    except Exception: pass
                # A valid BOQ row should contain at least a unit or a numeric quantity/rate.
                if not unit_s and qty==0 and rate==0: continue
                series_s="" if cs is None or cell(cs) is None else str(cell(cs)).strip()
                title_s="" if ct is None or cell(ct) is None else str(cell(ct)).strip()
                c.execute("INSERT INTO boq(project_id,item_no,description,unit,rate,contract_qty,source_sheet,series,title) VALUES(?,?,?,?,?,?,?,?,?) ON CONFLICT(project_id,item_no) DO UPDATE SET description=excluded.description,unit=excluded.unit,rate=excluded.rate,contract_qty=excluded.contract_qty,source_sheet=excluded.source_sheet,series=excluded.series,title=excluded.title",(pid,item_s,desc_s,unit_s,rate,qty,ws_val.title,series_s,title_s))
                count+=1; sheet_count+=1
            if sheet_count: sheets.append(f"{ws_val.title} ({sheet_count})")
        c.commit()
    except Exception:
        c.rollback(); raise
    finally:
        c.close(); wb_values.close(); wb_formulas.close()
    if not sheets:
        raise ValueError("No BOQ rows were detected. The file must contain columns similar to Item No/No., Description of Works, Unit, Quantity and Rate/Unit Price.")
    return count, ", ".join(sheets)

@app.route("/admin/boq/<int:pid>",methods=["GET","POST"])
@admin_required
def boq_admin(pid):
    c=db()
    if request.method=="POST":
        action=request.form.get("action")
        try:
            if action=="add":
                c.execute("INSERT INTO boq(project_id,item_no,description,unit,rate,contract_qty,series,title) VALUES(?,?,?,?,?,?,?,?)",(pid,request.form["item_no"].strip(),request.form["description"].strip(),request.form.get("unit","").strip(),parse_float(request.form.get("rate")),parse_float(request.form.get("contract_qty")),request.form.get("series","").strip(),request.form.get("title","").strip()))
            elif action=="settings":
                c.execute("INSERT INTO boq_settings(project_id,title,revision,effective_date) VALUES(?,?,?,?) ON CONFLICT(project_id) DO UPDATE SET title=excluded.title,revision=excluded.revision,effective_date=excluded.effective_date",(pid,request.form.get("boq_title","").strip(),request.form.get("revision","").strip(),request.form.get("effective_date") or None))
            elif action=="edit":
                bid=int(request.form["boq_id"]); c.execute("UPDATE boq SET item_no=?,description=?,unit=?,rate=?,contract_qty=?,series=?,title=? WHERE id=? AND project_id=?",(request.form["item_no"].strip(),request.form["description"].strip(),request.form.get("unit","").strip(),parse_float(request.form.get("rate")),parse_float(request.form.get("contract_qty")),request.form.get("series","").strip(),request.form.get("title","").strip(),bid,pid))
            elif action=="upload":
                f=request.files.get("file")
                if not f or not f.filename.lower().endswith((".xlsx",".xlsm")): raise ValueError("Upload an .xlsx or .xlsm BOQ file.")
                name=secure_filename(f.filename); path=os.path.join(UPLOADS,f"{dt.datetime.now().strftime('%Y%m%d%H%M%S')}_{name}"); f.save(path)
                c.commit(); c.close(); c=None
                try:
                    count,sheet=import_boq_xlsx(path,pid)
                    c=db(); c.execute("INSERT INTO boq_uploads(project_id,filename,uploaded_at,user_id,rows_imported) VALUES(?,?,?,?,?)",(pid,name,dt.datetime.now().isoformat(timespec="seconds"),current_user()["id"],count)); c.commit(); flash(f"📑 BOQ imported successfully: {count} items from {sheet}.","success")
                except Exception as e:
                    c=db(); c.rollback(); flash("BOQ import error: "+str(e),"error")
            if c is not None: c.commit()
        except Exception as e:
            if c is not None:
                try: c.rollback()
                except Exception: pass
            flash("BOQ import error: "+str(e),"error")
    rows=c.execute("SELECT * FROM boq WHERE project_id=? ORDER BY item_no",(pid,)).fetchall();uploads=c.execute("SELECT * FROM boq_uploads WHERE project_id=? ORDER BY uploaded_at DESC LIMIT 10",(pid,)).fetchall();settings=c.execute("SELECT * FROM boq_settings WHERE project_id=?",(pid,)).fetchone();c.close();return render_template("boq.html",pid=pid,rows=rows,uploads=uploads,settings=settings)

@app.route("/admin")
@admin_required
def admin():return redirect(url_for("users"))
@app.route("/health")
def health():return "OK"

@app.after_request
def security_headers(resp):
    resp.headers.setdefault("Cache-Control", "no-store")
    return resp

init_db()
if __name__=="__main__":app.run(host="0.0.0.0",port=int(os.environ.get("PORT",5000)),debug=True)
