import os, sqlite3, calendar, datetime as dt, json
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from openpyxl import load_workbook

BASE=os.path.dirname(os.path.abspath(__file__))
DB=os.path.join(BASE,"bagc.db")
UPLOADS=os.path.join(BASE,"uploads")
os.makedirs(UPLOADS,exist_ok=True)
app=Flask(__name__)
app.secret_key=os.environ.get("SECRET_KEY","bagc-change-this-secret")

DEPARTMENTS=["Administration","Design","Machinery","Finance","HR","Store","Project"]
MACHINE_TYPES=["Dozer","Excavator","Wheel Loader","Backhoe Loader","Motor Grader","Roller","Dump Truck","Water Truck","Crane","Forklift","Concrete Mixer","Concrete Pump","Batching Plant","Crusher","Asphalt Plant","Asphalt Paver","Bitumen Distributor","Road Sweeper","Generator","Welding Machine","Vibrator","Air Compressor","Pickup","Other"]
MATERIAL_CATEGORIES=["Common Construction","Concrete","Rebar","Formwork","Masonry","Finishing","Sanitary","Plumbing","Electrical","Aluminium","Glass","Road Works","Fuel & Oil","Spare Parts","Stationery & Cleaning","PPE","Other"]
MATERIAL_CATALOG=["Cement","Sand","Fine Aggregate","Coarse Aggregate 10mm","Coarse Aggregate 20mm","Coarse Aggregate 40mm","Water","Rebar Ø8","Rebar Ø10","Rebar Ø12","Rebar Ø16","Rebar Ø20","Binding Wire","Black Wire Roll","Nails 2in","Nails 3in","Nails 4in","Plywood 4x8","Eucalyptus Pole","Timber","Form Oil","Stone","Natural Stone Cladding","Aluminium Frame","Aluminium Sheet","Glass","PVC Pipe 110mm","PVC Pipe 160mm","HDPE Pipe","Electrical Cable","Conduit","Switch","Socket","Sanitary Fixture","Tile","Paint","Fuel Diesel","Engine Oil","Hydraulic Oil","Grease","Bitumen","Welding Rod","Other"]
DESIGN_STATUSES=["Draft","Submitted","Under Review","Approved","Approved with Comments","Revise & Resubmit","Rejected","As-Built","Handed Over"]


def db():
    c=sqlite3.connect(DB)
    c.row_factory=sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    return c


def init_db():
    c=db()
    c.executescript('''
    CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY,full_name TEXT,username TEXT UNIQUE,password_hash TEXT,department TEXT,location TEXT,role TEXT,active INTEGER DEFAULT 1,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
    CREATE TABLE IF NOT EXISTS projects(id INTEGER PRIMARY KEY,name TEXT UNIQUE,code TEXT,location TEXT,client TEXT,consultant TEXT,status TEXT DEFAULT 'Active',start_date TEXT,end_date TEXT);
    CREATE TABLE IF NOT EXISTS user_projects(user_id INTEGER,project_id INTEGER,UNIQUE(user_id,project_id));
    CREATE TABLE IF NOT EXISTS boq(id INTEGER PRIMARY KEY,project_id INTEGER,item_no TEXT,description TEXT,unit TEXT,rate REAL DEFAULT 0,contract_qty REAL DEFAULT 0,source_sheet TEXT,UNIQUE(project_id,item_no));
    CREATE TABLE IF NOT EXISTS daily_work(id INTEGER PRIMARY KEY,project_id INTEGER,date TEXT,boq_id INTEGER,quantity REAL,station_from TEXT,station_to TEXT,notes TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS machines(id INTEGER PRIMARY KEY,project_id INTEGER,machine_type TEXT,code TEXT,ownership TEXT,hourly_rate REAL DEFAULT 0,expected_fuel REAL DEFAULT 0,active INTEGER DEFAULT 1);
    CREATE TABLE IF NOT EXISTS machine_logs(id INTEGER PRIMARY KEY,project_id INTEGER,machine_id INTEGER,date TEXT,work_hours REAL DEFAULT 0,idle_hours REAL DEFAULT 0,idle_reason TEXT,down_hours REAL DEFAULT 0,down_reason TEXT,opening_gauge REAL DEFAULT 0,fuel_received REAL DEFAULT 0,closing_gauge REAL DEFAULT 0,notes TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS materials(id INTEGER PRIMARY KEY,project_id INTEGER,category TEXT,name TEXT,unit TEXT,min_stock REAL DEFAULT 0,active INTEGER DEFAULT 1,UNIQUE(project_id,name));
    CREATE TABLE IF NOT EXISTS store_logs(id INTEGER PRIMARY KEY,project_id INTEGER,material_id INTEGER,date TEXT,received REAL DEFAULT 0,issued REAL DEFAULT 0,unit_cost REAL DEFAULT 0,physical_balance REAL,reference TEXT,notes TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS manpower(id INTEGER PRIMARY KEY,project_id INTEGER,date TEXT,name TEXT,employment TEXT,position TEXT,present REAL DEFAULT 1,daily_rate REAL DEFAULT 0,overtime_hours REAL DEFAULT 0,overtime_rate REAL DEFAULT 0,notes TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS design_items(id INTEGER PRIMARY KEY,project_id INTEGER,drawing_no TEXT,title TEXT,discipline TEXT,revision TEXT,status TEXT,submitted TEXT,approved TEXT,comments TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS finance_logs(id INTEGER PRIMARY KEY,project_id INTEGER,date TEXT,category TEXT,kind TEXT,description TEXT,amount REAL DEFAULT 0,reference TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS boq_uploads(id INTEGER PRIMARY KEY,project_id INTEGER,filename TEXT,uploaded_at TEXT,user_id INTEGER,rows_imported INTEGER DEFAULT 0);
    CREATE TABLE IF NOT EXISTS performance_rates(id INTEGER PRIMARY KEY,project_id INTEGER,work_type TEXT,worker_type TEXT,unit TEXT,qty_per_hour REAL DEFAULT 0,notes TEXT,UNIQUE(project_id,work_type,worker_type));
    CREATE TABLE IF NOT EXISTS daily_activities(id INTEGER PRIMARY KEY,project_id INTEGER,date TEXT,boq_id INTEGER,work_type TEXT,executed_qty REAL DEFAULT 0,machine_id INTEGER,machine_hours REAL DEFAULT 0,manpower_position TEXT,manpower_qty REAL DEFAULT 0,manpower_hours REAL DEFAULT 0,material_id INTEGER,material_qty REAL DEFAULT 0,remarks TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS problems(id INTEGER PRIMARY KEY,project_id INTEGER,date TEXT,problem TEXT,remark TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS fuel_logs(id INTEGER PRIMARY KEY,project_id INTEGER,machine_id INTEGER,date TEXT,opening_gauge REAL DEFAULT 0,fuel_received REAL DEFAULT 0,closing_gauge REAL DEFAULT 0,fuel_price REAL DEFAULT 0,reference TEXT,notes TEXT,user_id INTEGER);
    CREATE TABLE IF NOT EXISTS report_settings(id INTEGER PRIMARY KEY,project_id INTEGER UNIQUE,contractor_role TEXT DEFAULT 'Main Contractor',phone TEXT,email TEXT,website TEXT,fax TEXT,address TEXT,logo_text TEXT);
    ''')
    # Safe migrations for databases created by earlier BAGC versions.
    existing=[r['name'] for r in c.execute("PRAGMA table_info(machines)").fetchall()]
    for col,typ in [('plate_no','TEXT'),('engine_no','TEXT'),('fuel_price','REAL DEFAULT 0')]:
        if col not in existing: c.execute(f"ALTER TABLE machines ADD COLUMN {col} {typ}")
    existing_mp=[r['name'] for r in c.execute("PRAGMA table_info(manpower)").fetchall()]
    for col,typ in [('working_hours','REAL DEFAULT 8'),('hourly_rate','REAL DEFAULT 0')]:
        if col not in existing_mp: c.execute(f"ALTER TABLE manpower ADD COLUMN {col} {typ}")
    existing_p=[r['name'] for r in c.execute("PRAGMA table_info(projects)").fetchall()]
    for col,typ in [('contractor_role',"TEXT DEFAULT 'Main Contractor'")]:
        if col not in existing_p: c.execute(f"ALTER TABLE projects ADD COLUMN {col} {typ}")
    # ENV-controlled Super Admin synchronization fixes an already-created SQLite DB.
    u=os.environ.get("ADMIN_USERNAME","admin").strip() or "admin"
    p=os.environ.get("ADMIN_PASSWORD","admin123")
    admin=c.execute("SELECT id FROM users WHERE role='SUPER_ADMIN' ORDER BY id LIMIT 1").fetchone()
    if not admin:
        c.execute("INSERT INTO users(full_name,username,password_hash,department,location,role) VALUES(?,?,?,?,?,?)",("System Administrator",u,generate_password_hash(p),"Administration","Head Office","SUPER_ADMIN"))
    else:
        c.execute("UPDATE users SET username=?,password_hash=?,active=1,department='Administration',location='Head Office' WHERE id=?",(u,generate_password_hash(p),admin["id"]))
    if not c.execute("SELECT id FROM projects").fetchone():
        c.execute("INSERT INTO projects(name,code,location,status) VALUES(?,?,?,?)",("Koye Feche","KOYE","Koye Feche","Active"))
    c.commit();c.close()


def current_user():
    if not session.get("user_id"): return None
    c=db(); u=c.execute("SELECT * FROM users WHERE id=? AND active=1",(session["user_id"],)).fetchone(); c.close(); return u

@app.context_processor
def inject():
    return {"me":current_user(),"machine_types":MACHINE_TYPES,"material_categories":MATERIAL_CATEGORIES,"material_catalog":MATERIAL_CATALOG,"design_statuses":DESIGN_STATUSES,"today":dt.date.today().isoformat()}


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


def dashboard_data(pid=None):
    c=db(); where="" if pid is None else " WHERE p.id=?"; args=() if pid is None else (pid,)
    projects=c.execute("SELECT p.* FROM projects p"+where+" ORDER BY p.name",args).fetchall()
    out=[]
    for p in projects:
        inc=c.execute("SELECT COALESCE(SUM(dw.quantity*b.rate),0) x FROM daily_work dw JOIN boq b ON b.id=dw.boq_id WHERE dw.project_id=?",(p["id"],)).fetchone()["x"]
        me=c.execute("SELECT COALESCE(SUM(ml.work_hours*m.hourly_rate),0) x FROM machine_logs ml JOIN machines m ON m.id=ml.machine_id WHERE ml.project_id=?",(p["id"],)).fetchone()["x"]
        pe=c.execute("SELECT COALESCE(SUM(CASE WHEN mp.hourly_rate>0 THEN mp.present*mp.working_hours*mp.hourly_rate ELSE mp.present*mp.daily_rate END+mp.overtime_hours*mp.overtime_rate),0) x FROM manpower mp WHERE mp.project_id=?",(p["id"],)).fetchone()["x"]
        se=c.execute("SELECT COALESCE(SUM(sl.issued*sl.unit_cost),0) x FROM store_logs sl WHERE sl.project_id=?",(p["id"],)).fetchone()["x"]
        other=c.execute("SELECT COALESCE(SUM(amount),0) x FROM finance_logs WHERE project_id=? AND kind='Expense'",(p["id"],)).fetchone()["x"]
        workers=c.execute("SELECT COUNT(*) x FROM manpower WHERE project_id=? AND present>0",(p["id"],)).fetchone()["x"]
        machines=c.execute("SELECT COUNT(*) x FROM machines WHERE project_id=? AND active=1",(p["id"],)).fetchone()["x"]
        total_exp=money(me+pe+se+other)
        daily_m=c.execute("SELECT ml.*,m.machine_type,m.code,m.plate_no,m.engine_no,m.ownership,(ml.work_hours*m.hourly_rate) expense,(ml.opening_gauge+ml.fuel_received-ml.closing_gauge) actual_fuel FROM machine_logs ml JOIN machines m ON m.id=ml.machine_id WHERE ml.project_id=? ORDER BY ml.date DESC,ml.id DESC LIMIT 8",(p["id"],)).fetchall()
        daily_mat=c.execute("SELECT sl.*,m.name,m.category,m.unit FROM store_logs sl JOIN materials m ON m.id=sl.material_id WHERE sl.project_id=? ORDER BY sl.date DESC,sl.id DESC LIMIT 8",(p["id"],)).fetchall()
        out.append({"p":p,"income":money(inc),"expense":total_exp,"expense_pct":money((total_exp/(inc+0.0001))*100) if inc else 0,"machine_expense":money(me),"manpower_expense":money(pe),"store_expense":money(se),"other_expense":money(other),"workers":workers,"machines":machines,"daily_machines":daily_m,"daily_materials":daily_mat})
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
    if not allowed_project(pid):flash("🚫 You do not have access to this project.","error");return redirect(url_for("dashboard"))
    c=db();p=c.execute("SELECT * FROM projects WHERE id=?",(pid,)).fetchone();boq_count=c.execute("SELECT COUNT(*) n FROM boq WHERE project_id=?",(pid,)).fetchone()["n"];machine_count=c.execute("SELECT COUNT(*) n FROM machines WHERE project_id=? AND active=1",(pid,)).fetchone()["n"];mat_count=c.execute("SELECT COUNT(*) n FROM materials WHERE project_id=? AND active=1",(pid,)).fetchone()["n"];c.close()
    return render_template("project.html",p=p,boq_count=boq_count,machine_count=machine_count,mat_count=mat_count)

@app.route("/projects/<int:pid>/daily",methods=["GET","POST"])
@login_required
def daily(pid):
    if not allowed_project(pid) or not can_module("Project"):
        flash("🚫 Project daily reporting access is not assigned to your account.","error");return redirect(url_for("project",pid=pid))
    c=db(); default_date=request.args.get("date",dt.date.today().isoformat())
    if request.method=="POST":
        d=request.form.get("date") or default_date
        # BOQ work
        if request.form.get("section")=="boq":
            c.execute("INSERT INTO daily_work(project_id,date,boq_id,quantity,station_from,station_to,notes,user_id) VALUES(?,?,?,?,?,?,?,?)",(pid,d,request.form["boq_id"],parse_float(request.form["quantity"]),request.form.get("station_from",""),request.form.get("station_to",""),request.form.get("notes",""),current_user()["id"]))
            flash("📐 BOQ work registered — income = quantity × BOQ rate.","success")
        elif request.form.get("section")=="machinery":
            mid=request.form["machine_id"];c.execute("INSERT INTO machine_logs(project_id,machine_id,date,work_hours,idle_hours,idle_reason,down_hours,down_reason,opening_gauge,fuel_received,closing_gauge,notes,user_id) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",(pid,mid,d,parse_float(request.form["work_hours"]),parse_float(request.form["idle_hours"]),request.form.get("idle_reason",""),parse_float(request.form["down_hours"]),request.form.get("down_reason",""),parse_float(request.form["opening_gauge"]),parse_float(request.form["fuel_received"]),parse_float(request.form["closing_gauge"]),request.form.get("notes",""),current_user()["id"]))
            flash("🚜 Machinery daily time saved.","success")
        elif request.form.get("section")=="manpower":
            c.execute("INSERT INTO manpower(project_id,date,name,employment,position,present,working_hours,hourly_rate,daily_rate,overtime_hours,overtime_rate,notes,user_id) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",(pid,d,request.form["name"],request.form["employment"],request.form["position"],parse_float(request.form["present"]),parse_float(request.form.get("working_hours",8)),parse_float(request.form.get("hourly_rate")),parse_float(request.form["daily_rate"]),parse_float(request.form["overtime_hours"]),parse_float(request.form["overtime_rate"]),request.form.get("notes",""),current_user()["id"]))
            flash("👷 Manpower daily attendance saved.","success")
        elif request.form.get("section")=="store":
            physical=request.form.get("physical_balance");physical=parse_float(physical) if physical not in (None,"") else None
            c.execute("INSERT INTO store_logs(project_id,material_id,date,received,issued,unit_cost,physical_balance,reference,notes,user_id) VALUES(?,?,?,?,?,?,?,?,?,?)",(pid,request.form["material_id"],d,parse_float(request.form["received"]),parse_float(request.form["issued"]),parse_float(request.form["unit_cost"]),physical,request.form.get("reference",""),request.form.get("notes",""),current_user()["id"]))
            flash("📦 Store receipt/issue saved.","success")
        elif request.form.get("section")=="finance":
            c.execute("INSERT INTO finance_logs(project_id,date,category,kind,description,amount,reference,user_id) VALUES(?,?,?,?,?,?,?,?)",(pid,d,request.form["category"],request.form["kind"],request.form["description"],parse_float(request.form["amount"]),request.form.get("reference",""),current_user()["id"]))
            flash("💰 Financial entry saved.","success")
        elif request.form.get("section")=="activity":
            c.execute("INSERT INTO daily_activities(project_id,date,boq_id,work_type,executed_qty,machine_id,machine_hours,manpower_position,manpower_qty,manpower_hours,material_id,material_qty,remarks,user_id) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(pid,d,request.form.get("boq_id") or None,request.form.get("work_type",""),parse_float(request.form.get("executed_qty")),request.form.get("machine_id") or None,parse_float(request.form.get("machine_hours")),request.form.get("manpower_position",""),parse_float(request.form.get("manpower_qty")),parse_float(request.form.get("manpower_hours")),request.form.get("material_id") or None,parse_float(request.form.get("material_qty")),request.form.get("remarks",""),current_user()["id"]))
            flash("🏗️ Detailed work activity registered.","success")
        elif request.form.get("section")=="problem":
            c.execute("INSERT INTO problems(project_id,date,problem,remark,user_id) VALUES(?,?,?,?,?)",(pid,d,request.form.get("problem",""),request.form.get("remark",""),current_user()["id"]))
            flash("⚠️ Problem and corrective remark saved.","success")
        c.commit()
    boq=c.execute("SELECT * FROM boq WHERE project_id=? ORDER BY item_no",(pid,)).fetchall();machines=c.execute("SELECT * FROM machines WHERE project_id=? AND active=1 ORDER BY machine_type,code",(pid,)).fetchall();materials=c.execute("SELECT * FROM materials WHERE project_id=? AND active=1 ORDER BY category,name",(pid,)).fetchall()
    recent=c.execute("SELECT dw.*,b.item_no,b.description,b.unit,b.rate,dw.quantity*b.rate amount FROM daily_work dw JOIN boq b ON b.id=dw.boq_id WHERE dw.project_id=? ORDER BY dw.date DESC,dw.id DESC LIMIT 20",(pid,)).fetchall();c.close()
    return render_template("daily.html",pid=pid,date=default_date,boq=boq,machines=machines,materials=materials,recent=recent)


@app.route("/projects/<int:pid>/fuel",methods=["GET","POST"])
@login_required
def fuel(pid):
    if not allowed_project(pid) or not can_module("Machinery"): return redirect(url_for("project",pid=pid))
    c=db()
    if request.method=="POST":
        c.execute("INSERT INTO fuel_logs(project_id,machine_id,date,opening_gauge,fuel_received,closing_gauge,fuel_price,reference,notes,user_id) VALUES(?,?,?,?,?,?,?,?,?,?)",(pid,request.form["machine_id"],request.form["date"],parse_float(request.form["opening_gauge"]),parse_float(request.form["fuel_received"]),parse_float(request.form["closing_gauge"]),parse_float(request.form["fuel_price"]),request.form.get("reference",""),request.form.get("notes",""),current_user()["id"])); c.commit(); flash("⛽ Fuel log saved.","success")
    machines=c.execute("SELECT * FROM machines WHERE project_id=? AND active=1 ORDER BY machine_type,code",(pid,)).fetchall()
    logs=c.execute("SELECT f.*,m.machine_type,m.code,m.plate_no,m.engine_no,m.ownership,(f.opening_gauge+f.fuel_received-f.closing_gauge) consumption,(f.fuel_received*f.fuel_price) cost,COALESCE((SELECT SUM(ml.work_hours) FROM machine_logs ml WHERE ml.machine_id=f.machine_id AND ml.date=f.date),0) work_hours,COALESCE((SELECT SUM(ml.work_hours) FROM machine_logs ml WHERE ml.machine_id=f.machine_id AND ml.date=f.date),0)*m.expected_fuel expected_consumption FROM fuel_logs f JOIN machines m ON m.id=f.machine_id WHERE f.project_id=? ORDER BY f.date DESC,f.id DESC LIMIT 100",(pid,)).fetchall(); c.close()
    return render_template("fuel.html",pid=pid,machines=machines,logs=logs)

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
    activities=c.execute("SELECT a.*,b.item_no,b.description,m.machine_type,m.code,m.plate_no,mat.name material,mat.unit material_unit FROM daily_activities a LEFT JOIN boq b ON b.id=a.boq_id LEFT JOIN machines m ON m.id=a.machine_id LEFT JOIN materials mat ON mat.id=a.material_id WHERE a.project_id=? AND a.date=? ORDER BY a.id",(pid,date)).fetchall()
    machines=c.execute("SELECT ml.*,m.machine_type,m.code,m.plate_no,m.ownership FROM machine_logs ml JOIN machines m ON m.id=ml.machine_id WHERE ml.project_id=? AND ml.date=?",(pid,date)).fetchall()
    manpower=c.execute("SELECT * FROM manpower WHERE project_id=? AND date=?",(pid,date)).fetchall(); store=c.execute("SELECT sl.*,m.name,m.unit FROM store_logs sl JOIN materials m ON m.id=sl.material_id WHERE sl.project_id=? AND sl.date=?",(pid,date)).fetchall(); problems=c.execute("SELECT * FROM problems WHERE project_id=? AND date=?",(pid,date)).fetchall(); c.close()
    return render_template("print_report.html",p=p,settings=settings,date=date,boq=boq,activities=activities,machines=machines,manpower=manpower,store=store,problems=problems)

@app.route("/projects/<int:pid>/machinery",methods=["GET","POST"])
@login_required
def machinery(pid):
    if not allowed_project(pid) or not can_module("Machinery"):flash("🚫 Machinery access is not assigned.","error");return redirect(url_for("project",pid=pid))
    c=db()
    if request.method=="POST":
        action=request.form.get("action")
        if action=="add":
            c.execute("INSERT INTO machines(project_id,machine_type,code,plate_no,engine_no,ownership,hourly_rate,expected_fuel,fuel_price) VALUES(?,?,?,?,?,?,?,?,?)",(pid,request.form["machine_type"],request.form["code"],request.form.get("plate_no",request.form["code"]),request.form.get("engine_no",""),request.form["ownership"],parse_float(request.form["hourly_rate"]),parse_float(request.form["expected_fuel"]),parse_float(request.form.get("fuel_price"))))
            flash("🚜 Machine added to this project's fleet.","success")
        elif action=="remove":c.execute("UPDATE machines SET active=0 WHERE id=? AND project_id=?",(request.form["machine_id"],pid));flash("Machine removed from active fleet.","success")
        elif action=="log":
            c.execute("INSERT INTO machine_logs(project_id,machine_id,date,work_hours,idle_hours,idle_reason,down_hours,down_reason,opening_gauge,fuel_received,closing_gauge,notes,user_id) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",(pid,request.form["machine_id"],request.form["date"],parse_float(request.form["work_hours"]),parse_float(request.form["idle_hours"]),request.form.get("idle_reason",""),parse_float(request.form["down_hours"]),request.form.get("down_reason",""),parse_float(request.form["opening_gauge"]),parse_float(request.form["fuel_received"]),parse_float(request.form["closing_gauge"]),request.form.get("notes",""),current_user()["id"]))
            flash("⏱️ Machine hours / idle / down / fuel log saved.","success")
        c.commit()
    machines=c.execute("SELECT * FROM machines WHERE project_id=? AND active=1 ORDER BY machine_type,code",(pid,)).fetchall();logs=c.execute("SELECT ml.*,m.machine_type,m.code,m.ownership,m.hourly_rate,m.expected_fuel,(ml.work_hours*m.hourly_rate) expense,(ml.opening_gauge+ml.fuel_received-ml.closing_gauge) actual_fuel,(ml.work_hours*m.expected_fuel) expected_fuel_qty,CASE WHEN (ml.work_hours+ml.idle_hours+ml.down_hours)>0 THEN ml.work_hours*100.0/(ml.work_hours+ml.idle_hours+ml.down_hours) ELSE 0 END utilization,CASE WHEN (ml.work_hours+ml.idle_hours+ml.down_hours)>0 THEN (ml.work_hours+ml.idle_hours)*100.0/(ml.work_hours+ml.idle_hours+ml.down_hours) ELSE 0 END availability,(ml.opening_gauge+ml.fuel_received-ml.closing_gauge)-(ml.work_hours*m.expected_fuel) fuel_discrepancy FROM machine_logs ml JOIN machines m ON m.id=ml.machine_id WHERE ml.project_id=? ORDER BY ml.date DESC,ml.id DESC LIMIT 50",(pid,)).fetchall();c.close()
    return render_template("machinery.html",pid=pid,machines=machines,logs=logs)

@app.route("/projects/<int:pid>/manpower",methods=["GET","POST"])
@login_required
def manpower(pid):
    if not allowed_project(pid) or not can_module("HR"):flash("🚫 HR/manpower access is not assigned.","error");return redirect(url_for("project",pid=pid))
    c=db()
    if request.method=="POST":
        c.execute("INSERT INTO manpower(project_id,date,name,employment,position,present,daily_rate,overtime_hours,overtime_rate,notes,user_id) VALUES(?,?,?,?,?,?,?,?,?,?,?)",(pid,request.form["date"],request.form["name"],request.form["employment"],request.form["position"],parse_float(request.form["present"]),parse_float(request.form["daily_rate"]),parse_float(request.form["overtime_hours"]),parse_float(request.form["overtime_rate"]),request.form.get("notes",""),current_user()["id"]));c.commit();flash("👷 Manpower record saved.","success")
    rows=c.execute("SELECT * FROM manpower WHERE project_id=? ORDER BY date DESC,id DESC LIMIT 100",(pid,)).fetchall();c.close();return render_template("manpower.html",pid=pid,rows=rows)

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
    if not allowed_project(pid):return redirect(url_for("dashboard"))
    period=request.args.get("period","month");anchor=request.args.get("date",dt.date.today().isoformat());start,end=period_bounds(period,anchor)
    c=db();rows=c.execute("SELECT * FROM boq WHERE project_id=? ORDER BY item_no",(pid,)).fetchall();out=[]
    for b in rows:
        prev=c.execute("SELECT COALESCE(SUM(quantity),0) q FROM daily_work WHERE boq_id=? AND date<?",(b["id"],start.isoformat())).fetchone()["q"]
        per=c.execute("SELECT COALESCE(SUM(quantity),0) q FROM daily_work WHERE boq_id=? AND date BETWEEN ? AND ?",(b["id"],start.isoformat(),end.isoformat())).fetchone()["q"]
        td=c.execute("SELECT COALESCE(SUM(quantity),0) q FROM daily_work WHERE boq_id=?",(b["id"],)).fetchone()["q"]
        out.append({"item_no":b["item_no"],"description":b["description"],"unit":b["unit"],"rate":b["rate"],"contract_qty":b["contract_qty"],"previous_qty":prev,"period_qty":per,"todate_qty":td,"previous_amount":prev*b["rate"],"period_amount":per*b["rate"],"todate_amount":td*b["rate"]})
    inc=c.execute("SELECT COALESCE(SUM(quantity*b.rate),0) x FROM daily_work dw JOIN boq b ON b.id=dw.boq_id WHERE dw.project_id=? AND date BETWEEN ? AND ?",(pid,start.isoformat(),end.isoformat())).fetchone()["x"]
    me=c.execute("SELECT COALESCE(SUM(ml.work_hours*m.hourly_rate),0) x FROM machine_logs ml JOIN machines m ON m.id=ml.machine_id WHERE ml.project_id=? AND ml.date BETWEEN ? AND ?",(pid,start.isoformat(),end.isoformat())).fetchone()["x"]
    pe=c.execute("SELECT COALESCE(SUM(present*daily_rate+overtime_hours*overtime_rate),0) x FROM manpower WHERE project_id=? AND date BETWEEN ? AND ?",(pid,start.isoformat(),end.isoformat())).fetchone()["x"]
    se=c.execute("SELECT COALESCE(SUM(issued*unit_cost),0) x FROM store_logs WHERE project_id=? AND date BETWEEN ? AND ?",(pid,start.isoformat(),end.isoformat())).fetchone()["x"]
    oe=c.execute("SELECT COALESCE(SUM(amount),0) x FROM finance_logs WHERE project_id=? AND kind='Expense' AND date BETWEEN ? AND ?",(pid,start.isoformat(),end.isoformat())).fetchone()["x"]
    c.close();return render_template("reports.html",pid=pid,rows=out,period=period,anchor=anchor,start=start,end=end,income=inc,expenses={"Machinery":me,"Manpower":pe,"Store":se,"Other":oe},total_expense=me+pe+se+oe)

@app.route("/admin/users")
@admin_required
def users():
    c=db();users=c.execute("SELECT * FROM users ORDER BY full_name").fetchall();projects=c.execute("SELECT * FROM projects ORDER BY name").fetchall();assign={u["id"]:[r["project_id"] for r in c.execute("SELECT project_id FROM user_projects WHERE user_id=?",(u["id"],)).fetchall()] for u in users};c.close();return render_template("users.html",users=users,projects=projects,assign=assign)

@app.route("/admin/users/add",methods=["POST"])
@admin_required
def add_user():
    c=db()
    try:
        role="SUPER_ADMIN" if request.form["role"]=="SUPER_ADMIN" else "STAFF";c.execute("INSERT INTO users(full_name,username,password_hash,department,location,role) VALUES(?,?,?,?,?,?)",(request.form["full_name"],request.form["username"],generate_password_hash(request.form["password"]),request.form["department"],request.form["location"],role));uid=c.execute("SELECT id FROM users WHERE username=?",(request.form["username"],)).fetchone()["id"]
        for pid in request.form.getlist("project_ids"):c.execute("INSERT OR IGNORE INTO user_projects(user_id,project_id) VALUES(?,?)",(uid,pid))
        c.commit();flash("👤 User created with project permissions.","success")
    except Exception as e:c.rollback();flash("Could not create user: "+str(e),"error")
    c.close();return redirect(url_for("users"))

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
        try:c.execute("INSERT INTO projects(name,code,location,client,consultant,status,start_date,end_date,contractor_role) VALUES(?,?,?,?,?,?,?,?,?)",(request.form["name"],request.form["code"],request.form["location"],request.form.get("client",""),request.form.get("consultant",""),request.form.get("status","Active"),request.form.get("start_date",""),request.form.get("end_date",""),request.form.get("contractor_role","Main Contractor")));c.commit();flash("🏗️ Project created.","success")
        except Exception as e:c.rollback();flash(str(e),"error")
    projects=c.execute("SELECT * FROM projects ORDER BY name").fetchall();c.close();return render_template("projects_admin.html",projects=projects)


def import_boq_xlsx(path,pid):
    wb=load_workbook(path,data_only=True,read_only=True)
    # Prefer BOQ sheet; otherwise locate a sheet with the expected header labels.
    candidates=[ws for ws in wb.worksheets if "boq" in ws.title.lower()]
    candidates += [ws for ws in wb.worksheets if ws not in candidates]
    target=None;header=None
    for ws in candidates:
        for row in ws.iter_rows(min_row=1,max_row=min(ws.max_row,100),values_only=True):
            vals=[str(v).strip().lower() if v is not None else "" for v in row]
            if any("item no" in v for v in vals) and any("description" in v for v in vals) and any("unit"==v or v.endswith("unit") for v in vals):
                target=ws;header=vals;break
        if target:break
    if not target:raise ValueError("Could not find a BOQ header row containing Item No, Description and Unit.")
    def col(names):
        for i,v in enumerate(header):
            if any(n in v for n in names):return i
        return None
    ci=col(["item no","item number","item"]);cd=col(["description","general activity"]);cu=col(["unit"]);cq=col(["total contract quantity","contract quantity","quantity"]);cr=col(["contract rate","unit rate","rate"])
    if None in (ci,cd,cu,cr):raise ValueError("BOQ needs item no, description, unit and rate columns.")
    c=db();count=0
    for row in target.iter_rows(min_row=target._current_row+1 if hasattr(target,'_current_row') else 2,values_only=True):
        if ci>=len(row) or cd>=len(row):continue
        item=row[ci];desc=row[cd]
        if item is None or desc is None:continue
        try:rate=parse_float(row[cr])
        except:rate=0
        qty=parse_float(row[cq]) if cq is not None and cq<len(row) else 0
        unit=str(row[cu]).strip() if cu<len(row) and row[cu] is not None else ""
        item=str(item).strip();desc=str(desc).strip()
        if not item or not desc:continue
        c.execute("INSERT INTO boq(project_id,item_no,description,unit,rate,contract_qty,source_sheet) VALUES(?,?,?,?,?,?,?) ON CONFLICT(project_id,item_no) DO UPDATE SET description=excluded.description,unit=excluded.unit,rate=excluded.rate,contract_qty=excluded.contract_qty,source_sheet=excluded.source_sheet",(pid,item,desc,unit,rate,qty,target.title));count+=1
    c.commit();c.close();return count,target.title

@app.route("/admin/boq/<int:pid>",methods=["GET","POST"])
@admin_required
def boq_admin(pid):
    c=db()
    if request.method=="POST":
        action=request.form.get("action")
        try:
            if action=="add":c.execute("INSERT INTO boq(project_id,item_no,description,unit,rate,contract_qty) VALUES(?,?,?,?,?,?)",(pid,request.form["item_no"],request.form["description"],request.form["unit"],parse_float(request.form["rate"]),parse_float(request.form["contract_qty"])))
            elif action=="upload":
                f=request.files.get("file")
                if not f or not f.filename.lower().endswith((".xlsx",".xlsm")):raise ValueError("Upload an .xlsx or .xlsm BOQ file.")
                name=secure_filename(f.filename);path=os.path.join(UPLOADS,name);f.save(path);count,sheet=import_boq_xlsx(path,pid);c=db();c.execute("INSERT INTO boq_uploads(project_id,filename,uploaded_at,user_id,rows_imported) VALUES(?,?,?,?,?)",(pid,name,dt.datetime.now().isoformat(timespec="seconds"),current_user()["id"],count));c.commit();flash(f"📑 BOQ imported: {count} items from sheet '{sheet}'.","success")
            c.commit()
        except Exception as e:c.rollback();flash("BOQ import error: "+str(e),"error")
    rows=c.execute("SELECT * FROM boq WHERE project_id=? ORDER BY item_no",(pid,)).fetchall();uploads=c.execute("SELECT * FROM boq_uploads WHERE project_id=? ORDER BY uploaded_at DESC LIMIT 10",(pid,)).fetchall();c.close();return render_template("boq.html",pid=pid,rows=rows,uploads=uploads)

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
