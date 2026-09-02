from pathlib import Path
p=Path('/mnt/data/v21src/app.py')
s=p.read_text()
# Remove per-connection WAL PRAGMA block, replace with timeout only.
old='''    c.execute("PRAGMA busy_timeout=30000")\n    c.execute("PRAGMA foreign_keys=ON")\n    try:\n        c.execute("PRAGMA journal_mode=WAL")\n        c.execute("PRAGMA synchronous=NORMAL")\n    except sqlite3.OperationalError:\n        # Keep working if an existing deployment temporarily cannot change journal mode.\n        pass\n'''
new='''    c.execute("PRAGMA busy_timeout=30000")\n    c.execute("PRAGMA foreign_keys=ON")\n'''
s=s.replace(old,new)
# Add WAL initialization after c=db in init_db
s=s.replace('''    c=db()\n    c.executescript('''','''    c=db()\n    # Set WAL once during database initialization, not on every request/connection.\n    # Re-running PRAGMA journal_mode=WAL on every connection can itself cause SQLite lock contention.\n    try:\n        c.execute("PRAGMA journal_mode=WAL")\n        c.execute("PRAGMA synchronous=NORMAL")\n    except sqlite3.OperationalError:\n        pass\n    c.executescript(''',1)
# Add unit columns to table create definitions
s=s.replace('CREATE TABLE IF NOT EXISTS daily_work(id INTEGER PRIMARY KEY,project_id INTEGER,date TEXT,boq_id INTEGER,quantity REAL,station_from TEXT,station_to TEXT,notes TEXT,user_id INTEGER);',
'''CREATE TABLE IF NOT EXISTS daily_work(id INTEGER PRIMARY KEY,project_id INTEGER,date TEXT,boq_id INTEGER,quantity REAL,unit TEXT,station_from TEXT,station_to TEXT,notes TEXT,user_id INTEGER);''')
s=s.replace('CREATE TABLE IF NOT EXISTS daily_activities(id INTEGER PRIMARY KEY,project_id INTEGER,date TEXT,boq_id INTEGER,work_type TEXT,executed_qty REAL DEFAULT 0,machine_id INTEGER,machine_hours REAL DEFAULT 0,manpower_position TEXT,manpower_qty REAL DEFAULT 0,manpower_hours REAL DEFAULT 0,material_id INTEGER,material_qty REAL DEFAULT 0,remarks TEXT,user_id INTEGER);',
'''CREATE TABLE IF NOT EXISTS daily_activities(id INTEGER PRIMARY KEY,project_id INTEGER,date TEXT,boq_id INTEGER,work_type TEXT,executed_qty REAL DEFAULT 0,unit TEXT,machine_id INTEGER,machine_hours REAL DEFAULT 0,manpower_position TEXT,manpower_qty REAL DEFAULT 0,manpower_hours REAL DEFAULT 0,material_id INTEGER,material_qty REAL DEFAULT 0,remarks TEXT,user_id INTEGER);''')
# Add migration for unit columns after existing daily tables are available.
needle="""    existing=[r['name'] for r in c.execute(\"PRAGMA table_info(machines)\").fetchall()]\n"""
insert="""    existing_dw=[r['name'] for r in c.execute(\"PRAGMA table_info(daily_work)\").fetchall()]\n    if 'unit' not in existing_dw: c.execute(\"ALTER TABLE daily_work ADD COLUMN unit TEXT\")\n    existing_da=[r['name'] for r in c.execute(\"PRAGMA table_info(daily_activities)\").fetchall()]\n    if 'unit' not in existing_da: c.execute(\"ALTER TABLE daily_activities ADD COLUMN unit TEXT\")\n"""
s=s.replace(needle,insert+needle,1)
# Add machine assignment routes before machinery route.
marker='@app.route("/projects/<int:pid>/machinery",methods=["GET","POST"])\n'
routes=r'''@app.route("/projects/<int:pid>/machinery/assign", methods=["POST"])
@login_required
def assign_machine(pid):
    if not allowed_project(pid) or not can_module("Machinery"):
        flash("🚫 Machinery access is not assigned.", "error")
        return redirect(url_for("project", pid=pid))
    c=db()
    try:
        mid=int(request.form.get("machine_id"))
        m=c.execute("SELECT * FROM machines WHERE id=? AND project_id=? AND active=1",(mid,pid)).fetchone()
        if not m: raise ValueError("Machine not found in this project fleet.")
        active=c.execute("SELECT id FROM machine_assignments WHERE machine_id=? AND project_id=? AND status='ACTIVE'",(mid,pid)).fetchone()
        if active: raise ValueError("This machine already has an active signed assignment. End it or use it until its signed hours are completed.")
        start_date=request.form.get("start_date") or dt.date.today().isoformat()
        start_hour=parse_float(request.form.get("start_hour"))
        total=parse_float(request.form.get("total_hours"))
        if total<=0: raise ValueError("Total Signed Hours must be greater than zero.")
        c.execute("INSERT INTO machine_assignments(machine_id,project_id,start_date,start_hour,total_signed_hours,hours_used,status,assigned_by,notes) VALUES(?,?,?,?,?,0,'ACTIVE',?,?)",(mid,pid,start_date,start_hour,total,current_user()["id"],request.form.get("notes","")))
        c.execute("UPDATE machines SET assignment_start_date=?,assignment_start_hour=?,assignment_end_date=NULL,assignment_end_hour=NULL,total_signed_hours=?,hours_used=0,lifecycle_status='ACTIVE',assignment_signed_by=?,assignment_ended_by=NULL,assignment_ended_at=NULL WHERE id=? AND project_id=?",(start_date,start_hour,total,current_user()["id"],mid,pid))
        c.commit(); flash("✍️ Machine assignment signed successfully. Daily hours will automatically close it when the signed total is reached.","success")
    except Exception as e:
        c.rollback(); flash("Could not sign machine assignment: "+str(e),"error")
    finally: c.close()
    return redirect(url_for("machinery",pid=pid))

@app.route("/projects/<int:pid>/machinery/end-assignment", methods=["POST"])
@login_required
def end_machine_assignment(pid):
    if not allowed_project(pid) or not can_module("Machinery"):
        flash("🚫 Machinery access is not assigned.", "error")
        return redirect(url_for("project", pid=pid))
    c=db()
    try:
        mid=int(request.form.get("machine_id"))
        a=c.execute("SELECT * FROM machine_assignments WHERE machine_id=? AND project_id=? AND status='ACTIVE' ORDER BY id DESC LIMIT 1",(mid,pid)).fetchone()
        if not a: raise ValueError("No active assignment exists for this machine.")
        end_date=request.form.get("end_date") or dt.date.today().isoformat()
        end_hour=parse_float(request.form.get("end_hour"))
        c.execute("UPDATE machine_assignments SET status='ENDED',end_date=?,end_hour=?,ended_by=?,ended_at=CURRENT_TIMESTAMP WHERE id=?",(end_date,end_hour,current_user()["id"],a["id"]))
        c.execute("UPDATE machines SET lifecycle_status='ENDED',assignment_end_date=?,assignment_end_hour=?,assignment_ended_by=?,assignment_ended_at=CURRENT_TIMESTAMP WHERE id=? AND project_id=?",(end_date,end_hour,current_user()["id"],mid,pid))
        c.commit(); flash("🛑 Assignment ended. A new signed assignment is required before the machine can be logged again.","success")
    except Exception as e:
        c.rollback(); flash("Could not end machine assignment: "+str(e),"error")
    finally: c.close()
    return redirect(url_for("machinery",pid=pid))

'''
s=s.replace(marker,routes+marker,1)
# Fix machinery function robust transaction and missing render variables not needed. Replace whole function section through manpower route.
start=s.index('@app.route("/projects/<int:pid>/machinery",methods=["GET","POST"])')
end=s.index('@app.route("/projects/<int:pid>/manpower",methods=["GET","POST"])',start)
newmach=r'''@app.route("/projects/<int:pid>/machinery",methods=["GET","POST"])
@login_required
def machinery(pid):
    if not allowed_project(pid) or not can_module("Machinery"):
        flash("🚫 Machinery access is not assigned.","error"); return redirect(url_for("project",pid=pid))
    c=db()
    try:
        if request.method=="POST":
            action=request.form.get("action")
            if action=="add":
                code=request.form.get("code","").strip()
                if not code: raise ValueError("Plate / Fleet Code is required.")
                c.execute("INSERT INTO machines(project_id,machine_type,code,plate_no,engine_no,ownership,hourly_rate,expected_fuel,fuel_price,lifecycle_status) VALUES(?,?,?,?,?,?,?,?,?,?)",(pid,request.form["machine_type"],code,request.form.get("plate_no",code),request.form.get("engine_no",""),request.form["ownership"],parse_float(request.form.get("hourly_rate")),parse_float(request.form.get("expected_fuel")),parse_float(request.form.get("fuel_price")),"UNASSIGNED"))
                c.commit(); flash("🚜 Machine added to this project's fleet.","success")
            elif action=="remove":
                c.execute("UPDATE machines SET active=0 WHERE id=? AND project_id=?",(request.form["machine_id"],pid)); c.commit(); flash("Machine removed from active fleet.","success")
            elif action=="log":
                mid=int(request.form['machine_id'])
                active_assignment=c.execute("SELECT * FROM machine_assignments WHERE machine_id=? AND project_id=? AND status='ACTIVE' ORDER BY id DESC LIMIT 1",(mid,pid)).fetchone()
                if not active_assignment: raise ValueError('Machine has no active signed assignment. Machinery Admin must sign a new start date/hour first.')
                work=parse_float(request.form.get("work_hours")); idle=parse_float(request.form.get("idle_hours")); down=parse_float(request.form.get("down_hours"))
                vals=(pid,mid,request.form["date"],work,idle,request.form.get("idle_reason",""),1 if request.form.get("idle_payable")=="1" else 0,down,request.form.get("down_reason",""),parse_float(request.form.get("opening_gauge")),parse_float(request.form.get("fuel_received")),parse_float(request.form.get("closing_gauge")),request.form.get("notes",""),current_user()["id"])
                c.execute("INSERT INTO machine_logs(project_id,machine_id,date,work_hours,idle_hours,idle_reason,idle_payable,down_hours,down_reason,opening_gauge,fuel_received,closing_gauge,notes,user_id) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",vals)
                if any(parse_float(request.form.get(k)) for k in ("opening_gauge","fuel_received","closing_gauge")):
                    mp=c.execute("SELECT fuel_price FROM machines WHERE id=?",(mid,)).fetchone()
                    c.execute("INSERT INTO fuel_logs(project_id,machine_id,date,opening_gauge,fuel_received,closing_gauge,fuel_price,reference,notes,user_id,source) VALUES(?,?,?,?,?,?,?,?,?,?,?)",(pid,mid,request.form['date'],parse_float(request.form.get('opening_gauge')),parse_float(request.form.get('fuel_received')),parse_float(request.form.get('closing_gauge')),mp['fuel_price'] if mp else 0,'MCH-'+request.form['date'],request.form.get('notes',''),current_user()['id'],'Machinery Log'))
                used=work+idle+down
                new_used=(active_assignment['hours_used'] or 0)+used
                reached=(active_assignment['total_signed_hours'] or 0)>0 and new_used >= active_assignment['total_signed_hours']
                if reached:
                    c.execute("UPDATE machine_assignments SET hours_used=?,status='ENDED',end_date=?,end_hour=?,ended_by=?,ended_at=CURRENT_TIMESTAMP WHERE id=?",(new_used,request.form['date'],parse_float(request.form.get('closing_gauge')),current_user()['id'],active_assignment['id']))
                    c.execute("UPDATE machines SET hours_used=?,lifecycle_status='ENDED',assignment_end_date=?,assignment_end_hour=?,assignment_ended_by=?,assignment_ended_at=CURRENT_TIMESTAMP WHERE id=? AND project_id=?",(new_used,request.form['date'],parse_float(request.form.get('closing_gauge')),current_user()['id'],mid,pid))
                    flash("⏱️ Machine log saved and signed assignment automatically ended because the total signed hours were reached.","success")
                else:
                    c.execute("UPDATE machine_assignments SET hours_used=? WHERE id=?",(new_used,active_assignment['id']))
                    c.execute("UPDATE machines SET hours_used=? WHERE id=? AND project_id=?",(new_used,mid,pid))
                    c.commit(); flash("⏱️ Machine hours / idle / down / gauge saved successfully.","success")
                    return redirect(url_for("machinery",pid=pid))
                c.commit()
    except Exception as e:
        try: c.rollback()
        except Exception: pass
        flash("Machinery save failed: "+str(e),"error")
    finally:
        c.close()
    c=db()
    try:
        machines=c.execute("SELECT * FROM machines WHERE project_id=? AND active=1 ORDER BY machine_type,code",(pid,)).fetchall()
        assignments=c.execute("SELECT ma.*,m.machine_type,m.code,m.plate_no FROM machine_assignments ma JOIN machines m ON m.id=ma.machine_id WHERE ma.project_id=? ORDER BY ma.id DESC LIMIT 100",(pid,)).fetchall()
        logs=c.execute("SELECT ml.*,m.machine_type,m.code,m.ownership,m.hourly_rate,m.expected_fuel,((ml.work_hours + CASE WHEN ml.idle_payable=1 THEN ml.idle_hours ELSE 0 END)*m.hourly_rate) expense,(ml.opening_gauge+ml.fuel_received-ml.closing_gauge) actual_fuel,(ml.work_hours*m.expected_fuel) expected_fuel_qty,CASE WHEN (ml.work_hours+ml.idle_hours+ml.down_hours)>0 THEN ml.work_hours*100.0/(ml.work_hours+ml.idle_hours+ml.down_hours) ELSE 0 END utilization,CASE WHEN (ml.work_hours+ml.idle_hours+ml.down_hours)>0 THEN (ml.work_hours+ml.idle_hours)*100.0/(ml.work_hours+ml.idle_hours+ml.down_hours) ELSE 0 END availability,(ml.opening_gauge+ml.fuel_received-ml.closing_gauge)-(ml.work_hours*m.expected_fuel) fuel_discrepancy FROM machine_logs ml JOIN machines m ON m.id=ml.machine_id WHERE ml.project_id=? ORDER BY ml.date DESC,ml.id DESC LIMIT 50",(pid,)).fetchall()
        return render_template("machinery.html",pid=pid,machines=machines,assignments=assignments,logs=logs)
    finally: c.close()

'''
s=s[:start]+newmach+s[end:]
# Fix daily inserts and allocation parsing. Replace relevant snippets.
s=s.replace('''            machine_data=request.form.getlist("wp_machine_json[]")\n            manpower_data=request.form.getlist("wp_manpower_json[]")\n            crew_data=request.form.getlist("wp_crew_json[]")\n            store_data=request.form.getlist("wp_store_json[]")\n            fuel_data=request.form.getlist("wp_fuel_json[]")\n            finance_data=request.form.getlist("wp_finance_json[]")\n''','''            machine_data=request.form.getlist("wp_machine_json[]")\n            manpower_data=request.form.getlist("wp_manpower_json[]")\n            crew_data=request.form.getlist("wp_crew_json[]")\n            store_data=request.form.getlist("wp_store_json[]")\n            fuel_data=request.form.getlist("wp_fuel_json[]")\n            finance_data=request.form.getlist("wp_finance_json[]")\n            units=request.form.getlist("wp_unit[]")\n''')
s=s.replace('''            def json_ids(items,i):\n                raw=arr(items,i,"[]")\n                try: value=_json.loads(raw) if raw else []\n                except Exception: value=[]\n                return [int(x) for x in value if str(x).strip().isdigit()]\n''','''            def json_alloc(items,i):\n                raw=arr(items,i,"[]")\n                try: value=_json.loads(raw) if raw else []\n                except Exception: value=[]\n                out=[]\n                if isinstance(value,dict):\n                    for k,v in value.items():\n                        if str(k).strip().isdigit():\n                            if isinstance(v,dict): out.append((int(k),v))\n                            else: out.append((int(k),{"qty":parse_float(v)}))\n                elif isinstance(value,list):\n                    for x in value:\n                        if isinstance(x,dict) and str(x.get("id","")).isdigit(): out.append((int(x["id"]),x))\n                        elif str(x).strip().isdigit(): out.append((int(x),{}))\n                return out\n''')
s=s.replace('''                bid=int(bid_s); qty=parse_float(arr(quantities,i))\n                if qty<=0: continue\n                # Variation warning is raised only when cumulative executed quantity exceeds BOQ quantity.\n                msg=variation_check(c,pid,bid,qty,d)\n                c.execute("INSERT INTO daily_work(project_id,date,boq_id,quantity,station_from,station_to,notes,user_id) VALUES(?,?,?,?,?,?,?,?)",(pid,d,bid,qty,arr(station_froms,i),arr(station_tos,i),arr(remarks,i),u["id"]))\n                c.execute("INSERT INTO daily_activities(project_id,date,boq_id,work_type,executed_qty,machine_id,machine_hours,manpower_position,manpower_qty,manpower_hours,material_id,material_qty,remarks,user_id) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(pid,d,bid,arr(work_types,i),qty,None,0,"",0,0,None,0,arr(remarks,i),u["id"]))\n''','''                bid=int(bid_s); qty=parse_float(arr(quantities,i))\n                if qty<=0: continue\n                b=c.execute("SELECT unit FROM boq WHERE id=? AND project_id=?",(bid,pid)).fetchone()\n                selected_unit=arr(units,i,(b["unit"] if b else "")) or (b["unit"] if b else "")\n                # Variation warning is raised only when cumulative executed quantity exceeds BOQ quantity.\n                msg=variation_check(c,pid,bid,qty,d)\n                c.execute("INSERT INTO daily_work(project_id,date,boq_id,quantity,unit,station_from,station_to,notes,user_id) VALUES(?,?,?,?,?,?,?,?,?)",(pid,d,bid,qty,selected_unit,arr(station_froms,i),arr(station_tos,i),arr(remarks,i),u["id"]))\n                c.execute("INSERT INTO daily_activities(project_id,date,boq_id,work_type,executed_qty,unit,machine_id,machine_hours,manpower_position,manpower_qty,manpower_hours,material_id,material_qty,remarks,user_id) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(pid,d,bid,arr(work_types,i),qty,selected_unit,None,0,"",0,0,None,0,arr(remarks,i),u["id"]))\n''')
# Replace loops with allocation-aware.
s=s.replace('''                for lid in json_ids(machine_data,i):\n                    c.execute("INSERT INTO activity_machines(activity_id,machine_log_id,machine_id,hours) SELECT ?,id,machine_id,work_hours+idle_hours+down_hours FROM machine_logs WHERE id=? AND project_id=?",(aid,lid,pid))\n                for mid in json_ids(manpower_data,i):\n                    c.execute("INSERT INTO activity_manpower(activity_id,manpower_id,crew_id,qty,hours) SELECT ?,id,crew_id,present,working_hours FROM manpower WHERE id=? AND project_id=?",(aid,mid,pid))\n                for cid in json_ids(crew_data,i):\n                    c.execute("INSERT INTO crew_evaluations(activity_id,crew_id,evaluation,remarks,score) SELECT ?,id,?,?,? FROM project_crews WHERE id=? AND project_id=?",(aid,arr(evaluations,i),arr(eval_remarks,i),parse_float(arr(scores,i)),cid,pid))\n                for sid in json_ids(store_data,i):\n                    c.execute("INSERT INTO activity_store(activity_id,store_log_id,material_id,qty) SELECT ?,id,material_id,issued FROM store_logs WHERE id=? AND project_id=?",(aid,sid,pid))\n                for fid in json_ids(fuel_data,i):\n                    c.execute("INSERT INTO activity_fuel(activity_id,fuel_log_id,litres) SELECT ?,id,opening_gauge+fuel_received-closing_gauge FROM fuel_logs WHERE id=? AND project_id=?",(aid,fid,pid))\n                for xid in json_ids(finance_data,i):\n                    c.execute("INSERT INTO activity_finance(activity_id,finance_log_id,amount) SELECT ?,id,amount FROM finance_logs WHERE id=? AND project_id=?",(aid,xid,pid))\n''','''                for lid,meta in json_alloc(machine_data,i):\n                    hours=parse_float(meta.get("hours")) if isinstance(meta,dict) and meta.get("hours") not in (None,"") else 0\n                    c.execute("INSERT INTO activity_machines(activity_id,machine_log_id,machine_id,hours) SELECT ?,id,machine_id,? FROM machine_logs WHERE id=? AND project_id=?",(aid,hours,lid,pid))\n                for mid,meta in json_alloc(manpower_data,i):\n                    q=parse_float(meta.get("qty")) if isinstance(meta,dict) else 0; hrs=parse_float(meta.get("hours")) if isinstance(meta,dict) else 0\n                    c.execute("INSERT INTO activity_manpower(activity_id,manpower_id,crew_id,qty,hours) SELECT ?,id,crew_id,?,? FROM manpower WHERE id=? AND project_id=?",(aid,q,hrs,mid,pid))\n                for cid,meta in json_alloc(crew_data,i):\n                    c.execute("INSERT INTO crew_evaluations(activity_id,crew_id,evaluation,remarks,score) SELECT ?,id,?,?,? FROM project_crews WHERE id=? AND project_id=?",(aid,arr(evaluations,i),arr(eval_remarks,i),parse_float(arr(scores,i)),cid,pid))\n                for sid,meta in json_alloc(store_data,i):\n                    q=parse_float(meta.get("qty")) if isinstance(meta,dict) else 0\n                    c.execute("INSERT INTO activity_store(activity_id,store_log_id,material_id,qty) SELECT ?,id,material_id,? FROM store_logs WHERE id=? AND project_id=?",(aid,q,sid,pid))\n                for fid,meta in json_alloc(fuel_data,i):\n                    q=parse_float(meta.get("litres")) if isinstance(meta,dict) else 0\n                    c.execute("INSERT INTO activity_fuel(activity_id,fuel_log_id,litres) SELECT ?,id,? FROM fuel_logs WHERE id=? AND project_id=?",(aid,q,fid,pid))\n                for xid,meta in json_alloc(finance_data,i):\n                    q=parse_float(meta.get("amount")) if isinstance(meta,dict) else 0\n                    c.execute("INSERT INTO activity_finance(activity_id,finance_log_id,amount) SELECT ?,id,? FROM finance_logs WHERE id=? AND project_id=?",(aid,q,xid,pid))\n''')
# Remove unused first list gets harmless but add report unit display later.
p.write_text(s)
