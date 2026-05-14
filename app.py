#!/usr/bin/env python3
"""
MAG Weekly Tasks Web App
Multi-user task tracker with natural language input + weekly Excel reports
"""
import os, re, io, sqlite3
from datetime import datetime, date, timedelta
from flask import (Flask, render_template, request, redirect,
                   url_for, session, send_file, flash, jsonify)
from werkzeug.security import generate_password_hash, check_password_hash
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.fills import PatternFill as PF

app = Flask(__name__)
app.secret_key = 'mag-weekly-tasks-2026-secret'
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(days=30)  # stay logged in for 30 days

# ── DB path: prefer ~/Library/Application Support/ (no TCC issues on macOS),
#    fall back to app directory, then /tmp as last resort ──────────────────────
def _resolve_db_path():
    candidates = []
    # 1. macOS Application Support — safest, no TCC restriction
    if os.uname().sysname == 'Darwin':
        app_support = os.path.expanduser('~/Library/Application Support/MAGWeeklyTasks')
        candidates.append(os.path.join(app_support, 'tasks.db'))
    # 2. Same directory as app.py
    candidates.append(os.path.join(os.path.abspath(os.path.dirname(__file__)), 'tasks.db'))
    # 3. /tmp fallback (works everywhere, not persistent across reboots but fine for demo)
    candidates.append('/tmp/mag_weeklytasks.db')

    for path in candidates:
        folder = os.path.dirname(path)
        try:
            os.makedirs(folder, exist_ok=True)
            # Quick write test
            test = path + '.writetest'
            with open(test, 'w') as f:
                f.write('ok')
            os.remove(test)
            print(f"  DB location: {path}")
            return path
        except Exception:
            continue
    raise RuntimeError("No writable location found for the database.")

DB = _resolve_db_path()

# ── Database ──────────────────────────────────────────────────────────────────

def db():
    c = sqlite3.connect(DB, check_same_thread=False)
    c.row_factory = sqlite3.Row
    return c

def init_db():
    conn = db()
    try:
        conn.execute('''CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            full_name TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )''')
        conn.execute('''CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            task_id TEXT,
            type TEXT DEFAULT 'General & Admin',
            description TEXT,
            requested_by TEXT,
            assignee TEXT,
            creation_date DATE,
            resolved_on DATE,
            priority TEXT DEFAULT 'Low',
            status TEXT DEFAULT 'Completed',
            hours_spent REAL DEFAULT 0,
            notes TEXT,
            raw_input TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id)
        )''')
        conn.commit()
        # Auto-create default account if no users exist (survives DB resets)
        existing = conn.execute('SELECT COUNT(*) FROM users').fetchone()[0]
        if existing == 0:
            from werkzeug.security import generate_password_hash as _gph
            conn.execute(
                'INSERT INTO users (username, password_hash, full_name) VALUES (?,?,?)',
                ('ashokdy', _gph('Mag2026', method='pbkdf2:sha256'), 'Ashok Devangam Yerra')
            )
            conn.commit()
            print("  Default user created: ashokdy / Mag2026")
        # Migration: fix full_name for ashokdy if it was stored as a short/incorrect value
        row = conn.execute("SELECT id, full_name FROM users WHERE username='ashokdy'").fetchone()
        if row and row['full_name'] != 'Ashok Devangam Yerra':
            old_name = row['full_name']
            conn.execute("UPDATE users SET full_name='Ashok Devangam Yerra' WHERE username='ashokdy'")
            conn.execute(
                "UPDATE tasks SET requested_by='Ashok Devangam Yerra' WHERE user_id=? AND requested_by=?",
                (row['id'], old_name)
            )
            conn.execute(
                "UPDATE tasks SET assignee='Ashok Devangam Yerra' WHERE user_id=? AND assignee=?",
                (row['id'], old_name)
            )
            conn.commit()
            print(f"  Migrated full_name from '{old_name}' to 'Ashok Devangam Yerra'")
        print(f"  Database ready: {DB}")
    finally:
        conn.close()

# Always init DB when module loads (covers both direct run and service launch)
init_db()

# ── Natural Language Parser ───────────────────────────────────────────────────

MONTHS = {
    'jan':1,'january':1,'feb':2,'february':2,'mar':3,'march':3,
    'apr':4,'april':4,'may':5,'jun':6,'june':6,'jul':7,'july':7,
    'aug':8,'august':8,'sep':9,'september':9,'oct':10,'october':10,
    'nov':11,'november':11,'dec':12,'december':12
}

def group_task_lines(raw):
    """
    Handles both input formats:
      Multi-line: Task ID alone on one line, description on the next line(s)
      Single-line: Everything on one line per task
    Returns a list of single-line strings, one per task, ready for parse_input().
    """
    tid_start = re.compile(r'^([A-Za-z]{2,6}\d+|Adhoc)\b', re.I)
    lines = [l.strip() for l in raw.strip().splitlines() if l.strip()]
    groups, current = [], []
    for line in lines:
        if tid_start.match(line):
            if current:
                groups.append(' '.join(current))
            current = [line]
        else:
            current.append(line)
    if current:
        groups.append(' '.join(current))
    return groups if groups else lines

def parse_date_str(s):
    s = s.strip().lower()
    yr = datetime.now().year
    m = re.match(r'^(\d{1,2})\s+(\w+)(?:\s+(\d{4}))?$', s)
    if m:
        day, mon, y = m.groups()
        if mon in MONTHS:
            return date(int(y or yr), MONTHS[mon], int(day))
    m = re.match(r'^(\w+)\s+(\d{1,2})(?:\s+(\d{4}))?$', s)
    if m:
        mon, day, y = m.groups()
        if mon in MONTHS:
            return date(int(y or yr), MONTHS[mon], int(day))
    return None

def extract_date_hours(text):
    """Return list of (date, hours) pairs and cleaned description text."""
    tl = text.lower()
    mon_re = '|'.join(sorted(MONTHS, key=len, reverse=True))
    date_re = rf'(?:(\d{{1,2}})\s+(?:{mon_re})|(?:{mon_re})\s+(\d{{1,2}}))'
    pairs, spans = [], []

    # Pattern A: "on DATE for N [hours]"
    pA = rf'on\s+({date_re})\s+for\s+(\d+(?:\.\d+)?)\s*(?:hours?)?'
    for m in re.finditer(pA, tl):
        d = parse_date_str(m.group(1))
        h = float(m.group(m.lastindex))
        if d:
            pairs.append((d, h, m.span()))
            spans.append(m.span())

    # Pattern B: "N [hours] on DATE"
    pB = rf'(\d+(?:\.\d+)?)\s*(?:hours?)?\s+on\s+({date_re})'
    for m in re.finditer(pB, tl):
        if any(s[0] <= m.start() <= s[1] for s in spans):
            continue
        d = parse_date_str(m.group(2))
        h = float(m.group(1))
        if d:
            pairs.append((d, h, m.span()))
            spans.append(m.span())

    # Pattern C: "DATE for N [hours]"  (no "on" prefix)
    pC = rf'({date_re})\s+for\s+(\d+(?:\.\d+)?)\s*(?:hours?)?'
    for m in re.finditer(pC, tl):
        if any(s[0] <= m.start() <= s[1] for s in spans):
            continue
        d = parse_date_str(m.group(1))
        h = float(m.group(m.lastindex))
        if d:
            pairs.append((d, h, m.span()))
            spans.append(m.span())

    pairs.sort(key=lambda x: x[0])

    # Strip matched spans from text
    clean = text
    for span in sorted([p[2] for p in pairs], reverse=True):
        clean = clean[:span[0]] + ' ' + clean[span[1]:]

    return [(d, h) for d, h, _ in pairs], re.sub(r'\s+', ' ', clean).strip()

def parse_input(raw, full_name, default_date=None):
    """Parse natural language into list of task dicts (one per date)."""
    text = raw.strip()

    # Task ID
    tid_m = re.match(r'^([A-Za-z]{1,4}\d+|Adhoc|adhoc|ADHOC)\s*', text, re.I)
    task_id = None
    if tid_m:
        task_id = 'Adhoc' if tid_m.group(1).lower() == 'adhoc' else tid_m.group(1).upper()
        text = text[tid_m.end():]

    # Type
    type_val = 'General & Admin'
    tm = re.search(r'\btype\s*[-:]\s*([A-Za-z&/ ]+?)(?=\s{2,}|\s+\w+\s+\w|\s*$)', text, re.I)
    if tm:
        type_val = tm.group(1).strip()
        text = text[:tm.start()] + text[tm.end():]
    elif re.search(r'\bR&D\b|\bRnD\b', text, re.I):
        type_val = 'R&D'
        text = re.sub(r'\bR&D\b|\bRnD\b', '', text, flags=re.I)

    # Status
    status = 'Completed'
    sm = re.search(r'\b(in\s+progress|completed?|pending|on\s+hold)\b', text, re.I)
    if sm:
        sl = sm.group(1).lower()
        if 'progress' in sl: status = 'In Progress'
        elif 'complet' in sl: status = 'Completed'
        elif 'pending' in sl: status = 'Pending'
        elif 'hold' in sl: status = 'On Hold'
        text = text[:sm.start()] + text[sm.end():]

    # Priority
    priority = 'Low'
    pm = re.search(r'\b(low|medium|high|critical)\s*priority\b', text, re.I)
    if pm:
        priority = pm.group(1).capitalize()
        text = text[:pm.start()] + text[pm.end():]

    # Notes / comments
    notes = None
    nm = re.search(r'(?:with\s+)?(?:comments?|notes?)\s+(?:as\s+)?["\']([^"\']+)["\']', text, re.I)
    if not nm:
        nm = re.search(r'(?:comments?|notes?)\s*[:\-]\s*(.+?)(?=\s+(?:on|and)\s+\d|\s*$)', text, re.I)
    if nm:
        notes = nm.group(1).strip().strip('"\'')
        text = text[:nm.start()] + text[nm.end():]

    # Date-hour pairs
    date_hours, remaining = extract_date_hours(text)

    # Clean description — strip any number of trailing/leading connector words
    desc = re.sub(r'(\s*\b(and|with|for|on)\b)+\s*$', '', remaining, flags=re.I)
    desc = re.sub(r'^\s*(\b(and|with|for|on)\b\s*)+', '', desc, flags=re.I)
    desc = re.sub(r'\s+', ' ', desc).strip(' ,-')

    if not date_hours:
        date_hours = [(default_date or date.today(), 0)]

    # Build one entry per date, sorted by creation_date
    entries = []
    for i, (cdate, hours) in enumerate(date_hours):
        resolved = cdate + timedelta(days=1)
        entries.append({
            'task_id':      task_id or 'Adhoc',
            'type':         type_val,
            'description':  desc,
            'requested_by': full_name,
            'assignee':     full_name,
            'creation_date': cdate.isoformat(),
            'resolved_on':  resolved.isoformat(),
            'priority':     priority,
            'status':       status,
            'hours_spent':  hours,
            'notes':        notes if i == len(date_hours) - 1 else None,
            'raw_input':    raw,
        })
    return entries

# ── Week helpers ──────────────────────────────────────────────────────────────

def week_bounds(for_date=None):
    """Return (monday, friday) of the week containing for_date."""
    d = for_date or date.today()
    mon = d - timedelta(days=d.weekday())
    fri = mon + timedelta(days=4)
    return mon, fri

def friday_of(d=None):
    d = d or date.today()
    return d + timedelta(days=(4 - d.weekday()) % 7)

# ── Excel report ──────────────────────────────────────────────────────────────

COLS = ['Task ID','Type','Description','Requested By','Assignee',
        'Creation Date','Resolved On','Priority','Status','Hours Spent','CR Number / Notes']
COL_W = [10, 16, 55, 22, 22, 14, 14, 10, 15, 13, 45]

def make_report(week_fri):
    """Generate Excel workbook with one sheet per user for the given week."""
    mon, fri = week_bounds(week_fri)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    hdr_font  = Font(bold=True, size=11, name='Calibri')
    hdr_fill  = PF(fill_type='solid', fgColor='2E75B6')
    hdr_font2 = Font(bold=True, size=11, name='Calibri', color='FFFFFF')
    center    = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin      = Side(style='thin', color='BFBFBF')
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)
    alt_fill  = PF(fill_type='solid', fgColor='EBF3FB')

    conn = db()
    users = conn.execute('SELECT * FROM users ORDER BY full_name').fetchall()
    for u in users:
        rows = conn.execute('''
            SELECT * FROM tasks
            WHERE user_id=? AND creation_date BETWEEN ? AND ?
            ORDER BY creation_date ASC, id ASC
        ''', (u['id'], mon.isoformat(), fri.isoformat())).fetchall()

        if not rows:
            continue

        sheet_name = u['full_name'][:31]
        ws = wb.create_sheet(title=sheet_name)

        # Header
        ws.append(COLS)
        for ci, cell in enumerate(ws[1], 1):
            cell.font   = hdr_font2
            cell.fill   = hdr_fill
            cell.alignment = center
            cell.border = border
            ws.column_dimensions[cell.column_letter].width = COL_W[ci-1]
        ws.row_dimensions[1].height = 22

        # Data rows
        total_hours = 0
        for ri, row in enumerate(rows, 2):
            ws.append([
                row['task_id'], row['type'], row['description'],
                row['requested_by'], row['assignee'],
                row['creation_date'], row['resolved_on'],
                row['priority'], row['status'],
                row['hours_spent'], row['notes']
            ])
            for ci, cell in enumerate(ws[ri], 1):
                cell.border = border
                cell.alignment = Alignment(vertical='center', wrap_text=(ci==3))
                if ri % 2 == 0:
                    cell.fill = alt_fill
                if ci in (6, 7) and cell.value:
                    cell.number_format = 'DD/MM/YYYY'
            ws.row_dimensions[ri].height = 18
            total_hours += row['hours_spent'] or 0

        # Totals row
        tr = len(rows) + 2
        ws.cell(tr, 9, 'TOTAL HOURS').font  = Font(bold=True)
        ws.cell(tr, 9).fill   = PF(fill_type='solid', fgColor='D9E1F2')
        ws.cell(tr, 10, total_hours).font = Font(bold=True)
        ws.cell(tr, 10).fill  = PF(fill_type='solid', fgColor='D9E1F2')
        ws.freeze_panes = 'A2'

    conn.close()
    if not wb.sheetnames:
        ws = wb.create_sheet('No Data')
        ws['A1'] = 'No tasks recorded for this week.'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Routes ────────────────────────────────────────────────────────────────────

def login_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect('/login')
        # Always refresh full_name from DB so stale session cookies pick up name changes
        conn = db()
        row = conn.execute('SELECT full_name FROM users WHERE id=?', (session['user_id'],)).fetchone()
        conn.close()
        if row:
            session['full_name'] = row['full_name']
        return f(*args, **kwargs)
    return decorated

@app.route('/')
@login_required
def dashboard():
    # Build list of available weeks (any week that has tasks + current week)
    conn = db()
    rows = conn.execute('''
        SELECT DISTINCT creation_date FROM tasks WHERE user_id=?
        ORDER BY creation_date DESC
    ''', (session['user_id'],)).fetchall()
    conn.close()

    seen, weeks = set(), []
    for r in rows:
        d = date.fromisoformat(r['creation_date'])
        fri = friday_of(d).isoformat()
        if fri not in seen:
            seen.add(fri)
            weeks.append(fri)
    # Always include current week and last week
    cur_fri  = friday_of().isoformat()
    last_fri = friday_of(date.today() - timedelta(days=7)).isoformat()
    if cur_fri not in seen:
        weeks.insert(0, cur_fri)
        seen.add(cur_fri)
    if last_fri not in seen:
        # Insert right after current week
        idx = weeks.index(cur_fri) + 1
        weeks.insert(idx, last_fri)
        seen.add(last_fri)

    # Selected week — default to the week with the most recent task data
    sel_week = request.args.get('week', weeks[0])
    sel_fri = date.fromisoformat(sel_week)
    sel_mon, _ = week_bounds(sel_fri)

    conn = db()
    tasks = conn.execute('''
        SELECT * FROM tasks WHERE user_id=?
        AND creation_date BETWEEN ? AND ?
        ORDER BY creation_date ASC, id ASC
    ''', (session['user_id'], sel_mon.isoformat(), sel_fri.isoformat())).fetchall()
    conn.close()
    total = sum(t['hours_spent'] or 0 for t in tasks)
    return render_template('dashboard.html',
        tasks=tasks, mon=sel_mon, fri=sel_fri,
        total_hours=round(total, 1), weeks=weeks, sel_week=sel_week)

@app.route('/add', methods=['POST'])
@login_required
def add_task():
    raw = request.form.get('raw', '').strip()
    if not raw:
        flash('Please enter task details.', 'error')
        return redirect('/')

    # Use the week the user is currently viewing as the default date for undated entries
    try:
        default_date = date.fromisoformat(request.form.get('default_date', ''))
    except ValueError:
        default_date = date.today()

    added = 0
    skipped = 0
    errors = []
    all_dates = []  # collect all dates parsed from input to redirect correctly

    for line in group_task_lines(raw):
        line = line.strip()
        if not line:
            continue
        try:
            entries = parse_input(line, session['full_name'], default_date=default_date)
            conn = db()
            for e in entries:
                all_dates.append(date.fromisoformat(e['creation_date']))
                conn.execute('''
                    INSERT INTO tasks
                    (user_id,task_id,type,description,requested_by,assignee,
                     creation_date,resolved_on,priority,status,hours_spent,notes,raw_input)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                ''', (session['user_id'], e['task_id'], e['type'], e['description'],
                      e['requested_by'], e['assignee'], e['creation_date'], e['resolved_on'],
                      e['priority'], e['status'], e['hours_spent'], e['notes'], e['raw_input']))
                added += 1
            conn.commit()
            conn.close()
        except Exception as ex:
            errors.append(f"Could not parse: {line[:60]}… ({ex})")

    if added:
        flash(f'✓ Added {added} task entr{"y" if added==1 else "ies"} successfully.', 'success')
    for err in errors:
        flash(err, 'error')

    # Always redirect to the week the input data belongs to (even if all were duplicates)
    if all_dates:
        target_fri = friday_of(max(all_dates)).isoformat()
        return redirect(f'/?week={target_fri}')
    return redirect('/')

@app.route('/delete/<int:task_id>')
@login_required
def delete_task(task_id):
    conn = db()
    conn.execute('DELETE FROM tasks WHERE id=? AND user_id=?', (task_id, session['user_id']))
    conn.commit()
    conn.close()
    flash('Entry deleted.', 'success')
    return redirect(request.referrer or '/')

@app.route('/clone/<int:task_id>')
@login_required
def clone_task(task_id):
    conn = db()
    row = conn.execute('SELECT * FROM tasks WHERE id=? AND user_id=?',
                       (task_id, session['user_id'])).fetchone()
    if row:
        conn.execute('''
            INSERT INTO tasks (user_id,task_id,type,description,requested_by,assignee,
                               creation_date,resolved_on,priority,status,hours_spent,notes,raw_input)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
        ''', (session['user_id'], row['task_id'], row['type'], row['description'],
              row['requested_by'], row['assignee'], row['creation_date'], row['resolved_on'],
              row['priority'], row['status'], row['hours_spent'], row['notes'], row['raw_input']))
        conn.commit()
        flash('✓ Row cloned — original kept, copy added below.', 'success')
        # Redirect to the exact week so both original and clone are visible
        week = request.args.get('week') or friday_of(
            date.fromisoformat(row['creation_date'])).isoformat()
    else:
        flash('Row not found.', 'error')
        week = request.args.get('week', friday_of().isoformat())
    conn.close()
    return redirect(f'/?week={week}')

@app.route('/all-tasks')
@login_required
def all_tasks():
    conn = db()
    rows = conn.execute('''
        SELECT DISTINCT creation_date FROM tasks WHERE user_id=?
        ORDER BY creation_date DESC
    ''', (session['user_id'],)).fetchall()
    conn.close()

    # Build list of unique friday dates for weeks that have tasks
    seen, weeks = set(), []
    for r in rows:
        d = date.fromisoformat(r['creation_date'])
        fri = friday_of(d).isoformat()
        if fri not in seen:
            seen.add(fri)
            weeks.append(fri)

    if not weeks:
        weeks = [friday_of().isoformat()]

    sel_week = request.args.get('week', weeks[0])
    sel_fri = date.fromisoformat(sel_week)
    sel_mon, _ = week_bounds(sel_fri)

    conn = db()
    tasks = conn.execute('''
        SELECT * FROM tasks WHERE user_id=?
        AND creation_date BETWEEN ? AND ?
        ORDER BY creation_date ASC, id ASC
    ''', (session['user_id'], sel_mon.isoformat(), sel_fri.isoformat())).fetchall()
    conn.close()
    total = sum(t['hours_spent'] or 0 for t in tasks)
    return render_template('all_tasks.html',
        tasks=tasks, weeks=weeks, sel_week=sel_week, total_hours=round(total, 1))

@app.route('/report')
@login_required
def report():
    week_str = request.args.get('week', friday_of().isoformat())
    week_fri = date.fromisoformat(week_str)
    buf = make_report(week_fri)
    fname = f"MAG_WeeklyTasks_{week_fri.isoformat()}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        u = request.form.get('username', '').strip()
        p = request.form.get('password', '')
        conn = db()
        user = conn.execute('SELECT * FROM users WHERE username=?', (u,)).fetchone()
        conn.close()
        if user and check_password_hash(user['password_hash'], p):
            session.permanent   = True   # keep logged in for 30 days
            session['user_id']   = user['id']
            session['full_name'] = user['full_name']
            session['username']  = user['username']
            return redirect('/')
        flash('Invalid username or password.', 'error')
    return render_template('login.html')

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        fn = request.form.get('full_name', '').strip()
        u  = request.form.get('username', '').strip()
        p  = request.form.get('password', '')
        if not fn or not u or not p:
            flash('All fields are required.', 'error')
            return render_template('register.html')
        try:
            conn = db()
            conn.execute(
                'INSERT INTO users (username, password_hash, full_name) VALUES (?,?,?)',
                (u, generate_password_hash(p, method='pbkdf2:sha256'), fn)
            )
            conn.commit()
            conn.close()
            flash('Account created! Please sign in.', 'success')
            return redirect('/login')
        except sqlite3.IntegrityError:
            flash('Username already taken. Try a different username.', 'error')
        except Exception as e:
            app.logger.error(f"Register error: {e}", exc_info=True)
            flash(f'Registration failed: {e}', 'error')
    return render_template('register.html')

@app.route('/clear-week/<week_fri>')
@login_required
def clear_week(week_fri):
    fri = date.fromisoformat(week_fri)
    mon, _ = week_bounds(fri)
    conn = db()
    conn.execute('''
        DELETE FROM tasks
        WHERE user_id=? AND creation_date BETWEEN ? AND ?
    ''', (session['user_id'], mon.isoformat(), fri.isoformat()))
    conn.commit()
    conn.close()
    flash(f'✓ All your entries for week {mon.strftime("%d %b")}–{fri.strftime("%d %b")} cleared.', 'success')
    return redirect(f'/?week={week_fri}')

@app.route('/update/<int:task_id>', methods=['POST'])
@login_required
def update_task(task_id):
    field = request.form.get('field')
    value = request.form.get('value', '').strip()
    allowed = {'hours_spent', 'notes', 'task_id', 'type', 'description',
               'status', 'creation_date', 'resolved_on'}
    if field not in allowed:
        return jsonify(ok=False, error='Invalid field'), 400
    if field == 'hours_spent':
        try:
            value = float(value)
            if value < 0 or value > 24:
                return jsonify(ok=False, error='Hours must be 0–24'), 400
        except ValueError:
            return jsonify(ok=False, error='Not a valid number'), 400
    elif field in ('creation_date', 'resolved_on'):
        try:
            date.fromisoformat(value)
        except ValueError:
            return jsonify(ok=False, error='Invalid date'), 400
    elif field == 'status':
        if value not in ('Completed', 'In Progress', 'Pending', 'On Hold'):
            return jsonify(ok=False, error='Invalid status'), 400
    else:
        value = value.strip() if value and value.strip() else None
    conn = db()
    conn.execute(f'UPDATE tasks SET {field}=? WHERE id=? AND user_id=?',
                 (value, task_id, session['user_id']))
    conn.commit()
    conn.close()
    return jsonify(ok=True)

@app.route('/add-row', methods=['POST'])
@login_required
def add_row():
    """Save a row entered directly in the inline table editor."""
    full_name = session['full_name']
    task_id   = request.form.get('task_id', '').strip().upper() or 'Adhoc'
    type_val  = request.form.get('type', 'General & Admin').strip() or 'General & Admin'
    desc      = request.form.get('description', '').strip()
    creation_date = request.form.get('creation_date', date.today().isoformat())
    resolved_on   = request.form.get('resolved_on', '').strip()
    priority  = request.form.get('priority', 'Low').strip()
    status    = request.form.get('status', 'Completed').strip()
    notes     = request.form.get('notes', '').strip() or None
    try:
        hours = float(request.form.get('hours_spent', 0) or 0)
    except ValueError:
        hours = 0.0
    if not resolved_on:
        cdate = date.fromisoformat(creation_date)
        resolved_on = (cdate + timedelta(days=1)).isoformat()
    conn = db()
    conn.execute('''
        INSERT INTO tasks (user_id,task_id,type,description,requested_by,assignee,
                           creation_date,resolved_on,priority,status,hours_spent,notes,raw_input)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
    ''', (session['user_id'], task_id, type_val, desc, full_name, full_name,
          creation_date, resolved_on, priority, status, hours, notes, 'inline'))
    conn.commit()
    conn.close()
    week = friday_of(date.fromisoformat(creation_date)).isoformat()
    return jsonify(ok=True, week=week)

@app.route('/delete-selected', methods=['POST'])
@login_required
def delete_selected():
    ids = request.form.getlist('ids')
    if ids:
        conn = db()
        for tid in ids:
            conn.execute('DELETE FROM tasks WHERE id=? AND user_id=?', (tid, session['user_id']))
        conn.commit()
        conn.close()
        flash(f'✓ Deleted {len(ids)} entr{"y" if len(ids)==1 else "ies"}.', 'success')
    else:
        flash('No entries selected.', 'error')
    return redirect(request.referrer or '/')

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/login')

# ── Run ───────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    print("\n  MAG Weekly Tasks App running at: http://localhost:5000\n")
    app.run(debug=True, host='0.0.0.0', port=5000)
