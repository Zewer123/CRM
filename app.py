from flask import Flask, render_template, request, jsonify, session, redirect, url_for, send_file
from werkzeug.security import check_password_hash, generate_password_hash
import sqlite3
import os
import io
import csv
from datetime import datetime, timedelta
from functools import wraps

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    import cloudinary
    import cloudinary.uploader
    CLOUDINARY_AVAILABLE = True
    cloudinary.config(
        cloud_name=os.getenv('CLOUDINARY_CLOUD_NAME', ''),
        api_key=os.getenv('CLOUDINARY_API_KEY', ''),
        api_secret=os.getenv('CLOUDINARY_API_SECRET', ''),
    )
except ImportError:
    CLOUDINARY_AVAILABLE = False

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'zewer-aml-crm-secret-2026')

# SQLite database path - /app always exists in Railway container
DB = os.getenv('DB_PATH', os.path.join(os.path.dirname(__file__), 'aml_crm.db'))

@app.context_processor
def inject_user():
    return dict(
        user_name=session.get('user_name', ''),
        user_role=session.get('user_role', ''),
        user_email=session.get('user_email', ''),
        current_user_id=session.get('user_id'),
    )

def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA foreign_keys = ON')
    return conn

def days_left(date_str):
    if not date_str: return None
    try:
        return (datetime.strptime(str(date_str)[:10], '%Y-%m-%d').date() - datetime.now().date()).days
    except: return None

def expiry_status(d):
    if d is None: return 'unknown'
    if d < 0: return 'expired'
    if d <= 30: return 'critical'
    if d <= 90: return 'warning'
    return 'valid'

def get_dropdowns():
    conn = get_db()
    rows = conn.execute("SELECT field_name, value FROM dropdowns WHERE is_active=1 ORDER BY field_name, value").fetchall()
    conn.close()
    dd = {}
    for r in rows:
        dd.setdefault(r['field_name'], []).append(r['value'])
    return dd

def setup_db():
    conn = get_db()
    conn.executescript('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL,
            name TEXT NOT NULL, role TEXT DEFAULT 'staff',
            contact_number TEXT, mobile TEXT, is_active INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ac_code TEXT UNIQUE NOT NULL, client_name TEXT NOT NULL,
            ac_opening_date DATE, ac_status TEXT DEFAULT 'Active', active_till_year TEXT,
            nature TEXT, type_of_client TEXT, name_of_freezone TEXT, mode_of_ac TEXT,
            country_of_incorporation TEXT, region TEXT, address TEXT,
            telephone TEXT, mobile TEXT, whatsapp_number TEXT, email_id TEXT,
            contact_person_name TEXT, contact_person_number TEXT, account_manager TEXT,
            address_proof_type TEXT, address_proof_expiry DATE, kyc_status TEXT,
            trade_license_no TEXT, issuing_authority TEXT, legal_type TEXT,
            incorporation_date DATE, trade_license_expiry DATE,
            tax_no_trn TEXT, vat_cert TEXT, vat_declaration TEXT, deal_after_vat TEXT,
            num_beneficial_owners INTEGER DEFAULT 0, moa TEXT, pep TEXT,
            undertaking TEXT, source_of_fund TEXT, software_updation TEXT,
            doc_status TEXT DEFAULT 'Incompleted', screening_date DATE,
            registration_screening_tool TEXT, risk_status TEXT DEFAULT 'Unspecified',
            verified_by TEXT, verified_date DATE, followup_details TEXT,
            crowe_feedback TEXT, zewer_comments TEXT, created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS ubos (
            id INTEGER PRIMARY KEY AUTOINCREMENT, company_id INTEGER NOT NULL,
            position TEXT, share_percentage REAL, person_name TEXT NOT NULL,
            nationality TEXT, residential_status TEXT, passport_no TEXT,
            passport_expiry DATE, emirates_id TEXT, emirates_id_expiry DATE,
            doc_status TEXT DEFAULT 'Incompleted', verified_by TEXT,
            verified_date DATE, followup_details TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS dropdowns (
            id INTEGER PRIMARY KEY AUTOINCREMENT, field_name TEXT NOT NULL,
            value TEXT NOT NULL, description TEXT, is_active INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, UNIQUE(field_name, value)
        );
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL,
            description TEXT, assigned_to INTEGER, created_by INTEGER,
            company_id INTEGER, priority TEXT DEFAULT 'normal',
            status TEXT DEFAULT 'todo', due_date DATE,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT, company_id INTEGER NOT NULL,
            doc_type TEXT NOT NULL, file_name TEXT NOT NULL, file_url TEXT NOT NULL,
            public_id TEXT, uploaded_by INTEGER, notes TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    ''')
    # Migrations for existing databases
    migrations = [
        ('users', 'contact_number', 'TEXT'),
        ('users', 'mobile', 'TEXT'),
        ('companies', 'contact_person_name', 'TEXT'),
        ('companies', 'contact_person_number', 'TEXT'),
        ('companies', 'account_manager', 'TEXT'),
        ('companies', 'whatsapp_number', 'TEXT'),
        ('companies', 'deal_after_vat', 'TEXT'),
        ('companies', 'registration_screening_tool', 'TEXT'),
        ('dropdowns', 'description', 'TEXT'),
    ]
    for table, col, typ in migrations:
        try: conn.execute(f'ALTER TABLE {table} ADD COLUMN {col} {typ}')
        except: pass

    # Seed admin user
    if not conn.execute('SELECT id FROM users WHERE email=?', ('admin@zewer.ae',)).fetchone():
        conn.execute('INSERT INTO users (email,password_hash,name,role) VALUES (?,?,?,?)',
            ('admin@zewer.ae', generate_password_hash('Admin@123'), 'Administrator', 'admin'))
        conn.execute('INSERT INTO users (email,password_hash,name,role) VALUES (?,?,?,?)',
            ('compliance@zewer.ae', generate_password_hash('Compliance@123'), 'Compliance Officer', 'compliance'))

    # Seed dropdowns
    if not conn.execute('SELECT id FROM dropdowns LIMIT 1').fetchone():
        dd = {
            'AC STATUS': ['Active','Inactive'],
            'NATURE': ['Individual','Legal entity'],
            'TYPE OF CLIENT': ['MainLand','Free Zone','Abroad','International Corporate'],
            'MODE OF AC': ['Supplier','Customer','Bullion','Refinery','Logistics Co','Exchange','Bank','Insurance','Investor','Technical Services'],
            'RISK STATUS': ['High','Medium','Low','Unspecified'],
            'DOC STATUS': ['Completed','Incompleted'],
            'KYC STATUS': ['New Kyc Updated','Kyc 2025 Updated','Kyc 2024 Updated','Kyc 2023 Updated','Kyc 2022 Updated','Kyc 2021 Updated','Kyc 2020 Updated','Kyc 2019 Updated','Kyc 2018 Updated','Kyc 2017 Updated','Kyc 2016 Updated','Not Updated'],
            'REGION': ['Dubai','Abu Dhabi','Sharjah','Ajman','Ras Al Khaimah','Fujairah','Umm Al Quwain','Al Ain','West Bengal','United Kingdom','Canada','Pakistan','Malaysia','Singapore','Bahrain','Italy','REPUBLIC OF CONGO','HONG KONG','Saudi Arabia','India'],
            'FREEZONE': ['Jabel Ali FZ','DMCC','Sharjah Saif Zone','Ajman FZ','Fujairah','Dubai Production City','N/A','Ras Al Khaimah Fz','Dubai Free Zone','Dubai Gold & Diamond Park'],
            'LEGAL TYPE': ['Limited Liability Company(LLC)','Limited Liability Company (WLL)','Civil Company Professional','Foreign Company','Public Company','Privet Company','DMCC','Free Zone Limited Liability Company (FZ-LLC)','Sole Establishment','Partnership Company','Services Agency','Limited (LTD)','Individual Institution','Establishment','FZCO','FZE','FZC'],
            'ISSUING AUTHORITY': ['Dubai Economy & Tourism','Department of Economic Development Ajman','Abu-Dhabi Department of Economic Development','DMCC','Saif Zone','Ajman FZ','Jebel Ali FZ','Dubai Development Authority','Government Of Sharjah Economic Development Department','Government of Ras Al-Khaimah Department of Economic Development','Department of Economic Development Dubai','Dubai Integrated Economic Zones Authority','Fujairah Municipality','Trade Development Authority Of Pakistan','UK HMRC','The Registrar Of Companies For England And Wales','Canada Revenue Agency','Kolkata Municipal Corporation','Abroad'],
            'ADDRESS PROOF TYPE': ['Ejari','Tenancy','Electricity Bill','Gst Registration Certificate','Certification Of Incorporation','Certificate Of Registration For Value Added Tax','Vat Certificate','Telephone Bill','Title Deed','Certificate Of Enlistment','Association Of Article Details','Warehouse Lease Agreement','Not Required'],
            'VAT CERT': ['Yes','No','Not Required'],
            'VAT DECLARATION': ['Yes','No','Not Required'],
            'MOA': ['Yes','No'],
            'PEP': ['Yes','No'],
            'UNDERTAKING': ['Yes','No'],
            'SOURCE OF FUND': ['Yes','No'],
            'POSITION': ['UBO','Authorized Person','Director','Manager','Partner'],
            'RESIDENTIAL STATUS': ['Resident','Non Resident'],
            'COUNTRY': ['United Arab Emirates','Saudi Arabia','Kuwait','Qatar','Bahrain','Oman','India','Pakistan','Bangladesh','Sri Lanka','Philippines','Malaysia','Singapore','China','Hong Kong','Jordan','Lebanon','Syria','Iraq','Yemen','Egypt','Libya','Nigeria','Ethiopia','Republic Of Congo','Turkey','Iran','Afghanistan','Algeria','Canada','United Kingdom','United States of America','France','Ireland','Italy','Germany','Armenia','Belize'],
            'TASK TEMPLATE': ['Collect Updated Trade License','KYC Update Required','Address Proof Renewal','Passport Renewal Follow-up','Emirates ID Update','VAT Certificate Collection','Screening Review','MOA Collection','Undertaking Form','Source of Funds Verification','Risk Assessment Review','Annual KYC Review'],
        }
        for field, vals in dd.items():
            for v in vals:
                try: conn.execute('INSERT OR IGNORE INTO dropdowns (field_name,value,is_active) VALUES (?,?,1)', (field,v))
                except: pass
        # Task templates with descriptions
        templates = [
            ('Collect Updated Trade License','Request client to provide renewed trade license document'),
            ('KYC Update Required','Client KYC documents are outdated - collect updated forms'),
            ('Address Proof Renewal','Collect updated Ejari/Tenancy/Utility bill for address proof'),
            ('Passport Renewal Follow-up','UBO or authorized person passport has expired or expiring'),
            ('Emirates ID Update','Collect renewed Emirates ID for UBO or authorized person'),
            ('VAT Certificate Collection','Request updated VAT registration certificate from client'),
            ('Screening Review','Conduct AML screening review for client'),
            ('MOA Collection','Obtain signed Memorandum of Association from client'),
            ('Undertaking Form','Get signed undertaking form from client'),
            ('Source of Funds Verification','Verify and document source of funds for this client'),
            ('Risk Assessment Review','Review and update risk rating for this client'),
            ('Annual KYC Review','Perform annual KYC review and update all documents'),
        ]
        for val, desc in templates:
            try: conn.execute('INSERT OR REPLACE INTO dropdowns (field_name,value,description,is_active) VALUES (?,?,?,1)', ('TASK TEMPLATE',val,desc))
            except: pass

    conn.commit()
    conn.close()

setup_db()

# ── AUTH DECORATORS ─────────────────────────────────────────

def login_required(f):
    @wraps(f)
    def dec(*a, **kw):
        if 'user_id' not in session: return redirect(url_for('login'))
        return f(*a, **kw)
    return dec

def admin_required(f):
    @wraps(f)
    def dec(*a, **kw):
        if 'user_id' not in session: return redirect(url_for('login'))
        if session.get('user_role') != 'admin': return redirect(url_for('dashboard'))
        return f(*a, **kw)
    return dec

def compliance_required(f):
    @wraps(f)
    def dec(*a, **kw):
        if 'user_id' not in session: return redirect(url_for('login'))
        if session.get('user_role') not in ('admin','compliance'): return redirect(url_for('tasks'))
        return f(*a, **kw)
    return dec

# ── AUTH ────────────────────────────────────────────────────

@app.route('/')
def index():
    return redirect(url_for('dashboard') if 'user_id' in session else url_for('login'))

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        d = request.get_json()
        conn = get_db()
        u = conn.execute('SELECT * FROM users WHERE email=?', (d.get('email'),)).fetchone()
        conn.close()
        if u and check_password_hash(u['password_hash'], d.get('password','')) and u['is_active']:
            session.update(user_id=u['id'], user_email=u['email'], user_name=u['name'], user_role=u['role'])
            return jsonify({'success': True})
        return jsonify({'success': False, 'error': 'Invalid credentials'}), 401
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ── DASHBOARD ───────────────────────────────────────────────

@app.route('/dashboard')
@login_required
def dashboard():
    conn = get_db()
    today = datetime.now().date()
    def c(sql, p=None):
        r = conn.execute(sql, p or []).fetchone()
        return r[0] if r else 0
    total = c('SELECT COUNT(*) FROM companies')
    active = c('SELECT COUNT(*) FROM companies WHERE ac_status="Active"')
    expired_tl = c('SELECT COUNT(*) FROM companies WHERE trade_license_expiry<?', (today,))
    expiring_30_tl = c('SELECT COUNT(*) FROM companies WHERE trade_license_expiry BETWEEN ? AND ?', (today, today+timedelta(days=30)))
    expiring_90_tl = c('SELECT COUNT(*) FROM companies WHERE trade_license_expiry BETWEEN ? AND ?', (today, today+timedelta(days=90)))
    expired_ap = c('SELECT COUNT(*) FROM companies WHERE address_proof_expiry<?', (today,))
    expiring_30_ap = c('SELECT COUNT(*) FROM companies WHERE address_proof_expiry BETWEEN ? AND ?', (today, today+timedelta(days=30)))
    expired_pass = c('SELECT COUNT(*) FROM ubos WHERE passport_expiry<?', (today,))
    expiring_30_pass = c('SELECT COUNT(*) FROM ubos WHERE passport_expiry BETWEEN ? AND ?', (today, today+timedelta(days=30)))
    risk_data = conn.execute('SELECT risk_status, COUNT(*) as c FROM companies GROUP BY risk_status').fetchall()
    risk_breakdown = {r['risk_status']: r['c'] for r in risk_data}
    doc_data = conn.execute('SELECT doc_status, COUNT(*) as c FROM companies GROUP BY doc_status').fetchall()
    doc_breakdown = {r['doc_status']: r['c'] for r in doc_data}
    urgent = conn.execute('''SELECT id,ac_code,client_name,trade_license_expiry,address_proof_expiry,
        risk_status,mobile,whatsapp_number FROM companies
        WHERE (trade_license_expiry BETWEEN ? AND ?) OR (address_proof_expiry BETWEEN ? AND ?)
        ORDER BY trade_license_expiry LIMIT 10''',
        (today, today+timedelta(days=90), today, today+timedelta(days=90))).fetchall()
    open_tasks = c("SELECT COUNT(*) FROM tasks WHERE status NOT IN ('done')")
    overdue_tasks = c("SELECT COUNT(*) FROM tasks WHERE status NOT IN ('done','pending_close') AND due_date<?", (today,))
    due_today = c("SELECT COUNT(*) FROM tasks WHERE status NOT IN ('done') AND due_date=?", (today,))
    pending_close = c("SELECT COUNT(*) FROM tasks WHERE status='pending_close'")
    urgent_tasks = conn.execute('''SELECT t.*, u.name as assigned_name FROM tasks t
        LEFT JOIN users u ON t.assigned_to=u.id
        WHERE t.status NOT IN ('done') AND (t.priority IN ('urgent','high') OR t.due_date<=?)
        ORDER BY CASE t.priority WHEN 'urgent' THEN 1 WHEN 'high' THEN 2 ELSE 3 END, t.due_date LIMIT 8''',
        (today+timedelta(days=2),)).fetchall()
    conn.close()
    return render_template('dashboard.html',
        total_companies=total, active_companies=active,
        expired_tl=expired_tl, expiring_30_tl=expiring_30_tl, expiring_90_tl=expiring_90_tl,
        expired_ap=expired_ap, expiring_30_ap=expiring_30_ap,
        expired_pass=expired_pass, expiring_30_pass=expiring_30_pass,
        risk_breakdown=risk_breakdown, doc_breakdown=doc_breakdown,
        urgent_companies=urgent, open_tasks=open_tasks, overdue_tasks=overdue_tasks,
        due_today=due_today, pending_close=pending_close, urgent_tasks=urgent_tasks,
        today=str(today), days_left=days_left)

# ── COMPANIES ───────────────────────────────────────────────

@app.route('/companies')
@compliance_required
def companies():
    conn = get_db()
    search = request.args.get('search','')
    risk_f = request.args.get('risk','')
    status_f = request.args.get('status','')
    region_f = request.args.get('region','')
    q = 'SELECT * FROM companies WHERE 1=1'; p = []
    if search:
        q += ' AND (client_name LIKE ? OR ac_code LIKE ? OR mobile LIKE ? OR trade_license_no LIKE ?)'
        s = f'%{search}%'; p += [s,s,s,s]
    if risk_f: q += ' AND risk_status=?'; p.append(risk_f)
    if status_f: q += ' AND ac_status=?'; p.append(status_f)
    if region_f: q += ' AND region=?'; p.append(region_f)
    rows = conn.execute(q+' ORDER BY created_at DESC', p).fetchall()
    dd = get_dropdowns()
    conn.close()
    cl = []
    for c in rows:
        tl = days_left(c['trade_license_expiry']); ap = days_left(c['address_proof_expiry'])
        cl.append(dict(id=c['id'], ac_code=c['ac_code'], client_name=c['client_name'],
            region=c['region'], type_of_client=c['type_of_client'], mode_of_ac=c['mode_of_ac'],
            mobile=c['mobile'], whatsapp_number=c['whatsapp_number'], risk_status=c['risk_status'],
            ac_status=c['ac_status'], doc_status=c['doc_status'], kyc_status=c['kyc_status'],
            account_manager=c['account_manager'], contact_person_name=c['contact_person_name'],
            trade_license_expiry=c['trade_license_expiry'], tl_days=tl, tl_status=expiry_status(tl),
            address_proof_expiry=c['address_proof_expiry'], ap_days=ap, ap_status=expiry_status(ap)))
    return render_template('companies.html', companies=cl, search=search,
        risk_filter=risk_f, status_filter=status_f, region_filter=region_f,
        regions=dd.get('REGION', []))

@app.route('/company/new')
@compliance_required
def company_new():
    conn = get_db()
    staff_users = conn.execute('SELECT id,name FROM users WHERE is_active=1 ORDER BY name').fetchall()
    conn.close()
    return render_template('company_form.html', dropdown_data=get_dropdowns(),
        company=None, ubos=[], edit=False, staff_users=staff_users)

@app.route('/company/<int:id>')
@compliance_required
def company_detail(id):
    conn = get_db()
    company = conn.execute('SELECT * FROM companies WHERE id=?', (id,)).fetchone()
    if not company: conn.close(); return redirect(url_for('companies'))
    ubos = conn.execute('SELECT * FROM ubos WHERE company_id=? ORDER BY share_percentage DESC', (id,)).fetchall()
    conn.close()
    tl = days_left(company['trade_license_expiry']); ap = days_left(company['address_proof_expiry'])
    ubo_list = []
    for u in ubos:
        pd = days_left(u['passport_expiry']); ed = days_left(u['emirates_id_expiry'])
        ubo_list.append(dict(id=u['id'], position=u['position'], share_percentage=u['share_percentage'],
            person_name=u['person_name'], nationality=u['nationality'], residential_status=u['residential_status'],
            passport_no=u['passport_no'], passport_expiry=u['passport_expiry'], p_days=pd, p_status=expiry_status(pd),
            emirates_id=u['emirates_id'], emirates_id_expiry=u['emirates_id_expiry'], e_days=ed, e_status=expiry_status(ed),
            doc_status=u['doc_status'], verified_by=u['verified_by']))
    return render_template('company_detail.html', company=company, ubos=ubo_list,
        tl_days=tl, tl_status=expiry_status(tl), ap_days=ap, ap_status=expiry_status(ap),
        today=str(datetime.now().date()))

@app.route('/company/<int:id>/edit')
@compliance_required
def company_edit(id):
    conn = get_db()
    company = conn.execute('SELECT * FROM companies WHERE id=?', (id,)).fetchone()
    if not company: conn.close(); return redirect(url_for('companies'))
    ubos = conn.execute('SELECT * FROM ubos WHERE company_id=? ORDER BY share_percentage DESC', (id,)).fetchall()
    staff_users = conn.execute('SELECT id,name FROM users WHERE is_active=1 ORDER BY name').fetchall()
    conn.close()
    return render_template('company_form.html', dropdown_data=get_dropdowns(),
        company=company, ubos=ubos, edit=True, staff_users=staff_users)

def _company_vals(data):
    return (data.get('ac_opening_date') or None, data.get('ac_status','Active'),
        data.get('active_till_year'), data.get('nature'), data.get('type_of_client'),
        data.get('name_of_freezone'), data.get('mode_of_ac'), data.get('country_of_incorporation'),
        data.get('region'), data.get('address'), data.get('telephone'), data.get('mobile'),
        data.get('whatsapp_number'), data.get('email_id'), data.get('contact_person_name'),
        data.get('contact_person_number'), data.get('account_manager'),
        data.get('address_proof_type'), data.get('address_proof_expiry') or None,
        data.get('kyc_status'), data.get('trade_license_no'), data.get('issuing_authority'),
        data.get('legal_type'), data.get('incorporation_date') or None,
        data.get('trade_license_expiry') or None, data.get('tax_no_trn'),
        data.get('vat_cert'), data.get('vat_declaration'), data.get('deal_after_vat'),
        int(data.get('num_beneficial_owners') or 0), data.get('moa'), data.get('pep'),
        data.get('undertaking'), data.get('source_of_fund'), data.get('software_updation'),
        data.get('doc_status','Incompleted'), data.get('screening_date') or None,
        data.get('registration_screening_tool'), data.get('risk_status','Unspecified'),
        data.get('verified_by'), data.get('verified_date') or None,
        data.get('followup_details'), data.get('crowe_feedback'), data.get('zewer_comments'))

def _save_ubos(conn, company_id, ubos):
    conn.execute('DELETE FROM ubos WHERE company_id=?', (company_id,))
    for u in ubos:
        if not u.get('person_name'): continue
        conn.execute('''INSERT INTO ubos (company_id,position,share_percentage,person_name,nationality,
            residential_status,passport_no,passport_expiry,emirates_id,emirates_id_expiry,
            doc_status,verified_by,verified_date,followup_details) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (company_id, u.get('position'), u.get('share_percentage') or None, u.get('person_name'),
             u.get('nationality'), u.get('residential_status'), u.get('passport_no'),
             u.get('passport_expiry') or None, u.get('emirates_id'), u.get('emirates_id_expiry') or None,
             u.get('doc_status','Incompleted'), u.get('verified_by'), u.get('verified_date') or None,
             u.get('followup_details')))

@app.route('/api/company/add', methods=['POST'])
@compliance_required
def api_add_company():
    data = request.get_json()
    try:
        conn = get_db()
        conn.execute('''INSERT INTO companies (ac_code,client_name,ac_opening_date,ac_status,active_till_year,
            nature,type_of_client,name_of_freezone,mode_of_ac,country_of_incorporation,region,address,
            telephone,mobile,whatsapp_number,email_id,contact_person_name,contact_person_number,account_manager,
            address_proof_type,address_proof_expiry,kyc_status,trade_license_no,issuing_authority,legal_type,
            incorporation_date,trade_license_expiry,tax_no_trn,vat_cert,vat_declaration,deal_after_vat,
            num_beneficial_owners,moa,pep,undertaking,source_of_fund,software_updation,doc_status,
            screening_date,registration_screening_tool,risk_status,verified_by,verified_date,
            followup_details,crowe_feedback,zewer_comments,created_by)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (data.get('ac_code'), data.get('client_name')) + _company_vals(data) + (session.get('user_id'),))
        cid = conn.execute('SELECT last_insert_rowid()').fetchone()[0]
        _save_ubos(conn, cid, data.get('ubos', []))
        conn.commit(); conn.close()
        return jsonify({'success': True, 'id': cid})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/company/<int:id>/edit', methods=['POST'])
@compliance_required
def api_edit_company(id):
    data = request.get_json()
    try:
        conn = get_db()
        conn.execute('''UPDATE companies SET client_name=?,ac_opening_date=?,ac_status=?,active_till_year=?,
            nature=?,type_of_client=?,name_of_freezone=?,mode_of_ac=?,country_of_incorporation=?,region=?,
            address=?,telephone=?,mobile=?,whatsapp_number=?,email_id=?,contact_person_name=?,
            contact_person_number=?,account_manager=?,address_proof_type=?,address_proof_expiry=?,
            kyc_status=?,trade_license_no=?,issuing_authority=?,legal_type=?,incorporation_date=?,
            trade_license_expiry=?,tax_no_trn=?,vat_cert=?,vat_declaration=?,deal_after_vat=?,
            num_beneficial_owners=?,moa=?,pep=?,undertaking=?,source_of_fund=?,software_updation=?,
            doc_status=?,screening_date=?,registration_screening_tool=?,risk_status=?,verified_by=?,
            verified_date=?,followup_details=?,crowe_feedback=?,zewer_comments=?,
            updated_at=CURRENT_TIMESTAMP WHERE id=?''',
            (data.get('client_name'),) + _company_vals(data) + (id,))
        _save_ubos(conn, id, data.get('ubos', []))
        conn.commit(); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/company/<int:id>/delete', methods=['POST'])
@compliance_required
def api_delete_company(id):
    try:
        conn = get_db()
        conn.execute('DELETE FROM companies WHERE id=?', (id,))
        conn.commit(); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ── ALERTS ──────────────────────────────────────────────────

@app.route('/alerts')
@compliance_required
def alerts():
    conn = get_db()
    today = datetime.now().date()
    cutoff = today - timedelta(days=30)
    tl = conn.execute('SELECT id,ac_code,client_name,mobile,whatsapp_number,trade_license_expiry,risk_status,region FROM companies WHERE trade_license_expiry>=? ORDER BY trade_license_expiry', (cutoff,)).fetchall()
    ap = conn.execute('SELECT id,ac_code,client_name,mobile,whatsapp_number,address_proof_expiry,address_proof_type,risk_status FROM companies WHERE address_proof_expiry>=? ORDER BY address_proof_expiry', (cutoff,)).fetchall()
    ubos = conn.execute('''SELECT u.id,u.person_name,u.passport_no,u.passport_expiry,u.emirates_id,u.emirates_id_expiry,
        c.id as company_id,c.client_name,c.ac_code FROM ubos u JOIN companies c ON u.company_id=c.id
        WHERE u.passport_expiry>=? OR u.emirates_id_expiry>=? ORDER BY u.passport_expiry''', (cutoff,cutoff)).fetchall()
    conn.close()
    return render_template('alerts.html', tl_expiring=tl, ap_expiring=ap, ubo_expiring=ubos,
        today=str(today), days_left=days_left, expiry_status=expiry_status)

# ── REPORTS ─────────────────────────────────────────────────

@app.route('/reports')
@compliance_required
def reports():
    conn = get_db()
    today = datetime.now().date()
    date_from = request.args.get('from','')
    date_to = request.args.get('to','')
    where = '1=1'; params = []
    if date_from: where += ' AND created_at >= ?'; params.append(date_from)
    if date_to: where += ' AND created_at <= ?'; params.append(date_to+' 23:59:59')
    total = conn.execute(f'SELECT COUNT(*) FROM companies WHERE {where}', params).fetchone()[0]
    def q(sql): return conn.execute(sql, params).fetchall()
    result = render_template('reports.html',
        risk_data=q(f'SELECT risk_status,COUNT(*) as c FROM companies WHERE {where} GROUP BY risk_status'),
        doc_data=q(f'SELECT doc_status,COUNT(*) as c FROM companies WHERE {where} GROUP BY doc_status'),
        type_data=q(f'SELECT type_of_client,COUNT(*) as c FROM companies WHERE {where} GROUP BY type_of_client ORDER BY c DESC'),
        region_data=q(f'SELECT region,COUNT(*) as c FROM companies WHERE {where} GROUP BY region ORDER BY c DESC LIMIT 10'),
        kyc_data=q(f'SELECT kyc_status,COUNT(*) as c FROM companies WHERE {where} GROUP BY kyc_status ORDER BY c DESC'),
        mode_data=q(f'SELECT mode_of_ac,COUNT(*) as c FROM companies WHERE {where} GROUP BY mode_of_ac ORDER BY c DESC'),
        total=total, date_from=date_from, date_to=date_to,
        expired_tl=conn.execute('SELECT COUNT(*) FROM companies WHERE trade_license_expiry<?',(today,)).fetchone()[0],
        expiring_30=conn.execute('SELECT COUNT(*) FROM companies WHERE trade_license_expiry BETWEEN ? AND ?',(today,today+timedelta(days=30))).fetchone()[0],
        expiring_90=conn.execute('SELECT COUNT(*) FROM companies WHERE trade_license_expiry BETWEEN ? AND ?',(today,today+timedelta(days=90))).fetchone()[0])
    conn.close()
    return result

# ── SETTINGS ────────────────────────────────────────────────

@app.route('/settings')
@admin_required
def settings():
    conn = get_db()
    users = conn.execute('SELECT id,email,name,role,is_active,contact_number FROM users ORDER BY role,name').fetchall()
    dds = conn.execute('SELECT * FROM dropdowns WHERE is_active=1 ORDER BY field_name,value').fetchall()
    conn.close()
    groups = {}
    for d in dds: groups.setdefault(d['field_name'], []).append(d)
    return render_template('settings.html', users=users, dropdown_groups=groups)

@app.route('/api/user/add', methods=['POST'])
@admin_required
def api_add_user():
    d = request.get_json()
    try:
        conn = get_db()
        conn.execute('INSERT INTO users (email,password_hash,name,role,contact_number,is_active) VALUES (?,?,?,?,?,1)',
            (d.get('email'), generate_password_hash(d.get('password','')), d.get('name'), d.get('role'), d.get('contact_number')))
        conn.commit(); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/user/<int:id>/edit', methods=['POST'])
@admin_required
def api_edit_user(id):
    d = request.get_json()
    try:
        conn = get_db()
        if d.get('password'):
            conn.execute('UPDATE users SET name=?,email=?,role=?,contact_number=?,password_hash=? WHERE id=?',
                (d.get('name'),d.get('email'),d.get('role'),d.get('contact_number'),generate_password_hash(d.get('password')),id))
        else:
            conn.execute('UPDATE users SET name=?,email=?,role=?,contact_number=? WHERE id=?',
                (d.get('name'),d.get('email'),d.get('role'),d.get('contact_number'),id))
        conn.commit(); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/user/<int:id>/toggle', methods=['POST'])
@admin_required
def api_toggle_user(id):
    try:
        conn = get_db()
        cur = conn.execute('SELECT is_active FROM users WHERE id=?',(id,)).fetchone()
        if cur: conn.execute('UPDATE users SET is_active=? WHERE id=?',(0 if cur['is_active'] else 1,id))
        conn.commit(); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/user/<int:id>/delete', methods=['POST'])
@admin_required
def api_delete_user(id):
    try:
        conn = get_db()
        conn.execute('DELETE FROM users WHERE id=?',(id,))
        conn.commit(); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/dropdown/add', methods=['POST'])
@admin_required
def api_add_dropdown():
    d = request.get_json()
    try:
        conn = get_db()
        conn.execute('INSERT INTO dropdowns (field_name,value,is_active) VALUES (?,?,1)', (d.get('field_name'),d.get('value')))
        conn.commit(); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/dropdown/<int:id>/delete', methods=['POST'])
@admin_required
def api_delete_dropdown(id):
    try:
        conn = get_db()
        conn.execute('DELETE FROM dropdowns WHERE id=?',(id,))
        conn.commit(); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ── TASKS ───────────────────────────────────────────────────

@app.route('/tasks')
@login_required
def tasks():
    conn = get_db()
    uid = session.get('user_id')
    role = session.get('user_role')
    base = '''SELECT t.*,u.name as assigned_name,u.mobile as assigned_mobile,
        cu.name as created_by_name,c.client_name as company_name,c.ac_code
        FROM tasks t LEFT JOIN users u ON t.assigned_to=u.id
        LEFT JOIN users cu ON t.created_by=cu.id
        LEFT JOIN companies c ON t.company_id=c.id'''
    order = " ORDER BY CASE t.priority WHEN 'urgent' THEN 1 WHEN 'high' THEN 2 WHEN 'normal' THEN 3 ELSE 4 END, t.due_date"
    if role == 'staff':
        rows = conn.execute(base+' WHERE t.assigned_to=?'+order, (uid,)).fetchall()
    else:
        rows = conn.execute(base+order).fetchall()
    all_users = conn.execute('SELECT id,name,role,mobile FROM users WHERE is_active=1 ORDER BY name').fetchall()
    all_companies = conn.execute('SELECT id,ac_code,client_name FROM companies ORDER BY client_name').fetchall()
    task_templates = conn.execute("SELECT value,description FROM dropdowns WHERE field_name='TASK TEMPLATE' AND is_active=1 ORDER BY value").fetchall()
    conn.close()
    tlist = []
    for t in rows:
        d = days_left(t['due_date'])
        tlist.append(dict(id=t['id'],title=t['title'],description=t['description'],
            assigned_to=t['assigned_to'],assigned_name=t['assigned_name'],assigned_mobile=t['assigned_mobile'],
            created_by=t['created_by'],created_by_name=t['created_by_name'],
            company_id=t['company_id'],company_name=t['company_name'],ac_code=t['ac_code'],
            priority=t['priority'] or 'normal',status=t['status'] or 'todo',
            due_date=t['due_date'],days_until_due=d,is_overdue=(d is not None and d<0)))
    return render_template('tasks.html', tasks=tlist, all_users=all_users,
        all_companies=all_companies, task_templates=task_templates)

@app.route('/api/task/add', methods=['POST'])
@login_required
def api_add_task():
    d = request.get_json()
    try:
        conn = get_db()
        conn.execute('INSERT INTO tasks (title,description,assigned_to,created_by,company_id,priority,due_date,status) VALUES (?,?,?,?,?,?,?,?)',
            (d.get('title'),d.get('description'),d.get('assigned_to') or None,session.get('user_id'),
             d.get('company_id') or None,d.get('priority','normal'),d.get('due_date') or None,'todo'))
        conn.commit(); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/task/<int:id>/edit', methods=['POST'])
@login_required
def api_edit_task(id):
    d = request.get_json()
    try:
        conn = get_db()
        conn.execute('UPDATE tasks SET title=?,description=?,assigned_to=?,company_id=?,priority=?,due_date=?,updated_at=CURRENT_TIMESTAMP WHERE id=?',
            (d.get('title'),d.get('description'),d.get('assigned_to') or None,
             d.get('company_id') or None,d.get('priority','normal'),d.get('due_date') or None,id))
        conn.commit(); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/task/<int:id>/status', methods=['POST'])
@login_required
def api_task_status(id):
    d = request.get_json()
    try:
        conn = get_db()
        conn.execute('UPDATE tasks SET status=?,updated_at=CURRENT_TIMESTAMP WHERE id=?',(d.get('status'),id))
        conn.commit(); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/task/<int:id>/delete', methods=['POST'])
@login_required
def api_delete_task(id):
    try:
        conn = get_db()
        conn.execute('DELETE FROM tasks WHERE id=?',(id,))
        conn.commit(); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ── DOCUMENTS ───────────────────────────────────────────────

@app.route('/api/company/<int:company_id>/documents')
@compliance_required
def api_get_documents(company_id):
    conn = get_db()
    docs = conn.execute('''SELECT d.*,u.name as uploader_name FROM documents d
        LEFT JOIN users u ON d.uploaded_by=u.id WHERE d.company_id=? ORDER BY d.created_at DESC''',(company_id,)).fetchall()
    conn.close()
    return jsonify([dict(d) for d in docs])

@app.route('/api/company/<int:company_id>/upload', methods=['POST'])
@compliance_required
def api_upload_document(company_id):
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No file'}), 400
    file = request.files['file']
    doc_type = request.form.get('doc_type','General')
    notes = request.form.get('notes','')
    if not file.filename:
        return jsonify({'success': False, 'error': 'No file selected'}), 400
    try:
        if CLOUDINARY_AVAILABLE and os.getenv('CLOUDINARY_CLOUD_NAME'):
            result = cloudinary.uploader.upload(file, folder=f'zewer_crm/company_{company_id}', resource_type='auto', use_filename=True, unique_filename=True)
            file_url = result['secure_url']; public_id = result['public_id']
        else:
            return jsonify({'success': False, 'error': 'Cloudinary not configured. Add CLOUDINARY_CLOUD_NAME, CLOUDINARY_API_KEY, CLOUDINARY_API_SECRET to Railway environment variables.'}), 500
        conn = get_db()
        conn.execute('INSERT INTO documents (company_id,doc_type,file_name,file_url,public_id,uploaded_by,notes) VALUES (?,?,?,?,?,?,?)',
            (company_id,doc_type,file.filename,file_url,public_id,session.get('user_id'),notes))
        conn.commit(); conn.close()
        return jsonify({'success': True, 'url': file_url, 'name': file.filename})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/document/<int:doc_id>/delete', methods=['POST'])
@compliance_required
def api_delete_document(doc_id):
    try:
        conn = get_db()
        doc = conn.execute('SELECT public_id FROM documents WHERE id=?',(doc_id,)).fetchone()
        if doc and doc['public_id'] and CLOUDINARY_AVAILABLE and os.getenv('CLOUDINARY_CLOUD_NAME'):
            try: cloudinary.uploader.destroy(doc['public_id'], resource_type='raw')
            except: pass
        conn.execute('DELETE FROM documents WHERE id=?',(doc_id,))
        conn.commit(); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ── EXPORT / IMPORT ─────────────────────────────────────────

@app.route('/export/template')
@compliance_required
def export_template():
    if not OPENPYXL_AVAILABLE: return "openpyxl not installed", 500
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Companies"
    headers = ['ac_code','client_name','ac_opening_date','ac_status','nature','type_of_client',
               'name_of_freezone','mode_of_ac','country_of_incorporation','region','address',
               'telephone','mobile','whatsapp_number','email_id','contact_person_name',
               'contact_person_number','account_manager','address_proof_type','address_proof_expiry',
               'trade_license_no','issuing_authority','legal_type','incorporation_date',
               'trade_license_expiry','tax_no_trn','vat_cert','vat_declaration','num_beneficial_owners',
               'moa','pep','undertaking','source_of_fund','doc_status','risk_status','kyc_status',
               'verified_by','followup_details','zewer_comments']
    ws.append(headers)
    ws.append(['TJ5003','SAMPLE COMPANY LLC','2020-01-01','Active','Legal entity','MainLand','N/A',
               'Supplier','United Arab Emirates','Dubai','Unit 101, Gold Souq, Dubai','+97142258019',
               '+971506594165','+971506594165','info@sample.com','Ahmed Ali','+971501234567','Jaseel',
               'Ejari','2025-12-31','534230','Dubai Economy & Tourism','Limited Liability Company(LLC)',
               '2019-06-01','2025-12-31','100003063300003','Yes','Yes',2,'Yes','Yes','Yes','Yes',
               'Completed','Medium','Kyc 2025 Updated','Jaseel','',''])
    out = io.BytesIO(); wb.save(out); out.seek(0)
    return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True, download_name='zewer_company_template.xlsx')

@app.route('/export/companies')
@compliance_required
def export_companies():
    scope = request.args.get('scope','all')
    fmt = request.args.get('format','csv')
    conn = get_db()
    q = 'SELECT * FROM companies WHERE 1=1'
    if scope == 'active': q += ' AND ac_status="Active"'
    elif scope == 'inactive': q += ' AND ac_status="Inactive"'
    elif scope == 'high': q += ' AND risk_status="High"'
    elif scope == 'medium': q += ' AND risk_status="Medium"'
    elif scope == 'low': q += ' AND risk_status="Low"'
    elif scope == 'incomplete': q += ' AND doc_status="Incompleted"'
    rows = conn.execute(q).fetchall(); conn.close()
    fname = f'companies_{scope}_{datetime.now().strftime("%Y%m%d_%H%M%S")}'
    if fmt == 'xlsx' and OPENPYXL_AVAILABLE:
        wb = openpyxl.Workbook(); ws = wb.active
        if rows: ws.append(list(rows[0].keys())); [ws.append(list(r)) for r in rows]
        out = io.BytesIO(); wb.save(out); out.seek(0)
        return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True, download_name=fname+'.xlsx')
    out = io.StringIO(); w = csv.writer(out)
    if rows: w.writerow(rows[0].keys()); [w.writerow(tuple(r)) for r in rows]
    out.seek(0)
    return send_file(io.BytesIO(out.getvalue().encode()), mimetype='text/csv',
        as_attachment=True, download_name=fname+'.csv')

@app.route('/export/report')
@compliance_required
def export_report():
    rtype = request.args.get('type','all'); fmt = request.args.get('format','xlsx')
    conn = get_db()
    if rtype == 'expiry':
        rows = conn.execute('SELECT ac_code,client_name,trade_license_no,trade_license_expiry,address_proof_expiry,risk_status,region,account_manager FROM companies ORDER BY trade_license_expiry').fetchall()
    elif rtype == 'kyc':
        rows = conn.execute('SELECT ac_code,client_name,kyc_status,doc_status,risk_status,verified_by,verified_date FROM companies ORDER BY kyc_status').fetchall()
    elif rtype == 'risk':
        rows = conn.execute('SELECT ac_code,client_name,risk_status,doc_status,kyc_status,region,account_manager FROM companies ORDER BY risk_status').fetchall()
    else:
        rows = conn.execute('SELECT * FROM companies ORDER BY client_name').fetchall()
    conn.close()
    fname = f'report_{rtype}_{datetime.now().strftime("%Y%m%d_%H%M%S")}'
    if fmt == 'xlsx' and OPENPYXL_AVAILABLE:
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = rtype.upper()
        if rows: ws.append(list(rows[0].keys())); [ws.append([str(v) if v else '' for v in r]) for r in rows]
        out = io.BytesIO(); wb.save(out); out.seek(0)
        return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True, download_name=fname+'.xlsx')
    out = io.StringIO(); w = csv.writer(out)
    if rows: w.writerow(rows[0].keys()); [w.writerow(tuple(r)) for r in rows]
    out.seek(0)
    return send_file(io.BytesIO(out.getvalue().encode()), mimetype='text/csv',
        as_attachment=True, download_name=fname+'.csv')

@app.route('/api/import/companies', methods=['POST'])
@compliance_required
def api_import_companies():
    if not OPENPYXL_AVAILABLE: return jsonify({'success': False, 'error': 'openpyxl not installed'}), 500
    if 'file' not in request.files: return jsonify({'success': False, 'error': 'No file'}), 400
    try:
        wb = openpyxl.load_workbook(request.files['file']); ws = wb.active
        headers = [str(c.value).strip() if c.value else '' for c in ws[1]]
        imported = 0; skipped = 0; conn = get_db()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(row): continue
            rd = {headers[i]: (str(row[i]).strip() if row[i] is not None else '') for i in range(min(len(headers),len(row)))}
            ac = rd.get('ac_code','').strip(); name = rd.get('client_name','').strip()
            if not ac or not name: continue
            if conn.execute('SELECT id FROM companies WHERE ac_code=?',(ac,)).fetchone():
                skipped += 1; continue
            try:
                conn.execute('''INSERT INTO companies (ac_code,client_name,ac_status,risk_status,doc_status,
                    nature,type_of_client,region,telephone,mobile,email_id,trade_license_no,
                    trade_license_expiry,address_proof_expiry,kyc_status,account_manager,created_by)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                    (ac,name,rd.get('ac_status','Active'),rd.get('risk_status','Unspecified'),
                     rd.get('doc_status','Incompleted'),rd.get('nature'),rd.get('type_of_client'),
                     rd.get('region'),rd.get('telephone'),rd.get('mobile'),rd.get('email_id'),
                     rd.get('trade_license_no'),rd.get('trade_license_expiry') or None,
                     rd.get('address_proof_expiry') or None,rd.get('kyc_status'),
                     rd.get('account_manager'),session.get('user_id')))
                imported += 1
            except: skipped += 1
        conn.commit(); conn.close()
        return jsonify({'success': True, 'imported': imported, 'skipped': skipped})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.getenv('PORT', 8000))
    app.run(debug=False, host='0.0.0.0', port=port)
