from flask import Flask, render_template, request, jsonify, session, redirect, url_for, send_file
from werkzeug.security import check_password_hash, generate_password_hash
import os, io, csv, sqlite3
from datetime import datetime, timedelta
from functools import wraps

try:
    import openpyxl
    HAS_XL = True
except: HAS_XL = False

try:
    import cloudinary, cloudinary.uploader
    cloudinary.config(cloud_name=os.getenv('CLOUDINARY_CLOUD_NAME',''),
                      api_key=os.getenv('CLOUDINARY_API_KEY',''),
                      api_secret=os.getenv('CLOUDINARY_API_SECRET',''))
    HAS_CLD = True
except: HAS_CLD = False

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'zewer-aml-secret-2026')

DATABASE_URL = os.getenv('DATABASE_URL', '')

@app.context_processor
def inject_user():
    return dict(user_name=session.get('user_name',''), user_role=session.get('user_role',''),
                user_email=session.get('user_email',''), current_user_id=session.get('user_id'))

# ── DATABASE ────────────────────────────────────────────────

def use_pg():
    return bool(DATABASE_URL)

def get_db():
    if use_pg():
        try:
            import psycopg2, psycopg2.extras
            url = DATABASE_URL.replace('postgres://', 'postgresql://', 1)
            c = psycopg2.connect(url, connect_timeout=10,
                                 cursor_factory=psycopg2.extras.RealDictCursor)
            c.autocommit = False
            return c
        except ImportError:
            print("psycopg2 not installed, using SQLite")
        except Exception as e:
            print(f"PG failed: {e}, using SQLite")
    c = sqlite3.connect('/tmp/aml_crm.db')
    c.row_factory = sqlite3.Row
    c.execute('PRAGMA foreign_keys = ON')
    return c

def is_pg(conn):
    return not isinstance(conn, sqlite3.Connection)

def P():
    return '%s' if use_pg() else '?'

def x(conn, sql, p=None):
    """Execute, auto-converting ? to %s for postgres"""
    if is_pg(conn):
        sql = sql.replace('?', '%s')
        cur = conn.cursor()
        cur.execute(sql, p or [])
        return cur
    return conn.execute(sql, p or [])

def one(conn, sql, p=None):
    r = x(conn, sql, p).fetchone()
    return dict(r) if r else None

def all_(conn, sql, p=None):
    rows = x(conn, sql, p).fetchall()
    return [dict(r) for r in rows]

def cnt(conn, sql, p=None):
    r = x(conn, sql, p).fetchone()
    if r is None: return 0
    v = dict(r) if is_pg(conn) else r
    return list(v.values())[0] if isinstance(v, dict) else v[0]

def lastid(conn):
    if is_pg(conn):
        return dict(x(conn, 'SELECT lastval() as id').fetchone())['id']
    return conn.execute('SELECT last_insert_rowid()').fetchone()[0]

def commit(conn):
    conn.commit()

def setup_db():
    conn = get_db()
    if is_pg(conn):
        _pg_schema(conn)
        commit(conn)
    else:
        _sqlite_schema(conn)
    conn.close()

def _pg_schema(conn):
    stmts = [
        """CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY, email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL, name TEXT NOT NULL,
            role TEXT DEFAULT 'staff', contact_number TEXT, mobile TEXT,
            is_active INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS companies (
            id SERIAL PRIMARY KEY, ac_code TEXT UNIQUE NOT NULL,
            client_name TEXT NOT NULL, ac_opening_date DATE,
            ac_status TEXT DEFAULT 'Active', active_till_year TEXT,
            nature TEXT, type_of_client TEXT, name_of_freezone TEXT,
            mode_of_ac TEXT, country_of_incorporation TEXT, region TEXT,
            address TEXT, telephone TEXT, mobile TEXT, whatsapp_number TEXT,
            email_id TEXT, contact_person_name TEXT, contact_person_number TEXT,
            account_manager TEXT, address_proof_type TEXT, address_proof_expiry DATE,
            kyc_status TEXT, trade_license_no TEXT, issuing_authority TEXT,
            legal_type TEXT, incorporation_date DATE, trade_license_expiry DATE,
            tax_no_trn TEXT, vat_cert TEXT, vat_declaration TEXT, deal_after_vat TEXT,
            num_beneficial_owners INTEGER DEFAULT 0, moa TEXT, pep TEXT,
            undertaking TEXT, source_of_fund TEXT, software_updation TEXT,
            doc_status TEXT DEFAULT 'Incompleted', screening_date DATE,
            registration_screening_tool TEXT, risk_status TEXT DEFAULT 'Unspecified',
            verified_by TEXT, verified_date DATE, followup_details TEXT,
            crowe_feedback TEXT, zewer_comments TEXT, created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS ubos (
            id SERIAL PRIMARY KEY, company_id INTEGER NOT NULL,
            position TEXT, share_percentage REAL, person_name TEXT NOT NULL,
            nationality TEXT, residential_status TEXT, passport_no TEXT,
            passport_expiry DATE, emirates_id TEXT, emirates_id_expiry DATE,
            doc_status TEXT DEFAULT 'Incompleted', verified_by TEXT,
            verified_date DATE, followup_details TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS dropdowns (
            id SERIAL PRIMARY KEY, field_name TEXT NOT NULL, value TEXT NOT NULL,
            description TEXT, is_active INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(field_name, value))""",
        """CREATE TABLE IF NOT EXISTS tasks (
            id SERIAL PRIMARY KEY, title TEXT NOT NULL, description TEXT,
            assigned_to INTEGER, created_by INTEGER, company_id INTEGER,
            priority TEXT DEFAULT 'normal', status TEXT DEFAULT 'todo',
            due_date DATE, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS documents (
            id SERIAL PRIMARY KEY, company_id INTEGER NOT NULL,
            doc_type TEXT NOT NULL, file_name TEXT NOT NULL,
            file_url TEXT NOT NULL, public_id TEXT, uploaded_by INTEGER,
            notes TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""",
    ]
    for s in stmts:
        x(conn, s)
    # Seed admin
    if not one(conn, 'SELECT id FROM users WHERE email=%s', ('admin@zewer.ae',)):
        x(conn, 'INSERT INTO users (email,password_hash,name,role) VALUES (%s,%s,%s,%s)',
          ('admin@zewer.ae', generate_password_hash('Admin@123'), 'Administrator', 'admin'))
        x(conn, 'INSERT INTO users (email,password_hash,name,role) VALUES (%s,%s,%s,%s)',
          ('compliance@zewer.ae', generate_password_hash('Compliance@123'), 'Compliance Officer', 'compliance'))
    if cnt(conn, 'SELECT COUNT(*) FROM dropdowns') == 0:
        _seed(conn)

def _sqlite_schema(conn):
    conn.executescript('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT, email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL, name TEXT NOT NULL,
            role TEXT DEFAULT 'staff', contact_number TEXT, mobile TEXT,
            is_active INTEGER DEFAULT 1, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT, ac_code TEXT UNIQUE NOT NULL,
            client_name TEXT NOT NULL, ac_opening_date DATE,
            ac_status TEXT DEFAULT 'Active', active_till_year TEXT, nature TEXT,
            type_of_client TEXT, name_of_freezone TEXT, mode_of_ac TEXT,
            country_of_incorporation TEXT, region TEXT, address TEXT,
            telephone TEXT, mobile TEXT, whatsapp_number TEXT, email_id TEXT,
            contact_person_name TEXT, contact_person_number TEXT, account_manager TEXT,
            address_proof_type TEXT, address_proof_expiry DATE, kyc_status TEXT,
            trade_license_no TEXT, issuing_authority TEXT, legal_type TEXT,
            incorporation_date DATE, trade_license_expiry DATE, tax_no_trn TEXT,
            vat_cert TEXT, vat_declaration TEXT, deal_after_vat TEXT,
            num_beneficial_owners INTEGER DEFAULT 0, moa TEXT, pep TEXT,
            undertaking TEXT, source_of_fund TEXT, software_updation TEXT,
            doc_status TEXT DEFAULT 'Incompleted', screening_date DATE,
            registration_screening_tool TEXT, risk_status TEXT DEFAULT 'Unspecified',
            verified_by TEXT, verified_date DATE, followup_details TEXT,
            crowe_feedback TEXT, zewer_comments TEXT, created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS ubos (
            id INTEGER PRIMARY KEY AUTOINCREMENT, company_id INTEGER NOT NULL,
            position TEXT, share_percentage REAL, person_name TEXT NOT NULL,
            nationality TEXT, residential_status TEXT, passport_no TEXT,
            passport_expiry DATE, emirates_id TEXT, emirates_id_expiry DATE,
            doc_status TEXT DEFAULT 'Incompleted', verified_by TEXT,
            verified_date DATE, followup_details TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS dropdowns (
            id INTEGER PRIMARY KEY AUTOINCREMENT, field_name TEXT NOT NULL,
            value TEXT NOT NULL, description TEXT, is_active INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, UNIQUE(field_name, value));
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL,
            description TEXT, assigned_to INTEGER, created_by INTEGER,
            company_id INTEGER, priority TEXT DEFAULT 'normal',
            status TEXT DEFAULT 'todo', due_date DATE,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT, company_id INTEGER NOT NULL,
            doc_type TEXT NOT NULL, file_name TEXT NOT NULL, file_url TEXT NOT NULL,
            public_id TEXT, uploaded_by INTEGER, notes TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
    ''')
    for col in ['contact_number','mobile']:
        try: conn.execute(f'ALTER TABLE users ADD COLUMN {col} TEXT')
        except: pass
    for col in ['contact_person_name','contact_person_number','account_manager',
                'whatsapp_number','deal_after_vat','registration_screening_tool']:
        try: conn.execute(f'ALTER TABLE companies ADD COLUMN {col} TEXT')
        except: pass
    try: conn.execute('ALTER TABLE dropdowns ADD COLUMN description TEXT')
    except: pass
    if not conn.execute('SELECT id FROM users WHERE email=?',('admin@zewer.ae',)).fetchone():
        conn.execute('INSERT INTO users (email,password_hash,name,role) VALUES (?,?,?,?)',
            ('admin@zewer.ae', generate_password_hash('Admin@123'), 'Administrator', 'admin'))
        conn.execute('INSERT INTO users (email,password_hash,name,role) VALUES (?,?,?,?)',
            ('compliance@zewer.ae', generate_password_hash('Compliance@123'), 'Compliance Officer', 'compliance'))
    if not conn.execute('SELECT id FROM dropdowns LIMIT 1').fetchone():
        _seed(conn)
    conn.commit()

def _seed(conn):
    data = {
        'AC STATUS':['Active','Inactive'],
        'NATURE':['Individual','Legal entity'],
        'TYPE OF CLIENT':['MainLand','Free Zone','Abroad','International Corporate'],
        'MODE OF AC':['Supplier','Customer','Bullion','Refinery','Logistics Co','Exchange','Bank','Insurance','Investor','Technical Services'],
        'RISK STATUS':['High','Medium','Low','Unspecified'],
        'DOC STATUS':['Completed','Incompleted'],
        'KYC STATUS':['New Kyc Updated','Kyc 2025 Updated','Kyc 2024 Updated','Kyc 2023 Updated','Kyc 2022 Updated','Kyc 2021 Updated','Kyc 2020 Updated','Kyc 2019 Updated','Kyc 2018 Updated','Kyc 2017 Updated','Kyc 2016 Updated','Not Updated'],
        'REGION':['Dubai','Abu Dhabi','Sharjah','Ajman','Ras Al Khaimah','Fujairah','Umm Al Quwain','Al Ain','West Bengal','United Kingdom','Canada','Pakistan','Malaysia','Singapore','Bahrain','Italy','REPUBLIC OF CONGO','HONG KONG','Saudi Arabia','India'],
        'FREEZONE':['Jabel Ali FZ','DMCC','Sharjah Saif Zone','Ajman FZ','Fujairah','Dubai Production City','N/A','Ras Al Khaimah Fz','Dubai Free Zone','Dubai Gold & Diamond Park'],
        'LEGAL TYPE':['Limited Liability Company(LLC)','Limited Liability Company (WLL)','Civil Company Professional','Foreign Company','Public Company','Privet Company','DMCC','Free Zone Limited Liability Company (FZ-LLC)','Sole Establishment','Partnership Company','Services Agency','Limited (LTD)','Individual Institution','Establishment','FZCO','FZE','FZC'],
        'ISSUING AUTHORITY':['Dubai Economy & Tourism','Department of Economic Development Ajman','Abu-Dhabi Department of Economic Development','DMCC','Saif Zone','Ajman FZ','Jebel Ali FZ','Dubai Development Authority','Government Of Sharjah Economic Development Department','Government of Ras Al-Khaimah Department of Economic Development','Department of Economic Development Dubai','Dubai Integrated Economic Zones Authority','Fujairah Municipality','Trade Development Authority Of Pakistan','UK HMRC','The Registrar Of Companies For England And Wales','Canada Revenue Agency','Kolkata Municipal Corporation','Abroad'],
        'ADDRESS PROOF TYPE':['Ejari','Tenancy','Electricity Bill','Gst Registration Certificate','Certification Of Incorporation','Certificate Of Registration For Value Added Tax','Vat Certificate','Telephone Bill','Title Deed','Certificate Of Enlistment','Association Of Article Details','Warehouse Lease Agreement','Not Required'],
        'VAT CERT':['Yes','No','Not Required'],
        'VAT DECLARATION':['Yes','No','Not Required'],
        'MOA':['Yes','No'], 'PEP':['Yes','No'],
        'UNDERTAKING':['Yes','No'], 'SOURCE OF FUND':['Yes','No'],
        'POSITION':['UBO','Authorized Person','Director','Manager','Partner'],
        'RESIDENTIAL STATUS':['Resident','Non Resident'],
        'COUNTRY':['United Arab Emirates','Saudi Arabia','Kuwait','Qatar','Bahrain','Oman','India','Pakistan','Bangladesh','Sri Lanka','Philippines','Malaysia','Singapore','China','Hong Kong','Jordan','Lebanon','Syria','Iraq','Yemen','Egypt','Libya','Nigeria','Ethiopia','Republic Of Congo','Turkey','Iran','Afghanistan','Algeria','Canada','United Kingdom','United States of America','France','Ireland','Italy','Germany','Armenia','Belize'],
        'TASK TEMPLATE':['Collect Updated Trade License','KYC Update Required','Address Proof Renewal','Passport Renewal Follow-up','Emirates ID Update','VAT Certificate Collection','Screening Review','MOA Collection','Undertaking Form','Source of Funds Verification','Risk Assessment Review','Annual KYC Review'],
    }
    pg = is_pg(conn)
    for field, vals in data.items():
        for v in vals:
            try:
                if pg:
                    x(conn, 'INSERT INTO dropdowns (field_name,value,is_active) VALUES (%s,%s,1) ON CONFLICT DO NOTHING', (field,v))
                else:
                    conn.execute('INSERT OR IGNORE INTO dropdowns (field_name,value,is_active) VALUES (?,?,1)', (field,v))
            except: pass

try:
    setup_db()
    print("DB ready")
except Exception as e:
    print(f"DB setup warning: {e}")

# ── HELPERS ─────────────────────────────────────────────────

def days_left(d):
    if not d: return None
    try: return (datetime.strptime(str(d)[:10],'%Y-%m-%d').date()-datetime.now().date()).days
    except: return None

def exp_status(d):
    if d is None: return 'unknown'
    if d<0: return 'expired'
    if d<=30: return 'critical'
    if d<=90: return 'warning'
    return 'valid'

def dropdowns():
    conn=get_db()
    rows=all_(conn,"SELECT field_name,value FROM dropdowns WHERE is_active=1 ORDER BY field_name,value")
    conn.close()
    dd={}
    for r in rows: dd.setdefault(r['field_name'],[]).append(r['value'])
    return dd

def login_required(f):
    @wraps(f)
    def d(*a,**k):
        if 'user_id' not in session: return redirect(url_for('login'))
        return f(*a,**k)
    return d

def admin_required(f):
    @wraps(f)
    def d(*a,**k):
        if 'user_id' not in session: return redirect(url_for('login'))
        if session.get('user_role')!='admin': return redirect(url_for('dashboard'))
        return f(*a,**k)
    return d

def compliance_required(f):
    @wraps(f)
    def d(*a,**k):
        if 'user_id' not in session: return redirect(url_for('login'))
        if session.get('user_role') not in ('admin','compliance'): return redirect(url_for('tasks'))
        return f(*a,**k)
    return d

# ── ROUTES ──────────────────────────────────────────────────

@app.route('/')
def index(): return redirect(url_for('dashboard') if 'user_id' in session else url_for('login'))

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method=='POST':
        d=request.get_json()
        conn=get_db()
        u=one(conn,'SELECT * FROM users WHERE email=?',(d.get('email'),))
        conn.close()
        if u and check_password_hash(u['password_hash'],d.get('password','')) and u['is_active']:
            session.update(user_id=u['id'],user_email=u['email'],user_name=u['name'],user_role=u['role'])
            return jsonify({'success':True})
        return jsonify({'success':False,'error':'Invalid credentials'}),401
    return render_template('login.html')

@app.route('/logout')
def logout(): session.clear(); return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    conn=get_db(); today=datetime.now().date()
    def c(sql,p=None): return cnt(conn,sql,p or [])
    total=c('SELECT COUNT(*) FROM companies')
    active=c('SELECT COUNT(*) FROM companies WHERE ac_status=?',('Active',))
    etl=c('SELECT COUNT(*) FROM companies WHERE trade_license_expiry<?',(today,))
    e30tl=c('SELECT COUNT(*) FROM companies WHERE trade_license_expiry BETWEEN ? AND ?',(today,today+timedelta(days=30)))
    e90tl=c('SELECT COUNT(*) FROM companies WHERE trade_license_expiry BETWEEN ? AND ?',(today,today+timedelta(days=90)))
    eap=c('SELECT COUNT(*) FROM companies WHERE address_proof_expiry<?',(today,))
    e30ap=c('SELECT COUNT(*) FROM companies WHERE address_proof_expiry BETWEEN ? AND ?',(today,today+timedelta(days=30)))
    epass=c('SELECT COUNT(*) FROM ubos WHERE passport_expiry<?',(today,))
    e30p=c('SELECT COUNT(*) FROM ubos WHERE passport_expiry BETWEEN ? AND ?',(today,today+timedelta(days=30)))
    risk_rows=all_(conn,'SELECT risk_status,COUNT(*) as c FROM companies GROUP BY risk_status')
    risk_bd={r['risk_status']:r['c'] for r in risk_rows}
    doc_rows=all_(conn,'SELECT doc_status,COUNT(*) as c FROM companies GROUP BY doc_status')
    doc_bd={r['doc_status']:r['c'] for r in doc_rows}
    urgent=all_(conn,'''SELECT id,ac_code,client_name,trade_license_expiry,address_proof_expiry,
        risk_status,mobile,whatsapp_number FROM companies
        WHERE (trade_license_expiry BETWEEN ? AND ?) OR (address_proof_expiry BETWEEN ? AND ?)
        ORDER BY trade_license_expiry LIMIT 10''',(today,today+timedelta(days=90),today,today+timedelta(days=90)))
    otasks=c("SELECT COUNT(*) FROM tasks WHERE status NOT IN ('done')")
    overtasks=c("SELECT COUNT(*) FROM tasks WHERE status NOT IN ('done','pending_close') AND due_date<?", (today,))
    dtasks=c("SELECT COUNT(*) FROM tasks WHERE status NOT IN ('done') AND due_date=?", (today,))
    ptasks=c("SELECT COUNT(*) FROM tasks WHERE status='pending_close'")
    utasks=all_(conn,"""SELECT t.*,u.name as assigned_name FROM tasks t
        LEFT JOIN users u ON t.assigned_to=u.id
        WHERE t.status NOT IN ('done') AND (t.priority IN ('urgent','high') OR t.due_date<=?)
        ORDER BY CASE t.priority WHEN 'urgent' THEN 1 WHEN 'high' THEN 2 ELSE 3 END,t.due_date LIMIT 8""",
        (today+timedelta(days=2),))
    conn.close()
    return render_template('dashboard.html',total_companies=total,active_companies=active,
        expired_tl=etl,expiring_30_tl=e30tl,expiring_90_tl=e90tl,
        expired_ap=eap,expiring_30_ap=e30ap,expired_pass=epass,expiring_30_pass=e30p,
        risk_breakdown=risk_bd,doc_breakdown=doc_bd,urgent_companies=urgent,
        open_tasks=otasks,overdue_tasks=overtasks,due_today=dtasks,pending_close=ptasks,
        urgent_tasks=utasks,today=str(today),days_left=days_left)

@app.route('/companies')
@compliance_required
def companies():
    conn=get_db()
    s=request.args.get('search',''); rf=request.args.get('risk','')
    sf=request.args.get('status',''); rgf=request.args.get('region','')
    q='SELECT * FROM companies WHERE 1=1'; p=[]
    if s:
        q+=' AND (client_name LIKE ? OR ac_code LIKE ? OR mobile LIKE ? OR trade_license_no LIKE ?)'
        p+=[f'%{s}%']*4
    if rf: q+=' AND risk_status=?'; p.append(rf)
    if sf: q+=' AND ac_status=?'; p.append(sf)
    if rgf: q+=' AND region=?'; p.append(rgf)
    rows=all_(conn,q+' ORDER BY created_at DESC',p or None)
    dd=dropdowns(); conn.close()
    cl=[]
    for c in rows:
        tl=days_left(c['trade_license_expiry']); ap=days_left(c['address_proof_expiry'])
        cl.append({**c,'tl_days':tl,'tl_status':exp_status(tl),'ap_days':ap,'ap_status':exp_status(ap),
                   'trade_license_expiry':str(c['trade_license_expiry']) if c['trade_license_expiry'] else None,
                   'address_proof_expiry':str(c['address_proof_expiry']) if c['address_proof_expiry'] else None})
    return render_template('companies.html',companies=cl,search=s,risk_filter=rf,
        status_filter=sf,region_filter=rgf,regions=dd.get('REGION',[]))

@app.route('/company/new')
@compliance_required
def company_new():
    conn=get_db()
    staff=all_(conn,'SELECT id,name FROM users WHERE is_active=1 ORDER BY name')
    conn.close()
    return render_template('company_form.html',dropdown_data=dropdowns(),company=None,ubos=[],edit=False,staff_users=staff)

@app.route('/company/<int:id>')
@compliance_required
def company_detail(id):
    conn=get_db()
    co=one(conn,'SELECT * FROM companies WHERE id=?',(id,))
    if not co: conn.close(); return redirect(url_for('companies'))
    ubos=all_(conn,'SELECT * FROM ubos WHERE company_id=? ORDER BY share_percentage DESC',(id,))
    conn.close()
    tl=days_left(co['trade_license_expiry']); ap=days_left(co['address_proof_expiry'])
    ul=[]
    for u in ubos:
        pd=days_left(u['passport_expiry']); ed=days_left(u['emirates_id_expiry'])
        ul.append({**u,'p_days':pd,'p_status':exp_status(pd),'e_days':ed,'e_status':exp_status(ed),
                   'passport_expiry':str(u['passport_expiry']) if u['passport_expiry'] else None,
                   'emirates_id_expiry':str(u['emirates_id_expiry']) if u['emirates_id_expiry'] else None})
    return render_template('company_detail.html',company=co,ubos=ul,
        tl_days=tl,tl_status=exp_status(tl),ap_days=ap,ap_status=exp_status(ap),
        today=str(datetime.now().date()))

@app.route('/company/<int:id>/edit')
@compliance_required
def company_edit(id):
    conn=get_db()
    co=one(conn,'SELECT * FROM companies WHERE id=?',(id,))
    if not co: conn.close(); return redirect(url_for('companies'))
    ubos=all_(conn,'SELECT * FROM ubos WHERE company_id=? ORDER BY share_percentage DESC',(id,))
    staff=all_(conn,'SELECT id,name FROM users WHERE is_active=1 ORDER BY name')
    conn.close()
    return render_template('company_form.html',dropdown_data=dropdowns(),company=co,ubos=ubos,edit=True,staff_users=staff)

def _cv(d):
    return (d.get('ac_opening_date') or None,d.get('ac_status','Active'),d.get('active_till_year'),
        d.get('nature'),d.get('type_of_client'),d.get('name_of_freezone'),d.get('mode_of_ac'),
        d.get('country_of_incorporation'),d.get('region'),d.get('address'),d.get('telephone'),
        d.get('mobile'),d.get('whatsapp_number'),d.get('email_id'),d.get('contact_person_name'),
        d.get('contact_person_number'),d.get('account_manager'),d.get('address_proof_type'),
        d.get('address_proof_expiry') or None,d.get('kyc_status'),d.get('trade_license_no'),
        d.get('issuing_authority'),d.get('legal_type'),d.get('incorporation_date') or None,
        d.get('trade_license_expiry') or None,d.get('tax_no_trn'),d.get('vat_cert'),
        d.get('vat_declaration'),d.get('deal_after_vat'),int(d.get('num_beneficial_owners') or 0),
        d.get('moa'),d.get('pep'),d.get('undertaking'),d.get('source_of_fund'),
        d.get('software_updation'),d.get('doc_status','Incompleted'),d.get('screening_date') or None,
        d.get('registration_screening_tool'),d.get('risk_status','Unspecified'),d.get('verified_by'),
        d.get('verified_date') or None,d.get('followup_details'),d.get('crowe_feedback'),d.get('zewer_comments'))

def _save_ubos(conn,cid,ubos):
    x(conn,'DELETE FROM ubos WHERE company_id=?',(cid,))
    for u in ubos:
        if not u.get('person_name'): continue
        x(conn,'''INSERT INTO ubos (company_id,position,share_percentage,person_name,nationality,
            residential_status,passport_no,passport_expiry,emirates_id,emirates_id_expiry,
            doc_status,verified_by,verified_date,followup_details) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (cid,u.get('position'),u.get('share_percentage') or None,u.get('person_name'),
             u.get('nationality'),u.get('residential_status'),u.get('passport_no'),
             u.get('passport_expiry') or None,u.get('emirates_id'),u.get('emirates_id_expiry') or None,
             u.get('doc_status','Incompleted'),u.get('verified_by'),u.get('verified_date') or None,u.get('followup_details')))

@app.route('/api/company/add', methods=['POST'])
@compliance_required
def api_add_company():
    d=request.get_json()
    try:
        conn=get_db()
        x(conn,'''INSERT INTO companies (ac_code,client_name,ac_opening_date,ac_status,active_till_year,
            nature,type_of_client,name_of_freezone,mode_of_ac,country_of_incorporation,region,address,
            telephone,mobile,whatsapp_number,email_id,contact_person_name,contact_person_number,account_manager,
            address_proof_type,address_proof_expiry,kyc_status,trade_license_no,issuing_authority,legal_type,
            incorporation_date,trade_license_expiry,tax_no_trn,vat_cert,vat_declaration,deal_after_vat,
            num_beneficial_owners,moa,pep,undertaking,source_of_fund,software_updation,doc_status,
            screening_date,registration_screening_tool,risk_status,verified_by,verified_date,
            followup_details,crowe_feedback,zewer_comments,created_by)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (d.get('ac_code'),d.get('client_name'))+_cv(d)+(session.get('user_id'),))
        cid=lastid(conn)
        _save_ubos(conn,cid,d.get('ubos',[]))
        commit(conn); conn.close()
        return jsonify({'success':True,'id':cid})
    except Exception as e: return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/company/<int:id>/edit', methods=['POST'])
@compliance_required
def api_edit_company(id):
    d=request.get_json()
    try:
        conn=get_db()
        x(conn,'''UPDATE companies SET client_name=?,ac_opening_date=?,ac_status=?,active_till_year=?,
            nature=?,type_of_client=?,name_of_freezone=?,mode_of_ac=?,country_of_incorporation=?,region=?,
            address=?,telephone=?,mobile=?,whatsapp_number=?,email_id=?,contact_person_name=?,
            contact_person_number=?,account_manager=?,address_proof_type=?,address_proof_expiry=?,
            kyc_status=?,trade_license_no=?,issuing_authority=?,legal_type=?,incorporation_date=?,
            trade_license_expiry=?,tax_no_trn=?,vat_cert=?,vat_declaration=?,deal_after_vat=?,
            num_beneficial_owners=?,moa=?,pep=?,undertaking=?,source_of_fund=?,software_updation=?,
            doc_status=?,screening_date=?,registration_screening_tool=?,risk_status=?,verified_by=?,
            verified_date=?,followup_details=?,crowe_feedback=?,zewer_comments=?,
            updated_at=CURRENT_TIMESTAMP WHERE id=?''',
            (d.get('client_name'),)+_cv(d)+(id,))
        _save_ubos(conn,id,d.get('ubos',[]))
        commit(conn); conn.close()
        return jsonify({'success':True})
    except Exception as e: return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/company/<int:id>/delete', methods=['POST'])
@compliance_required
def api_delete_company(id):
    try:
        conn=get_db(); x(conn,'DELETE FROM companies WHERE id=?',(id,))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e: return jsonify({'success':False,'error':str(e)}),500

@app.route('/alerts')
@compliance_required
def alerts():
    conn=get_db(); today=datetime.now().date(); cutoff=today-timedelta(days=30)
    tl=all_(conn,'SELECT id,ac_code,client_name,mobile,whatsapp_number,trade_license_expiry,risk_status,region FROM companies WHERE trade_license_expiry>=? ORDER BY trade_license_expiry',(cutoff,))
    ap=all_(conn,'SELECT id,ac_code,client_name,mobile,whatsapp_number,address_proof_expiry,address_proof_type,risk_status FROM companies WHERE address_proof_expiry>=? ORDER BY address_proof_expiry',(cutoff,))
    ubos=all_(conn,'''SELECT u.id,u.person_name,u.passport_no,u.passport_expiry,u.emirates_id,u.emirates_id_expiry,
        c.id as company_id,c.client_name,c.ac_code FROM ubos u JOIN companies c ON u.company_id=c.id
        WHERE u.passport_expiry>=? OR u.emirates_id_expiry>=? ORDER BY u.passport_expiry''',(cutoff,cutoff))
    conn.close()
    return render_template('alerts.html',tl_expiring=tl,ap_expiring=ap,ubo_expiring=ubos,
        today=str(today),days_left=days_left,expiry_status=exp_status)

@app.route('/reports')
@compliance_required
def reports():
    conn=get_db(); today=datetime.now().date()
    df=request.args.get('from',''); dt=request.args.get('to','')
    w='1=1'; p=[]
    if df: w+=' AND created_at >= ?'; p.append(df)
    if dt: w+=' AND created_at <= ?'; p.append(dt+' 23:59:59')
    total=cnt(conn,f'SELECT COUNT(*) FROM companies WHERE {w}',p or None)
    def q(sql): return all_(conn,sql,p or None)
    res=render_template('reports.html',
        risk_data=q(f'SELECT risk_status,COUNT(*) as c FROM companies WHERE {w} GROUP BY risk_status'),
        doc_data=q(f'SELECT doc_status,COUNT(*) as c FROM companies WHERE {w} GROUP BY doc_status'),
        type_data=q(f'SELECT type_of_client,COUNT(*) as c FROM companies WHERE {w} GROUP BY type_of_client ORDER BY c DESC'),
        region_data=q(f'SELECT region,COUNT(*) as c FROM companies WHERE {w} GROUP BY region ORDER BY c DESC LIMIT 10'),
        kyc_data=q(f'SELECT kyc_status,COUNT(*) as c FROM companies WHERE {w} GROUP BY kyc_status ORDER BY c DESC'),
        mode_data=q(f'SELECT mode_of_ac,COUNT(*) as c FROM companies WHERE {w} GROUP BY mode_of_ac ORDER BY c DESC'),
        total=total,date_from=df,date_to=dt,
        expired_tl=cnt(conn,'SELECT COUNT(*) FROM companies WHERE trade_license_expiry<?',(today,)),
        expiring_30=cnt(conn,'SELECT COUNT(*) FROM companies WHERE trade_license_expiry BETWEEN ? AND ?',(today,today+timedelta(days=30))),
        expiring_90=cnt(conn,'SELECT COUNT(*) FROM companies WHERE trade_license_expiry BETWEEN ? AND ?',(today,today+timedelta(days=90))))
    conn.close(); return res

@app.route('/settings')
@admin_required
def settings():
    conn=get_db()
    users=all_(conn,'SELECT id,email,name,role,is_active,contact_number FROM users ORDER BY role,name')
    dds=all_(conn,'SELECT * FROM dropdowns WHERE is_active=1 ORDER BY field_name,value')
    conn.close()
    groups={}
    for d in dds: groups.setdefault(d['field_name'],[]).append(d)
    return render_template('settings.html',users=users,dropdown_groups=groups)

@app.route('/api/user/add',methods=['POST'])
@admin_required
def api_add_user():
    d=request.get_json()
    try:
        conn=get_db()
        x(conn,'INSERT INTO users (email,password_hash,name,role,contact_number,is_active) VALUES (?,?,?,?,?,1)',
          (d.get('email'),generate_password_hash(d.get('password','')),d.get('name'),d.get('role'),d.get('contact_number')))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e: return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/user/<int:id>/edit',methods=['POST'])
@admin_required
def api_edit_user(id):
    d=request.get_json()
    try:
        conn=get_db()
        if d.get('password'):
            x(conn,'UPDATE users SET name=?,email=?,role=?,contact_number=?,password_hash=? WHERE id=?',
              (d.get('name'),d.get('email'),d.get('role'),d.get('contact_number'),generate_password_hash(d.get('password')),id))
        else:
            x(conn,'UPDATE users SET name=?,email=?,role=?,contact_number=? WHERE id=?',
              (d.get('name'),d.get('email'),d.get('role'),d.get('contact_number'),id))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e: return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/user/<int:id>/toggle',methods=['POST'])
@admin_required
def api_toggle_user(id):
    try:
        conn=get_db(); u=one(conn,'SELECT is_active FROM users WHERE id=?',(id,))
        if u: x(conn,'UPDATE users SET is_active=? WHERE id=?',(0 if u['is_active'] else 1,id))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e: return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/user/<int:id>/delete',methods=['POST'])
@admin_required
def api_delete_user(id):
    try:
        conn=get_db(); x(conn,'DELETE FROM users WHERE id=?',(id,))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e: return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/dropdown/add',methods=['POST'])
@admin_required
def api_add_dropdown():
    d=request.get_json()
    try:
        conn=get_db()
        x(conn,'INSERT INTO dropdowns (field_name,value,is_active) VALUES (?,?,1)',(d.get('field_name'),d.get('value')))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e: return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/dropdown/<int:id>/delete',methods=['POST'])
@admin_required
def api_delete_dropdown(id):
    try:
        conn=get_db(); x(conn,'DELETE FROM dropdowns WHERE id=?',(id,))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e: return jsonify({'success':False,'error':str(e)}),500

@app.route('/tasks')
@login_required
def tasks():
    conn=get_db(); uid=session.get('user_id'); role=session.get('user_role')
    base='''SELECT t.*,u.name as assigned_name,u.mobile as assigned_mobile,
        cu.name as created_by_name,c.client_name as company_name,c.ac_code
        FROM tasks t LEFT JOIN users u ON t.assigned_to=u.id
        LEFT JOIN users cu ON t.created_by=cu.id
        LEFT JOIN companies c ON t.company_id=c.id'''
    order=" ORDER BY CASE t.priority WHEN 'urgent' THEN 1 WHEN 'high' THEN 2 WHEN 'normal' THEN 3 ELSE 4 END,t.due_date"
    rows=all_(conn,base+(' WHERE t.assigned_to=?' if role=='staff' else '')+order,(uid,) if role=='staff' else None)
    users=all_(conn,'SELECT id,name,role,mobile FROM users WHERE is_active=1 ORDER BY name')
    cos=all_(conn,'SELECT id,ac_code,client_name FROM companies ORDER BY client_name')
    tmpls=all_(conn,"SELECT value,description FROM dropdowns WHERE field_name='TASK TEMPLATE' AND is_active=1 ORDER BY value")
    conn.close()
    tl=[]
    for t in rows:
        d=days_left(t['due_date'])
        tl.append({**t,'priority':t['priority'] or 'normal','status':t['status'] or 'todo',
                   'due_date':str(t['due_date']) if t['due_date'] else None,
                   'days_until_due':d,'is_overdue':(d is not None and d<0)})
    return render_template('tasks.html',tasks=tl,all_users=users,all_companies=cos,task_templates=tmpls)

@app.route('/api/task/add',methods=['POST'])
@login_required
def api_add_task():
    d=request.get_json()
    try:
        conn=get_db()
        x(conn,'INSERT INTO tasks (title,description,assigned_to,created_by,company_id,priority,due_date,status) VALUES (?,?,?,?,?,?,?,?)',
          (d.get('title'),d.get('description'),d.get('assigned_to') or None,session.get('user_id'),
           d.get('company_id') or None,d.get('priority','normal'),d.get('due_date') or None,'todo'))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e: return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/task/<int:id>/edit',methods=['POST'])
@login_required
def api_edit_task(id):
    d=request.get_json()
    try:
        conn=get_db()
        x(conn,'UPDATE tasks SET title=?,description=?,assigned_to=?,company_id=?,priority=?,due_date=?,updated_at=CURRENT_TIMESTAMP WHERE id=?',
          (d.get('title'),d.get('description'),d.get('assigned_to') or None,d.get('company_id') or None,
           d.get('priority','normal'),d.get('due_date') or None,id))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e: return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/task/<int:id>/status',methods=['POST'])
@login_required
def api_task_status(id):
    d=request.get_json()
    try:
        conn=get_db()
        x(conn,'UPDATE tasks SET status=?,updated_at=CURRENT_TIMESTAMP WHERE id=?',(d.get('status'),id))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e: return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/task/<int:id>/delete',methods=['POST'])
@login_required
def api_delete_task(id):
    try:
        conn=get_db(); x(conn,'DELETE FROM tasks WHERE id=?',(id,))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e: return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/company/<int:cid>/documents')
@compliance_required
def api_get_documents(cid):
    conn=get_db()
    docs=all_(conn,'''SELECT d.*,u.name as uploader_name FROM documents d
        LEFT JOIN users u ON d.uploaded_by=u.id WHERE d.company_id=? ORDER BY d.created_at DESC''',(cid,))
    conn.close(); return jsonify(docs)

@app.route('/api/company/<int:cid>/upload',methods=['POST'])
@compliance_required
def api_upload_document(cid):
    if 'file' not in request.files: return jsonify({'success':False,'error':'No file'}),400
    file=request.files['file']
    if not file.filename: return jsonify({'success':False,'error':'No filename'}),400
    if not HAS_CLD or not os.getenv('CLOUDINARY_CLOUD_NAME'):
        return jsonify({'success':False,'error':'Cloudinary not configured'}),500
    try:
        r=cloudinary.uploader.upload(file,folder=f'zewer_crm/co_{cid}',resource_type='auto',use_filename=True,unique_filename=True)
        conn=get_db()
        x(conn,'INSERT INTO documents (company_id,doc_type,file_name,file_url,public_id,uploaded_by,notes) VALUES (?,?,?,?,?,?,?)',
          (cid,request.form.get('doc_type','General'),file.filename,r['secure_url'],r['public_id'],session.get('user_id'),request.form.get('notes','')))
        commit(conn); conn.close()
        return jsonify({'success':True,'url':r['secure_url'],'name':file.filename})
    except Exception as e: return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/document/<int:did>/delete',methods=['POST'])
@compliance_required
def api_delete_document(did):
    try:
        conn=get_db(); doc=one(conn,'SELECT public_id FROM documents WHERE id=?',(did,))
        if doc and doc.get('public_id') and HAS_CLD and os.getenv('CLOUDINARY_CLOUD_NAME'):
            try: cloudinary.uploader.destroy(doc['public_id'],resource_type='raw')
            except: pass
        x(conn,'DELETE FROM documents WHERE id=?',(did,))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e: return jsonify({'success':False,'error':str(e)}),500

@app.route('/export/template')
@compliance_required
def export_template():
    if not HAS_XL: return "openpyxl not installed",500
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Companies"
    ws.append(['ac_code','client_name','ac_opening_date','ac_status','nature','type_of_client',
               'name_of_freezone','mode_of_ac','country_of_incorporation','region','address',
               'telephone','mobile','whatsapp_number','email_id','contact_person_name',
               'contact_person_number','account_manager','address_proof_type','address_proof_expiry',
               'trade_license_no','issuing_authority','legal_type','incorporation_date',
               'trade_license_expiry','tax_no_trn','vat_cert','vat_declaration','num_beneficial_owners',
               'moa','pep','undertaking','source_of_fund','doc_status','risk_status','kyc_status',
               'verified_by','followup_details','zewer_comments'])
    ws.append(['TJ5003','SAMPLE COMPANY LLC','2020-01-01','Active','Legal entity','MainLand','N/A',
               'Supplier','United Arab Emirates','Dubai','Unit 101, Gold Souq, Dubai','+97142258019',
               '+971506594165','+971506594165','info@sample.com','Ahmed Ali','+971501234567','Jaseel',
               'Ejari','2025-12-31','534230','Dubai Economy & Tourism','Limited Liability Company(LLC)',
               '2019-06-01','2025-12-31','100003063300003','Yes','Yes',2,'Yes','Yes','Yes','Yes',
               'Completed','Medium','Kyc 2025 Updated','Jaseel','',''])
    out=io.BytesIO(); wb.save(out); out.seek(0)
    return send_file(out,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,download_name='zewer_company_template.xlsx')

@app.route('/export/companies')
@compliance_required
def export_companies():
    scope=request.args.get('scope','all'); fmt=request.args.get('format','csv')
    conn=get_db()
    q='SELECT * FROM companies WHERE 1=1'
    if scope=='active': q+=' AND ac_status=\'Active\''
    elif scope=='inactive': q+=' AND ac_status=\'Inactive\''
    elif scope=='high': q+=' AND risk_status=\'High\''
    elif scope=='medium': q+=' AND risk_status=\'Medium\''
    elif scope=='low': q+=' AND risk_status=\'Low\''
    elif scope=='incomplete': q+=' AND doc_status=\'Incompleted\''
    rows=all_(conn,q); conn.close()
    fname=f'companies_{scope}_{datetime.now().strftime("%Y%m%d_%H%M%S")}'
    if fmt=='xlsx' and HAS_XL:
        wb=openpyxl.Workbook(); ws=wb.active
        if rows: ws.append(list(rows[0].keys())); [ws.append([str(v) if v else '' for v in r.values()]) for r in rows]
        out=io.BytesIO(); wb.save(out); out.seek(0)
        return send_file(out,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,download_name=fname+'.xlsx')
    out=io.StringIO(); w=csv.writer(out)
    if rows: w.writerow(rows[0].keys()); [w.writerow(list(r.values())) for r in rows]
    out.seek(0)
    return send_file(io.BytesIO(out.getvalue().encode()),mimetype='text/csv',as_attachment=True,download_name=fname+'.csv')

@app.route('/export/report')
@compliance_required
def export_report():
    rt=request.args.get('type','all'); fmt=request.args.get('format','xlsx')
    conn=get_db()
    if rt=='expiry': rows=all_(conn,'SELECT ac_code,client_name,trade_license_no,trade_license_expiry,address_proof_expiry,risk_status,region,account_manager FROM companies ORDER BY trade_license_expiry')
    elif rt=='kyc': rows=all_(conn,'SELECT ac_code,client_name,kyc_status,doc_status,risk_status,verified_by,verified_date FROM companies ORDER BY kyc_status')
    elif rt=='risk': rows=all_(conn,'SELECT ac_code,client_name,risk_status,doc_status,kyc_status,region,account_manager FROM companies ORDER BY risk_status')
    else: rows=all_(conn,'SELECT * FROM companies ORDER BY client_name')
    conn.close()
    fname=f'report_{rt}_{datetime.now().strftime("%Y%m%d_%H%M%S")}'
    if fmt=='xlsx' and HAS_XL:
        wb=openpyxl.Workbook(); ws=wb.active; ws.title=rt.upper()
        if rows: ws.append(list(rows[0].keys())); [ws.append([str(v) if v else '' for v in r.values()]) for r in rows]
        out=io.BytesIO(); wb.save(out); out.seek(0)
        return send_file(out,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,download_name=fname+'.xlsx')
    out=io.StringIO(); w=csv.writer(out)
    if rows: w.writerow(rows[0].keys()); [w.writerow(list(r.values())) for r in rows]
    out.seek(0)
    return send_file(io.BytesIO(out.getvalue().encode()),mimetype='text/csv',as_attachment=True,download_name=fname+'.csv')

@app.route('/api/import/companies',methods=['POST'])
@compliance_required
def api_import_companies():
    if not HAS_XL: return jsonify({'success':False,'error':'openpyxl not installed'}),500
    if 'file' not in request.files: return jsonify({'success':False,'error':'No file'}),400
    try:
        wb=openpyxl.load_workbook(request.files['file']); ws=wb.active
        headers=[str(c.value).strip() if c.value else '' for c in ws[1]]
        imported=0; skipped=0; conn=get_db()
        for row in ws.iter_rows(min_row=2,values_only=True):
            if not any(row): continue
            rd={headers[i]:(str(row[i]).strip() if row[i] is not None else '') for i in range(min(len(headers),len(row)))}
            ac=rd.get('ac_code','').strip(); nm=rd.get('client_name','').strip()
            if not ac or not nm: continue
            if one(conn,'SELECT id FROM companies WHERE ac_code=?',(ac,)): skipped+=1; continue
            try:
                x(conn,'''INSERT INTO companies (ac_code,client_name,ac_status,risk_status,doc_status,
                    nature,type_of_client,region,telephone,mobile,email_id,trade_license_no,
                    trade_license_expiry,address_proof_expiry,kyc_status,account_manager,created_by)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                    (ac,nm,rd.get('ac_status','Active'),rd.get('risk_status','Unspecified'),
                     rd.get('doc_status','Incompleted'),rd.get('nature'),rd.get('type_of_client'),
                     rd.get('region'),rd.get('telephone'),rd.get('mobile'),rd.get('email_id'),
                     rd.get('trade_license_no'),rd.get('trade_license_expiry') or None,
                     rd.get('address_proof_expiry') or None,rd.get('kyc_status'),
                     rd.get('account_manager'),session.get('user_id')))
                imported+=1
            except: skipped+=1
        commit(conn); conn.close()
        return jsonify({'success':True,'imported':imported,'skipped':skipped})
    except Exception as e: return jsonify({'success':False,'error':str(e)}),500

if __name__=='__main__':
    app.run(debug=False,host='0.0.0.0',port=int(os.getenv('PORT',8000)))
