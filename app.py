from flask import Flask, render_template, request, jsonify, session, redirect, url_for, send_file
from werkzeug.security import check_password_hash, generate_password_hash
import os, io, csv, sqlite3, logging
from datetime import datetime, timedelta, timezone

DUBAI_TZ = timezone(timedelta(hours=4))
def dubai_today():
    """Returns today's date in Dubai (UTC+4) regardless of server timezone."""
    return datetime.now(DUBAI_TZ).date()
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

# ── RISK ASSESSMENT LOOKUPS ──────────────────────────────────
RISK_LOOKUPS = {
    'countries': {'AFGHANISTAN':3,'ALAND ISLANDS':2,'ALBANIA':2,'ALGERIA':3,'AMERICAN SAMOA':2,'ANDORRA':2,'ANGOLA':3,'ANGUILLA':2,'ANTIGUA & BARBUDA':2,'ARGENTINA':2,'ARMENIA':2,'ARUBA':2,'AUSTRALIA':1,'AUSTRIA':2,'AZERBAIJAN':2,'BAHAMAS':2,'BAHRAIN':2,'BANGLADESH':2,'BARBADOS':2,'BELARUS':3,'BELGIUM':2,'BELIZE':2,'BENIN':2,'BERMUDA':2,'BHUTAN':1,'BOLIVIA':3,'BOSNIA AND HERZEGOVINA':2,'BOTSWANA':2,'BRAZIL':2,'BRITISH VIRGIN ISLANDS':2,'BRUNEI':1,'BULGARIA':2,'BURKINA FASO':3,'BURUNDI':3,'CAMBODIA':3,'CAMEROON':2,'CANADA':1,'CAPE VERDE':2,'CAYMAN ISLANDS':2,'CENTRAL AFRICAN REPUBLIC':3,'CHAD':3,'CHILE':2,'CHINA':2,'CHRISTMAS ISLAND':2,'COCOS (KEELING) ISLANDS':2,'COLOMBIA':2,'COMOROS':3,'CONGO':3,'COOK ISLANDS':2,'COSTA RICA':2,'COTE D IVOIRE':3,'CROATIA':2,'CUBA':3,'CURACAO':2,'CYPRUS':2,'CZECH REPUBLIC':2,'DEMOCRATIC REPUBLIC OF CONGO':3,'DENMARK':1,'DJIBOUTI':2,'DOMINICA':2,'DOMINICAN REPUBLIC':2,'EAST TIMOR':2,'ECUADOR':2,'EGYPT':3,'EL SALVADOR':3,'EQUATORIAL GUINEA':3,'ERITREA':3,'ESTONIA':2,'ETHIOPIA':3,'FALKLAND ISLANDS':2,'FAROE ISLANDS':2,'FIJI':2,'FINLAND':1,'FRANCE':1,'FRENCH GUIANA':2,'FRENCH POLYNESIA':2,'GABON':3,'GAMBIA':3,'GEORGIA':2,'GERMANY':1,'GHANA':3,'GIBRALTAR':2,'GREECE':2,'GREENLAND':2,'GRENADA':2,'GUADELOUPE':2,'GUAM':2,'GUATEMALA':3,'GUERNSEY':2,'GUINEA':3,'GUINEA BISSAU':3,'GUYANA':2,'HAITI':3,'HONDURAS':3,'HONG KONG':1,'HUNGARY':2,'ICELAND':1,'INDIA':2,'INDONESIA':2,'IRAN':3,'IRAQ':3,'IRELAND':1,'ISLE OF MAN':2,'ISRAEL':2,'ITALY':2,'JAMAICA':2,'JAPAN':1,'JERSEY':2,'JORDAN':2,'KAZAKHSTAN':3,'KENYA':3,'KIRIBATI':2,'KOREA NORTH':3,'KOREA SOUTH':1,'KOSOVO':2,'KUWAIT':2,'KYRGYZSTAN':3,'LAOS':3,'LATVIA':2,'LEBANON':3,'LESOTHO':2,'LIBERIA':3,'LIBYA':3,'LIECHTENSTEIN':2,'LITHUANIA':2,'LUXEMBOURG':1,'MACAO':2,'MACEDONIA':2,'MADAGASCAR':2,'MALAWI':2,'MALAYSIA':2,'MALDIVES':2,'MALI':3,'MALTA':2,'MARSHALL ISLANDS':2,'MARTINIQUE':2,'MAURITANIA':3,'MAURITIUS':2,'MAYOTTE':2,'MEXICO':3,'MICRONESIA':2,'MOLDOVA':3,'MONACO':2,'MONGOLIA':2,'MONTENEGRO':2,'MONTSERRAT':2,'MOROCCO':2,'MOZAMBIQUE':3,'MYANMAR':3,'NAMIBIA':2,'NAURU':2,'NEPAL':2,'NETHERLANDS':1,'NEW CALEDONIA':2,'NEW ZEALAND':1,'NICARAGUA':3,'NIGER':3,'NIGERIA':3,'NIUE':2,'NORFOLK ISLAND':2,'NORTHERN MARIANA ISLANDS':2,'NORWAY':1,'OMAN':2,'PAKISTAN':3,'PALAU':2,'PALESTINE':3,'PANAMA':3,'PAPUA NEW GUINEA':3,'PARAGUAY':2,'PERU':3,'PHILIPPINES':2,'PITCAIRN':2,'POLAND':2,'PORTUGAL':2,'PUERTO RICO':2,'QATAR':2,'REUNION':2,'ROMANIA':2,'RUSSIAN FEDERATION':3,'RWANDA':3,'SAINT BARTHELEMY':2,'SAINT HELENA':2,'SAINT KITTS & NEVIS':2,'SAINT LUCIA':2,'SAINT MARTIN':2,'SAINT PIERRE & MIQUELON':2,'SAINT VINCENT & THE GRENADINES':2,'SAMOA':2,'SAN MARINO':2,'SAO TOME & PRINCIPE':2,'SAUDI ARABIA':2,'SENEGAL':2,'SERBIA':2,'SEYCHELLES':2,'SIERRA LEONE':3,'SINGAPORE':1,'SINT MAARTEN':2,'SLOVAKIA':2,'SLOVENIA':2,'SOLOMON ISLANDS':2,'SOMALIA':3,'SOUTH AFRICA':2,'SOUTH SUDAN':3,'SPAIN':2,'SRI LANKA':2,'SUDAN':3,'SURINAME':2,'SWAZILAND':2,'SWEDEN':1,'SWITZERLAND':1,'SYRIA':3,'TAIWAN':2,'TAJIKISTAN':3,'TANZANIA':3,'THAILAND':2,'TIMOR-LESTE':2,'TOGO':3,'TOKELAU':2,'TONGA':2,'TRINIDAD AND TOBAGO':2,'TUNISIA':2,'TURKEY':2,'TURKMENISTAN':3,'TURKS & CAICOS ISLANDS':2,'TUVALU':2,'UGANDA':3,'UKRAINE':3,'UNITED ARAB EMIRATES':2,'UNITED KINGDOM':1,'UNITED STATES':1,'URUGUAY':2,'UZBEKISTAN':3,'VANUATU':2,'VATICAN':2,'VENEZUELA':3,'VIETNAM':2,'VIRGIN ISLANDS':2,'WALLIS & FUTUNA':2,'WESTERN SAHARA':3,'YEMEN':3,'ZAMBIA':3,'ZIMBABWE':3,'NOT APPLICABLE':1},
    'nature_of_business': {'Accounting Firm':2,'Advertising':2,'Airlines':2,'Arms Dealer':3,'Auction':2,'Auditors / Audit Firm':2,'Automobile Service':1,'Banking':3,'Brokerage Business':3,'Building Contracting':2,'Building Maintenance':1,'Building Management':1,'Car Business (New & Used Sales)':2,'Cargo & Logistics':3,'Carpentry':1,'Carpets Trading':1,'Catering':1,'Charities/Non-Profit Organizations':3,'Chemicals Trading':3,'Cleaning and Maintenance':1,'Commission Agent':2,'Communication Department':2,'Computer/Laptop Sales':1,'Computer/Laptop Service & Repairs':1,'Construction Items Shop':1,'Consultancy Services':3,'Corporate Services':3,'Cosmetics Trading':1,'Credit provider':2,'Dealers in Precious Metals & Stones':3,'Engineering':2,'Exchange Service':3,'Fast Moving Consumer Goods (FMCG)':1,'Financial Consultant':3,'Fish Trading':1,'Food Supplier':1,'Foreign Exchange':3,'Fruits & Vegetable Trading':1,'Fuel Station':1,'Furniture Trading':2,'General Trading Company':2,'Gift Shop':1,'Gold Smith':3,'Grocery':1,'Healthcare':1,'Heavy Equipment Leasing':2,'HMO/Travel Clinic':1,'Hotel':1,'HR / Recruitment Consultant':2,'Import/Export':2,'Insurance':2,'Interior Decoration':1,'Internet Service Provider':1,'Investment Company':3,'Jewellery Trading':3,'Laundry Services':1,'Law Firm':2,'Leather Items Trading':1,'Leasing Company':2,'Logistics':3,'Machinery Supplier':1,'Marble Trading':1,'Marketing':1,'Medicals/Pharmaceutical':1,'Metals & Minerals Trading':2,'Mining':1,'Mobile Phone Sales':1,'Mobile Phone Service & Repairs':1,'Money Lender':3,'Motor Vehicle Agent':2,'Motor Vehicle Spare Parts':1,'Nursery/School':1,'Nut Trading':1,'Oil Gas Supply':3,'Packaging Material Supplier':1,'Paper Trading':1,'Perfumes Trading':1,'Petrol Station':1,'Pharmaceutical':1,'Photography':1,'Plastic Item Trading':1,'Printing Press':1,'Professional Services':2,'Property Leasing':1,'Property Rental':1,'Publishing':1,'Real Estate':2,'Restaurant':1,'Retail':1,'Retail Gems & Jewellery':3,'Retail Supermarket':1,'Sales Agent/Distributor':2,'Sanitary Ware Trading':1,'Security Services':1,'Shell Company':3,'Shoe Shop':1,'Showroom':1,'Spare Parts Trading':1,'Sports':1,'Stationery Shop':1,'Steel Trading':2,'Stone Supplier':2,'Transportation':1,'Travel & Tourism':1,'Travel Agency':2,'Trading Company':2,'Tube Well Contractor':1,'Wholesale Jewellery':3,'Wholesale Trade':2,'Window Material Trading':1},
    'products': {'Hotels & Resorts':3,'Luxury Villas':3,'Penthouses':3,'Residential Apartments':2,'Retail Shops':2,'Townhouses':2,'Vacant Land':3,'Warehouses':1,'Office Spaces':2,'Permanent Desk':2,'Virtual Office':3,'PO Box':3,'Antique & Vintage Jewellery':2,'Bullion':3,'Collectible Coins':2,'Designer & Branded Jewellery':3,'Fine Jewellery':2,'Investment-Grade Precious Metals':3,'Luxury Watches':3,'Polished Diamonds':3,'Raw Precious Metals':2,'Uncut & Rough Gemstones':3,'Wholesale Jewellery':3},
    'amount_ranges_company': {'Amount up to AED 55,000':1,'Amount upto AED 500,000':2,'Amount more than AED 500,000':3},
    'amount_ranges_individual': {'Amount up to AED 100,000':1,'Amount more than AED 100,000':3},
    'payment_methods': {'Cash':3,'Cheque':1,'Crypto':3,'Bank Transfer (Local)':2,'Bank Transfer (International)':3,'Debit/Credit Card':2,'Gold to Gold Exchange':3,'All of the Above':3},
    'third_party': {'Yes':3,'No':1},
    'yes_no': {'Yes':3,'No':1},
}

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'zewer-aml-secret-2026')

# ── SESSION & SECURITY CONFIG ────────────────────────────────
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=8)  # auto logout after 8h idle
app.config['SESSION_COOKIE_HTTPONLY'] = True   # JS cannot read session cookie
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'  # CSRF protection

# ── LOGGING ──────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
logger = logging.getLogger('zewer_crm')

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
    try:
        r = x(conn, sql, p).fetchone()
        if r is None: return 0
        if isinstance(r, dict): return list(r.values())[0]
        try: return r[0]
        except: return 0
    except: return 0

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
    _run_migrations(conn)
    commit(conn)
    conn.close()

def _run_migrations(conn):
    """Universal migrations that run for BOTH PostgreSQL and SQLite on every startup."""
    pg = is_pg(conn)

    # New tables (use x() which handles both DB types correctly)
    new_tables = [
        "company_groups (id {pk}, group_name TEXT UNIQUE NOT NULL, description TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)",
        "clients (id {pk}, name TEXT NOT NULL, phone TEXT, whatsapp_number TEXT, email TEXT, date_of_birth DATE, profession TEXT, address TEXT, notes TEXT, nationality TEXT, passport_no TEXT, passport_expiry DATE, emirates_id TEXT, emirates_id_expiry DATE, address_proof TEXT, created_by INTEGER, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)",
        "app_settings (key TEXT PRIMARY KEY, value TEXT, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)",
        "regular_task_templates (id {pk}, title TEXT NOT NULL, description TEXT, frequency TEXT DEFAULT 'daily', assigned_role TEXT DEFAULT 'all', assigned_user_id INTEGER, created_by INTEGER, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)",
        "regular_task_logs (id {pk}, template_id INTEGER NOT NULL, user_id INTEGER NOT NULL, notes TEXT, status TEXT DEFAULT 'done', logged_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)",
        "internal_documents (id {pk}, doc_name TEXT NOT NULL, doc_category TEXT DEFAULT 'Staff', person_name TEXT, issuing_authority TEXT, issue_date DATE, expiry_date DATE, notes TEXT, added_by INTEGER, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)",
        "client_documents (id {pk}, client_id INTEGER NOT NULL, doc_type TEXT NOT NULL, file_name TEXT NOT NULL, file_url TEXT NOT NULL, public_id TEXT, uploaded_by INTEGER, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)",
        "additional_tasks (id {pk}, title TEXT NOT NULL, task_details TEXT, remarks TEXT, from_datetime TIMESTAMP NOT NULL, to_datetime TIMESTAMP NOT NULL, created_by INTEGER NOT NULL, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)",
        "aml_tracker (id {pk}, company_id INTEGER, individual_id INTEGER, transaction_date DATE NOT NULL, period TEXT, due_date DATE, vc_no TEXT, payment_mode TEXT, ac_type TEXT, client_name TEXT, transaction_currency TEXT, usd_amount NUMERIC, aed_amount NUMERIC, payment_remarks TEXT, invoice_no TEXT, invoice_amount NUMERIC, invoice_currency TEXT, goaml_submission_date DATE, goaml_status TEXT DEFAULT 'pending', goaml_ref_no TEXT, submitted_by INTEGER NOT NULL, checked_by INTEGER, comment TEXT, verified_ledger BOOLEAN DEFAULT FALSE, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)",
        "company_risk_assessments (id {pk}, company_id INTEGER NOT NULL, jurisdiction_score NUMERIC, ownership_score NUMERIC, delivery_channel_score NUMERIC, payment_method_score NUMERIC, transaction_volume_score NUMERIC, product_score NUMERIC, pep_status_score NUMERIC, nationality_score NUMERIC, years_relationship_score NUMERIC, years_operation_score NUMERIC, third_party_score NUMERIC, sanctions_score NUMERIC, final_score NUMERIC, risk_rating TEXT, assessment_date DATE, notes TEXT, assessed_by INTEGER, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)",
        "individual_risk_assessments (id {pk}, individual_id INTEGER NOT NULL, nationality_score NUMERIC, residence_status_score NUMERIC, pep_status_score NUMERIC, profession_score NUMERIC, product_score NUMERIC, delivery_channel_score NUMERIC, payment_method_score NUMERIC, transaction_amount_score NUMERIC, years_relationship_score NUMERIC, place_of_birth_score NUMERIC, third_party_score NUMERIC, sanctions_score NUMERIC, final_score NUMERIC, risk_rating TEXT, assessment_date DATE, notes TEXT, assessed_by INTEGER, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)",
    ]
    pk = "SERIAL PRIMARY KEY" if pg else "INTEGER PRIMARY KEY AUTOINCREMENT"
    for t in new_tables:
        sql = "CREATE TABLE IF NOT EXISTS " + t.format(pk=pk)
        try:
            x(conn, sql)
            commit(conn)
        except Exception as e:
            print(f"Migration table error ({sql[:50]}...): {e}")
            conn.rollback() if pg else None

    # ALTER TABLE additions — each wrapped individually so one failure doesn't block the rest
    def safe_alter(table, col, coltype='TEXT'):
        try:
            if pg:
                x(conn, f'ALTER TABLE {table} ADD COLUMN IF NOT EXISTS {col} {coltype}')
            else:
                x(conn, f'ALTER TABLE {table} ADD COLUMN {col} {coltype}')
            commit(conn)
        except Exception:
            if pg: conn.rollback()

    for col in ['contact_number','mobile','username','permissions']:
        safe_alter('users', col)
    # Ensure username column exists with explicit PG check
    if use_pg():
        try:
            x(conn, "ALTER TABLE users ADD COLUMN IF NOT EXISTS username TEXT")
            x(conn, "ALTER TABLE users ADD COLUMN IF NOT EXISTS permissions TEXT DEFAULT ''")
            commit(conn)
        except: conn.rollback()
    safe_alter('companies', 'group_name')
    safe_alter('companies', 'kyc_expiry_date', 'DATE')
    safe_alter('companies', 'id_type')
    safe_alter('companies', 'health_note')
    for col in ['contact_person_name','contact_person_number','account_manager',
                'whatsapp_number','deal_after_vat','registration_screening_tool']:
        safe_alter('companies', col)
    safe_alter('dropdowns', 'description')
    for col in ['nationality','passport_no','passport_expiry','emirates_id','emirates_id_expiry','address_proof','emirate','location','account_number','mode_of_ac','ac_status','id_type','pep_status','kyc_status','kyc_expiry_date','risk_status','screening_status','screening_date','is_resident']:
        coltype = 'DATE' if ('expiry' in col or col == 'screening_date') else ('BOOLEAN DEFAULT FALSE' if col == 'is_resident' else 'TEXT')
        safe_alter('clients', col, coltype)
    # Force-insert new dropdown values on existing DBs
    new_dd = [
        ('ID TYPE','National ID'),('ID TYPE','Passport'),('ID TYPE','Emirates ID'),('ID TYPE','Visa No'),('ID TYPE','Other'),
        ('SCREENING REGISTRATION STATUS','Yes'),('SCREENING REGISTRATION STATUS','No'),
        ('SCREENING REGISTRATION STATUS','Not Required'),('SCREENING REGISTRATION STATUS','Pending'),
    ]
    for field, val in new_dd:
        try:
            if use_pg():
                x(conn, "INSERT INTO dropdowns (field_name,value,is_active) VALUES (%s,%s,1) ON CONFLICT DO NOTHING", (field, val))
                commit(conn)
            else:
                conn.execute("INSERT OR IGNORE INTO dropdowns (field_name,value,is_active) VALUES (?,?,1)", (field, val))
        except: pass

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
            crowe_feedback TEXT, zewer_comments TEXT, group_name TEXT, kyc_expiry_date DATE, id_type TEXT, created_by INTEGER,
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
        """CREATE TABLE IF NOT EXISTS regular_task_templates (
            id SERIAL PRIMARY KEY, title TEXT NOT NULL, description TEXT,
            frequency TEXT DEFAULT 'daily', assigned_role TEXT DEFAULT 'all',
            assigned_user_id INTEGER, created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS regular_task_logs (
            id SERIAL PRIMARY KEY, template_id INTEGER NOT NULL,
            user_id INTEGER NOT NULL, notes TEXT, status TEXT DEFAULT 'done',
            logged_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS internal_documents (
            id SERIAL PRIMARY KEY, doc_name TEXT NOT NULL,
            doc_category TEXT DEFAULT 'Staff', person_name TEXT,
            issuing_authority TEXT, issue_date DATE, expiry_date DATE,
            notes TEXT, added_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS company_groups (
            id SERIAL PRIMARY KEY, group_name TEXT UNIQUE NOT NULL,
            description TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS clients (
            id SERIAL PRIMARY KEY, name TEXT NOT NULL, phone TEXT,
            whatsapp_number TEXT, email TEXT, date_of_birth DATE,
            profession TEXT, address TEXT, notes TEXT,
            passport_file TEXT, emirates_id_file TEXT, address_proof_file TEXT,
            created_by INTEGER, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS app_settings (
            key TEXT PRIMARY KEY, value TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""",
        """CREATE TABLE IF NOT EXISTS additional_tasks (
            id SERIAL PRIMARY KEY, title TEXT NOT NULL,
            task_details TEXT, remarks TEXT,
            from_datetime TIMESTAMP NOT NULL, to_datetime TIMESTAMP NOT NULL,
            created_by INTEGER NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""",
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
            crowe_feedback TEXT, zewer_comments TEXT, group_name TEXT, kyc_expiry_date DATE, id_type TEXT, created_by INTEGER,
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
        CREATE TABLE IF NOT EXISTS regular_task_templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL, description TEXT,
            frequency TEXT DEFAULT 'daily', assigned_role TEXT DEFAULT 'all',
            assigned_user_id INTEGER, created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS regular_task_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT, template_id INTEGER NOT NULL,
            user_id INTEGER NOT NULL, notes TEXT, status TEXT DEFAULT 'done',
            logged_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS internal_documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT, doc_name TEXT NOT NULL,
            doc_category TEXT DEFAULT 'Staff', person_name TEXT,
            issuing_authority TEXT, issue_date DATE, expiry_date DATE,
            notes TEXT, added_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS company_groups (
            id INTEGER PRIMARY KEY AUTOINCREMENT, group_name TEXT UNIQUE NOT NULL,
            description TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS clients (
            id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, phone TEXT,
            whatsapp_number TEXT, email TEXT, date_of_birth DATE,
            profession TEXT, address TEXT, notes TEXT,
            passport_file TEXT, emirates_id_file TEXT, address_proof_file TEXT,
            created_by INTEGER, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS app_settings (
            key TEXT PRIMARY KEY, value TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
    ''')

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
        'ID TYPE':['National ID','Passport','Emirates ID','Visa No'],
        'SCREENING REGISTRATION STATUS':['Yes','No','Not Required','Pending'],
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
    try:
        if hasattr(d, 'year'):  # already a date object
            return (d - dubai_today()).days
        return (datetime.strptime(str(d)[:10],'%Y-%m-%d').date()-dubai_today()).days
    except: return None

def exp_status(d):
    if d is None: return 'unknown'
    if d<0: return 'expired'
    if d<=30: return 'critical'
    if d<=90: return 'warning'
    return 'ok'

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
        login_id = d.get('username') or d.get('email','')
        u=one(conn,'SELECT * FROM users WHERE username=? OR email=?',(login_id,login_id))
        conn.close()
        if u and check_password_hash(u['password_hash'],d.get('password','')) and u['is_active']:
            session.permanent = True  # enables PERMANENT_SESSION_LIFETIME timeout
            session.update(user_id=u['id'],user_email=u['email'],user_name=u['name'],user_role=u['role'],user_permissions=(u.get('permissions') or ''))
            logger.info(f"Login: {u['email']} (role={u['role']})")
            return jsonify({'success':True})
        logger.warning(f"Failed login attempt for: {login_id}")
        return jsonify({'success':False,'error':'Invalid credentials'}),401
    return render_template('login.html')

@app.route('/logout')
def logout(): session.clear(); return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    # Staff get their own task dashboard
    if session.get('user_role') == 'staff':
        uid = session.get('user_id')
        conn = get_db()
        today = dubai_today()
        my_tasks = all_(conn, '''SELECT t.*,c.client_name as company_name,c.ac_code
            FROM tasks t LEFT JOIN companies c ON t.company_id=c.id
            WHERE t.assigned_to=? AND t.status NOT IN ('done')
            ORDER BY CASE t.priority WHEN 'urgent' THEN 1 WHEN 'high' THEN 2 WHEN 'normal' THEN 3 ELSE 4 END,t.due_date''', (uid,))
        tl = []
        for t in my_tasks:
            d = days_left(t['due_date'])
            tl.append({**t,'priority':t['priority'] or 'normal','status':t['status'] or 'todo',
                       'due_date':str(t['due_date']) if t['due_date'] else None,
                       'days_until_due':d,'is_overdue':(d is not None and d<0)})
        todo_c = sum(1 for t in tl if t['status']=='todo')
        inprog_c = sum(1 for t in tl if t['status']=='inprogress')
        pending_c = sum(1 for t in tl if t['status']=='pending_close')
        overdue_c = sum(1 for t in tl if t['is_overdue'])
        conn.close()
        return render_template('staff_dashboard.html', my_tasks=tl,
            todo_count=todo_c, inprogress_count=inprog_c,
            pending_count=pending_c, overdue_count=overdue_c, today=str(today))

    conn=get_db(); today=dubai_today()
    def c(sql,p=None): return cnt(conn,sql,p or [])
    total=c('SELECT COUNT(*) FROM companies')
    active=c('SELECT COUNT(*) FROM companies WHERE ac_status=?',('Active',))
    total_clients=c('SELECT COUNT(*) FROM clients')
    pep_count=(c("SELECT COUNT(*) FROM companies WHERE pep='Yes'") +
               c("SELECT COUNT(*) FROM clients WHERE pep_status='Yes'"))
    etl=c('SELECT COUNT(*) FROM companies WHERE trade_license_expiry<?',(today,))
    e30tl=c('SELECT COUNT(*) FROM companies WHERE trade_license_expiry BETWEEN ? AND ?',(today,today+timedelta(days=30)))
    eap=c('SELECT COUNT(*) FROM companies WHERE address_proof_expiry<?',(today,))
    e30ap=c('SELECT COUNT(*) FROM companies WHERE address_proof_expiry BETWEEN ? AND ?',(today,today+timedelta(days=30)))
    epass=c('SELECT COUNT(*) FROM ubos WHERE passport_expiry<?',(today,))
    e30p=c('SELECT COUNT(*) FROM ubos WHERE passport_expiry BETWEEN ? AND ?',(today,today+timedelta(days=30)))
    eid_exp=c('SELECT COUNT(*) FROM ubos WHERE emirates_id_expiry<?',(today,))
    risk_rows=all_(conn,'SELECT risk_status,COUNT(*) as c FROM companies GROUP BY risk_status')
    risk_bd={r['risk_status']:r['c'] for r in risk_rows}
    doc_rows=all_(conn,'SELECT doc_status,COUNT(*) as c FROM companies GROUP BY doc_status')
    doc_bd={r['doc_status']:r['c'] for r in doc_rows}
    # Count ALL expiring documents (TL + AP + Passport + EID) by window
    def doc_count(days):
        t1 = cnt(conn,'SELECT COUNT(*) FROM companies WHERE trade_license_expiry BETWEEN ? AND ?',(today,today+timedelta(days=days)))
        t2 = cnt(conn,'SELECT COUNT(*) FROM companies WHERE address_proof_expiry BETWEEN ? AND ?',(today,today+timedelta(days=days)))
        t3 = cnt(conn,'SELECT COUNT(*) FROM ubos WHERE passport_expiry BETWEEN ? AND ?',(today,today+timedelta(days=days)))
        t4 = cnt(conn,'SELECT COUNT(*) FROM ubos WHERE emirates_id_expiry BETWEEN ? AND ?',(today,today+timedelta(days=days)))
        return t1 + t2 + t3 + t4
    exp_30 = doc_count(30)
    exp_60 = doc_count(60)
    exp_90 = doc_count(90)
    urgent = []  # no longer used in template
    otasks=c("SELECT COUNT(*) FROM tasks WHERE status NOT IN ('done')")
    overtasks=c("SELECT COUNT(*) FROM tasks WHERE status NOT IN ('done','pending_close') AND due_date<?", (today,))
    dtasks=c("SELECT COUNT(*) FROM tasks WHERE status NOT IN ('done') AND due_date=?", (today,))
    ptasks=c("SELECT COUNT(*) FROM tasks WHERE status='pending_close'")
    utasks_raw=all_(conn,"""SELECT t.id,t.title,t.priority,t.status,t.due_date,u.name as assigned_name
        FROM tasks t LEFT JOIN users u ON t.assigned_to=u.id
        WHERE t.status NOT IN ('done') AND (t.priority IN ('urgent','high') OR t.due_date<=?)
        ORDER BY CASE t.priority WHEN 'urgent' THEN 1 WHEN 'high' THEN 2 ELSE 3 END,t.due_date LIMIT 8""",
        (today+timedelta(days=2),))
    # Convert date objects to strings
    utasks = []
    for t in utasks_raw:
        row = dict(t)
        row['due_date'] = str(row['due_date'])[:10] if row.get('due_date') else None
        utasks.append(row)
    try:
        staff_task_counts=all_(conn,"SELECT u.name,u.id,COUNT(t.id) as pending FROM users u LEFT JOIN tasks t ON t.assigned_to=u.id AND t.status NOT IN ('done') WHERE u.is_active=1 GROUP BY u.id,u.name ORDER BY pending DESC")
    except: staff_task_counts=[]

    # Upcoming client birthdays (next 30 days)
    upcoming_birthdays = []
    try:
        all_clients = all_(conn, "SELECT id,name,whatsapp_number,phone,date_of_birth FROM clients WHERE date_of_birth IS NOT NULL")
        for cl in all_clients:
            dob = cl.get('date_of_birth')
            if not dob: continue
            try:
                d = datetime.strptime(str(dob)[:10], '%Y-%m-%d').date()
                this_year = d.replace(year=today.year)
                if this_year < today:
                    this_year = d.replace(year=today.year+1)
                days_away = (this_year - today).days
                if days_away <= 30:
                    upcoming_birthdays.append({
                        'id': cl['id'], 'name': cl['name'],
                        'whatsapp_number': cl.get('whatsapp_number') or cl.get('phone'),
                        'days_away': days_away,
                        'is_today': days_away == 0,
                        'date_display': d.strftime('%b %d')
                    })
            except: pass
        upcoming_birthdays.sort(key=lambda x: x['days_away'])
    except: pass

    # Client KYC expiry alerts (expired or expiring within 90 days)
    kyc_alerts = []
    try:
        kyc_clients = all_(conn, "SELECT id,name,kyc_expiry_date,kyc_status FROM clients WHERE kyc_expiry_date IS NOT NULL")
        # Also add companies with expiring KYC
        kyc_cos = all_(conn, "SELECT id,client_name as name,kyc_expiry_date,kyc_status FROM companies WHERE kyc_expiry_date IS NOT NULL")
        for cl in kyc_clients:
            ke = cl.get('kyc_expiry_date')
            if not ke: continue
            dl = days_left(ke)
            if dl is not None and dl <= 90:
                kyc_alerts.append({
                    'id': cl['id'], 'name': cl['name'],
                    'kyc_expiry_date': str(ke)[:10],
                    'days_left': dl, 'is_expired': dl < 0,
                    'kyc_status': cl.get('kyc_status')
                })
        for co in kyc_cos:
            ke = co.get('kyc_expiry_date')
            if not ke: continue
            dl = days_left(ke)
            if dl is not None and dl <= 90:
                kyc_alerts.append({'id': co['id'], 'name': co['name'] + ' (Co)',
                    'kyc_expiry_date': str(ke)[:10], 'days_left': dl,
                    'is_expired': dl < 0, 'kyc_status': co.get('kyc_status'), 'type': 'company'})
        kyc_alerts.sort(key=lambda x: x['days_left'])
    except: pass

    conn.close()
    return render_template('dashboard.html',total_companies=total,active_companies=active,total_clients=total_clients,pep_count=pep_count,eid_exp=eid_exp,
        expired_tl=etl,expiring_30_tl=e30tl,
        expired_ap=eap,expiring_30_ap=e30ap,expired_pass=epass,expiring_30_pass=e30p,
        exp_30=exp_30,exp_60=exp_60,exp_90=exp_90,
        risk_breakdown=risk_bd,doc_breakdown=doc_bd,urgent_companies=urgent,
        open_tasks=otasks,overdue_tasks=overtasks,due_today=dtasks,pending_close=ptasks,
        urgent_tasks=utasks,today=str(today),days_left=days_left,staff_task_counts=staff_task_counts,
        upcoming_birthdays=upcoming_birthdays,kyc_alerts=kyc_alerts)

@app.route('/companies')
@compliance_required
def companies():
    conn=get_db()
    s=request.args.get('search',''); sf=request.args.get('status','')
    rgf=request.args.get('region',''); modf=request.args.get('mode',''); grpf=request.args.get('group','')
    q='SELECT * FROM companies WHERE 1=1'; p=[]
    if s:
        q+=' AND (client_name LIKE ? OR ac_code LIKE ? OR mobile LIKE ? OR trade_license_no LIKE ?)'
        p+=[f'%{s}%']*4
    if sf: q+=' AND ac_status=?'; p.append(sf)
    if rgf: q+=' AND region=?'; p.append(rgf)
    if modf: q+=' AND mode_of_ac=?'; p.append(modf)
    if grpf: q+=' AND group_name=?'; p.append(grpf)

    page = max(1, int(request.args.get('page', 1)))
    per_page = 100
    total_count = cnt(conn, f"SELECT COUNT(*) FROM companies WHERE {q.split('WHERE')[1]}", p or None)
    rows=all_(conn,q+f' ORDER BY created_at DESC LIMIT {per_page} OFFSET {(page-1)*per_page}',p or None)
    total_pages = max(1, (total_count + per_page - 1) // per_page)
    dd=dropdowns()
    try: groups=all_(conn,'SELECT group_name FROM company_groups ORDER BY group_name')
    except: groups=[]
    conn.close()
    cl=[]
    for c in rows:
        tl=days_left(c['trade_license_expiry']); ap=days_left(c['address_proof_expiry'])
        cl.append({**c,'tl_days':tl,'tl_status':exp_status(tl),'ap_days':ap,'ap_status':exp_status(ap),
                   'trade_license_expiry':str(c['trade_license_expiry']) if c['trade_license_expiry'] else None,
                   'address_proof_expiry':str(c['address_proof_expiry']) if c['address_proof_expiry'] else None})
    return render_template('companies.html',companies=cl,search=s,
        status_filter=sf,region_filter=rgf,mode_filter=modf,group_filter=grpf,
        regions=dd.get('REGION',[]),modes=dd.get('MODE OF AC',[]),
        groups=[g['group_name'] for g in groups],
        page=page, total_pages=total_pages, total_count=total_count, per_page=per_page)

@app.route('/company/new')
@compliance_required
def company_new():
    conn=get_db()
    staff=all_(conn,'SELECT id,name FROM users WHERE is_active=1 ORDER BY name')
    conn.close()
    max_kyc = (dubai_today().replace(year=dubai_today().year+2)).isoformat()
    return render_template('company_form.html',dropdown_data=dropdowns(),company=None,ubos=[],edit=False,staff_users=staff,max_kyc_date=max_kyc)

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
        today=str(dubai_today()))

@app.route('/company/<int:id>/edit')
@compliance_required
def company_edit(id):
    conn=get_db()
    co=one(conn,'SELECT * FROM companies WHERE id=?',(id,))
    if not co: conn.close(); return redirect(url_for('companies'))
    ubos=all_(conn,'SELECT * FROM ubos WHERE company_id=? ORDER BY share_percentage DESC',(id,))
    staff=all_(conn,'SELECT id,name FROM users WHERE is_active=1 ORDER BY name')
    conn.close()
    max_kyc = (dubai_today().replace(year=dubai_today().year+2)).isoformat()
    return render_template('company_form.html',dropdown_data=dropdowns(),company=co,ubos=ubos,edit=True,staff_users=staff,max_kyc_date=max_kyc)

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
        d.get('verified_date') or None,d.get('followup_details'),d.get('crowe_feedback'),d.get('zewer_comments'),
           d.get('kyc_expiry_date') or None, d.get('id_type'))

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
    kyc_err = _validate_kyc_expiry(d.get('kyc_expiry_date'))
    if kyc_err: return jsonify({'success':False,'error':kyc_err}),400
    try:
        conn=get_db()
        x(conn,'''INSERT INTO companies (ac_code,client_name,ac_opening_date,ac_status,active_till_year,
            nature,type_of_client,name_of_freezone,mode_of_ac,country_of_incorporation,region,address,
            telephone,mobile,whatsapp_number,email_id,contact_person_name,contact_person_number,account_manager,
            address_proof_type,address_proof_expiry,kyc_status,trade_license_no,issuing_authority,legal_type,
            incorporation_date,trade_license_expiry,tax_no_trn,vat_cert,vat_declaration,deal_after_vat,
            num_beneficial_owners,moa,pep,undertaking,source_of_fund,software_updation,doc_status,
            screening_date,registration_screening_tool,risk_status,verified_by,verified_date,
            followup_details,crowe_feedback,zewer_comments,kyc_expiry_date,id_type,created_by)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (d.get('ac_code'),d.get('client_name'))+_cv(d)+(session.get('user_id'),))
        cid=lastid(conn)
        _save_ubos(conn,cid,d.get('ubos',[]))
        commit(conn); conn.close()
        return jsonify({'success':True,'id':cid})
    except Exception as e:
        logger.error(f'Error in %s: {e}', request.path)
        return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/company/<int:id>/edit', methods=['POST'])
@compliance_required
def api_edit_company(id):
    d=request.get_json()
    kyc_err = _validate_kyc_expiry(d.get('kyc_expiry_date'))
    if kyc_err: return jsonify({'success':False,'error':kyc_err}),400
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
            verified_date=?,followup_details=?,crowe_feedback=?,zewer_comments=?,kyc_expiry_date=?,id_type=?,
            updated_at=CURRENT_TIMESTAMP WHERE id=?''',
            (d.get('client_name'),)+_cv(d)+(id,))
        _save_ubos(conn,id,d.get('ubos',[]))
        commit(conn); conn.close()
        return jsonify({'success':True})
    except Exception as e:
        logger.error(f'Error in %s: {e}', request.path)
        return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/company/<int:id>/delete', methods=['POST'])
@compliance_required
def api_delete_company(id):
    try:
        conn=get_db(); x(conn,'DELETE FROM companies WHERE id=?',(id,))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e:
        logger.error(f'Error in %s: {e}', request.path)
        return jsonify({'success':False,'error':str(e)}),500

@app.route('/alerts')
@compliance_required
def alerts():
    conn=get_db(); today=dubai_today()
    # Build unified alert list from all document types
    all_alerts = []

    # Trade licenses
    tl_rows = all_(conn,'''SELECT id,ac_code,client_name,mobile,whatsapp_number,
        trade_license_expiry,risk_status,region,account_manager,
        contact_person_name,contact_person_number
        FROM companies WHERE trade_license_expiry IS NOT NULL ORDER BY trade_license_expiry''')
    for r in tl_rows:
        d = days_left(r['trade_license_expiry'])
        if d is None: continue
        all_alerts.append({'company_id':r['id'],'ac_code':r['ac_code'],
            'client_name':r['client_name'],'mobile':r['mobile'],
            'whatsapp_number':r['whatsapp_number'],'account_manager':r.get('account_manager'),
            'contact_person_name':r.get('contact_person_name'),
            'contact_person_number':r.get('contact_person_number'),
            'doc_type':'Trade License','doc_subtype':None,
            'expiry_date':str(r['trade_license_expiry'])[:10],'days':d})

    # Address proofs
    ap_rows = all_(conn,'''SELECT id,ac_code,client_name,mobile,whatsapp_number,
        address_proof_expiry,address_proof_type,account_manager,
        contact_person_name,contact_person_number
        FROM companies WHERE address_proof_expiry IS NOT NULL ORDER BY address_proof_expiry''')
    for r in ap_rows:
        d = days_left(r['address_proof_expiry'])
        if d is None: continue
        all_alerts.append({'company_id':r['id'],'ac_code':r['ac_code'],
            'client_name':r['client_name'],'mobile':r['mobile'],
            'whatsapp_number':r['whatsapp_number'],'account_manager':r.get('account_manager'),
            'contact_person_name':r.get('contact_person_name'),
            'contact_person_number':r.get('contact_person_number'),
            'doc_type':'Address Proof','doc_subtype':r.get('address_proof_type'),
            'expiry_date':str(r['address_proof_expiry'])[:10],'days':d})

    # Passports
    ubo_rows = all_(conn,'''SELECT u.person_name,u.passport_no,u.passport_expiry,
        u.emirates_id,u.emirates_id_expiry,
        c.id as company_id,c.client_name,c.ac_code,c.mobile,c.whatsapp_number,
        c.account_manager,c.contact_person_name,c.contact_person_number
        FROM ubos u JOIN companies c ON u.company_id=c.id
        WHERE u.passport_expiry IS NOT NULL OR u.emirates_id_expiry IS NOT NULL''')
    for r in ubo_rows:
        if r.get('passport_expiry'):
            d = days_left(r['passport_expiry'])
            if d is not None:
                all_alerts.append({'company_id':r['company_id'],'ac_code':r['ac_code'],
                    'client_name':r['client_name'],'mobile':r['mobile'],
                    'whatsapp_number':r['whatsapp_number'],'account_manager':r.get('account_manager'),
                    'contact_person_name':r['person_name'],'contact_person_number':None,
                    'doc_type':'Passport','doc_subtype':r.get('passport_no'),
                    'expiry_date':str(r['passport_expiry'])[:10],'days':d})
        if r.get('emirates_id_expiry'):
            d = days_left(r['emirates_id_expiry'])
            if d is not None:
                all_alerts.append({'company_id':r['company_id'],'ac_code':r['ac_code'],
                    'client_name':r['client_name'],'mobile':r['mobile'],
                    'whatsapp_number':r['whatsapp_number'],'account_manager':r.get('account_manager'),
                    'contact_person_name':r['person_name'],'contact_person_number':None,
                    'doc_type':'Emirates ID','doc_subtype':r.get('emirates_id'),
                    'expiry_date':str(r['emirates_id_expiry'])[:10],'days':d})

    # Clients (Individuals) documents
    client_rows = all_(conn,'''SELECT id,name,phone,whatsapp_number,
        passport_expiry,emirates_id_expiry,account_number
        FROM clients WHERE passport_expiry IS NOT NULL OR emirates_id_expiry IS NOT NULL''')
    for r in client_rows:
        if r.get('passport_expiry'):
            d = days_left(r['passport_expiry'])
            if d is not None:
                all_alerts.append({'company_id':None,'ac_code':r.get('account_number','N/A'),
                    'client_name':r['name'],'mobile':r.get('phone'),'whatsapp_number':r.get('whatsapp_number'),
                    'account_manager':None,'contact_person_name':r['name'],'contact_person_number':r.get('phone'),
                    'doc_type':'Passport (Individual)','doc_subtype':None,
                    'expiry_date':str(r['passport_expiry'])[:10],'days':d})
        if r.get('emirates_id_expiry'):
            d = days_left(r['emirates_id_expiry'])
            if d is not None:
                all_alerts.append({'company_id':None,'ac_code':r.get('account_number','N/A'),
                    'client_name':r['name'],'mobile':r.get('phone'),'whatsapp_number':r.get('whatsapp_number'),
                    'account_manager':None,'contact_person_name':r['name'],'contact_person_number':r.get('phone'),
                    'doc_type':'Emirates ID (Individual)','doc_subtype':None,
                    'expiry_date':str(r['emirates_id_expiry'])[:10],'days':d})

    # Sort by days (most urgent first)
    all_alerts.sort(key=lambda x: x['days'])

    # Get unique managers for filter
    managers = sorted(set(a['account_manager'] for a in all_alerts if a.get('account_manager')))
    
    # Get users for task assignment
    all_users = all_(conn,'SELECT id,name,role FROM users WHERE is_active=1 ORDER BY name')
    
    conn.close()
    return render_template('alerts.html', all_alerts=all_alerts, managers=managers,
        all_users=all_users, today=str(today))

@app.route('/health-check')
@compliance_required
def health_check():
    conn = get_db()
    companies = all_(conn, 'SELECT id, ac_code, client_name FROM companies ORDER BY client_name')
    conn.close()
    return render_template('health_check.html', companies=companies)

@app.route('/health-check/<int:id>')
@compliance_required
def health_check_detail(id):
    conn = get_db()
    co = one(conn, 'SELECT * FROM companies WHERE id=?', (id,))
    if not co: conn.close(); return redirect(url_for('health_check'))
    if 'health_note' not in co: co['health_note'] = ''
    ubos = all_(conn, 'SELECT * FROM ubos WHERE company_id=? ORDER BY share_percentage DESC', (id,))
    
    # Get latest risk assessment
    risk_assessment = one(conn, """
        SELECT final_score, risk_rating FROM company_risk_assessments 
        WHERE company_id = ? ORDER BY assessment_date DESC LIMIT 1
    """, (id,))
    
    conn.close()
    today = dubai_today()

    def doc_status(expiry):
        dl = days_left(expiry)
        if expiry is None: return 'missing', None
        if dl is None: return 'missing', None
        if dl < 0: return 'expired', dl
        if dl <= 30: return 'critical', dl
        if dl <= 90: return 'warning', dl
        return 'ok', dl

    # Build document list with icons
    docs = []

    # Trade License
    tl_st, tl_dl = doc_status(co.get('trade_license_expiry'))
    docs.append({'name': 'Trade License', 'icon': '📜', 'ref': co.get('trade_license_no'), 'expiry': str(co.get('trade_license_expiry',''))[:10] if co.get('trade_license_expiry') else None, 'days': tl_dl, 'status': tl_st, 'category': 'Company'})

    # Address Proof
    ap_st, ap_dl = doc_status(co.get('address_proof_expiry'))
    docs.append({'name': 'Address Proof', 'icon': '🏠', 'ref': co.get('address_proof_type'), 'expiry': str(co.get('address_proof_expiry',''))[:10] if co.get('address_proof_expiry') else None, 'days': ap_dl, 'status': ap_st, 'category': 'Company'})

    # VAT / TRN
    vat_present = bool(co.get('tax_no_trn'))
    docs.append({'name': 'VAT / TRN', 'icon': '🧾', 'ref': co.get('tax_no_trn'), 'expiry': None, 'days': None, 'status': 'ok' if vat_present else 'missing', 'category': 'Tax'})

    # KYC Review
    kyc_st, kyc_dl = doc_status(co.get('kyc_expiry_date'))
    docs.append({'name': 'KYC Review', 'icon': '🛡️', 'ref': co.get('kyc_status'), 'expiry': str(co.get('kyc_expiry_date',''))[:10] if co.get('kyc_expiry_date') else None, 'days': kyc_dl, 'status': kyc_st, 'category': 'Compliance'})

    # UBO documents
    for u in ubos:
        pp_st, pp_dl = doc_status(u.get('passport_expiry'))
        docs.append({'name': f"Passport", 'icon': '🛂', 'ref': u['person_name'], 'expiry': str(u.get('passport_expiry',''))[:10] if u.get('passport_expiry') else None, 'days': pp_dl, 'status': pp_st, 'category': u.get('position') or 'UBO'})
        eid_st, eid_dl = doc_status(u.get('emirates_id_expiry'))
        docs.append({'name': f"Emirates ID", 'icon': '🪪', 'ref': u['person_name'], 'expiry': str(u.get('emirates_id_expiry',''))[:10] if u.get('emirates_id_expiry') else None, 'days': eid_dl, 'status': eid_st, 'category': u.get('position') or 'UBO'})

    # ── SUMMARY COUNTS ──
    summary = {
        'total': len(docs),
        'valid': sum(1 for d in docs if d['status'] == 'ok'),
        'expiring': sum(1 for d in docs if d['status'] in ('warning','critical')),
        'expired': sum(1 for d in docs if d['status'] == 'expired'),
        'missing': sum(1 for d in docs if d['status'] == 'missing'),
    }

    # ── HEALTH SCORE CALCULATION ──
    score = 100
    deductions = []
    for d in docs:
        if d['status'] == 'expired':
            score -= 15; deductions.append(f"{d['name']} expired")
        elif d['status'] == 'critical':
            score -= 8; deductions.append(f"{d['name']} expiring in {d['days']}d")
        elif d['status'] == 'warning':
            score -= 4; deductions.append(f"{d['name']} expiring in {d['days']}d")
        elif d['status'] == 'missing':
            score -= 5; deductions.append(f"{d['name']} missing")

    risk = (co.get('risk_status') or '').lower()
    if risk == 'high': score -= 10; deductions.append('High risk client')
    elif risk == 'medium': score -= 5; deductions.append('Medium risk client')
    
    # Factor in risk assessment if available
    if risk_assessment:
        if risk_assessment.get('risk_rating') == 'High':
            score -= 10; deductions.append(f"Risk Assessment: High ({risk_assessment.get('final_score')})")
        elif risk_assessment.get('risk_rating') == 'Medium':
            score -= 5; deductions.append(f"Risk Assessment: Medium ({risk_assessment.get('final_score')})")
        # Low risk doesn't deduct

    kyc = (co.get('kyc_status') or '').lower()
    if 'not' in kyc or 'pending' in kyc: score -= 8; deductions.append('KYC not completed')
    elif 'expired' in kyc: score -= 10; deductions.append('KYC expired')

    if co.get('doc_status') == 'Incompleted': score -= 5; deductions.append('Documents incomplete')
    if (co.get('pep') or '').lower() == 'yes': score -= 5; deductions.append('PEP flagged')

    score = max(0, min(100, score))

    if score >= 90: grade = 'Excellent'; grade_color = '#22c55e'
    elif score >= 70: grade = 'Good'; grade_color = '#86efac'
    elif score >= 50: grade = 'Average'; grade_color = '#f59e0b'
    else: grade = 'Poor'; grade_color = '#ef4444'

    return render_template('health_check_detail.html',
        company=co, ubos=ubos, docs=docs, summary=summary,
        score=score, grade=grade, grade_color=grade_color,
        deductions=deductions, today=str(today), risk_assessment=risk_assessment)

@app.route('/api/company/<int:id>/health-note', methods=['POST'])
@compliance_required
def api_save_health_note(id):
    d = request.get_json()
    try:
        conn = get_db()
        # Ensure column exists (Railway PG may not have run migration)
        try:
            if is_pg(conn):
                x(conn, 'ALTER TABLE companies ADD COLUMN IF NOT EXISTS health_note TEXT')
            else:
                x(conn, 'ALTER TABLE companies ADD COLUMN health_note TEXT')
            commit(conn)
        except:
            try: conn.rollback()
            except: pass
        x(conn, 'UPDATE companies SET health_note=? WHERE id=?', (d.get('note',''), id))
        commit(conn); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        logger.error(f'Error saving health note: {e}')
        try: conn.close()
        except: pass
        return jsonify({'success': False, 'error': str(e)}), 500

# ──────────── RISK ASSESSMENT ────────────
def calculate_risk_score(scores):
    """Calculate final risk score and rating from individual factor scores."""
    if not scores:
        return None, 'Unspecified'
    valid = [s for s in scores if s is not None and s > 0]
    if not valid:
        return None, 'Unspecified'
    avg = sum(valid) / len(valid)
    if avg <= 1.00:
        return round(avg, 2), 'Low'
    elif avg <= 2.00:
        return round(avg, 2), 'Medium'
    else:
        return round(avg, 2), 'High'

@app.route('/api/risk-lookup', methods=['POST'])
@compliance_required
def api_risk_lookup():
    """Lookup risk score for a value in a category"""
    data = request.get_json()
    category = data.get('category')
    value = data.get('value')
    
    if not category or not value:
        return jsonify({'score': None}), 400
    
    score = RISK_LOOKUPS.get(category, {}).get(value)
    return jsonify({'score': score})

@app.route('/api/send-risk-email', methods=['POST'])
@compliance_required
def api_send_risk_email():
    """Send risk assessment via email"""
    data = request.get_json()
    email = data.get('email')
    company = data.get('company')
    rating = data.get('rating')
    score = data.get('score')
    
    if not email or not company:
        return jsonify({'success': False, 'error': 'Missing email or company'}), 400
    
    try:
        # For now, just return success (email sending requires SMTP setup)
        logger.info(f'Risk assessment email requested for {company} to {email}')
        # TODO: Integrate actual email sending (SendGrid, SMTP, etc.)
        return jsonify({'success': True, 'message': f'Risk assessment for {company} ({rating} - {score}) ready to send'})
    except Exception as e:
        logger.error(f'Error sending risk email: {e}')
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/risk-assessment-list')
@compliance_required
def risk_assessment_list():
    """View all risk assessments with filters"""
    try:
        conn = get_db()
        
        # Ensure tables exist on Railway
        if is_pg(conn):
            x(conn, '''CREATE TABLE IF NOT EXISTS company_risk_assessments (
                id SERIAL PRIMARY KEY, company_id INTEGER NOT NULL,
                jurisdiction_score FLOAT, ownership_score FLOAT, delivery_channel_score FLOAT,
                payment_method_score FLOAT, transaction_volume_score FLOAT, product_score FLOAT,
                pep_status_score FLOAT, nationality_score FLOAT, years_relationship_score FLOAT,
                years_operation_score FLOAT, third_party_score FLOAT, sanctions_score FLOAT,
                final_score FLOAT, risk_rating TEXT, assessment_date DATE, notes TEXT, assessed_by INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
            x(conn, '''CREATE TABLE IF NOT EXISTS individual_risk_assessments (
                id SERIAL PRIMARY KEY, individual_id INTEGER NOT NULL,
                nationality_score FLOAT, residence_status_score FLOAT, pep_status_score FLOAT,
                profession_score FLOAT, product_score FLOAT, delivery_channel_score FLOAT,
                payment_method_score FLOAT, transaction_amount_score FLOAT, years_relationship_score FLOAT,
                place_of_birth_score FLOAT, third_party_score FLOAT, sanctions_score FLOAT,
                final_score FLOAT, risk_rating TEXT, assessment_date DATE, notes TEXT, assessed_by INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
            commit(conn)
        
        # Get company assessments
        co_assessments = all_(conn, """
            SELECT ca.id, ca.company_id, c.client_name as name, ca.final_score, 
                   ca.risk_rating, ca.assessment_date, 'company' as type
            FROM company_risk_assessments ca
            JOIN companies c ON ca.company_id = c.id
            ORDER BY ca.assessment_date DESC
        """) or []
        
        # Get individual assessments  
        ind_assessments = all_(conn, """
            SELECT ia.id, ia.individual_id, cl.name, ia.final_score,
                   ia.risk_rating, ia.assessment_date, 'individual' as type
            FROM individual_risk_assessments ia
            JOIN clients cl ON ia.individual_id = cl.id
            ORDER BY ia.assessment_date DESC
        """) or []
        
        # Combine and sort
        assessments = sorted(
            co_assessments + ind_assessments,
            key=lambda x: x.get('assessment_date', ''),
            reverse=True
        )
        
        # Count by rating
        low = sum(1 for a in assessments if a.get('risk_rating') == 'Low')
        medium = sum(1 for a in assessments if a.get('risk_rating') == 'Medium')
        high = sum(1 for a in assessments if a.get('risk_rating') == 'High')
        
        conn.close()
        
        return render_template('risk_assessment_list.html',
                             assessments=assessments,
                             total_count=len(assessments),
                             low_count=low,
                             medium_count=medium,
                             high_count=high)
    except Exception as e:
        logger.error(f'Error loading risk assessment list: {e}', exc_info=True)
        try: conn.close()
        except: pass
        return f'<h1>Error</h1><p>{str(e)}</p>', 500

@app.route('/risk-assessment')
@compliance_required
def risk_assessment():
    """Select company or individual for risk assessment"""
    try:
        conn = get_db()
        companies = all_(conn, 'SELECT id, ac_code, client_name FROM companies ORDER BY client_name')
        # Individuals are stored in clients table
        individuals = all_(conn, 'SELECT id, name FROM clients ORDER BY name')
        conn.close()
        logger.info(f'Risk assessment: {len(companies or [])} companies, {len(individuals or [])} individuals')
        return render_template('risk_assessment.html', 
                             companies=companies or [], 
                             individuals=individuals or [])
    except Exception as e:
        logger.error(f'Error in risk_assessment: {e}', exc_info=True)
        try: conn.close()
        except: pass
        return f'<h1>Error</h1><p>{str(e)}</p>', 500

@app.route('/risk-assessment/company/<int:id>', methods=['GET','POST'])
@compliance_required
def risk_assessment_company(id):
    """Risk assessment form for company"""
    conn = get_db()
    co = one(conn, 'SELECT * FROM companies WHERE id=?', (id,))
    if not co:
        conn.close()
        return redirect(url_for('risk_assessment'))
    
    if request.method == 'POST':
        data = request.form
        logger.info(f'Risk assessment form received for company {id}')
        
        try:
            # Ensure tables exist on Railway
            if is_pg(conn):
                x(conn, '''CREATE TABLE IF NOT EXISTS company_risk_assessments (
                    id SERIAL PRIMARY KEY, company_id INTEGER NOT NULL,
                    jurisdiction_score FLOAT, ownership_score FLOAT, delivery_channel_score FLOAT,
                    payment_method_score FLOAT, transaction_volume_score FLOAT, product_score FLOAT,
                    pep_status_score FLOAT, nationality_score FLOAT, years_relationship_score FLOAT,
                    years_operation_score FLOAT, third_party_score FLOAT, sanctions_score FLOAT,
                    final_score FLOAT, risk_rating TEXT, assessment_date DATE, notes TEXT, assessed_by INTEGER,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
                commit(conn)
            
            scores = [
                float(data.get('jurisdiction_score') or 0),
                float(data.get('ownership_score') or 0),
                float(data.get('delivery_channel_score') or 0),
                float(data.get('payment_method_score') or 0),
                float(data.get('transaction_volume_score') or 0),
                float(data.get('product_score') or 0),
                float(data.get('pep_status_score') or 0),
                float(data.get('nationality_score') or 0),
                float(data.get('years_relationship_score') or 0),
                float(data.get('years_operation_score') or 0),
                float(data.get('third_party_score') or 0),
                float(data.get('sanctions_score') or 0),
            ]
            logger.info(f'Parsed scores: {scores}')
            final_score, risk_rating = calculate_risk_score(scores)
            logger.info(f'Calculated risk: score={final_score}, rating={risk_rating}')
            
            
            x(conn, '''INSERT INTO company_risk_assessments 
               (company_id, jurisdiction_score, ownership_score, delivery_channel_score, payment_method_score,
                transaction_volume_score, product_score, pep_status_score, nationality_score, 
                years_relationship_score, years_operation_score, third_party_score, sanctions_score,
                final_score, risk_rating, assessment_date, notes, assessed_by)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
               (id, scores[0], scores[1], scores[2], scores[3], scores[4], scores[5], scores[6],
                scores[7], scores[8], scores[9], scores[10], scores[11], final_score, risk_rating,
                dubai_today(), data.get('notes',''), session.get('user_id')))
            commit(conn)
            logger.info(f'✅ Risk assessment saved for company {id}')
            conn.close()
            return redirect(url_for('risk_assessment_company_result', id=id))
        except Exception as e:
            logger.error(f'❌ Error saving company risk assessment: {str(e)}', exc_info=True)
            try: conn.close()
            except: pass
            return redirect(url_for('risk_assessment'))
    
    # Prepare pre-fill data
    prefill = {
        'jurisdiction': co.get('country_of_incorporation', 'UNITED ARAB EMIRATES'),
        'nature_of_business': co.get('nature', ''),
        'product': co.get('type_of_client', ''),
        'payment_method': '',
        'pep_status': 'No' if (co.get('pep') or '').lower() != 'yes' else 'Yes',
        'third_party': 'No',
    }
    
    # Calculate years in operation
    years_operation = 3  # default high risk
    if co.get('incorporation_date'):
        try:
            inc_date = co['incorporation_date'] if isinstance(co['incorporation_date'], str) else str(co['incorporation_date'])[:10]
            from datetime import datetime
            inc = datetime.strptime(inc_date, '%Y-%m-%d').date()
            years_diff = (dubai_today() - inc).days / 365.25
            if years_diff >= 5:
                years_operation = 1
            elif years_diff >= 1:
                years_operation = 2
        except:
            pass
    prefill['years_operation'] = years_operation
    
    # Calculate years in relationship
    years_relationship = 3  # default high risk
    if co.get('ac_opening_date'):
        try:
            opening = co['ac_opening_date'] if isinstance(co['ac_opening_date'], str) else str(co['ac_opening_date'])[:10]
            from datetime import datetime
            open_dt = datetime.strptime(opening, '%Y-%m-%d').date()
            years_diff = (dubai_today() - open_dt).days / 365.25
            if years_diff >= 5:
                years_relationship = 1
            elif years_diff >= 1:
                years_relationship = 2
        except:
            pass
    prefill['years_relationship'] = years_relationship
    
    conn.close()
    return render_template('risk_assessment_company.html', company=co, prefill=prefill, lookups=RISK_LOOKUPS)

@app.route('/risk-assessment/company/<int:id>/result')
@compliance_required
def risk_assessment_company_result(id):
    """Display company risk assessment result"""
    conn = get_db()
    co = one(conn, 'SELECT * FROM companies WHERE id=?', (id,))
    assessment = one(conn, '''SELECT * FROM company_risk_assessments 
                             WHERE company_id=? ORDER BY created_at DESC LIMIT 1''', (id,))
    conn.close()
    
    if not co or not assessment:
        return redirect(url_for('risk_assessment'))
    
    # Determine color based on risk rating
    color_map = {'Low': '#22c55e', 'Medium': '#f59e0b', 'High': '#ef4444'}
    risk_color = color_map.get(assessment.get('risk_rating', 'Unspecified'), '#6b7280')
    
    return render_template('risk_assessment_result.html',
        company=co, assessment=assessment, risk_color=risk_color, assessment_type='company')

@app.route('/risk-assessment/individual/<int:id>', methods=['GET','POST'])
@compliance_required
def risk_assessment_individual(id):
    """Risk assessment form for individual/client"""
    conn = get_db()
    ind = one(conn, 'SELECT * FROM clients WHERE id=?', (id,))
    if not ind:
        conn.close()
        return redirect(url_for('risk_assessment'))
    
    if request.method == 'POST':
        data = request.form
        try:
            # Ensure tables exist on Railway
            if is_pg(conn):
                x(conn, '''CREATE TABLE IF NOT EXISTS individual_risk_assessments (
                    id SERIAL PRIMARY KEY, individual_id INTEGER NOT NULL,
                    nationality_score FLOAT, residence_status_score FLOAT, pep_status_score FLOAT,
                    profession_score FLOAT, product_score FLOAT, delivery_channel_score FLOAT,
                    payment_method_score FLOAT, transaction_amount_score FLOAT, years_relationship_score FLOAT,
                    place_of_birth_score FLOAT, third_party_score FLOAT, sanctions_score FLOAT,
                    final_score FLOAT, risk_rating TEXT, assessment_date DATE, notes TEXT, assessed_by INTEGER,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
                commit(conn)
            
            scores = [
                float(data.get('nationality_score') or 0),
                float(data.get('residence_status_score') or 0),
                float(data.get('pep_status_score') or 0),
                float(data.get('profession_score') or 0),
                float(data.get('product_score') or 0),
                float(data.get('delivery_channel_score') or 0),
                float(data.get('payment_method_score') or 0),
                float(data.get('transaction_amount_score') or 0),
                float(data.get('years_relationship_score') or 0),
                float(data.get('place_of_birth_score') or 0),
                float(data.get('third_party_score') or 0),
                float(data.get('sanctions_score') or 0),
            ]
            final_score, risk_rating = calculate_risk_score(scores)
            
            
            x(conn, '''INSERT INTO individual_risk_assessments 
               (individual_id, nationality_score, residence_status_score, pep_status_score, profession_score,
                product_score, delivery_channel_score, payment_method_score, transaction_amount_score,
                years_relationship_score, place_of_birth_score, third_party_score, sanctions_score,
                final_score, risk_rating, assessment_date, notes, assessed_by)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
               (id, scores[0], scores[1], scores[2], scores[3], scores[4], scores[5], scores[6],
                scores[7], scores[8], scores[9], scores[10], scores[11], final_score, risk_rating,
                dubai_today(), data.get('notes',''), session.get('user_id')))
            commit(conn)
            conn.close()
            return redirect(url_for('risk_assessment_individual_result', id=id))
        except Exception as e:
            logger.error(f'Error saving individual risk assessment: {e}')
            conn.close()
            return redirect(url_for('risk_assessment'))
    
    # Prepare pre-fill data
    prefill = {
        'nationality': ind.get('nationality', ''),
        'residence_status': 'Resident' if ind.get('is_resident') else 'Non-Resident',
        'pep_status': ind.get('pep_status', 'No'),
        'profession': ind.get('profession', ''),
        'place_of_birth': ind.get('nationality', ''),
        'payment_method': '',
        'third_party': 'No',
    }
    
    # Calculate years in relationship
    years_relationship = 3  # default high risk
    if ind.get('created_at'):
        try:
            created = ind['created_at'] if isinstance(ind['created_at'], str) else str(ind['created_at'])[:10]
            from datetime import datetime
            created_dt = datetime.strptime(created, '%Y-%m-%d').date()
            years_diff = (dubai_today() - created_dt).days / 365.25
            if years_diff >= 5:
                years_relationship = 1
            elif years_diff >= 1:
                years_relationship = 2
        except:
            pass
    prefill['years_relationship'] = years_relationship
    
    conn.close()
    return render_template('risk_assessment_individual.html', individual=ind, prefill=prefill, lookups=RISK_LOOKUPS)

@app.route('/risk-assessment/individual/<int:id>/result')
@compliance_required
def risk_assessment_individual_result(id):
    """Display individual risk assessment result"""
    conn = get_db()
    ind = one(conn, 'SELECT * FROM clients WHERE id=?', (id,))
    assessment = one(conn, '''SELECT * FROM individual_risk_assessments 
                             WHERE individual_id=? ORDER BY created_at DESC LIMIT 1''', (id,))
    conn.close()
    
    if not ind or not assessment:
        return redirect(url_for('risk_assessment'))
    
    color_map = {'Low': '#22c55e', 'Medium': '#f59e0b', 'High': '#ef4444'}
    risk_color = color_map.get(assessment.get('risk_rating', 'Unspecified'), '#6b7280')
    
    return render_template('risk_assessment_result.html',
        individual=ind, assessment=assessment, risk_color=risk_color, assessment_type='individual')

@app.route('/reports')
@compliance_required
def reports():
    conn=get_db(); today=dubai_today()
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

# ──────────── AML TRACKER ────────────
@app.route('/aml-tracker')
@login_required
def aml_tracker():
    """List all AML Tracker records with filters"""
    try:
        conn = get_db()
        
        # Fallback: Ensure table exists (on-first-use pattern for Railway PostgreSQL)
        if is_pg(conn):
            x(conn, '''CREATE TABLE IF NOT EXISTS aml_tracker (
                id SERIAL PRIMARY KEY, company_id INTEGER, individual_id INTEGER,
                transaction_date DATE NOT NULL, period TEXT, due_date DATE, vc_no TEXT,
                payment_mode TEXT, ac_type TEXT, client_name TEXT, transaction_currency TEXT,
                usd_amount NUMERIC, aed_amount NUMERIC, payment_remarks TEXT, invoice_no TEXT,
                invoice_amount NUMERIC, invoice_currency TEXT, goaml_submission_date DATE,
                goaml_status TEXT DEFAULT 'pending', goaml_ref_no TEXT, submitted_by INTEGER NOT NULL,
                checked_by INTEGER, comment TEXT, verified_ledger BOOLEAN DEFAULT FALSE,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
            commit(conn)
        else:
            x(conn, '''CREATE TABLE IF NOT EXISTS aml_tracker (
                id INTEGER PRIMARY KEY AUTOINCREMENT, company_id INTEGER, individual_id INTEGER,
                transaction_date DATE NOT NULL, period TEXT, due_date DATE, vc_no TEXT,
                payment_mode TEXT, ac_type TEXT, client_name TEXT, transaction_currency TEXT,
                usd_amount NUMERIC, aed_amount NUMERIC, payment_remarks TEXT, invoice_no TEXT,
                invoice_amount NUMERIC, invoice_currency TEXT, goaml_submission_date DATE,
                goaml_status TEXT DEFAULT 'pending', goaml_ref_no TEXT, submitted_by INTEGER NOT NULL,
                checked_by INTEGER, comment TEXT, verified_ledger BOOLEAN DEFAULT FALSE,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
        status_filter = request.args.get('status', '')
        company_filter = request.args.get('company', '')
        vc_no_filter = request.args.get('vc_no', '')
        submitted_by_filter = request.args.get('submitted_by', '')
        
        # Build WHERE clause
        where = []
        params = []
        
        if status_filter:
            where.append(f'goaml_status = {P()}')
            params.append(status_filter)
        if company_filter:
            where.append(f'company_id = {P()}')
            params.append(company_filter)
        if vc_no_filter:
            where.append(f'vc_no LIKE {P()}')
            params.append(f'%{vc_no_filter}%')
        if submitted_by_filter:
            where.append(f'submitted_by = {P()}')
            params.append(submitted_by_filter)
        
        where_clause = ' AND '.join(where) if where else '1=1'
        
        # Get AML records with company/individual/user names
        sql = f'''SELECT a.*, c.client_name as company_name, cli.name as individual_name, u.email as submitted_by_email, 
                  u2.email as checked_by_email
                  FROM aml_tracker a
                  LEFT JOIN companies c ON a.company_id = c.id
                  LEFT JOIN clients cli ON a.individual_id = cli.id
                  LEFT JOIN users u ON a.submitted_by = u.id
                  LEFT JOIN users u2 ON a.checked_by = u2.id
                  WHERE {where_clause}
                  ORDER BY a.transaction_date DESC'''
        
        records = all_(conn, sql, params or None)
        
        # Get dropdown options for filters
        companies = all_(conn, 'SELECT id, client_name as company_name FROM companies ORDER BY client_name')
        individuals = all_(conn, 'SELECT id, name FROM clients ORDER BY name')
        users = all_(conn, 'SELECT id, email FROM users WHERE is_active=1 ORDER BY email')
        statuses = ['pending', 'submitted', 'approved', 'rejected']
        
        conn.close()
        return render_template('aml_tracker.html', 
                             records=records,
                             companies=companies,
                             users=users,
                             statuses=statuses,
                             status_filter=status_filter,
                             company_filter=company_filter,
                             vc_no_filter=vc_no_filter,
                             submitted_by_filter=submitted_by_filter)
    except Exception as e:
        logger.error(f'Error in aml_tracker: {e}')
        return render_template('aml_tracker.html', records=[], companies=[], 
                             users=[], statuses=[], error=str(e))

@app.route('/aml-tracker/add', methods=['GET', 'POST'])
@login_required
def aml_tracker_add():
    """Add new AML Tracker record"""
    try:
        conn = get_db()
        companies = all_(conn, 'SELECT id, client_name as company_name FROM companies ORDER BY client_name')
        individuals = all_(conn, 'SELECT id, name FROM clients ORDER BY name')
        users = all_(conn, 'SELECT id, email FROM users WHERE is_active=1 ORDER BY email')
        
        if request.method == 'POST':
            data = request.get_json() if request.is_json else request.form
            
            transaction_date = data.get('transaction_date')
            if not transaction_date:
                conn.close()
                return jsonify({'success': False, 'error': 'Transaction date is required'}), 400
            
            # Auto-calculate PERIOD and DUE DATE
            tdate = datetime.strptime(transaction_date, '%Y-%m-%d').date()
            month_num = tdate.month
            quarter_num = (month_num - 1) // 3 + 1
            week_num = tdate.isocalendar()[1]
            period = data.get('period', f'Month {month_num}, Q{quarter_num}, W{week_num}')
            due_date = tdate + timedelta(days=14)
            
            # Insert AML record
            x(conn, '''INSERT INTO aml_tracker 
               (company_id, individual_id, transaction_date, period, due_date, vc_no, payment_mode,
                ac_type, client_name, transaction_currency, usd_amount, aed_amount, payment_remarks,
                invoice_no, invoice_amount, invoice_currency, goaml_submission_date, goaml_status,
                goaml_ref_no, submitted_by, checked_by, comment, verified_ledger, created_at, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)''',
              (data.get('company_id') or None,
               data.get('individual_id') or None,
               transaction_date,
               period,
               str(due_date),
               data.get('vc_no'),
               data.get('payment_mode'),
               data.get('ac_type'),
               data.get('client_name'),
               data.get('transaction_currency'),
               data.get('usd_amount') or None,
               data.get('aed_amount') or None,
               data.get('payment_remarks'),
               data.get('invoice_no'),
               data.get('invoice_amount') or None,
               data.get('invoice_currency'),
               data.get('goaml_submission_date') or None,
               'pending',
               None,
               session.get('user_id'),
               data.get('checked_by') or None,
               data.get('comment'),
               'true' if data.get('verified_ledger') else 'false'))
            
            aml_id = lastid(conn)
            
            # Create task for checked_by user if specified
            checked_by_id = data.get('checked_by')
            if checked_by_id:
                x(conn, '''INSERT INTO tasks 
                   (title, description, assigned_to, created_by, status, due_date)
                   VALUES (?, ?, ?, ?, ?, ?)''',
                  (f'AML Verification - {data.get("client_name")}',
                   f'Review and verify AML Tracker record #{aml_id}. VC No: {data.get("vc_no")}',
                   int(checked_by_id),
                   session.get('user_id'),
                   'todo',
                   str(due_date)))
            
            commit(conn)
            conn.close()
            return jsonify({'success': True, 'message': 'AML record created', 'id': aml_id})
        
        conn.close()
        return render_template('aml_tracker_form.html', 
                             companies=companies, 
                             individuals=individuals,
                             users=users,
                             record=None)
    except Exception as e:
        logger.error(f'Error in aml_tracker_add: {e}')
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/aml-tracker/<int:id>/edit', methods=['GET', 'POST'])
@login_required
def aml_tracker_edit(id):
    """Edit AML Tracker record"""
    try:
        conn = get_db()
        record = one(conn, 'SELECT * FROM aml_tracker WHERE id=?', (id,))
        
        if not record:
            conn.close()
            return redirect(url_for('aml_tracker'))
        
        companies = all_(conn, 'SELECT id, client_name as company_name FROM companies ORDER BY client_name')
        individuals = all_(conn, 'SELECT id, name FROM clients ORDER BY name')
        users = all_(conn, 'SELECT id, email FROM users WHERE is_active=1 ORDER BY email')
        
        if request.method == 'POST':
            data = request.get_json() if request.is_json else request.form
            
            # Update record
            x(conn, '''UPDATE aml_tracker 
               SET vc_no=?, payment_mode=?, ac_type=?, client_name=?, transaction_currency=?,
                   usd_amount=?, aed_amount=?, payment_remarks=?, invoice_no=?, invoice_amount=?,
                   invoice_currency=?, goaml_submission_date=?, goaml_status=?, goaml_ref_no=?,
                   checked_by=?, comment=?, verified_ledger=?, updated_at=CURRENT_TIMESTAMP
               WHERE id=?''',
              (data.get('vc_no'),
               data.get('payment_mode'),
               data.get('ac_type'),
               data.get('client_name'),
               data.get('transaction_currency'),
               data.get('usd_amount') or None,
               data.get('aed_amount') or None,
               data.get('payment_remarks'),
               data.get('invoice_no'),
               data.get('invoice_amount') or None,
               data.get('invoice_currency'),
               data.get('goaml_submission_date') or None,
               data.get('goaml_status'),
               data.get('goaml_ref_no'),
               data.get('checked_by') or None,
               data.get('comment'),
               'true' if data.get('verified_ledger') else 'false',
               id))
            
            commit(conn)
            conn.close()
            return jsonify({'success': True, 'message': 'AML record updated'})
        
        # Format dates for template
        for field in ['transaction_date', 'due_date', 'goaml_submission_date']:
            if record.get(field):
                record[field] = str(record[field])[:10] if record[field] else None
        
        conn.close()
        return render_template('aml_tracker_form.html',
                             companies=companies,
                             individuals=individuals,
                             users=users,
                             record=record)
    except Exception as e:
        logger.error(f'Error in aml_tracker_edit: {e}')
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/aml-tracker/<int:id>/delete', methods=['POST'])
@login_required
def api_aml_tracker_delete(id):
    """Delete AML Tracker record"""
    try:
        conn = get_db()
        x(conn, 'DELETE FROM aml_tracker WHERE id=?', (id,))
        commit(conn)
        conn.close()
        return jsonify({'success': True, 'message': 'AML record deleted'})
    except Exception as e:
        logger.error(f'Error deleting AML record: {e}')
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/settings')
@admin_required
def settings():
    conn=get_db()
    users=all_(conn,'SELECT id,email,name,role,is_active,contact_number,username,permissions FROM users ORDER BY role,name')
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
        username = d.get('username','').strip().lower()
        if username:
            existing = one(conn,'SELECT id FROM users WHERE username=?',(username,))
            if existing: return jsonify({'success':False,'error':'Username already taken'}),400
        x(conn,'INSERT INTO users (email,password_hash,name,role,contact_number,username,permissions,is_active) VALUES (?,?,?,?,?,?,?,1)',
          (d.get('email',''),generate_password_hash(d.get('password','')),d.get('name'),d.get('role','staff'),d.get('contact_number'),username or None,d.get('permissions','')))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e:
        logger.error(f'Error in %s: {e}', request.path)
        return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/user/<int:id>/edit',methods=['POST'])
@admin_required
def api_edit_user(id):
    d=request.get_json()
    try:
        conn=get_db()
        username = d.get('username','').strip().lower() or None
        if username:
            existing = one(conn,'SELECT id FROM users WHERE username=? AND id!=?',(username,id))
            if existing: return jsonify({'success':False,'error':'Username already taken'}),400
        if d.get('password'):
            x(conn,'UPDATE users SET name=?,email=?,role=?,contact_number=?,username=?,permissions=?,password_hash=? WHERE id=?',
              (d.get('name'),d.get('email'),d.get('role'),d.get('contact_number'),username,d.get('permissions',''),generate_password_hash(d.get('password')),id))
        else:
            x(conn,'UPDATE users SET name=?,email=?,role=?,contact_number=?,username=?,permissions=? WHERE id=?',
              (d.get('name'),d.get('email'),d.get('role'),d.get('contact_number'),username,d.get('permissions',''),id))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e:
        logger.error(f'Error in %s: {e}', request.path)
        return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/user/<int:id>/toggle',methods=['POST'])
@admin_required
def api_toggle_user(id):
    try:
        conn=get_db(); u=one(conn,'SELECT is_active FROM users WHERE id=?',(id,))
        if u: x(conn,'UPDATE users SET is_active=? WHERE id=?',(0 if u['is_active'] else 1,id))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e:
        logger.error(f'Error in %s: {e}', request.path)
        return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/user/<int:id>/delete',methods=['POST'])
@admin_required
def api_delete_user(id):
    try:
        conn=get_db(); x(conn,'DELETE FROM users WHERE id=?',(id,))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e:
        logger.error(f'Error in %s: {e}', request.path)
        return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/dropdown/add',methods=['POST'])
@admin_required
def api_add_dropdown():
    d=request.get_json()
    try:
        conn=get_db()
        x(conn,'INSERT INTO dropdowns (field_name,value,is_active) VALUES (?,?,1)',(d.get('field_name'),d.get('value')))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e:
        logger.error(f'Error in %s: {e}', request.path)
        return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/dropdown/<int:id>/delete',methods=['POST'])
@admin_required
def api_delete_dropdown(id):
    try:
        conn=get_db(); x(conn,'DELETE FROM dropdowns WHERE id=?',(id,))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e:
        logger.error(f'Error in %s: {e}', request.path)
        return jsonify({'success':False,'error':str(e)}),500

@app.route('/tasks')
@login_required
def tasks():
    conn=get_db(); uid=session.get('user_id'); role=session.get('user_role')
    today=dubai_today()
    active_tab=request.args.get('tab','temp')

    # ── TEMP TASKS ──
    base='''SELECT t.*,u.name as assigned_name,u.mobile as assigned_mobile,
        cu.name as created_by_name,c.client_name as company_name,c.ac_code
        FROM tasks t LEFT JOIN users u ON t.assigned_to=u.id
        LEFT JOIN users cu ON t.created_by=cu.id
        LEFT JOIN companies c ON t.company_id=c.id'''
    order=" ORDER BY CASE t.priority WHEN 'urgent' THEN 1 WHEN 'high' THEN 2 WHEN 'normal' THEN 3 ELSE 4 END,t.due_date"
    if role == 'admin':
        rows=all_(conn,base+order)
    else:
        rows=all_(conn,base+" WHERE (t.assigned_to=? OR t.created_by=?)"+order,(uid,uid))
    users=all_(conn,'SELECT id,name,role,mobile FROM users WHERE is_active=1 ORDER BY name')
    cos=all_(conn,'SELECT id,ac_code,client_name FROM companies ORDER BY client_name')
    tmpls=all_(conn,"SELECT value,description FROM dropdowns WHERE field_name='TASK TEMPLATE' AND is_active=1 ORDER BY value")
    tl=[]
    for t in rows:
        d=days_left(t['due_date'])
        tl.append({**t,'priority':t['priority'] or 'normal','status':t['status'] or 'todo',
                   'due_date':str(t['due_date']) if t['due_date'] else None,
                   'days_until_due':d,'is_overdue':(d is not None and d<0)})

    # ── REGULAR TASKS ──
    def due_dates_for_frequency(freq, since_date):
        dates=[]
        if freq=='daily':
            d=since_date
            while d<=today:
                dates.append(d); d+=timedelta(days=1)
        elif freq=='weekly':
            d=since_date
            days_ahead=(4-d.weekday())%7
            d=d+timedelta(days=days_ahead)
            while d<=today:
                dates.append(d); d+=timedelta(weeks=1)
        elif freq=='monthly':
            d=since_date.replace(day=1)
            while True:
                try: due=d.replace(day=25)
                except ValueError: due=None
                if due and due>=since_date and due<=today: dates.append(due)
                if d.month==12: d=d.replace(year=d.year+1,month=1)
                else: d=d.replace(month=d.month+1)
                if d>today: break
        return dates

    try:
        if role in ['admin','compliance']:
            rt_templates=all_(conn,"""SELECT rt.*,u.name as created_by_name,au.name as assigned_user_name
                FROM regular_task_templates rt LEFT JOIN users u ON rt.created_by=u.id
                LEFT JOIN users au ON rt.assigned_user_id=au.id
                ORDER BY rt.frequency,rt.title""")
            rt_logs=all_(conn,"""SELECT l.*,u.name as staff_name,rt.title as task_title,rt.frequency
                FROM regular_task_logs l JOIN users u ON l.user_id=u.id
                JOIN regular_task_templates rt ON l.template_id=rt.id
                ORDER BY l.logged_at DESC LIMIT 200""")
        else:
            rt_templates=all_(conn,"""SELECT rt.*,u.name as created_by_name,au.name as assigned_user_name
                FROM regular_task_templates rt LEFT JOIN users u ON rt.created_by=u.id
                LEFT JOIN users au ON rt.assigned_user_id=au.id
                WHERE rt.assigned_role='all' OR rt.assigned_user_id=? OR rt.assigned_role=?
                ORDER BY rt.frequency,rt.title""", (uid,role))
            rt_logs=all_(conn,"""SELECT l.*,u.name as staff_name,rt.title as task_title,rt.frequency
                FROM regular_task_logs l JOIN users u ON l.user_id=u.id
                JOIN regular_task_templates rt ON l.template_id=rt.id
                WHERE l.user_id=? ORDER BY l.logged_at DESC LIMIT 100""", (uid,))
    except: rt_templates=[]; rt_logs=[]

    task_status={}
    for t in rt_templates:
        try:
            user_logs=all_(conn,"""SELECT DATE(logged_at) as log_date FROM regular_task_logs
                WHERE template_id=? AND user_id=? ORDER BY logged_at ASC""", (t['id'],uid))
            logged_dates=set(r['log_date'] for r in user_logs)
            created=t.get('created_at')
            try: start=datetime.strptime(str(created)[:10],'%Y-%m-%d').date()
            except: start=today
            all_due=due_dates_for_frequency(t['frequency'],start)
            missed=[d for d in all_due if str(d) not in logged_dates]
            freq=t['frequency']
            if freq=='daily': next_due=today+timedelta(days=1)
            elif freq=='weekly':
                days_ahead=(4-today.weekday())%7
                next_due=today+timedelta(days=(days_ahead if days_ahead>0 else 7))
            elif freq=='monthly':
                if today.day<25: next_due=today.replace(day=25)
                elif today.month==12: next_due=today.replace(year=today.year+1,month=1,day=25)
                else: next_due=today.replace(month=today.month+1,day=25)
            else: next_due=today
            is_due_today=str(today) in [str(d) for d in all_due] and str(today) not in logged_dates
            task_status[t['id']]={
                'overdue_count':len(missed),'is_due_today':is_due_today,
                'next_due':str(next_due),'last_logged':user_logs[-1]['log_date'] if user_logs else None,
                'missed_dates':[str(d) for d in missed[-3:]]
            }
        except: task_status[t['id']]={'overdue_count':0,'is_due_today':False,'next_due':str(today),'last_logged':None,'missed_dates':[]}

    # ── ADDITIONAL TASKS ──
    try:
        # Ensure table exists on Railway PG
        if is_pg(conn):
            x(conn, '''CREATE TABLE IF NOT EXISTS additional_tasks (
                id SERIAL PRIMARY KEY, title TEXT NOT NULL,
                task_details TEXT, remarks TEXT,
                from_datetime TIMESTAMP NOT NULL, to_datetime TIMESTAMP NOT NULL,
                created_by INTEGER NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
        else:
            x(conn, '''CREATE TABLE IF NOT EXISTS additional_tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL,
                task_details TEXT, remarks TEXT,
                from_datetime TIMESTAMP NOT NULL, to_datetime TIMESTAMP NOT NULL,
                created_by INTEGER NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
        commit(conn)
    except: pass
    try:
        if role == 'admin':
            add_tasks = all_(conn, '''SELECT at.*,u.name as staff_name
                FROM additional_tasks at LEFT JOIN users u ON at.created_by=u.id
                ORDER BY at.from_datetime DESC''')
        else:
            add_tasks = all_(conn, '''SELECT at.*,u.name as staff_name
                FROM additional_tasks at LEFT JOIN users u ON at.created_by=u.id
                WHERE at.created_by=?
                ORDER BY at.from_datetime DESC''', (uid,))
        add_tasks = [{**t,
            'from_datetime': str(t['from_datetime'])[:16] if t['from_datetime'] else '',
            'to_datetime': str(t['to_datetime'])[:16] if t['to_datetime'] else '',
            'created_at': str(t['created_at'])[:10] if t['created_at'] else ''
        } for t in add_tasks]
    except: add_tasks = []

    conn.close()
    return render_template('tasks.html',tasks=tl,all_users=users,all_companies=cos,task_templates=tmpls,
                           rt_templates=rt_templates,rt_logs=rt_logs,task_status=task_status,
                           active_tab=active_tab,today=str(today),add_tasks=add_tasks,
                           current_user_id=uid)

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
    except Exception as e:
        logger.error(f'Error in %s: {e}', request.path)
        return jsonify({'success':False,'error':str(e)}),500

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
    except Exception as e:
        logger.error(f'Error in %s: {e}', request.path)
        return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/task/<int:id>/status',methods=['POST'])
@login_required
def api_task_status(id):
    d=request.get_json()
    try:
        conn=get_db()
        x(conn,'UPDATE tasks SET status=?,updated_at=CURRENT_TIMESTAMP WHERE id=?',(d.get('status'),id))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e:
        logger.error(f'Error in %s: {e}', request.path)
        return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/task/<int:id>/delete',methods=['POST'])
@login_required
def api_delete_task(id):
    if session.get('user_role') == 'staff':
        return jsonify({'success':False,'error':'Staff cannot delete tasks'}),403
    try:
        conn=get_db(); x(conn,'DELETE FROM tasks WHERE id=?',(id,))
        commit(conn); conn.close(); return jsonify({'success':True})
    except Exception as e:
        logger.error(f'Error in %s: {e}', request.path)
        return jsonify({'success':False,'error':str(e)}),500

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
    except Exception as e:
        logger.error(f'Error in %s: {e}', request.path)
        return jsonify({'success':False,'error':str(e)}),500

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
    except Exception as e:
        logger.error(f'Error in %s: {e}', request.path)
        return jsonify({'success':False,'error':str(e)}),500

@app.route('/export/template')
@compliance_required
def export_template():
    if not HAS_XL: return "openpyxl not installed",500
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = openpyxl.Workbook()

    # ── SHEET 1: DATA ENTRY ──────────────────────────────────
    ws = wb.active
    ws.title = "Companies"

    headers = ['ac_code','client_name','ac_opening_date','ac_status','nature','type_of_client',
               'name_of_freezone','mode_of_ac','country_of_incorporation','region','address',
               'telephone','mobile','whatsapp_number','email_id','contact_person_name',
               'contact_person_number','account_manager','address_proof_type','address_proof_expiry',
               'trade_license_no','issuing_authority','legal_type','incorporation_date',
               'trade_license_expiry','tax_no_trn','vat_cert','vat_declaration','num_beneficial_owners',
               'moa','pep','undertaking','source_of_fund','doc_status','risk_status','kyc_status',
               'verified_by','followup_details','zewer_comments']

    # Style header row
    header_fill = PatternFill(start_color="1C1917", end_color="1C1917", fill_type="solid")
    header_font = Font(color="D97706", bold=True, size=11)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h.replace('_',' ').title())
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align

    ws.row_dimensions[1].height = 30

    # Sample row
    sample = ['TJ5003','SAMPLE COMPANY LLC','2020-01-01','Active','Legal entity','MainLand','N/A',
              'Supplier','United Arab Emirates','Dubai','Unit 101, Gold Souq, Dubai','+97142258019',
              '+971506594165','+971506594165','info@sample.com','Ahmed Ali','+971501234567','Jaseel',
              'Ejari','2025-12-31','534230','Dubai Economy & Tourism','Limited Liability Company(LLC)',
              '2019-06-01','2025-12-31','100003063300003','Yes','Yes',2,'Yes','Yes','Yes','Yes',
              'Completed','Medium','Kyc 2025 Updated','Jaseel','','']
    ws.append(sample)

    # Style sample row
    sample_fill = PatternFill(start_color="292524", end_color="292524", fill_type="solid")
    sample_font = Font(color="A8A29E", size=10, italic=True)
    for col in range(1, len(headers)+1):
        cell = ws.cell(row=2, column=col)
        cell.fill = sample_fill
        cell.font = sample_font
        cell.alignment = Alignment(horizontal="center")

    # Column widths
    col_widths = {'ac_code':12,'client_name':30,'ac_opening_date':16,'ac_status':12,'nature':16,
                  'type_of_client':18,'name_of_freezone':20,'mode_of_ac':16,'country_of_incorporation':22,
                  'region':16,'address':30,'telephone':18,'mobile':18,'whatsapp_number':18,
                  'email_id':25,'contact_person_name':22,'contact_person_number':20,'account_manager':18,
                  'address_proof_type':22,'address_proof_expiry':18,'trade_license_no':18,
                  'issuing_authority':28,'legal_type':30,'incorporation_date':18,'trade_license_expiry':18,
                  'tax_no_trn':16,'vat_cert':12,'vat_declaration':16,'num_beneficial_owners':12,
                  'moa':8,'pep':8,'undertaking':14,'source_of_fund':16,'doc_status':14,
                  'risk_status':12,'kyc_status':22,'verified_by':16,'followup_details':30,'zewer_comments':30}
    for col, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(col)].width = col_widths.get(h, 16)

    # Freeze top 2 rows (header + sample)
    ws.freeze_panes = 'A3'

    # ── CELL DROPDOWN VALIDATION ──────────────────────────────
    conn = get_db()
    dd_rows = all_(conn, "SELECT field_name, value FROM dropdowns WHERE is_active=1 ORDER BY field_name, value")
    conn.close()

    dd = {}
    for r in dd_rows:
        dd.setdefault(r['field_name'], []).append(r['value'])

    # Map column field names to DB dropdown field_names
    dropdown_map = {
        'ac_status': 'AC STATUS',
        'nature': 'NATURE',
        'type_of_client': 'TYPE OF CLIENT',
        'name_of_freezone': 'NAME OF FREEZONE',
        'mode_of_ac': 'MODE OF AC',
        'country_of_incorporation': 'COUNTRY',
        'region': 'REGION',
        'address_proof_type': 'ADDRESS PROOF TYPE',
        'account_manager': 'ACCOUNT MANAGER',
        'issuing_authority': 'ISSUING AUTHORITY',
        'legal_type': 'LEGAL TYPE',
        'vat_cert': 'VAT CERT',
        'vat_declaration': 'VAT DECLARATION',
        'moa': 'MOA',
        'pep': 'PEP',
        'undertaking': 'UNDERTAKING',
        'source_of_fund': 'SOURCE OF FUND',
        'doc_status': 'DOC STATUS',
        'risk_status': 'RISK STATUS',
        'kyc_status': 'KYC STATUS',
    }

    for col_idx, h in enumerate(headers, 1):
        db_field = dropdown_map.get(h)
        if not db_field:
            continue
        values = dd.get(db_field, [])
        if not values:
            continue
        joined = ','.join(values)
        if len(joined) > 250:
            joined = joined[:250].rsplit(',', 1)[0]  # trim to fit Excel limit
        formula = '\"' + joined.replace(',', '\",\"') + '\"'
        dv = DataValidation(
            type="list",
            formula1='"' + joined + '"',
            allow_blank=True,
            showDropDown=False,
            showErrorMessage=True,
            errorTitle="Invalid Value",
            error="Please select a value from the dropdown list."
        )
        col_letter = get_column_letter(col_idx)
        dv.sqref = f"{col_letter}3:{col_letter}1000"
        ws.add_data_validation(dv)

    # ── SHEET 2: INSTRUCTIONS ───────────────────────────────
    ws3 = wb.create_sheet("Instructions")
    instructions = [
        ["ZEWER AML CRM — Company Import Template"],
        [""],
        ["HOW TO USE:"],
        ["1. Fill in company data in the 'Companies' sheet starting from Row 3"],
        ["2. Row 2 is a sample — you can delete it before importing"],
        ["3. AC Code and Client Name are REQUIRED — all others are optional"],
        ["4. For dropdown fields, click the cell — a dropdown arrow will appear to select valid values"],
        ["5. Dates must be in YYYY-MM-DD format (e.g. 2025-12-31)"],
        ["6. Phone numbers should include country code (e.g. +97142258019)"],
        ["7. Duplicate AC Codes will be skipped on import"],
        [""],
        ["REQUIRED FIELDS:"],
        ["  • ac_code — Unique account code (e.g. TJ5003)"],
        ["  • client_name — Full legal company name"],
        [""],
        ["DATE FORMAT:"],
        ["  • ac_opening_date, address_proof_expiry, incorporation_date,"],
        ["    trade_license_expiry — all use YYYY-MM-DD"],
    ]
    for r, row_data in enumerate(instructions, 1):
        cell = ws3.cell(row=r, column=1, value=row_data[0] if row_data else '')
        if r == 1:
            cell.font = Font(bold=True, size=14, color="D97706")
        elif row_data and row_data[0].startswith(('HOW','REQUIRED','DATE')):
            cell.font = Font(bold=True, size=11, color="F5F5F4")
        else:
            cell.font = Font(size=10, color="A8A29E")
    ws3.column_dimensions['A'].width = 70
    ws3.sheet_view.showGridLines = False

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True, download_name='zewer_company_template.xlsx')

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
    conn=get_db(); today=dubai_today()
    
    if rt=='expiry':
        # Use the same logic as alerts page
        all_alerts = []
        
        # Trade licenses
        tl_rows = all_(conn,'''SELECT id,ac_code,client_name,mobile,whatsapp_number,
            trade_license_expiry,risk_status,region,account_manager,
            contact_person_name,contact_person_number
            FROM companies WHERE trade_license_expiry IS NOT NULL ORDER BY trade_license_expiry''')
        for r in tl_rows:
            d = days_left(r['trade_license_expiry'])
            if d is None: continue
            all_alerts.append({'ac_code':r['ac_code'],'client_name':r['client_name'],
                'doc_type':'Trade License','expiry_date':str(r['trade_license_expiry'])[:10],'days':d})
        
        # Address proofs
        ap_rows = all_(conn,'''SELECT id,ac_code,client_name,address_proof_type,address_proof_expiry,account_manager
            FROM companies WHERE address_proof_expiry IS NOT NULL ORDER BY address_proof_expiry''')
        for r in ap_rows:
            d = days_left(r['address_proof_expiry'])
            if d is None: continue
            all_alerts.append({'ac_code':r['ac_code'],'client_name':r['client_name'],
                'doc_type':'Address Proof','expiry_date':str(r['address_proof_expiry'])[:10],'days':d})
        
        # Passports from UBOs
        ubo_rows = all_(conn,'''SELECT u.person_name,u.passport_no,u.passport_expiry,
            u.emirates_id,u.emirates_id_expiry,
            c.id as company_id,c.client_name,c.ac_code
            FROM ubos u JOIN companies c ON u.company_id=c.id
            WHERE u.passport_expiry IS NOT NULL OR u.emirates_id_expiry IS NOT NULL''')
        for r in ubo_rows:
            if r.get('passport_expiry'):
                d = days_left(r['passport_expiry'])
                if d is not None:
                    all_alerts.append({'ac_code':r['ac_code'],'client_name':r['client_name'],
                        'doc_type':'Passport','expiry_date':str(r['passport_expiry'])[:10],'days':d})
            if r.get('emirates_id_expiry'):
                d = days_left(r['emirates_id_expiry'])
                if d is not None:
                    all_alerts.append({'ac_code':r['ac_code'],'client_name':r['client_name'],
                        'doc_type':'Emirates ID','expiry_date':str(r['emirates_id_expiry'])[:10],'days':d})
        
        # Clients documents
        client_rows = all_(conn,'''SELECT id,name,phone,account_number,
            passport_expiry,emirates_id_expiry FROM clients 
            WHERE passport_expiry IS NOT NULL OR emirates_id_expiry IS NOT NULL''')
        for r in client_rows:
            if r.get('passport_expiry'):
                d = days_left(r['passport_expiry'])
                if d is not None:
                    all_alerts.append({'ac_code':r.get('account_number','N/A'),'client_name':r['name'],
                        'doc_type':'Passport (Individual)','expiry_date':str(r['passport_expiry'])[:10],'days':d})
            if r.get('emirates_id_expiry'):
                d = days_left(r['emirates_id_expiry'])
                if d is not None:
                    all_alerts.append({'ac_code':r.get('account_number','N/A'),'client_name':r['name'],
                        'doc_type':'Emirates ID (Individual)','expiry_date':str(r['emirates_id_expiry'])[:10],'days':d})
        
        # Apply filters
        expiry_filter = request.args.get('expiry_filter', 'all')
        doc_type_filter = request.args.get('doc_type', 'all')
        search_filter = request.args.get('search', '').lower()
        
        filtered = []
        for row in all_alerts:
            # Expiry filter
            if expiry_filter == 'expired' and row['days'] >= 0:
                continue
            if expiry_filter == '30' and (row['days'] >= 30 or row['days'] < 0):
                continue
            if expiry_filter == '60' and (row['days'] >= 60 or row['days'] < 0):
                continue
            if expiry_filter == '90' and (row['days'] >= 90 or row['days'] < 0):
                continue
            
            # Doc type filter
            if doc_type_filter == 'trade' and 'Trade License' not in row.get('doc_type', ''):
                continue
            if doc_type_filter == 'address' and 'Address Proof' not in row.get('doc_type', ''):
                continue
            if doc_type_filter == 'passport' and 'Passport' not in row.get('doc_type', ''):
                continue
            if doc_type_filter == 'eid' and 'Emirates ID' not in row.get('doc_type', ''):
                continue
            
            # Search filter
            if search_filter and search_filter not in str(row.get('ac_code', '')).lower() and search_filter not in str(row.get('client_name', '')).lower():
                continue
            
            filtered.append(row)
        
        rows = filtered
    elif rt=='kyc':
        w='1=1'; p=[]
        df=request.args.get('from',''); dt=request.args.get('to','')
        if df: w+=' AND created_at >= ?'; p.append(df)
        if dt: w+=' AND created_at <= ?'; p.append(dt+' 23:59:59')
        rows=all_(conn,f'SELECT ac_code,client_name,kyc_status,kyc_expiry_date,doc_status,risk_status,verified_by,verified_date,followup_details FROM companies WHERE {w} ORDER BY kyc_status',p or None)
    elif rt=='risk':
        w='1=1'; p=[]
        df=request.args.get('from',''); dt=request.args.get('to','')
        if df: w+=' AND created_at >= ?'; p.append(df)
        if dt: w+=' AND created_at <= ?'; p.append(dt+' 23:59:59')
        rows=all_(conn,f'SELECT ac_code,client_name,risk_status,doc_status,kyc_status,pep,source_of_fund,moa,undertaking,registration_screening_tool,region,account_manager FROM companies WHERE {w} ORDER BY risk_status',p or None)
    elif rt=='ubo':
        rows=all_(conn,'''SELECT c.ac_code,c.client_name,u.person_name,u.position,u.share_percentage,
            u.nationality,u.residential_status,u.passport_no,u.passport_expiry,
            u.emirates_id,u.emirates_id_expiry,u.doc_status,u.verified_by,u.verified_date
            FROM ubos u JOIN companies c ON u.company_id=c.id
            ORDER BY c.ac_code,u.share_percentage DESC''')
    elif rt=='tasks':
        status_f=request.args.get('status','all')
        w='1=1'; p=[]
        if status_f and status_f!='all': w+=' AND t.status=?'; p.append(status_f)
        df=request.args.get('from',''); dt=request.args.get('to','')
        if df: w+=' AND t.created_at >= ?'; p.append(df)
        if dt: w+=' AND t.created_at <= ?'; p.append(dt+' 23:59:59')
        rows=all_(conn,f'''SELECT t.title,t.description,t.priority,t.status,t.due_date,
            t.created_at,u.name as assigned_to,cu.name as created_by,c.ac_code,c.client_name as company
            FROM tasks t LEFT JOIN users u ON t.assigned_to=u.id
            LEFT JOIN users cu ON t.created_by=cu.id
            LEFT JOIN companies c ON t.company_id=c.id
            WHERE {w} ORDER BY t.priority,t.due_date''', p or None)
    elif rt=='all':
        w='1=1'; p=[]
        df=request.args.get('from',''); dt=request.args.get('to','')
        if df: w+=' AND created_at >= ?'; p.append(df)
        if dt: w+=' AND created_at <= ?'; p.append(dt+' 23:59:59')
        rows=all_(conn,f'SELECT ac_code,client_name,ac_status,nature,type_of_client,mode_of_ac,region,telephone,mobile,email_id,contact_person_name,account_manager,risk_status,doc_status,kyc_status,trade_license_no,trade_license_expiry,address_proof_expiry,created_at FROM companies WHERE {w} ORDER BY client_name',p or None)
    else:
        rows=all_(conn,'SELECT * FROM companies ORDER BY client_name')
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
        headers=[str(c.value).strip().lower().replace(' ','_') if c.value else '' for c in ws[1]]
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
    except Exception as e:
        logger.error(f'Error in %s: {e}', request.path)
        return jsonify({'success':False,'error':str(e)}),500

if __name__=='__main__':
    app.run(debug=False,host='0.0.0.0',port=int(os.getenv('PORT',8000)))


# ════════════════════════════════════════════════════════════════
# NEW FEATURES BLOCK
# ════════════════════════════════════════════════════════════════

# ── REGULAR TASKS → redirects to /tasks?tab=regular ──────────
@app.route('/regular-tasks')
@login_required
def regular_tasks():
    return redirect(url_for('tasks', tab='regular'))


@app.route('/api/regular-task/add', methods=['POST'])
@compliance_required
def api_add_regular_task():
    d = request.get_json()
    try:
        conn = get_db()
        x(conn, '''INSERT INTO regular_task_templates
            (title, description, frequency, assigned_role, assigned_user_id, created_by)
            VALUES (?,?,?,?,?,?)''',
          (d.get('title'), d.get('description'), d.get('frequency', 'daily'),
           d.get('assigned_role', 'all'), d.get('assigned_user_id') or None,
           session.get('user_id')))
        commit(conn); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        try: conn.close()
        except: pass
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/regular-task/<int:id>/delete', methods=['POST'])
@compliance_required
def api_delete_regular_task(id):
    try:
        conn = get_db()
        x(conn, 'DELETE FROM regular_task_logs WHERE template_id=?', (id,))
        x(conn, 'DELETE FROM regular_task_templates WHERE id=?', (id,))
        commit(conn); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        try: conn.close()
        except: pass
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/regular-task/<int:id>/log', methods=['POST'])
@login_required
def api_log_regular_task(id):
    d = request.get_json()
    try:
        conn = get_db()
        x(conn, '''INSERT INTO regular_task_logs (template_id, user_id, notes, status)
            VALUES (?,?,?,?)''',
          (id, session.get('user_id'), d.get('notes', ''), d.get('status', 'done')))
        commit(conn); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        try: conn.close()
        except: pass
        return jsonify({'success': False, 'error': str(e)}), 500

# ════════════════════════════════════════════════════════════
# ADDITIONAL TASKS
# ════════════════════════════════════════════════════════════
@app.route('/api/additional-task/add', methods=['POST'])
@login_required
def api_add_additional_task():
    d = request.get_json()
    try:
        conn = get_db()
        # Ensure table exists (Railway PG may not have run migration yet)
        try:
            if is_pg(conn):
                x(conn, '''CREATE TABLE IF NOT EXISTS additional_tasks (
                    id SERIAL PRIMARY KEY, title TEXT NOT NULL,
                    task_details TEXT, remarks TEXT,
                    from_datetime TIMESTAMP NOT NULL, to_datetime TIMESTAMP NOT NULL,
                    created_by INTEGER NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
            else:
                x(conn, '''CREATE TABLE IF NOT EXISTS additional_tasks (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL,
                    task_details TEXT, remarks TEXT,
                    from_datetime TIMESTAMP NOT NULL, to_datetime TIMESTAMP NOT NULL,
                    created_by INTEGER NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
            commit(conn)
        except: pass
        x(conn, '''INSERT INTO additional_tasks
            (title, task_details, remarks, from_datetime, to_datetime, created_by)
            VALUES (?,?,?,?,?,?)''',
          (d.get('title'), d.get('task_details'), d.get('remarks'),
           d.get('from_datetime'), d.get('to_datetime'), session.get('user_id')))
        commit(conn); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/additional-task/<int:id>/delete', methods=['POST'])
@login_required
def api_delete_additional_task(id):
    try:
        conn = get_db()
        uid = session.get('user_id'); role = session.get('user_role')
        task = one(conn, 'SELECT created_by FROM additional_tasks WHERE id=?', (id,))
        if not task: return jsonify({'success': False, 'error': 'Not found'}), 404
        if role != 'admin' and task['created_by'] != uid:
            return jsonify({'success': False, 'error': 'Not allowed'}), 403
        x(conn, 'DELETE FROM additional_tasks WHERE id=?', (id,))
        commit(conn); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ── INTERNAL DOCUMENTS (ZEWER STAFF / COMPANY DOCS) ──────────
@app.route('/internal-docs')
@compliance_required
def internal_docs():
    conn = get_db()
    try:
        docs = all_(conn, '''SELECT d.*,u.name as added_by_name
            FROM internal_documents d LEFT JOIN users u ON d.added_by=u.id
            ORDER BY CASE WHEN d.expiry_date IS NULL THEN 1 ELSE 0 END, d.expiry_date ASC, d.doc_category, d.doc_name''')
    except: docs = []
    today = dubai_today()
    doc_list = []
    for d in docs:
        dl = days_left(d['expiry_date'])
        doc_list.append({**d,
            'days_left': dl,
            'exp_status': exp_status(dl),
            'expiry_date': str(d['expiry_date']) if d['expiry_date'] else None})
    users = all_(conn, 'SELECT id,name FROM users WHERE is_active=1 ORDER BY name')
    conn.close()
    return render_template('internal_docs.html', docs=doc_list, today=str(today), all_users=users)

@app.route('/api/internal-doc/add', methods=['POST'])
@compliance_required
def api_add_internal_doc():
    d = request.get_json()
    try:
        conn = get_db()
        x(conn, '''INSERT INTO internal_documents
            (doc_name, doc_category, person_name, issuing_authority, issue_date, expiry_date, notes, added_by)
            VALUES (?,?,?,?,?,?,?,?)''',
          (d.get('doc_name'), d.get('doc_category'), d.get('person_name'),
           d.get('issuing_authority'), d.get('issue_date') or None,
           d.get('expiry_date') or None, d.get('notes'), session.get('user_id')))
        commit(conn); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/internal-doc/<int:id>/edit', methods=['POST'])
@compliance_required
def api_edit_internal_doc(id):
    d = request.get_json()
    try:
        conn = get_db()
        x(conn, '''UPDATE internal_documents SET doc_name=?,doc_category=?,person_name=?,
            issuing_authority=?,issue_date=?,expiry_date=?,notes=?,updated_at=CURRENT_TIMESTAMP
            WHERE id=?''',
          (d.get('doc_name'), d.get('doc_category'), d.get('person_name'),
           d.get('issuing_authority'), d.get('issue_date') or None,
           d.get('expiry_date') or None, d.get('notes'), id))
        commit(conn); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/internal-doc/<int:id>/delete', methods=['POST'])
@compliance_required
def api_delete_internal_doc(id):
    try:
        conn = get_db()
        x(conn, 'DELETE FROM internal_documents WHERE id=?', (id,))
        commit(conn); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/company/<int:cid>/docs-export')
@compliance_required
def api_export_company_docs(cid):
    if not HAS_XL:
        return "openpyxl not installed", 500
    from openpyxl.styles import Font, PatternFill, Alignment
    conn = get_db()
    co = one(conn, 'SELECT * FROM companies WHERE id=?', (cid,))
    ubos = all_(conn, 'SELECT * FROM ubos WHERE company_id=?', (cid,))
    conn.close()
    if not co:
        return "Company not found", 404
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Documents"
    hf = PatternFill(start_color="1C1917", end_color="1C1917", fill_type="solid")
    hfont = Font(color="D97706", bold=True, size=11)
    headers = ['Document Type', 'Reference / Number', 'Issuing Authority', 'Expiry Date', 'Status']
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = hfont; c.fill = hf; c.alignment = Alignment(horizontal='center')
    today = dubai_today()
    def status(exp):
        if not exp: return '—'
        try:
            d = datetime.strptime(str(exp)[:10], '%Y-%m-%d').date()
            diff = (d - today).days
            if diff < 0: return 'EXPIRED'
            elif diff <= 30: return f'{diff}d - CRITICAL'
            elif diff <= 90: return f'{diff}d - WARNING'
            else: return f'{diff}d - OK'
        except: return '—'
    rows = [
        ['Trade License', co['trade_license_no'], co['issuing_authority'], co['trade_license_expiry'], status(co['trade_license_expiry'])],
        ['Address Proof', co['address_proof_type'], '—', co['address_proof_expiry'], status(co['address_proof_expiry'])],
        ['VAT Certificate', co['tax_no_trn'], '—', '—', co['vat_cert'] or '—'],
    ]
    for ubo in ubos:
        rows.append([f"Passport ({ubo['person_name']})", ubo['passport_no'], ubo['nationality'], ubo['passport_expiry'], status(ubo['passport_expiry'])])
        rows.append([f"Emirates ID ({ubo['person_name']})", ubo['emirates_id'], '—', ubo['emirates_id_expiry'], status(ubo['emirates_id_expiry'])])
    for i, row in enumerate(rows, 2):
        for j, val in enumerate(row, 1):
            c = ws.cell(row=i, column=j, value=str(val) if val else '—')
            c.font = Font(size=10)
            c.alignment = Alignment(horizontal='left')
    for col, width in zip(['A','B','C','D','E'], [28, 22, 24, 16, 18]):
        ws.column_dimensions[col].width = width
    ws2 = wb.create_sheet("Company Info")
    info = [('AC Code', co['ac_code']), ('Company Name', co['client_name']),
            ('Status', co['ac_status']), ('Risk', co['risk_status']),
            ('Doc Status', co['doc_status']), ('Account Manager', co['account_manager']),
            ('KYC Status', co['kyc_status']), ('Region', co['region'])]
    for i, (k, v) in enumerate(info, 1):
        ws2.cell(row=i, column=1, value=k).font = Font(bold=True, color="D97706")
        ws2.cell(row=i, column=2, value=str(v) if v else '—')
    ws2.column_dimensions['A'].width = 22; ws2.column_dimensions['B'].width = 35
    out = io.BytesIO(); wb.save(out); out.seek(0)
    fname = f"{co['ac_code']}_documents.xlsx"
    return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=fname)





# ════════════════════════════════════════════════════════════
# ROLE PERMISSIONS API
# ════════════════════════════════════════════════════════════
@app.route('/api/settings/role-permissions', methods=['GET'])
@login_required
def api_get_role_permissions():
    try:
        conn = get_db()
        row = one(conn, "SELECT value FROM app_settings WHERE key='role_permissions'")
        conn.close()
        if row:
            import json as _json
            return jsonify({'permissions': _json.loads(row['value'])})
        return jsonify({'permissions': None})
    except Exception as e:
        return jsonify({'permissions': None, 'error': str(e)})

@app.route('/api/settings/role-permissions', methods=['POST'])
@login_required
def api_save_role_permissions():
    if session.get('user_role') != 'admin':
        return jsonify({'success': False, 'error': 'Admin only'}), 403
    import json as _json
    d = request.get_json()
    try:
        conn = get_db()
        val = _json.dumps(d.get('permissions', {}))
        x(conn, "INSERT INTO app_settings (key,value) VALUES ('role_permissions',?) ON CONFLICT(key) DO UPDATE SET value=?,updated_at=CURRENT_TIMESTAMP",
          (val, val))
        commit(conn); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ════════════════════════════════════════════════════════════
# CLIENTS PAGE
# ════════════════════════════════════════════════════════════
@app.route('/clients')
@login_required
def clients():
    conn = get_db()
    today = dubai_today()

    # SQL-level filtering — only fetch what's needed
    s = request.args.get('search', '').strip()
    resident_f = request.args.get('resident', '')
    mode_f = request.args.get('mode', '')
    page = max(1, int(request.args.get('page', 1)))
    per_page = 100

    q = 'SELECT * FROM clients WHERE 1=1'; p = []
    if s:
        q += ' AND (name LIKE ? OR phone LIKE ? OR account_number LIKE ? OR email LIKE ?)'
        p += [f'%{s}%'] * 4
    if resident_f == 'resident':
        q += ' AND is_resident=?'; p.append(True)
    elif resident_f == 'non-resident':
        q += ' AND (is_resident IS NULL OR is_resident=?)'; p.append(False)
    if mode_f:
        q += ' AND mode_of_ac=?'; p.append(mode_f)

    total_count = cnt(conn, f"SELECT COUNT(*) FROM clients WHERE {q.split('WHERE')[1]}", p or None)
    q += ' ORDER BY name LIMIT ? OFFSET ?'
    p += [per_page, (page - 1) * per_page]

    rows = all_(conn, q, p)
    cl = []
    for c in rows:
        dob = c.get('date_of_birth')
        birthday_today = False
        days_to_birthday = None
        if dob:
            try:
                d = datetime.strptime(str(dob)[:10], '%Y-%m-%d').date()
                this_year = d.replace(year=today.year)
                if this_year < today:
                    this_year = d.replace(year=today.year+1)
                days_to_birthday = (this_year - today).days
                birthday_today = (d.month == today.month and d.day == today.day)
            except: pass
        def expiry_status(d):
            dl = days_left(d)
            if dl is None: return 'ok'
            if dl < 0: return 'expired'
            if dl <= 90: return 'warning'
            return 'ok'
        cl.append({**c,
                   'date_of_birth': str(dob)[:10] if dob else None,
                   'passport_expiry': str(c.get('passport_expiry'))[:10] if c.get('passport_expiry') else None,
                   'emirates_id_expiry': str(c.get('emirates_id_expiry'))[:10] if c.get('emirates_id_expiry') else None,
                   'kyc_expiry_date': str(c.get('kyc_expiry_date'))[:10] if c.get('kyc_expiry_date') else None,
                   'screening_date': str(c.get('screening_date'))[:10] if c.get('screening_date') else None,
                   'birthday_today': birthday_today, 'days_to_birthday': days_to_birthday,
                   'passport_expiry_status': expiry_status(c.get('passport_expiry')),
                   'eid_expiry_status': expiry_status(c.get('emirates_id_expiry')),
                   'kyc_expiry_status': expiry_status(c.get('kyc_expiry_date'))})
    conn.close()
    birthdays_today = [c for c in cl if c['birthday_today']]
    dd = dropdowns()
    total_pages = max(1, (total_count + per_page - 1) // per_page)
    return render_template('clients.html', clients=cl, birthdays_today=birthdays_today, today=str(today), days_left=days_left,
        modes=dd.get('MODE OF AC',[]), ac_statuses=dd.get('AC STATUS',[]), id_types=dd.get('ID TYPE',[]),
        kyc_statuses=dd.get('KYC STATUS',[]), risk_statuses=dd.get('RISK STATUS',[]),
        screening_statuses=dd.get('SCREENING REGISTRATION STATUS',[]),
        search=s, resident_filter=resident_f, mode_filter=mode_f,
        page=page, total_pages=total_pages, total_count=total_count, per_page=per_page)

def _validate_kyc_expiry(kyc_expiry):
    """KYC expiry must not be more than 2 years from today."""
    if not kyc_expiry: return None
    try:
        d = datetime.strptime(str(kyc_expiry)[:10], '%Y-%m-%d').date()
        max_allowed = dubai_today().replace(year=dubai_today().year + 2)
        if d > max_allowed:
            return f"KYC expiry date cannot be more than 2 years from today (max: {max_allowed})"
    except Exception:
        return None
    return None

@app.route('/api/client/add', methods=['POST'])
@login_required
def api_add_client():
    d = request.get_json()
    err = _validate_kyc_expiry(d.get('kyc_expiry_date'))
    if err: return jsonify({'success': False, 'error': err}), 400
    try:
        conn = get_db()
        x(conn, '''INSERT INTO clients (name,phone,whatsapp_number,email,date_of_birth,
            profession,address,notes,nationality,passport_no,passport_expiry,
            emirates_id,emirates_id_expiry,address_proof,emirate,location,
            account_number,mode_of_ac,ac_status,id_type,pep_status,kyc_status,kyc_expiry_date,
            risk_status,screening_status,screening_date,is_resident,created_by)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
          (d.get('name'), d.get('phone'), d.get('whatsapp_number'), d.get('email'),
           d.get('date_of_birth') or None, d.get('profession'), d.get('address'),
           d.get('notes'), d.get('nationality'), d.get('passport_no'),
           d.get('passport_expiry') or None, d.get('emirates_id'),
           d.get('emirates_id_expiry') or None, d.get('address_proof'),
           d.get('emirate'), d.get('location'),
           d.get('account_number'), d.get('mode_of_ac'), d.get('ac_status'),
           d.get('id_type'), d.get('pep_status'), d.get('kyc_status'),
           d.get('kyc_expiry_date') or None, d.get('risk_status'),
           d.get('screening_status'), d.get('screening_date') or None,
           d.get('is_resident', False),
           session.get('user_id')))
        commit(conn)
        client_id = lastid(conn)
        conn.close()
        return jsonify({'success': True, 'client_id': client_id})
    except Exception as e:
        try: conn.close()
        except: pass
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/client/<int:id>/edit', methods=['POST'])
@login_required
def api_edit_client(id):
    d = request.get_json()
    err = _validate_kyc_expiry(d.get('kyc_expiry_date'))
    if err: return jsonify({'success': False, 'error': err}), 400
    try:
        conn = get_db()
        x(conn, '''UPDATE clients SET name=?,phone=?,whatsapp_number=?,email=?,date_of_birth=?,
            profession=?,address=?,notes=?,nationality=?,passport_no=?,passport_expiry=?,
            emirates_id=?,emirates_id_expiry=?,address_proof=?,emirate=?,location=?,
            account_number=?,mode_of_ac=?,ac_status=?,id_type=?,pep_status=?,kyc_status=?,
            kyc_expiry_date=?,risk_status=?,screening_status=?,screening_date=?,is_resident=?,
            updated_at=CURRENT_TIMESTAMP WHERE id=?''',
          (d.get('name'), d.get('phone'), d.get('whatsapp_number'), d.get('email'),
           d.get('date_of_birth') or None, d.get('profession'), d.get('address'),
           d.get('notes'), d.get('nationality'), d.get('passport_no'),
           d.get('passport_expiry') or None, d.get('emirates_id'),
           d.get('emirates_id_expiry') or None, d.get('address_proof'),
           d.get('emirate'), d.get('location'),
           d.get('account_number'), d.get('mode_of_ac'), d.get('ac_status'),
           d.get('id_type'), d.get('pep_status'), d.get('kyc_status'),
           d.get('kyc_expiry_date') or None, d.get('risk_status'),
           d.get('screening_status'), d.get('screening_date') or None,
           d.get('is_resident', False), id))
        commit(conn); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        try: conn.close()
        except: pass
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/client/<int:id>/delete', methods=['POST'])
@compliance_required
def api_delete_client(id):
    try:
        conn = get_db()
        x(conn, 'DELETE FROM clients WHERE id=?', (id,))
        commit(conn); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ════════════════════════════════════════════════════════════
# COMPANY GROUPS (admin panel)
# ════════════════════════════════════════════════════════════
@app.route('/api/groups', methods=['GET'])
@login_required
def api_get_groups():
    conn = get_db()
    groups = all_(conn, 'SELECT * FROM company_groups ORDER BY group_name')
    conn.close()
    return jsonify({'groups': groups})

@app.route('/api/group/add', methods=['POST'])
@compliance_required
def api_add_group():
    d = request.get_json()
    try:
        conn = get_db()
        x(conn, 'INSERT INTO company_groups (group_name, description) VALUES (?,?)',
          (d.get('group_name','').strip(), d.get('description','')))
        commit(conn); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/group/<int:id>/delete', methods=['POST'])
@compliance_required
def api_delete_group(id):
    try:
        conn = get_db()
        x(conn, 'DELETE FROM company_groups WHERE id=?', (id,))
        commit(conn); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ════════════════════════════════════════════════════════════
# APP SETTINGS (action password)
# ════════════════════════════════════════════════════════════
@app.route('/api/settings/action-password', methods=['POST'])
@compliance_required
def api_set_action_password():
    if session.get('user_role') != 'admin':
        return jsonify({'success': False, 'error': 'Admin only'}), 403
    d = request.get_json()
    pw = d.get('password', '').strip()
    if len(pw) < 4:
        return jsonify({'success': False, 'error': 'Password must be at least 4 characters'}), 400
    try:
        import hashlib
        hashed = hashlib.sha256(pw.encode()).hexdigest()
        conn = get_db()
        x(conn, "INSERT INTO app_settings (key,value) VALUES ('action_password',?) ON CONFLICT(key) DO UPDATE SET value=?,updated_at=CURRENT_TIMESTAMP",
          (hashed, hashed))
        commit(conn); conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/settings/verify-action-password', methods=['POST'])
@login_required
def api_verify_action_password():
    d = request.get_json()
    pw = d.get('password', '')
    try:
        import hashlib
        hashed = hashlib.sha256(pw.encode()).hexdigest()
        conn = get_db()
        s = one(conn, "SELECT value FROM app_settings WHERE key='action_password'")
        conn.close()
        if not s:
            return jsonify({'success': True, 'note': 'No action password set'})
        return jsonify({'success': s['value'] == hashed})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ════════════════════════════════════════════════════════════
# DATA BACKUP — full Excel export
# ════════════════════════════════════════════════════════════
@app.route('/admin/backup')
@login_required
def admin_backup():
    if session.get('user_role') != 'admin':
        return redirect(url_for('dashboard'))
    if not HAS_XL:
        return "openpyxl not installed", 500
    from openpyxl.styles import Font, PatternFill, Alignment
    conn = get_db()
    wb = openpyxl.Workbook()
    gold_fill = PatternFill(start_color="1C1917", end_color="1C1917", fill_type="solid")
    gold_font = Font(color="D97706", bold=True, size=10)

    def add_sheet(wb, name, rows, first=False):
        ws = wb.active if first else wb.create_sheet(name)
        ws.title = name
        if not rows: return ws
        headers = list(rows[0].keys())
        for i, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=i, value=h)
            c.font = gold_font; c.fill = gold_fill
        for ri, row in enumerate(rows, 2):
            for ci, h in enumerate(headers, 1):
                val = row.get(h)
                ws.cell(row=ri, column=ci, value=str(val) if val is not None else '')
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(40, max(12, len(str(col[0].value or ''))+4))
        ws.freeze_panes = 'A2'
        return ws

    add_sheet(wb, 'Companies',     all_(conn, 'SELECT * FROM companies ORDER BY ac_code'), first=True)
    add_sheet(wb, 'UBOs',          all_(conn, 'SELECT * FROM ubos ORDER BY company_id'))
    add_sheet(wb, 'Tasks',         all_(conn, 'SELECT * FROM tasks ORDER BY created_at DESC'))
    add_sheet(wb, 'Users',         all_(conn, 'SELECT id,email,name,role,is_active,created_at FROM users ORDER BY name'))
    add_sheet(wb, 'Clients',       all_(conn, 'SELECT * FROM clients ORDER BY name'))
    add_sheet(wb, 'InternalDocs',  all_(conn, 'SELECT * FROM internal_documents ORDER BY expiry_date'))
    add_sheet(wb, 'RegularTasks',  all_(conn, 'SELECT * FROM regular_task_templates'))
    add_sheet(wb, 'TaskLogs',      all_(conn, 'SELECT * FROM regular_task_logs ORDER BY logged_at DESC'))
    add_sheet(wb, 'Groups',        all_(conn, 'SELECT * FROM company_groups'))
    add_sheet(wb, 'Dropdowns',     all_(conn, 'SELECT * FROM dropdowns WHERE is_active=1 ORDER BY field_name,value'))
    add_sheet(wb, 'AdditionalTasks', all_(conn, 'SELECT * FROM additional_tasks ORDER BY from_datetime DESC'))
    conn.close()

    out = io.BytesIO()
    wb.save(out); out.seek(0)
    fname = f"ZewerCRM_Backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=fname)

# ════════════════════════════════════════════════════════════
# RESTORE FROM BACKUP
# ════════════════════════════════════════════════════════════
@app.route('/admin/restore', methods=['POST'])
@login_required
def admin_restore():
    if session.get('user_role') != 'admin':
        return jsonify({'success': False, 'error': 'Admin only'}), 403
    if not HAS_XL:
        return jsonify({'success': False, 'error': 'openpyxl not installed'}), 500
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No file uploaded'}), 400
    try:
        f = request.files['file']
        wb = openpyxl.load_workbook(f)
        conn = get_db()
        results = {}

        def restore_sheet(sheet_name, table, pk='id', skip_cols=None):
            if sheet_name not in wb.sheetnames:
                results[sheet_name] = 'Sheet not found — skipped'
                return
            ws = wb[sheet_name]
            headers = [cell.value for cell in ws[1]]
            if not headers or not headers[0]:
                results[sheet_name] = 'Empty — skipped'
                return
            skip = set(skip_cols or [])
            use_headers = [h for h in headers if h and h not in skip]
            inserted = 0; skipped = 0
            for row in ws.iter_rows(min_row=2, values_only=True):
                rd = {headers[i]: row[i] for i in range(len(headers)) if headers[i]}
                pk_val = rd.get(pk)
                if pk_val is None: continue
                # Check if record already exists
                try:
                    existing = one(conn, f'SELECT {pk} FROM {table} WHERE {pk}=?', (pk_val,))
                    if existing: skipped += 1; continue
                except: skipped += 1; continue
                cols = [h for h in use_headers if h != pk]
                vals = [rd.get(h) or None for h in cols]
                placeholders = ','.join(['?' for _ in cols])
                try:
                    x(conn, f"INSERT INTO {table} ({','.join(cols)}) VALUES ({placeholders})", vals)
                    inserted += 1
                except: skipped += 1
            commit(conn)
            results[sheet_name] = f'{inserted} restored, {skipped} skipped'

        restore_sheet('Companies',    'companies',              pk='id')
        restore_sheet('UBOs',         'ubos',                   pk='id')
        restore_sheet('Clients',      'clients',                pk='id')
        restore_sheet('Tasks',        'tasks',                  pk='id')
        restore_sheet('InternalDocs', 'internal_documents',     pk='id')
        restore_sheet('RegularTasks', 'regular_task_templates', pk='id')
        restore_sheet('TaskLogs',     'regular_task_logs',      pk='id')
        restore_sheet('Groups',       'company_groups',         pk='id')
        restore_sheet('Dropdowns',    'dropdowns',              pk='id')
        restore_sheet('AdditionalTasks', 'additional_tasks',   pk='id')
        # Skip Users sheet — don't overwrite passwords/accounts from backup

        conn.close()
        return jsonify({'success': True, 'results': results})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ════════════════════════════════════════════════════════════
# GROUPS API for companies form
# ════════════════════════════════════════════════════════════
@app.route('/api/groups/list')
@login_required
def api_groups_list():
    conn = get_db()
    try:
        groups = all_(conn, 'SELECT group_name FROM company_groups ORDER BY group_name')
    except: groups = []
    conn.close()
    return jsonify([g['group_name'] for g in groups])







# ════════════════════════════════════════════════════════════
# CLIENT DOCUMENTS (Cloudinary)
# ════════════════════════════════════════════════════════════
@app.route('/api/client/<int:cid>/documents')
@login_required
def api_client_documents(cid):
    conn = get_db()
    docs = all_(conn, 'SELECT * FROM client_documents WHERE client_id=? ORDER BY created_at DESC', (cid,))
    conn.close()
    return jsonify(docs)

@app.route('/api/client/<int:cid>/upload', methods=['POST'])
@login_required
def api_client_upload_document(cid):
    if 'file' not in request.files: return jsonify({'success':False,'error':'No file'}),400
    file = request.files['file']
    if not file.filename: return jsonify({'success':False,'error':'No filename'}),400
    if not HAS_CLD or not os.getenv('CLOUDINARY_CLOUD_NAME'):
        return jsonify({'success':False,'error':'Cloudinary not configured'}),500
    try:
        r = cloudinary.uploader.upload(file, folder=f'zewer_crm/client_{cid}', resource_type='auto',
                                        use_filename=True, unique_filename=True)
        conn = get_db()
        x(conn, 'INSERT INTO client_documents (client_id,doc_type,file_name,file_url,public_id,uploaded_by) VALUES (?,?,?,?,?,?)',
          (cid, request.form.get('doc_type','General'), file.filename, r['secure_url'], r['public_id'], session.get('user_id')))
        commit(conn); conn.close()
        return jsonify({'success':True,'url':r['secure_url'],'name':file.filename})
    except Exception as e:
        return jsonify({'success':False,'error':str(e)}),500

@app.route('/api/client-document/<int:did>/delete', methods=['POST'])
@login_required
def api_delete_client_document(did):
    try:
        conn = get_db()
        doc = one(conn, 'SELECT public_id FROM client_documents WHERE id=?', (did,))
        if doc and doc.get('public_id') and HAS_CLD and os.getenv('CLOUDINARY_CLOUD_NAME'):
            try: cloudinary.uploader.destroy(doc['public_id'], resource_type='raw')
            except: pass
        x(conn, 'DELETE FROM client_documents WHERE id=?', (did,))
        commit(conn); conn.close()
        return jsonify({'success':True})
    except Exception as e:
        return jsonify({'success':False,'error':str(e)}),500

@app.route('/client/<int:id>')
@login_required
def client_detail(id):
    conn = get_db()
    c = one(conn, 'SELECT * FROM clients WHERE id=?', (id,))
    if not c:
        conn.close()
        return redirect(url_for('clients'))
    docs = all_(conn, 'SELECT * FROM client_documents WHERE client_id=? ORDER BY created_at DESC', (id,))
    conn.close()
    dob = c.get('date_of_birth')
    c['date_of_birth'] = str(dob)[:10] if dob else None
    pe = c.get('passport_expiry')
    c['passport_expiry'] = str(pe)[:10] if pe else None
    ee = c.get('emirates_id_expiry')
    c['emirates_id_expiry'] = str(ee)[:10] if ee else None
    ke = c.get('kyc_expiry_date')
    c['kyc_expiry_date'] = str(ke)[:10] if ke else None
    sd = c.get('screening_date')
    c['screening_date'] = str(sd)[:10] if sd else None
    c['kyc_expiry_status'] = exp_status(days_left(ke)) if ke else 'unknown'
    return render_template('client_detail.html', client=c, documents=docs)









