from flask import Flask, render_template, request, jsonify, session, redirect, url_for, send_file
from werkzeug.security import check_password_hash, generate_password_hash
import sqlite3
import os
from datetime import datetime, timedelta
from functools import wraps
import csv
import io

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'zewer-aml-crm-secret-2026')

DB = 'aml_crm.db'

def init_db():
    conn = sqlite3.connect(DB)
    conn.execute('PRAGMA foreign_keys = ON')
    cursor = conn.cursor()

    cursor.executescript('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            name TEXT NOT NULL,
            role TEXT DEFAULT 'staff',
            is_active INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ac_code TEXT UNIQUE NOT NULL,
            client_name TEXT NOT NULL,
            ac_opening_date DATE,
            ac_status TEXT DEFAULT 'Active',
            active_till_year TEXT,
            nature TEXT,
            type_of_client TEXT,
            name_of_freezone TEXT,
            mode_of_ac TEXT,
            country_of_incorporation TEXT,
            region TEXT,
            address TEXT,
            telephone TEXT,
            mobile TEXT,
            whatsapp_number TEXT,
            email_id TEXT,
            address_proof_type TEXT,
            address_proof_expiry DATE,
            kyc_status TEXT,
            trade_license_no TEXT,
            issuing_authority TEXT,
            legal_type TEXT,
            incorporation_date DATE,
            trade_license_expiry DATE,
            tax_no_trn TEXT,
            vat_cert TEXT,
            vat_declaration TEXT,
            deal_after_vat TEXT,
            num_beneficial_owners INTEGER DEFAULT 0,
            moa TEXT,
            pep TEXT,
            undertaking TEXT,
            source_of_fund TEXT,
            software_updation TEXT,
            doc_status TEXT DEFAULT 'Incompleted',
            screening_date DATE,
            registration_screening_tool TEXT,
            risk_status TEXT DEFAULT 'Unspecified',
            verified_by TEXT,
            verified_date DATE,
            followup_details TEXT,
            crowe_feedback TEXT,
            zewer_comments TEXT,
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(created_by) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS ubos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            position TEXT,
            share_percentage REAL,
            person_name TEXT NOT NULL,
            nationality TEXT,
            residential_status TEXT,
            passport_no TEXT,
            passport_expiry DATE,
            emirates_id TEXT,
            emirates_id_expiry DATE,
            doc_status TEXT DEFAULT 'Incompleted',
            verified_by TEXT,
            verified_date DATE,
            followup_details TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(company_id) REFERENCES companies(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS dropdowns (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            field_name TEXT NOT NULL,
            value TEXT NOT NULL,
            is_active INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(field_name, value)
        );
    ''')

    # Admin user
    existing = cursor.execute('SELECT id FROM users WHERE email = ?', ('admin@zewer.ae',)).fetchone()
    if not existing:
        cursor.execute('INSERT INTO users (email, password_hash, name, role) VALUES (?, ?, ?, ?)',
            ('admin@zewer.ae', generate_password_hash('Admin@123'), 'Administrator', 'admin'))
        cursor.execute('INSERT INTO users (email, password_hash, name, role) VALUES (?, ?, ?, ?)',
            ('compliance@zewer.ae', generate_password_hash('Compliance@123'), 'Compliance Officer', 'compliance'))

    # Dropdowns
    dropdowns_data = {
        'AC STATUS': ['Active', 'Inactive'],
        'NATURE': ['Individual', 'Legal entity'],
        'TYPE OF CLIENT': ['MainLand', 'Free Zone', 'Abroad', 'International Corporate'],
        'MODE OF AC': ['Supplier', 'Customer', 'Bullion', 'Refinery', 'Logistics Co', 'Exchange', 'Bank',
                       'Insurance', 'Investor', 'Technical Services'],
        'RISK STATUS': ['High', 'Medium', 'Low', 'Unspecified'],
        'DOC STATUS': ['Completed', 'Incompleted'],
        'KYC STATUS': ['New Kyc Updated', 'Kyc 2025 Updated', 'Kyc 2024 Updated', 'Kyc 2023 Updated',
                       'Kyc 2022 Updated', 'Kyc 2021 Updated', 'Kyc 2020 Updated', 'Kyc 2019 Updated',
                       'Kyc 2018 Updated', 'Kyc 2017 Updated', 'Kyc 2016 Updated', 'Not Updated'],
        'REGION': ['Dubai', 'Abu Dhabi', 'Sharjah', 'Ajman', 'Ras Al Khaimah', 'Fujairah',
                   'Umm Al Quwain', 'West Bengal', 'United Kingdom', 'Canada', 'Pakistan', 'Malaysia',
                   'Singapore', 'Bahrain', 'Italy', 'Al Ain', 'REPUBLIC OF CONGO', 'HONG KONG',
                   'Saudi Arabia', 'India'],
        'FREEZONE': ['Jabel Ali FZ', 'DMCC', 'Sharjah Saif Zone', 'Ajman FZ', 'Fujairah',
                     'Dubai Production City', 'N/A', 'Ras Al Khaimah Fz', 'Dubai Free Zone',
                     'Dubai Gold & Diamond Park'],
        'LEGAL TYPE': ['Limited Liability Company(LLC)', 'Limited Liability Company (WLL)',
                       'Civil Company Professional', 'Foreign Company', 'Public Company', 'Privet Company',
                       'DMCC', 'Free Zone Limited Liability Company (FZ-LLC)', 'Sole Establishment',
                       'Partnership Company', 'Services Agency', 'Limited (LTD)', 'Individual Institution',
                       'Establishment', 'Limited Liability Company(LLC-SO )(Single Ownership)',
                       'FZCO', 'FZE', 'FZC'],
        'ISSUING AUTHORITY': ['Dubai Economy & Tourism', 'Department of Economic Development Ajman',
                              'Abu-Dhabi Department of Economic Development', 'DMCC', 'Saif Zone',
                              'Ajman FZ', 'Jebel Ali FZ', 'Dubai Development Authority',
                              'Government Of Sharjah Economic Development Department',
                              'Government of Ras Al-Khaimah Department of Economic Development',
                              'Department of Economic Development Dubai',
                              'Dubai Integrated Economic Zones Authority',
                              'Fujairah Municipality', 'Trade Development Authority Of Pakistan',
                              'UK HMRC', 'The Registrar Of Companies For England And Wales',
                              'Canada Revenue Agency', 'Kolkata Municipal Corporation', 'Abroad'],
        'ADDRESS PROOF TYPE': ['Ejari', 'Tenancy', 'Electricity Bill', 'Gst Registration Certificate',
                               'Certification Of Incorporation', 'Certificate Of Registration For Value Added Tax',
                               'Vat Certificate', 'Telephone Bill', 'Title Deed',
                               'Certificate Of Enlistment', 'Association Of Article Details',
                               'Warehouse Lease Agreement', 'Not Required'],
        'VAT CERT': ['Yes', 'No', 'Not Required'],
        'VAT DECLARATION': ['Yes', 'No', 'Not Required'],
        'MOA': ['Yes', 'No'],
        'PEP': ['Yes', 'No'],
        'UNDERTAKING': ['Yes', 'No'],
        'SOURCE OF FUND': ['Yes', 'No'],
        'POSITION': ['UBO', 'Authorized Person', 'Director', 'Manager', 'Partner'],
        'RESIDENTIAL STATUS': ['Resident', 'Non Resident'],
        'COUNTRY': ['United Arab Emirates', 'Saudi Arabia', 'Kuwait', 'Qatar', 'Bahrain', 'Oman',
                    'India', 'Pakistan', 'Bangladesh', 'Sri Lanka', 'Philippines', 'Malaysia',
                    'Singapore', 'China', 'Hong Kong', 'Jordan', 'Lebanon', 'Syria', 'Iraq', 'Yemen',
                    'Egypt', 'Libya', 'Nigeria', 'Ethiopia', 'Republic Of Congo', 'Turkey', 'Iran',
                    'Afghanistan', 'Algeria', 'Canada', 'United Kingdom', 'United States of America',
                    'France', 'Ireland', 'Italy', 'Germany', 'Armenia', 'Belize'],
    }

    existing_dd = cursor.execute('SELECT COUNT(*) FROM dropdowns').fetchone()[0]
    if existing_dd == 0:
        for field, values in dropdowns_data.items():
            for value in values:
                try:
                    cursor.execute('INSERT INTO dropdowns (field_name, value, is_active) VALUES (?, ?, 1)',
                        (field, value))
                except:
                    pass

    conn.commit()
    conn.close()

init_db()

def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA foreign_keys = ON')
    return conn

def days_left(date_str):
    if not date_str:
        return None
    try:
        exp = datetime.strptime(str(date_str)[:10], '%Y-%m-%d').date()
        return (exp - datetime.now().date()).days
    except:
        return None

def expiry_status(days):
    if days is None:
        return 'unknown'
    if days < 0:
        return 'expired'
    if days <= 30:
        return 'critical'
    if days <= 90:
        return 'warning'
    return 'valid'

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        conn = get_db()
        user = conn.execute('SELECT role FROM users WHERE id = ?', (session['user_id'],)).fetchone()
        conn.close()
        if not user or user['role'] != 'admin':
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated

def get_dropdowns():
    conn = get_db()
    fields = conn.execute('SELECT DISTINCT field_name FROM dropdowns WHERE is_active = 1 ORDER BY field_name').fetchall()
    dd = {}
    for f in fields:
        vals = conn.execute('SELECT value FROM dropdowns WHERE field_name = ? AND is_active = 1 ORDER BY value',
                            (f['field_name'],)).fetchall()
        dd[f['field_name']] = [v['value'] for v in vals]
    conn.close()
    return dd

# AUTH
@app.route('/')
def index():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        data = request.get_json()
        email = data.get('email')
        password = data.get('password')
        conn = get_db()
        user = conn.execute('SELECT * FROM users WHERE email = ?', (email,)).fetchone()
        conn.close()
        if user and check_password_hash(user['password_hash'], password) and user['is_active']:
            session['user_id'] = user['id']
            session['user_email'] = user['email']
            session['user_name'] = user['name']
            session['user_role'] = user['role']
            return jsonify({'success': True}), 200
        return jsonify({'success': False, 'error': 'Invalid credentials'}), 401
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# DASHBOARD
@app.route('/dashboard')
@login_required
def dashboard():
    conn = get_db()
    today = datetime.now().date()

    total = conn.execute('SELECT COUNT(*) as c FROM companies').fetchone()['c']
    active = conn.execute('SELECT COUNT(*) as c FROM companies WHERE ac_status = "Active"').fetchone()['c']

    expiring_30_tl = conn.execute('SELECT COUNT(*) as c FROM companies WHERE trade_license_expiry BETWEEN ? AND ?',
        (today, today + timedelta(days=30))).fetchone()['c']
    expiring_90_tl = conn.execute('SELECT COUNT(*) as c FROM companies WHERE trade_license_expiry BETWEEN ? AND ?',
        (today, today + timedelta(days=90))).fetchone()['c']
    expired_tl = conn.execute('SELECT COUNT(*) as c FROM companies WHERE trade_license_expiry < ?',
        (today,)).fetchone()['c']

    expiring_30_ap = conn.execute('SELECT COUNT(*) as c FROM companies WHERE address_proof_expiry BETWEEN ? AND ?',
        (today, today + timedelta(days=30))).fetchone()['c']
    expired_ap = conn.execute('SELECT COUNT(*) as c FROM companies WHERE address_proof_expiry < ?',
        (today,)).fetchone()['c']

    # UBO passport expiry
    expiring_30_pass = conn.execute('SELECT COUNT(*) as c FROM ubos WHERE passport_expiry BETWEEN ? AND ?',
        (today, today + timedelta(days=30))).fetchone()['c']
    expired_pass = conn.execute('SELECT COUNT(*) as c FROM ubos WHERE passport_expiry < ?',
        (today,)).fetchone()['c']

    expiring_30_eid = conn.execute('SELECT COUNT(*) as c FROM ubos WHERE emirates_id_expiry BETWEEN ? AND ?',
        (today, today + timedelta(days=30))).fetchone()['c']
    expired_eid = conn.execute('SELECT COUNT(*) as c FROM ubos WHERE emirates_id_expiry < ?',
        (today,)).fetchone()['c']

    risk_data = conn.execute('SELECT risk_status, COUNT(*) as c FROM companies GROUP BY risk_status').fetchall()
    risk_breakdown = {row['risk_status']: row['c'] for row in risk_data}

    doc_data = conn.execute('SELECT doc_status, COUNT(*) as c FROM companies GROUP BY doc_status').fetchall()
    doc_breakdown = {row['doc_status']: row['c'] for row in doc_data}

    kyc_data = conn.execute('SELECT kyc_status, COUNT(*) as c FROM companies GROUP BY kyc_status ORDER BY c DESC LIMIT 8').fetchall()

    # Recent alerts - companies expiring in next 90 days
    urgent_companies = conn.execute('''
        SELECT id, ac_code, client_name, trade_license_expiry, address_proof_expiry, risk_status, mobile, whatsapp_number
        FROM companies
        WHERE (trade_license_expiry BETWEEN ? AND ?) OR (address_proof_expiry BETWEEN ? AND ?)
        ORDER BY trade_license_expiry ASC LIMIT 10
    ''', (today, today + timedelta(days=90), today, today + timedelta(days=90))).fetchall()

    conn.close()

    return render_template('dashboard.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        total_companies=total,
        active_companies=active,
        expiring_30_tl=expiring_30_tl,
        expiring_90_tl=expiring_90_tl,
        expired_tl=expired_tl,
        expiring_30_ap=expiring_30_ap,
        expired_ap=expired_ap,
        expiring_30_pass=expiring_30_pass,
        expired_pass=expired_pass,
        expiring_30_eid=expiring_30_eid,
        expired_eid=expired_eid,
        risk_breakdown=risk_breakdown,
        doc_breakdown=doc_breakdown,
        kyc_data=kyc_data,
        urgent_companies=urgent_companies,
        today=str(today)
    )

# COMPANIES
@app.route('/companies')
@login_required
def companies():
    conn = get_db()
    search = request.args.get('search', '')
    risk_filter = request.args.get('risk', '')
    status_filter = request.args.get('status', '')
    region_filter = request.args.get('region', '')

    query = 'SELECT * FROM companies WHERE 1=1'
    params = []

    if search:
        query += ' AND (client_name LIKE ? OR ac_code LIKE ? OR mobile LIKE ? OR trade_license_no LIKE ?)'
        s = f'%{search}%'
        params += [s, s, s, s]
    if risk_filter:
        query += ' AND risk_status = ?'
        params.append(risk_filter)
    if status_filter:
        query += ' AND ac_status = ?'
        params.append(status_filter)
    if region_filter:
        query += ' AND region = ?'
        params.append(region_filter)

    query += ' ORDER BY created_at DESC'
    rows = conn.execute(query, params).fetchall()
    conn.close()

    today = datetime.now().date()
    companies_list = []
    for c in rows:
        tl_days = days_left(c['trade_license_expiry'])
        ap_days = days_left(c['address_proof_expiry'])
        companies_list.append({
            'id': c['id'],
            'ac_code': c['ac_code'],
            'client_name': c['client_name'],
            'region': c['region'],
            'type_of_client': c['type_of_client'],
            'mode_of_ac': c['mode_of_ac'],
            'mobile': c['mobile'],
            'whatsapp_number': c['whatsapp_number'],
            'risk_status': c['risk_status'],
            'ac_status': c['ac_status'],
            'doc_status': c['doc_status'],
            'kyc_status': c['kyc_status'],
            'trade_license_expiry': c['trade_license_expiry'],
            'tl_days': tl_days,
            'tl_status': expiry_status(tl_days),
            'address_proof_expiry': c['address_proof_expiry'],
            'ap_days': ap_days,
            'ap_status': expiry_status(ap_days),
        })

    dd = get_dropdowns()
    return render_template('companies.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        companies=companies_list,
        search=search,
        risk_filter=risk_filter,
        status_filter=status_filter,
        region_filter=region_filter,
        regions=dd.get('REGION', []),
        today=str(today)
    )

@app.route('/company/new')
@login_required
def company_new():
    dd = get_dropdowns()
    return render_template('company_form.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        dropdown_data=dd,
        company=None,
        edit=False
    )

@app.route('/company/<int:id>')
@login_required
def company_detail(id):
    conn = get_db()
    company = conn.execute('SELECT * FROM companies WHERE id = ?', (id,)).fetchone()
    ubos = conn.execute('SELECT * FROM ubos WHERE company_id = ? ORDER BY share_percentage DESC', (id,)).fetchall()
    conn.close()

    if not company:
        return redirect(url_for('companies'))

    today = datetime.now().date()
    tl_days = days_left(company['trade_license_expiry'])
    ap_days = days_left(company['address_proof_expiry'])

    ubo_list = []
    for u in ubos:
        p_days = days_left(u['passport_expiry'])
        e_days = days_left(u['emirates_id_expiry'])
        ubo_list.append({
            'id': u['id'],
            'position': u['position'],
            'share_percentage': u['share_percentage'],
            'person_name': u['person_name'],
            'nationality': u['nationality'],
            'residential_status': u['residential_status'],
            'passport_no': u['passport_no'],
            'passport_expiry': u['passport_expiry'],
            'p_days': p_days,
            'p_status': expiry_status(p_days),
            'emirates_id': u['emirates_id'],
            'emirates_id_expiry': u['emirates_id_expiry'],
            'e_days': e_days,
            'e_status': expiry_status(e_days),
            'doc_status': u['doc_status'],
            'verified_by': u['verified_by'],
        })

    return render_template('company_detail.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        company=company,
        ubos=ubo_list,
        tl_days=tl_days,
        tl_status=expiry_status(tl_days),
        ap_days=ap_days,
        ap_status=expiry_status(ap_days),
        today=str(today)
    )

@app.route('/company/<int:id>/edit')
@login_required
def company_edit(id):
    conn = get_db()
    company = conn.execute('SELECT * FROM companies WHERE id = ?', (id,)).fetchone()
    ubos = conn.execute('SELECT * FROM ubos WHERE company_id = ? ORDER BY share_percentage DESC', (id,)).fetchall()
    conn.close()

    if not company:
        return redirect(url_for('companies'))

    dd = get_dropdowns()
    return render_template('company_form.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        dropdown_data=dd,
        company=company,
        ubos=ubos,
        edit=True
    )

# API - ADD COMPANY
@app.route('/api/company/add', methods=['POST'])
@login_required
def api_add_company():
    data = request.get_json()
    try:
        conn = get_db()
        conn.execute('''INSERT INTO companies (
            ac_code, client_name, ac_opening_date, ac_status, active_till_year,
            nature, type_of_client, name_of_freezone, mode_of_ac,
            country_of_incorporation, region, address, telephone, mobile,
            whatsapp_number, email_id,
            address_proof_type, address_proof_expiry,
            kyc_status, trade_license_no, issuing_authority, legal_type,
            incorporation_date, trade_license_expiry,
            tax_no_trn, vat_cert, vat_declaration, deal_after_vat,
            num_beneficial_owners, moa, pep, undertaking, source_of_fund,
            software_updation, doc_status, screening_date,
            registration_screening_tool, risk_status,
            verified_by, verified_date, followup_details,
            crowe_feedback, zewer_comments, created_by
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (
            data.get('ac_code'), data.get('client_name'), data.get('ac_opening_date') or None,
            data.get('ac_status', 'Active'), data.get('active_till_year'),
            data.get('nature'), data.get('type_of_client'), data.get('name_of_freezone'),
            data.get('mode_of_ac'), data.get('country_of_incorporation'),
            data.get('region'), data.get('address'), data.get('telephone'), data.get('mobile'),
            data.get('whatsapp_number'), data.get('email_id'),
            data.get('address_proof_type'), data.get('address_proof_expiry') or None,
            data.get('kyc_status'), data.get('trade_license_no'), data.get('issuing_authority'),
            data.get('legal_type'), data.get('incorporation_date') or None,
            data.get('trade_license_expiry') or None,
            data.get('tax_no_trn'), data.get('vat_cert'), data.get('vat_declaration'),
            data.get('deal_after_vat'), data.get('num_beneficial_owners', 0),
            data.get('moa'), data.get('pep'), data.get('undertaking'), data.get('source_of_fund'),
            data.get('software_updation'), data.get('doc_status', 'Incompleted'),
            data.get('screening_date') or None, data.get('registration_screening_tool'),
            data.get('risk_status', 'Unspecified'), data.get('verified_by'),
            data.get('verified_date') or None, data.get('followup_details'),
            data.get('crowe_feedback'), data.get('zewer_comments'),
            session.get('user_id')
        ))
        company_id = conn.execute('SELECT last_insert_rowid()').fetchone()[0]

        # Insert UBOs
        ubos = data.get('ubos', [])
        for u in ubos:
            if u.get('person_name'):
                conn.execute('''INSERT INTO ubos (
                    company_id, position, share_percentage, person_name,
                    nationality, residential_status, passport_no, passport_expiry,
                    emirates_id, emirates_id_expiry, doc_status, verified_by, verified_date, followup_details
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                (company_id, u.get('position'), u.get('share_percentage'),
                 u.get('person_name'), u.get('nationality'), u.get('residential_status'),
                 u.get('passport_no'), u.get('passport_expiry') or None,
                 u.get('emirates_id'), u.get('emirates_id_expiry') or None,
                 u.get('doc_status', 'Incompleted'), u.get('verified_by'),
                 u.get('verified_date') or None, u.get('followup_details')))

        conn.commit()
        conn.close()
        return jsonify({'success': True, 'id': company_id}), 200
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# API - EDIT COMPANY
@app.route('/api/company/<int:id>/edit', methods=['POST'])
@login_required
def api_edit_company(id):
    data = request.get_json()
    try:
        conn = get_db()
        conn.execute('''UPDATE companies SET
            client_name=?, ac_opening_date=?, ac_status=?, active_till_year=?,
            nature=?, type_of_client=?, name_of_freezone=?, mode_of_ac=?,
            country_of_incorporation=?, region=?, address=?, telephone=?, mobile=?,
            whatsapp_number=?, email_id=?,
            address_proof_type=?, address_proof_expiry=?,
            kyc_status=?, trade_license_no=?, issuing_authority=?, legal_type=?,
            incorporation_date=?, trade_license_expiry=?,
            tax_no_trn=?, vat_cert=?, vat_declaration=?, deal_after_vat=?,
            num_beneficial_owners=?, moa=?, pep=?, undertaking=?, source_of_fund=?,
            software_updation=?, doc_status=?, screening_date=?,
            registration_screening_tool=?, risk_status=?,
            verified_by=?, verified_date=?, followup_details=?,
            crowe_feedback=?, zewer_comments=?,
            updated_at=CURRENT_TIMESTAMP
            WHERE id=?''',
        (
            data.get('client_name'), data.get('ac_opening_date') or None,
            data.get('ac_status', 'Active'), data.get('active_till_year'),
            data.get('nature'), data.get('type_of_client'), data.get('name_of_freezone'),
            data.get('mode_of_ac'), data.get('country_of_incorporation'),
            data.get('region'), data.get('address'), data.get('telephone'), data.get('mobile'),
            data.get('whatsapp_number'), data.get('email_id'),
            data.get('address_proof_type'), data.get('address_proof_expiry') or None,
            data.get('kyc_status'), data.get('trade_license_no'), data.get('issuing_authority'),
            data.get('legal_type'), data.get('incorporation_date') or None,
            data.get('trade_license_expiry') or None,
            data.get('tax_no_trn'), data.get('vat_cert'), data.get('vat_declaration'),
            data.get('deal_after_vat'), data.get('num_beneficial_owners', 0),
            data.get('moa'), data.get('pep'), data.get('undertaking'), data.get('source_of_fund'),
            data.get('software_updation'), data.get('doc_status', 'Incompleted'),
            data.get('screening_date') or None, data.get('registration_screening_tool'),
            data.get('risk_status', 'Unspecified'), data.get('verified_by'),
            data.get('verified_date') or None, data.get('followup_details'),
            data.get('crowe_feedback'), data.get('zewer_comments'),
            id
        ))

        # Rebuild UBOs
        conn.execute('DELETE FROM ubos WHERE company_id = ?', (id,))
        ubos = data.get('ubos', [])
        for u in ubos:
            if u.get('person_name'):
                conn.execute('''INSERT INTO ubos (
                    company_id, position, share_percentage, person_name,
                    nationality, residential_status, passport_no, passport_expiry,
                    emirates_id, emirates_id_expiry, doc_status, verified_by, verified_date, followup_details
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                (id, u.get('position'), u.get('share_percentage'),
                 u.get('person_name'), u.get('nationality'), u.get('residential_status'),
                 u.get('passport_no'), u.get('passport_expiry') or None,
                 u.get('emirates_id'), u.get('emirates_id_expiry') or None,
                 u.get('doc_status', 'Incompleted'), u.get('verified_by'),
                 u.get('verified_date') or None, u.get('followup_details')))

        conn.commit()
        conn.close()
        return jsonify({'success': True}), 200
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/company/<int:id>/delete', methods=['POST'])
@login_required
def api_delete_company(id):
    try:
        conn = get_db()
        conn.execute('DELETE FROM companies WHERE id = ?', (id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True}), 200
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# SETTINGS
@app.route('/settings')
@admin_required
def settings():
    conn = get_db()
    users = conn.execute('SELECT id, email, name, role, is_active FROM users').fetchall()
    dropdowns = conn.execute('SELECT * FROM dropdowns WHERE is_active = 1 ORDER BY field_name').fetchall()
    dropdown_groups = {}
    for dd in dropdowns:
        if dd['field_name'] not in dropdown_groups:
            dropdown_groups[dd['field_name']] = []
        dropdown_groups[dd['field_name']].append(dd)
    conn.close()
    return render_template('settings.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        users=users,
        dropdown_groups=dropdown_groups
    )

@app.route('/api/user/add', methods=['POST'])
@admin_required
def api_add_user():
    data = request.get_json()
    try:
        conn = get_db()
        conn.execute('INSERT INTO users (email, password_hash, name, role, is_active) VALUES (?, ?, ?, ?, 1)',
            (data.get('email'), generate_password_hash(data.get('password')), data.get('name'), data.get('role')))
        conn.commit()
        conn.close()
        return jsonify({'success': True}), 200
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/user/<int:id>/delete', methods=['POST'])
@admin_required
def api_delete_user(id):
    try:
        conn = get_db()
        conn.execute('DELETE FROM users WHERE id = ?', (id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True}), 200
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/dropdown/add', methods=['POST'])
@admin_required
def api_add_dropdown():
    data = request.get_json()
    try:
        conn = get_db()
        conn.execute('INSERT INTO dropdowns (field_name, value, is_active) VALUES (?, ?, 1)',
            (data.get('field_name'), data.get('value')))
        conn.commit()
        conn.close()
        return jsonify({'success': True}), 200
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/dropdown/<int:id>/delete', methods=['POST'])
@admin_required
def api_delete_dropdown(id):
    try:
        conn = get_db()
        conn.execute('DELETE FROM dropdowns WHERE id = ?', (id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True}), 200
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ALERTS
@app.route('/alerts')
@login_required
def alerts():
    conn = get_db()
    today = datetime.now().date()

    tl_expiring = conn.execute('''
        SELECT id, ac_code, client_name, mobile, whatsapp_number, trade_license_expiry, risk_status, region
        FROM companies WHERE trade_license_expiry IS NOT NULL AND trade_license_expiry >= ?
        ORDER BY trade_license_expiry ASC
    ''', (today - timedelta(days=30),)).fetchall()

    ap_expiring = conn.execute('''
        SELECT id, ac_code, client_name, mobile, whatsapp_number, address_proof_expiry, address_proof_type, risk_status
        FROM companies WHERE address_proof_expiry IS NOT NULL AND address_proof_expiry >= ?
        ORDER BY address_proof_expiry ASC
    ''', (today - timedelta(days=30),)).fetchall()

    ubo_expiring = conn.execute('''
        SELECT u.id, u.person_name, u.passport_no, u.passport_expiry, u.emirates_id, u.emirates_id_expiry,
               c.id as company_id, c.client_name, c.ac_code
        FROM ubos u JOIN companies c ON u.company_id = c.id
        WHERE (u.passport_expiry >= ? OR u.emirates_id_expiry >= ?)
        ORDER BY u.passport_expiry ASC
    ''', (today - timedelta(days=30), today - timedelta(days=30))).fetchall()

    def enrich(rows, date_field):
        result = []
        for r in rows:
            d = days_left(r[date_field])
            result.append({'row': r, 'days': d, 'status': expiry_status(d)})
        return result

    conn.close()
    return render_template('alerts.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        tl_expiring=tl_expiring,
        ap_expiring=ap_expiring,
        ubo_expiring=ubo_expiring,
        today=str(today),
        days_left=days_left,
        expiry_status=expiry_status
    )

# REPORTS
@app.route('/reports')
@login_required
def reports():
    conn = get_db()
    risk_data = conn.execute('SELECT risk_status, COUNT(*) as c FROM companies GROUP BY risk_status').fetchall()
    doc_data = conn.execute('SELECT doc_status, COUNT(*) as c FROM companies GROUP BY doc_status').fetchall()
    type_data = conn.execute('SELECT type_of_client, COUNT(*) as c FROM companies GROUP BY type_of_client ORDER BY c DESC').fetchall()
    region_data = conn.execute('SELECT region, COUNT(*) as c FROM companies GROUP BY region ORDER BY c DESC LIMIT 10').fetchall()
    kyc_data = conn.execute('SELECT kyc_status, COUNT(*) as c FROM companies GROUP BY kyc_status ORDER BY c DESC').fetchall()
    mode_data = conn.execute('SELECT mode_of_ac, COUNT(*) as c FROM companies GROUP BY mode_of_ac ORDER BY c DESC').fetchall()

    today = datetime.now().date()
    expired_tl = conn.execute('SELECT COUNT(*) as c FROM companies WHERE trade_license_expiry < ?', (today,)).fetchone()['c']
    expiring_30 = conn.execute('SELECT COUNT(*) as c FROM companies WHERE trade_license_expiry BETWEEN ? AND ?',
        (today, today + timedelta(days=30))).fetchone()['c']
    expiring_90 = conn.execute('SELECT COUNT(*) as c FROM companies WHERE trade_license_expiry BETWEEN ? AND ?',
        (today, today + timedelta(days=90))).fetchone()['c']

    conn.close()
    return render_template('reports.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        risk_data=risk_data,
        doc_data=doc_data,
        type_data=type_data,
        region_data=region_data,
        kyc_data=kyc_data,
        mode_data=mode_data,
        expired_tl=expired_tl,
        expiring_30=expiring_30,
        expiring_90=expiring_90
    )

# EXPORT
@app.route('/export/companies')
@login_required
def export_companies():
    conn = get_db()
    companies = conn.execute('SELECT * FROM companies').fetchall()
    conn.close()
    output = io.StringIO()
    writer = csv.writer(output)
    if companies:
        writer.writerow(companies[0].keys())
        for c in companies:
            writer.writerow(c)
    output.seek(0)
    return send_file(io.BytesIO(output.getvalue().encode()),
        mimetype='text/csv', as_attachment=True,
        download_name=f'companies_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv')

if __name__ == '__main__':
    port = int(os.getenv('PORT', 8000))
    app.run(debug=False, host='0.0.0.0', port=port)
