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
    if os.path.exists(DB):
        return
    
    conn = sqlite3.connect(DB)
    conn.execute('PRAGMA foreign_keys = ON')
    cursor = conn.cursor()
    
    cursor.executescript('''
        CREATE TABLE users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            name TEXT NOT NULL,
            role TEXT DEFAULT 'staff',
            is_active INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        
        CREATE TABLE companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ac_code TEXT UNIQUE NOT NULL,
            client_name TEXT NOT NULL,
            customer_name TEXT,
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
            whatsapp_link TEXT,
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
            num_beneficial_owners INTEGER DEFAULT 0,
            moa TEXT,
            pep TEXT,
            undertaking TEXT,
            source_of_fund TEXT,
            software_updation TEXT,
            doc_status TEXT DEFAULT 'Incomplete',
            screening_date DATE,
            screening_tool_registered TEXT,
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
        
        CREATE TABLE ubos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            ac_code TEXT,
            client_name TEXT,
            position TEXT,
            share_percentage DECIMAL(5,2),
            person_name TEXT NOT NULL,
            nationality TEXT,
            residential_status TEXT,
            passport_no TEXT,
            passport_expiry DATE,
            emirates_id TEXT,
            emirates_id_expiry DATE,
            doc_status TEXT DEFAULT 'Incomplete',
            verified_by TEXT,
            verified_date DATE,
            followup_details TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(company_id) REFERENCES companies(id) ON DELETE CASCADE
        );
        
        CREATE TABLE dropdowns (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            field_name TEXT NOT NULL,
            value TEXT NOT NULL,
            is_active INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(field_name, value)
        );
    ''')
    
    # Insert admin user
    admin_password = generate_password_hash('Admin@123')
    cursor.execute('INSERT INTO users (email, password_hash, name, role) VALUES (?, ?, ?, ?)',
        ('admin@zewer.ae', admin_password, 'Administrator', 'admin'))
    
    # Insert compliance user
    compliance_password = generate_password_hash('Compliance@123')
    cursor.execute('INSERT INTO users (email, password_hash, name, role) VALUES (?, ?, ?, ?)',
        ('compliance@zewer.ae', compliance_password, 'Compliance Officer', 'compliance'))
    
    # Insert dropdown values
    dropdowns_data = {
        'AC STATUS': ['Active', 'Inactive'],
        'NATURE': ['Individual', 'Legal entity'],
        'TYPE OF CLIENT': ['MainLand', 'Free Zone', 'Abroad', 'International Corporate'],
        'MODE OF AC': ['Supplier', 'Customer', 'Bullion', 'Refinery', 'Logistics Co'],
        'RISK STATUS': ['High', 'Medium', 'Low', 'Unspecified'],
        'DOC STATUS': ['Completed', 'Incompleted'],
        'KYC STATUS': ['Yes', 'No', 'Pending'],
        'COUNTRY': ['UAE', 'Saudi Arabia', 'Kuwait', 'Qatar', 'Bahrain', 'Oman'],
        'REGION': ['Dubai', 'Abu Dhabi', 'Sharjah', 'Ajman', 'Ras Al Khaimah', 'Fujairah'],
        'FREEZONE': ['Jafza', 'DMCC', 'DAFZA', 'RAKEZ', 'ICAD'],
        'LEGAL TYPE': ['Individual', 'Partnership', 'LLC', 'Corporation']
    }
    
    for field, values in dropdowns_data.items():
        for value in values:
            cursor.execute('INSERT INTO dropdowns (field_name, value, is_active) VALUES (?, ?, 1)',
                (field, value))
    
    conn.commit()
    conn.close()

init_db()

def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn

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
    
    total = conn.execute('SELECT COUNT(*) as c FROM companies').fetchone()['c']
    active = conn.execute('SELECT COUNT(*) as c FROM companies WHERE ac_status = "Active"').fetchone()['c']
    
    today = datetime.now().date()
    expiring_30 = conn.execute('SELECT COUNT(*) as c FROM companies WHERE trade_license_expiry BETWEEN ? AND ?',
        (today, today + timedelta(days=30))).fetchone()['c']
    
    risk_data = conn.execute('SELECT risk_status, COUNT(*) as c FROM companies GROUP BY risk_status').fetchall()
    risk_breakdown = {row['risk_status']: row['c'] for row in risk_data}
    
    conn.close()
    
    return render_template('dashboard.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        total_companies=total,
        active_companies=active,
        expiring_30=expiring_30,
        risk_breakdown=risk_breakdown
    )

# COMPANIES
@app.route('/companies')
@login_required
def companies():
    conn = get_db()
    search = request.args.get('search', '')
    query = 'SELECT * FROM companies WHERE 1=1'
    params = []
    
    if search:
        query += ' AND (client_name LIKE ? OR customer_name LIKE ? OR mobile LIKE ?)'
        search_term = f'%{search}%'
        params = [search_term, search_term, search_term]
    
    query += ' ORDER BY created_at DESC'
    companies = conn.execute(query, params).fetchall()
    conn.close()
    
    return render_template('companies.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        companies=companies
    )

@app.route('/company/new')
@login_required
def company_new():
    conn = get_db()
    fields = conn.execute('SELECT DISTINCT field_name FROM dropdowns WHERE is_active = 1 ORDER BY field_name').fetchall()
    dropdown_data = {}
    for f in fields:
        vals = conn.execute('SELECT value FROM dropdowns WHERE field_name = ? AND is_active = 1 ORDER BY value', (f['field_name'],)).fetchall()
        dropdown_data[f['field_name']] = [v['value'] for v in vals]
    conn.close()
    
    return render_template('company_form.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        dropdown_data=dropdown_data
    )

@app.route('/company/<int:id>')
@login_required
def company_detail(id):
    conn = get_db()
    company = conn.execute('SELECT * FROM companies WHERE id = ?', (id,)).fetchone()
    ubos = conn.execute('SELECT * FROM ubos WHERE company_id = ?', (id,)).fetchall()
    conn.close()
    
    if not company:
        return redirect(url_for('companies'))
    
    return render_template('company_detail.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        company=company,
        ubos=ubos
    )

@app.route('/api/company/add', methods=['POST'])
@login_required
def api_add_company():
    data = request.get_json()
    try:
        conn = get_db()
        conn.execute('''INSERT INTO companies (
            ac_code, client_name, customer_name, ac_status, nature, type_of_client,
            mode_of_ac, country_of_incorporation, region, address, telephone, mobile,
            whatsapp_link, email_id, trade_license_no, issuing_authority, legal_type,
            trade_license_expiry, address_proof_expiry, kyc_status, risk_status,
            doc_status, created_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
        (data.get('ac_code'), data.get('client_name'), data.get('customer_name'),
         data.get('ac_status', 'Active'), data.get('nature'), data.get('type_of_client'),
         data.get('mode_of_ac'), data.get('country'), data.get('region'), data.get('address'),
         data.get('telephone'), data.get('mobile'), data.get('whatsapp'), data.get('email'),
         data.get('tl_no'), data.get('issuing_auth'), data.get('legal_type'),
         data.get('tl_expiry'), data.get('ejari_expiry'), data.get('kyc'), data.get('risk'),
         data.get('doc_status', 'Incomplete'), session.get('user_id')))
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

# ALERTS & REPORTS
@app.route('/alerts')
@login_required
def alerts():
    conn = get_db()
    today = datetime.now().date()
    expiring = conn.execute('''SELECT id, client_name, trade_license_expiry, address_proof_expiry
        FROM companies WHERE trade_license_expiry BETWEEN ? AND ?
        ORDER BY trade_license_expiry ASC''',
        (today, today + timedelta(days=90))).fetchall()
    conn.close()
    
    return render_template('alerts.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        alerts=expiring
    )

@app.route('/reports')
@login_required
def reports():
    conn = get_db()
    risk_data = conn.execute('SELECT risk_status, COUNT(*) as c FROM companies GROUP BY risk_status').fetchall()
    conn.close()
    
    return render_template('reports.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        risk_data=risk_data
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
        for company in companies:
            writer.writerow(company)
    
    output.seek(0)
    return send_file(io.BytesIO(output.getvalue().encode()),
        mimetype='text/csv',
        as_attachment=True,
        download_name=f'companies_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv')

if __name__ == '__main__':
    port = int(os.getenv('PORT', 8000))
    app.run(debug=False, host='0.0.0.0', port=port)
