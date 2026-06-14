from flask import Flask, render_template, request, jsonify, session, redirect, url_for, send_file
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename
import sqlite3
from datetime import datetime, timedelta
from functools import wraps
import os
import csv
import io
from collections import defaultdict

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'zewer-aml-crm-secret-2026')

DB = 'aml_crm.db'
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf', 'doc', 'docx', 'jpg', 'jpeg', 'png', 'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        conn = get_db()
        user = conn.execute('SELECT role FROM users WHERE id = ?', (session['user_id'],)).fetchone()
        conn.close()
        if not user or user['role'] != 'admin':
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated_function

# ============ AUTH ROUTES ============

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

# ============ DASHBOARD ============

@app.route('/dashboard')
@login_required
def dashboard():
    conn = get_db()
    
    # Get stats
    total_companies = conn.execute('SELECT COUNT(*) as count FROM companies').fetchone()['count']
    active_companies = conn.execute('SELECT COUNT(*) as count FROM companies WHERE ac_status = "Active"').fetchone()['count']
    
    # Get expiring documents
    today = datetime.now().date()
    expiring_30 = conn.execute('''
        SELECT COUNT(*) as count FROM companies 
        WHERE (trade_license_expiry BETWEEN ? AND ?) 
        OR (address_proof_expiry BETWEEN ? AND ?)
    ''', (today, today + timedelta(days=30), today, today + timedelta(days=30))).fetchone()['count']
    
    expiring_60 = conn.execute('''
        SELECT COUNT(*) as count FROM companies 
        WHERE (trade_license_expiry BETWEEN ? AND ?) 
        OR (address_proof_expiry BETWEEN ? AND ?)
    ''', (today + timedelta(days=31), today + timedelta(days=60), today + timedelta(days=31), today + timedelta(days=60))).fetchone()['count']
    
    # Risk breakdown
    risk_data = conn.execute('''
        SELECT risk_status, COUNT(*) as count FROM companies 
        GROUP BY risk_status
    ''').fetchall()
    
    risk_breakdown = {row['risk_status']: row['count'] for row in risk_data}
    
    conn.close()
    
    return render_template('dashboard.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        total_companies=total_companies,
        active_companies=active_companies,
        expiring_30=expiring_30,
        expiring_60=expiring_60,
        risk_breakdown=risk_breakdown,
        today=today,
        timedelta=timedelta
    )

# ============ COMPANIES ============

@app.route('/companies')
@login_required
def companies():
    conn = get_db()
    
    search = request.args.get('search', '')
    filter_status = request.args.get('status', '')
    
    query = 'SELECT * FROM companies WHERE 1=1'
    params = []
    
    if search:
        query += ' AND (client_name LIKE ? OR customer_name LIKE ? OR ac_code LIKE ?)'
        params = [f'%{search}%', f'%{search}%', f'%{search}%']
    
    if filter_status:
        query += ' AND ac_status = ?'
        params.append(filter_status)
    
    query += ' ORDER BY created_at DESC'
    
    companies = conn.execute(query, params).fetchall()
    conn.close()
    
    return render_template('companies.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        companies=companies,
        today=datetime.now().date(),
        timedelta=timedelta
    )

@app.route('/company/new', methods=['GET'])
@login_required
def company_new():
    conn = get_db()
    dropdowns = conn.execute('SELECT DISTINCT field_name FROM dropdowns WHERE is_active = 1').fetchall()
    dropdown_data = {}
    for field in dropdowns:
        values = conn.execute('SELECT value FROM dropdowns WHERE field_name = ? AND is_active = 1', (field['field_name'],)).fetchall()
        dropdown_data[field['field_name']] = [v['value'] for v in values]
    conn.close()
    
    return render_template('company_form.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        dropdown_data=dropdown_data,
        company=None
    )

@app.route('/company/<int:id>', methods=['GET'])
@login_required
def company_detail(id):
    conn = get_db()
    company = conn.execute('SELECT * FROM companies WHERE id = ?', (id,)).fetchone()
    ubos = conn.execute('SELECT * FROM ubos WHERE company_id = ?', (id,)).fetchall()
    notes = conn.execute('SELECT * FROM followup_notes WHERE company_id = ? ORDER BY created_at DESC', (id,)).fetchall()
    dropdowns = conn.execute('SELECT DISTINCT field_name FROM dropdowns WHERE is_active = 1').fetchall()
    dropdown_data = {}
    for field in dropdowns:
        values = conn.execute('SELECT value FROM dropdowns WHERE field_name = ? AND is_active = 1', (field['field_name'],)).fetchall()
        dropdown_data[field['field_name']] = [v['value'] for v in values]
    conn.close()
    
    if not company:
        return redirect(url_for('companies'))
    
    return render_template('company_detail.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        company=company,
        ubos=ubos,
        notes=notes,
        dropdown_data=dropdown_data
    )

@app.route('/api/company/add', methods=['POST'])
@login_required
def api_add_company():
    data = request.get_json()
    
    try:
        conn = get_db()
        conn.execute('''
            INSERT INTO companies (
                ac_code, client_name, customer_name, ac_opening_date,
                ac_status, active_till_year, nature, type_of_client,
                name_of_freezone, mode_of_ac, country_of_incorporation,
                region, address, telephone, mobile, whatsapp_link, email_id,
                address_proof_type, address_proof_expiry, kyc_status,
                trade_license_no, issuing_authority, legal_type,
                incorporation_date, trade_license_expiry, tax_no_trn,
                vat_cert, vat_declaration, num_beneficial_owners,
                moa, pep, undertaking, source_of_fund, doc_status,
                risk_status, verified_by, verified_date, followup_details,
                crowe_feedback, zewer_comments, created_by
            ) VALUES (
                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?, ?, ?, ?
            )
        ''', (
            data.get('ac_code'), data.get('client_name'), data.get('customer_name'), data.get('ac_opening_date'),
            data.get('ac_status', 'Active'), data.get('active_till_year'), data.get('nature'), data.get('type_of_client'),
            data.get('name_of_freezone'), data.get('mode_of_ac'), data.get('country_of_incorporation'),
            data.get('region'), data.get('address'), data.get('telephone'), data.get('mobile'), data.get('whatsapp_link'), data.get('email_id'),
            data.get('address_proof_type'), data.get('address_proof_expiry'), data.get('kyc_status'),
            data.get('trade_license_no'), data.get('issuing_authority'), data.get('legal_type'),
            data.get('incorporation_date'), data.get('trade_license_expiry'), data.get('tax_no_trn'),
            data.get('vat_cert'), data.get('vat_declaration'), data.get('num_beneficial_owners'),
            data.get('moa'), data.get('pep'), data.get('undertaking'), data.get('source_of_fund'), data.get('doc_status', 'Incomplete'),
            data.get('risk_status', 'Unspecified'), data.get('verified_by'), data.get('verified_date'), data.get('followup_details'),
            data.get('crowe_feedback'), data.get('zewer_comments'), session.get('user_id')
        ))
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'Company added successfully'}), 200
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

# ============ UBOs ============

@app.route('/api/ubo/add/<int:company_id>', methods=['POST'])
@login_required
def api_add_ubo(company_id):
    data = request.get_json()
    
    try:
        conn = get_db()
        company = conn.execute('SELECT ac_code, client_name FROM companies WHERE id = ?', (company_id,)).fetchone()
        
        conn.execute('''
            INSERT INTO ubos (
                company_id, ac_code, client_name, position, share_percentage,
                person_name, nationality, residential_status, passport_no,
                passport_expiry, emirates_id, emirates_id_expiry, doc_status,
                verified_by, verified_date, group_of_companies
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            company_id, company['ac_code'], company['client_name'],
            data.get('position'), data.get('share_percentage'),
            data.get('person_name'), data.get('nationality'), data.get('residential_status'),
            data.get('passport_no'), data.get('passport_expiry'), data.get('emirates_id'),
            data.get('emirates_id_expiry'), data.get('doc_status', 'Incomplete'),
            data.get('verified_by'), data.get('verified_date'), data.get('group_of_companies')
        ))
        conn.commit()
        conn.close()
        return jsonify({'success': True}), 200
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/ubo/<int:id>/delete', methods=['POST'])
@login_required
def api_delete_ubo(id):
    try:
        conn = get_db()
        ubo = conn.execute('SELECT company_id FROM ubos WHERE id = ?', (id,)).fetchone()
        conn.execute('DELETE FROM ubos WHERE id = ?', (id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True}), 200
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ============ SETTINGS ============

@app.route('/settings')
@admin_required
def settings():
    conn = get_db()
    
    # Get all users
    users = conn.execute('SELECT id, email, name, role, is_active FROM users').fetchall()
    
    # Get all dropdowns
    dropdowns = conn.execute('SELECT * FROM dropdowns WHERE is_active = 1 ORDER BY field_name, value').fetchall()
    dropdown_groups = defaultdict(list)
    for dd in dropdowns:
        dropdown_groups[dd['field_name']].append(dd)
    
    conn.close()
    
    return render_template('settings.html',
        user_name=session.get('user_name'),
        users=users,
        dropdown_groups=dict(dropdown_groups)
    )

@app.route('/api/user/add', methods=['POST'])
@admin_required
def api_add_user():
    data = request.get_json()
    
    try:
        password_hash = generate_password_hash(data.get('password'))
        conn = get_db()
        conn.execute('''
            INSERT INTO users (email, password_hash, name, role, is_active)
            VALUES (?, ?, ?, ?, ?)
        ''', (data.get('email'), password_hash, data.get('name'), data.get('role'), 1))
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
        conn.execute('''
            INSERT INTO dropdowns (field_name, value, is_active)
            VALUES (?, ?, 1)
        ''', (data.get('field_name'), data.get('value')))
        conn.commit()
        conn.close()
        return jsonify({'success': True}), 200
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# ============ REPORTS & ALERTS ============

@app.route('/alerts')
@login_required
def alerts():
    conn = get_db()
    today = datetime.now().date()
    
    expiring = conn.execute('''
        SELECT id, client_name, trade_license_expiry, address_proof_expiry
        FROM companies
        WHERE (trade_license_expiry BETWEEN ? AND ?) 
        OR (address_proof_expiry BETWEEN ? AND ?)
        ORDER BY trade_license_expiry ASC
    ''', (today, today + timedelta(days=90), today, today + timedelta(days=90))).fetchall()
    
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
    today = datetime.now().date()
    
    # Expired documents
    expired = conn.execute('''
        SELECT client_name, trade_license_expiry, address_proof_expiry
        FROM companies
        WHERE trade_license_expiry < ? OR address_proof_expiry < ?
    ''', (today, today)).fetchall()
    
    # Risk breakdown
    risk_data = conn.execute('''
        SELECT risk_status, COUNT(*) as count
        FROM companies
        GROUP BY risk_status
    ''').fetchall()
    
    conn.close()
    
    return render_template('reports.html',
        user_name=session.get('user_name'),
        user_role=session.get('user_role'),
        expired=expired,
        risk_data=risk_data
    )

# ============ EXPORT ============

@app.route('/export/companies')
@login_required
def export_companies():
    conn = get_db()
    companies = conn.execute('SELECT * FROM companies').fetchall()
    conn.close()
    
    output = io.StringIO()
    writer = csv.writer(output)
    
    # Headers
    if companies:
        writer.writerow(companies[0].keys())
        for company in companies:
            writer.writerow(company)
    
    output.seek(0)
    return send_file(io.BytesIO(output.getvalue().encode()), mimetype='text/csv', as_attachment=True, download_name=f'companies_{datetime.now().strftime("%Y%m%d")}.csv')

@app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html'), 404

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000)
