from flask import Flask, render_template, request, jsonify, session, redirect, url_for
from werkzeug.security import check_password_hash
import pg8000.native
import os
from dotenv import load_dotenv
from functools import wraps

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'dev-secret-key-change')

DATABASE_URL = os.getenv('DATABASE_URL')

print("=" * 80)
print(f"DATABASE_URL: {DATABASE_URL[:60] if DATABASE_URL else 'NOT SET'}...")
print("=" * 80)

def get_db_connection():
    """Connect to PostgreSQL using pg8000"""
    try:
        if not DATABASE_URL:
            print("ERROR: DATABASE_URL not set!")
            return None
        
        # pg8000 can parse PostgreSQL URLs directly!
        print(f"Attempting connection to: {DATABASE_URL[:60]}...")
        conn = pg8000.native.connect(DATABASE_URL, timeout=10)
        print("✓ Connection successful!")
        return conn
        
    except Exception as e:
        print(f"✗ Connection failed: {type(e).__name__}: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

@app.route('/api/health')
def health_check():
    try:
        conn = get_db_connection()
        if conn:
            conn.run('SELECT 1')
            conn.close()
            return jsonify({'status': 'healthy', 'message': 'Database OK'}), 200
        else:
            return jsonify({'status': 'unhealthy', 'message': 'Connection failed'}), 500
    except Exception as e:
        print(f"Health check error: {e}")
        return jsonify({'status': 'unhealthy', 'error': str(e)[:100]}), 500

@app.route('/')
def index():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        data = request.get_json()
        email = data.get('email', '').strip()
        password = data.get('password', '')
        
        if not email or not password:
            return jsonify({'success': False, 'error': 'Email and password required'}), 400
        
        try:
            conn = get_db_connection()
            if not conn:
                return jsonify({'success': False, 'error': 'Database connection failed'}), 500
            
            # Query user
            rows = conn.run(
                'SELECT id, email, password_hash, name, role FROM users WHERE email = :email',
                email=email
            )
            
            if rows and len(rows) > 0:
                user_id, user_email, password_hash, user_name, user_role = rows[0]
                
                if check_password_hash(password_hash, password):
                    session['user_id'] = user_id
                    session['user_email'] = user_email
                    session['user_name'] = user_name
                    session['user_role'] = user_role
                    print(f"✓ Login success: {email}")
                    conn.close()
                    return jsonify({'success': True}), 200
                else:
                    conn.close()
                    return jsonify({'success': False, 'error': 'Invalid email or password'}), 401
            else:
                conn.close()
                return jsonify({'success': False, 'error': 'Invalid email or password'}), 401
        
        except Exception as e:
            print(f"Login error: {type(e).__name__}: {e}")
            import traceback
            traceback.print_exc()
            return jsonify({'success': False, 'error': f'Error: {str(e)[:100]}'}), 500
    
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html',
                          user_name=session.get('user_name'),
                          user_role=session.get('user_role'))

@app.route('/api/user')
@login_required
def get_user():
    return jsonify({'success': True, 'user': {
        'id': session.get('user_id'),
        'email': session.get('user_email'),
        'name': session.get('user_name'),
        'role': session.get('user_role')
    }}), 200

@app.errorhandler(404)
def not_found(e):
    return jsonify({'error': 'Not found'}), 404

@app.errorhandler(500)
def server_error(e):
    return jsonify({'error': 'Server error'}), 500

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
