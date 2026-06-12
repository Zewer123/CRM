from flask import Flask, render_template, request, jsonify, session, redirect, url_for
from werkzeug.security import check_password_hash
import pg8000.native
import os
from dotenv import load_dotenv
from functools import wraps
import urllib.parse

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'dev-secret-key-change-in-production')

# ============ DATABASE CONNECTION ============

DATABASE_URL = os.getenv('DATABASE_URL')

def get_db_connection():
    """Connect to PostgreSQL using pg8000 (pure Python, no system libs needed)"""
    try:
        if not DATABASE_URL:
            print("ERROR: DATABASE_URL not set")
            return None
        
        # Parse PostgreSQL URL
        # Format: postgresql://user:password@host:port/database
        url = DATABASE_URL.replace('postgresql://', '').replace('postgres://', '')
        
        # Split credentials and host
        if '@' in url:
            credentials, host_db = url.split('@')
            user, password = credentials.split(':')
        else:
            user = 'postgres'
            password = ''
            host_db = url
        
        # Split host and database
        if '/' in host_db:
            host_port, database = host_db.split('/', 1)
        else:
            host_port = host_db
            database = 'postgres'
        
        # Split host and port
        if ':' in host_port:
            host, port = host_port.split(':')
            port = int(port)
        else:
            host = host_port
            port = 5432
        
        # Decode password if URL encoded
        password = urllib.parse.unquote(password)
        
        print(f"✓ Connecting to PostgreSQL: {user}@{host}:{port}/{database}")
        
        # Connect using pg8000
        conn = pg8000.native.connect(
            user=user,
            password=password,
            host=host,
            port=port,
            database=database,
            timeout=10
        )
        
        print("✓ Database connection successful!")
        return conn
        
    except Exception as e:
        print(f"✗ Database connection error: {e}")
        import traceback
        traceback.print_exc()
        return None

# ============ ROUTES ============

@app.route('/api/health')
def health_check():
    """Health check - test database"""
    try:
        conn = get_db_connection()
        if conn:
            try:
                result = conn.run('SELECT 1')
                conn.close()
                return jsonify({'status': 'healthy', 'database': 'connected'}), 200
            except Exception as e:
                conn.close()
                return jsonify({'status': 'unhealthy', 'error': str(e)[:100]}), 500
        else:
            return jsonify({'status': 'unhealthy', 'error': 'Could not connect'}), 500
    except Exception as e:
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
        
        conn = get_db_connection()
        if not conn:
            return jsonify({'success': False, 'error': 'Database not available'}), 500
        
        try:
            # Query user
            result = conn.run(
                'SELECT id, email, password_hash, name, role FROM users WHERE email = :email',
                email=email
            )
            
            if result:
                user_id, user_email, password_hash, user_name, user_role = result[0]
                
                # Check password
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
            print(f"✗ Login error: {e}")
            conn.close()
            return jsonify({'success': False, 'error': str(e)[:100]}), 500
    
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
    return jsonify({
        'success': True,
        'user': {
            'id': session.get('user_id'),
            'email': session.get('user_email'),
            'name': session.get('user_name'),
            'role': session.get('user_role')
        }
    }), 200

@app.errorhandler(404)
def not_found(e):
    return jsonify({'error': 'Not found'}), 404

@app.errorhandler(500)
def server_error(e):
    return jsonify({'error': 'Server error'}), 500

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    print(f"Starting server on port {port}...")
    app.run(debug=False, host='0.0.0.0', port=port)
