from flask import Flask, render_template, request, jsonify, session, redirect, url_for
from werkzeug.security import check_password_hash
from sqlalchemy import create_engine, text, event
from sqlalchemy.exc import OperationalError
import os
from dotenv import load_dotenv
from functools import wraps
import sys

# Load environment variables
load_dotenv()

# Create Flask app
app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'dev-secret-key-change-in-production')

# ============ DATABASE SETUP ============

DATABASE_URL = os.getenv('DATABASE_URL')

print("=" * 80, file=sys.stderr)
print(f"DATABASE_URL from env: {DATABASE_URL[:60] if DATABASE_URL else 'NOT SET'}...", file=sys.stderr)
print("=" * 80, file=sys.stderr)

if not DATABASE_URL:
    print("ERROR: DATABASE_URL is not set!", file=sys.stderr)
    engine = None
else:
    # Fix URL format
    if DATABASE_URL.startswith('postgres://'):
        DATABASE_URL = DATABASE_URL.replace('postgres://', 'postgresql://', 1)
        print(f"Converted postgres:// to postgresql://", file=sys.stderr)
    
    try:
        print(f"Creating SQLAlchemy engine with URL: {DATABASE_URL[:60]}...", file=sys.stderr)
        
        engine = create_engine(
            DATABASE_URL,
            pool_pre_ping=True,
            pool_recycle=3600,
            echo=True  # Enable SQL logging
        )
        
        # Test the connection immediately
        with engine.connect() as conn:
            result = conn.execute(text('SELECT 1'))
            print("✓✓✓ DATABASE CONNECTION TEST SUCCESSFUL ✓✓✓", file=sys.stderr)
        
        print("✓ SQLAlchemy engine created and tested successfully", file=sys.stderr)
        
    except Exception as e:
        print(f"✗✗✗ ENGINE CREATION FAILED ✗✗✗", file=sys.stderr)
        print(f"Error type: {type(e).__name__}", file=sys.stderr)
        print(f"Error message: {str(e)}", file=sys.stderr)
        import traceback
        traceback.print_exc(file=sys.stderr)
        engine = None

# ============ ROUTES ============

@app.route('/api/health')
def health_check():
    """Health check endpoint"""
    try:
        if engine is None:
            return jsonify({
                'status': 'unhealthy',
                'message': 'Engine is None - check logs',
                'database_url_set': DATABASE_URL is not None
            }), 500
        
        with engine.connect() as conn:
            conn.execute(text('SELECT 1'))
            return jsonify({
                'status': 'healthy',
                'database': 'connected',
                'message': 'Database OK'
            }), 200
            
    except Exception as e:
        print(f"Health check error: {e}", file=sys.stderr)
        return jsonify({
            'status': 'unhealthy',
            'error': str(e)[:100]
        }), 500

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
            if engine is None:
                return jsonify({'success': False, 'error': 'Database not available'}), 500
            
            with engine.connect() as conn:
                result = conn.execute(
                    text('SELECT id, email, password_hash, name, role FROM users WHERE email = :email'),
                    {'email': email}
                )
                user = result.fetchone()
                
                if user:
                    user_id, user_email, password_hash, user_name, user_role = user
                    
                    if check_password_hash(password_hash, password):
                        session['user_id'] = user_id
                        session['user_email'] = user_email
                        session['user_name'] = user_name
                        session['user_role'] = user_role
                        print(f"✓ Login successful: {email}", file=sys.stderr)
                        return jsonify({'success': True}), 200
                    else:
                        return jsonify({'success': False, 'error': 'Invalid email or password'}), 401
                else:
                    return jsonify({'success': False, 'error': 'Invalid email or password'}), 401
                    
        except Exception as e:
            print(f"Login error: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)
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
    print(f"Starting server on port {port}...", file=sys.stderr)
    app.run(debug=False, host='0.0.0.0', port=port)
