from flask import Flask, render_template, request, jsonify, session, redirect, url_for
from werkzeug.security import check_password_hash, generate_password_hash
from sqlalchemy import create_engine, text
from sqlalchemy.exc import OperationalError
import os
from dotenv import load_dotenv
from functools import wraps

# Load environment variables
load_dotenv()

# Create Flask app
app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'dev-secret-key-change-in-production')

# ============ DATABASE SETUP ============

# Get DATABASE_URL from environment (Railway provides this)
DATABASE_URL = os.getenv('DATABASE_URL', 'postgresql://postgres:password@localhost:5432/aml_crm')

# Ensure we use the correct PostgreSQL URI format
if DATABASE_URL and DATABASE_URL.startswith('postgres://'):
    DATABASE_URL = DATABASE_URL.replace('postgres://', 'postgresql://', 1)

print(f"Connecting to database: {DATABASE_URL[:50]}...")

# Create SQLAlchemy engine
try:
    engine = create_engine(
        DATABASE_URL,
        pool_pre_ping=True,  # Test connection before using
        pool_recycle=3600,   # Recycle connections after 1 hour
        echo=False
    )
    print("✓ SQLAlchemy engine created")
except Exception as e:
    print(f"ERROR: Could not create database engine: {e}")
    engine = None

# ============ ROUTES ============

@app.route('/api/health')
def health_check():
    """Health check endpoint - tests database connection"""
    try:
        if engine is None:
            return jsonify({
                'status': 'unhealthy',
                'database': 'not_configured',
                'message': 'Database engine not initialized'
            }), 500
        
        # Try to execute a simple query
        with engine.connect() as conn:
            conn.execute(text('SELECT 1'))
            return jsonify({
                'status': 'healthy',
                'database': 'connected',
                'message': 'Database connection OK'
            }), 200
    except OperationalError as e:
        return jsonify({
            'status': 'unhealthy',
            'database': 'connection_error',
            'message': f'Database connection failed: {str(e)[:200]}',
            'error_type': 'OperationalError'
        }), 500
    except Exception as e:
        return jsonify({
            'status': 'unhealthy',
            'database': 'error',
            'message': f'Unexpected error: {str(e)[:200]}',
            'error_type': type(e).__name__
        }), 500

@app.route('/')
def index():
    """Home page - redirects to dashboard if logged in, else to login"""
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Login page - accepts email and password"""
    if request.method == 'POST':
        data = request.get_json()
        email = data.get('email', '').strip()
        password = data.get('password', '')
        
        # Validate input
        if not email or not password:
            return jsonify({'success': False, 'error': 'Email and password required'}), 400
        
        try:
            if engine is None:
                print("ERROR: Database not configured")
                return jsonify({'success': False, 'error': 'Database configuration error'}), 500
            
            # Query user by email
            with engine.connect() as conn:
                result = conn.execute(text('SELECT id, email, password_hash, name, role FROM users WHERE email = :email'), 
                                      {'email': email})
                user = result.fetchone()
                
                if user:
                    user_id, user_email, password_hash, user_name, user_role = user
                    
                    # Check password
                    if check_password_hash(password_hash, password):
                        # Login successful
                        session['user_id'] = user_id
                        session['user_email'] = user_email
                        session['user_name'] = user_name
                        session['user_role'] = user_role
                        print(f"✓ Login successful: {email}")
                        return jsonify({'success': True, 'message': 'Login successful'}), 200
                    else:
                        print(f"✗ Invalid password for: {email}")
                        return jsonify({'success': False, 'error': 'Invalid email or password'}), 401
                else:
                    print(f"✗ User not found: {email}")
                    return jsonify({'success': False, 'error': 'Invalid email or password'}), 401
                    
        except OperationalError as e:
            print(f"Database connection error: {e}")
            return jsonify({'success': False, 'error': 'Database connection failed. Please try again.'}), 500
        except Exception as e:
            print(f"Login error: {e}")
            import traceback
            traceback.print_exc()
            return jsonify({'success': False, 'error': f'Server error: {str(e)[:100]}'}), 500
    
    # GET request - show login form
    return render_template('login.html')

@app.route('/logout')
def logout():
    """Logout - clears session and goes to login page"""
    session.clear()
    return redirect(url_for('login'))

def login_required(f):
    """Checks if user is logged in - if not, sends to login page"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

@app.route('/dashboard')
@login_required
def dashboard():
    """Dashboard - only for logged in users"""
    return render_template('dashboard.html', 
                          user_name=session.get('user_name'),
                          user_role=session.get('user_role'))

@app.route('/api/user')
@login_required
def get_user():
    """Returns current logged in user info (for JavaScript)"""
    return jsonify({
        'success': True,
        'user': {
            'id': session.get('user_id'),
            'email': session.get('user_email'),
            'name': session.get('user_name'),
            'role': session.get('user_role')
        }
    }), 200

# ============ ERROR HANDLERS ============

@app.errorhandler(404)
def page_not_found(error):
    """If page doesn't exist, show error"""
    return jsonify({'error': 'Page not found'}), 404

@app.errorhandler(500)
def server_error(error):
    """If server error, show error message"""
    return jsonify({'error': 'Server error'}), 500

# ============ RUN SERVER ============

if __name__ == '__main__':
    # Get port from environment or use 5000
    port = int(os.getenv('PORT', 5000))
    print(f"Starting Flask server on port {port}...")
    app.run(debug=False, host='0.0.0.0', port=port)
