from flask import Flask, render_template, request, jsonify, session, redirect, url_for
from werkzeug.security import check_password_hash, generate_password_hash
import os
from dotenv import load_dotenv
from functools import wraps

try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
except ImportError:
    psycopg2 = None
    RealDictCursor = None

# Load environment variables (secret keys, database info)
load_dotenv()

# Create Flask app
app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'dev-secret-key-change-in-production')

# Database configuration - use DATABASE_URL if available (Railway), else use individual vars
DATABASE_URL = os.getenv('DATABASE_URL')

# Function to connect to database
def get_db_connection():
    """Creates connection to PostgreSQL database"""
    try:
        if psycopg2 is None:
            print("ERROR: psycopg2 not available")
            return None
        
        # Try to use DATABASE_URL from Railway
        if DATABASE_URL:
            # Handle both postgres:// and postgresql:// URLs
            db_url = DATABASE_URL.replace('postgres://', 'postgresql://')
            print(f"Connecting with DATABASE_URL...")
            conn = psycopg2.connect(db_url)
            return conn
        
        # Fallback to individual environment variables
        print("Using individual DB environment variables...")
        conn = psycopg2.connect(
            host=os.getenv('DB_HOST', 'localhost'),
            user=os.getenv('DB_USER', 'postgres'),
            password=os.getenv('DB_PASSWORD', 'password'),
            database=os.getenv('DB_NAME', 'aml_crm'),
            port=os.getenv('DB_PORT', '5432')
        )
        return conn
        
    except Exception as e:
        print(f"Database connection ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

# Login required decorator
def login_required(f):
    """Checks if user is logged in - if not, sends to login page"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

# ============ ROUTES (Pages) ============

@app.route('/api/health')
def health_check():
    """Health check endpoint - tests database connection"""
    try:
        conn = get_db_connection()
        if conn:
            cur = conn.cursor()
            cur.execute('SELECT 1')
            cur.close()
            conn.close()
            return jsonify({
                'status': 'healthy',
                'database': 'connected',
                'message': 'Database connection OK'
            }), 200
        else:
            return jsonify({
                'status': 'unhealthy',
                'database': 'disconnected',
                'message': 'Could not establish database connection'
            }), 500
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        return jsonify({
            'status': 'unhealthy',
            'database': 'error',
            'message': str(e),
            'error_type': type(e).__name__,
            'traceback': error_trace
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
        email = data.get('email')
        password = data.get('password')
        
        # Check if email and password provided
        if not email or not password:
            return jsonify({'success': False, 'error': 'Email and password required'}), 400
        
        # Connect to database
        conn = get_db_connection()
        if not conn:
            print(f"ERROR: Could not connect to database!")
            return jsonify({'success': False, 'error': 'Database connection failed. Please try again.'}), 500
        
        try:
            cur = conn.cursor(cursor_factory=RealDictCursor)
            # Look up user by email
            cur.execute('SELECT * FROM users WHERE email = %s', (email,))
            user = cur.fetchone()
            
            # Check password matches
            if user and check_password_hash(user['password_hash'], password):
                # Login successful - save to session
                session['user_id'] = user['id']
                session['user_email'] = user['email']
                session['user_name'] = user['name']
                session['user_role'] = user['role']
                print(f"Login successful for: {email}")
                return jsonify({'success': True, 'message': 'Login successful'}), 200
            else:
                print(f"Login failed for: {email} - invalid credentials")
                return jsonify({'success': False, 'error': 'Invalid email or password'}), 401
        except Exception as e:
            print(f"Login error: {str(e)}")
            import traceback
            traceback.print_exc()
            return jsonify({'success': False, 'error': f'Server error: {str(e)}'}), 500
        finally:
            try:
                cur.close()
                conn.close()
            except:
                pass
    
    # GET request - show login form
    return render_template('login.html')

@app.route('/logout')
def logout():
    """Logout - clears session and goes to login page"""
    session.clear()
    return redirect(url_for('login'))

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
    return render_template('404.html'), 404

@app.errorhandler(500)
def server_error(error):
    """If server error, show error message"""
    return render_template('500.html'), 500

# ============ RUN SERVER ============

if __name__ == '__main__':
    # Get port from environment or use 5000
    port = int(os.getenv('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)
