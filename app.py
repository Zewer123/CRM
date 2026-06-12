from flask import Flask, render_template, request, jsonify, session, redirect, url_for
from werkzeug.security import check_password_hash
import asyncio
import asyncpg
import os
from dotenv import load_dotenv
from functools import wraps

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'dev-secret-key-change')

DATABASE_URL = os.getenv('DATABASE_URL')

print("=" * 80)
print(f"DATABASE_URL: {'SET' if DATABASE_URL else 'NOT SET'}")
print("=" * 80)

async def get_db_pool():
    """Create asyncpg connection pool"""
    try:
        print("Creating asyncpg pool...")
        pool = await asyncpg.create_pool(DATABASE_URL, min_size=1, max_size=10)
        print("✓ Pool created!")
        return pool
    except Exception as e:
        print(f"✗ Pool creation failed: {e}")
        return None

def run_async(coro):
    """Helper to run async code in Flask"""
    return asyncio.run(coro)

@app.before_request
def create_pool():
    """Create pool on first request"""
    if not hasattr(app, 'db_pool'):
        app.db_pool = run_async(get_db_pool())

@app.route('/api/health')
def health_check():
    try:
        if hasattr(app, 'db_pool') and app.db_pool:
            async def test():
                async with app.db_pool.acquire() as conn:
                    await conn.fetchval('SELECT 1')
                return True
            
            result = run_async(test())
            if result:
                return jsonify({'status': 'healthy', 'database': 'connected'}), 200
        
        return jsonify({'status': 'unhealthy', 'error': 'No pool'}), 500
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
        
        if not hasattr(app, 'db_pool') or not app.db_pool:
            return jsonify({'success': False, 'error': 'Database not available'}), 500
        
        try:
            async def authenticate():
                async with app.db_pool.acquire() as conn:
                    row = await conn.fetchrow(
                        'SELECT id, email, password_hash, name, role FROM users WHERE email = $1',
                        email
                    )
                    return row
            
            user_row = run_async(authenticate())
            
            if user_row:
                user_id, user_email, password_hash, user_name, user_role = user_row
                
                if check_password_hash(password_hash, password):
                    session['user_id'] = user_id
                    session['user_email'] = user_email
                    session['user_name'] = user_name
                    session['user_role'] = user_role
                    print(f"✓ Login success: {email}")
                    return jsonify({'success': True}), 200
                else:
                    return jsonify({'success': False, 'error': 'Invalid email or password'}), 401
            else:
                return jsonify({'success': False, 'error': 'Invalid email or password'}), 401
        
        except Exception as e:
            print(f"Login error: {type(e).__name__}: {e}")
            import traceback
            traceback.print_exc()
            return jsonify({'success': False, 'error': f'Error: {str(e)[:80]}'}), 500
    
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
