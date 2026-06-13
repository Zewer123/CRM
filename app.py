from flask import Flask, redirect, url_for, session
import os

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'dev-key-12345')

@app.route('/')
def index():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login')
def login():
    return '<h1>Login Page</h1><p>This is the login page</p>'

@app.route('/dashboard')
def dashboard():
    return '<h1>Dashboard</h1><p>Welcome!</p>'

if __name__ == '__main__':
    port = int(os.getenv('PORT', 8000))
    app.run(host='0.0.0.0', port=port, debug=False)
