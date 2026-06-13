from flask import Flask, jsonify
import os
import traceback

app = Flask(__name__)

@app.route('/')
def index():
    try:
        # Check DATABASE_URL
        db_url = os.getenv('DATABASE_URL')
        if not db_url:
            return f'''
            <h1>❌ DATABASE_URL NOT SET</h1>
            <p>Environment variables:</p>
            <pre>{os.environ}</pre>
            '''
        
        # Try to import psycopg2
        try:
            import psycopg2
            print("✅ psycopg2 imported")
        except ImportError as e:
            return f"<h1>❌ psycopg2 import failed:</h1><pre>{str(e)}</pre>"
        
        # Try to connect
        try:
            conn = psycopg2.connect(db_url)
            conn.close()
            return "<h1>✅ APP WORKING - DB CONNECTED</h1>"
        except Exception as e:
            return f'''
            <h1>❌ DATABASE CONNECTION FAILED</h1>
            <pre>Error: {str(e)}
            
Traceback:
{traceback.format_exc()}
            </pre>
            '''
    
    except Exception as e:
        return f'''
        <h1>❌ UNEXPECTED ERROR</h1>
        <pre>{traceback.format_exc()}</pre>
        '''

@app.errorhandler(500)
def server_error(e):
    return f'''
    <h1>500 ERROR</h1>
    <pre>{traceback.format_exc()}</pre>
    ''', 500

if __name__ == '__main__':
    port = int(os.getenv('PORT', 8000))
    print(f"Starting on port {port}")
    app.run(host='0.0.0.0', port=port, debug=False)
