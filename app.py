import sys
import os
print("🔴 APP STARTING", flush=True)

from flask import Flask
print("✅ Flask imported", flush=True)

app = Flask(__name__)
app.secret_key = 'dev'

@app.route('/')
def index():
    return 'OK'

if __name__ == '__main__':
    port = int(os.getenv('PORT', 8000))
    print(f"🔴 Starting on port {port}", flush=True)
    sys.stdout.flush()
    sys.stderr.flush()
    app.run(host='0.0.0.0', port=port, debug=False)
