from flask import Flask
import sys
import os

print("STEP 1: Importing Flask", flush=True)
sys.stdout.flush()

app = Flask(__name__)

print("STEP 2: Flask app created", flush=True)
sys.stdout.flush()

@app.route('/')
def hello():
    return 'HELLO'

print("STEP 3: Route registered", flush=True)
sys.stdout.flush()

if __name__ == '__main__':
    port = int(os.getenv('PORT', 8000))
    print(f"STEP 4: Starting server on port {port}", flush=True)
    sys.stdout.flush()
    app.run(host='0.0.0.0', port=port, debug=False, use_reloader=False)
