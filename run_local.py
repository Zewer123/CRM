"""Runs Zewer CRM entirely on THIS computer, using a local database file.

Use this for an on-premise install (a shop's own PC/server). Unlike
run_local_service.py, it does NOT connect to any cloud database — all data
lives in data/aml_crm.db on this machine, and documents/backups go to local
folders you set in Admin -> Settings.

Other staff on the same office network can use it too: they just open
    http://<this-PC's-IP>:8000
in their browser while this window is running.

Started automatically by START.bat (double-click). Keep the window open (or
install it as an auto-start task) for the app to stay available.
"""
import os
import sys
import secrets
import socket

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 1) Force a purely local install: never fall through to a cloud database even
#    if a stray DATABASE_URL is present in the environment.
os.environ.pop('DATABASE_URL', None)

# 2) Keep the local database + secret beside the app so they back up together.
os.environ.setdefault('SQLITE_PATH', os.path.join(BASE_DIR, 'data', 'aml_crm.db'))

# 3) A per-install SECRET_KEY (signs login sessions). Generated once and stored
#    locally so logins survive restarts. This file is private — do not share it.
key_file = os.path.join(BASE_DIR, 'data', 'secret.key')
os.makedirs(os.path.dirname(key_file), exist_ok=True)
if not os.getenv('SECRET_KEY'):
    if os.path.exists(key_file):
        with open(key_file) as fh:
            os.environ['SECRET_KEY'] = fh.read().strip()
    if not os.getenv('SECRET_KEY'):
        new_key = secrets.token_hex(32)
        with open(key_file, 'w') as fh:
            fh.write(new_key)
        os.environ['SECRET_KEY'] = new_key

# 4) Enable the daily local auto-backup scheduler.
os.environ['LOCAL_SERVICE'] = '1'

from app import app  # noqa: E402  (env must be set before import)

def _lan_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return '127.0.0.1'

if __name__ == '__main__':
    port = int(os.getenv('LOCAL_PORT', 8000))
    ip = _lan_ip()
    print('=' * 60)
    print('  Zewer AML CRM is running.')
    print(f'  On THIS computer:      http://localhost:{port}')
    print(f'  From other office PCs: http://{ip}:{port}')
    print('  Data folder: ' + os.path.join(BASE_DIR, 'data'))
    print('  Keep this window open. Close it to stop the CRM.')
    print('=' * 60)
    try:
        from waitress import serve
        serve(app, host='0.0.0.0', port=port)
    except ImportError:
        # Fallback so it still runs if waitress isn't installed.
        app.run(host='0.0.0.0', port=port)
