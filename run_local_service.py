"""Runs Zewer CRM as an always-on local Windows service using waitress
(gunicorn doesn't work on native Windows). Intended to be started automatically
via Windows Task Scheduler so staff can just open a browser to this PC.

Requires DATABASE_URL to be set as a permanent environment variable (e.g. via
`setx DATABASE_URL "..."`), pointing at the SAME production database as the
Railway site — this local install is a second access point onto the live
data, not a separate copy, so scheduled backups and local document copies
reflect the real company records.
"""
import os
import sys
import logging

LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
os.makedirs(LOG_DIR, exist_ok=True)
logging.basicConfig(
    filename=os.path.join(LOG_DIR, 'local_service.log'),
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(message)s',
)

if not os.getenv('DATABASE_URL'):
    logging.error('DATABASE_URL is not set — refusing to start against a blank local database. '
                   'Set it with: setx DATABASE_URL "<value from Railway Variables tab>"')
    sys.exit(1)

from app import app
from waitress import serve

if __name__ == '__main__':
    port = int(os.getenv('LOCAL_PORT', 8000))
    logging.info(f'Starting Zewer CRM local service on http://127.0.0.1:{port}')
    serve(app, host='0.0.0.0', port=port)
