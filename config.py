import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Flask Settings
DEBUG = os.getenv('FLASK_ENV') == 'development'
SECRET_KEY = os.getenv('SECRET_KEY', 'dev-key-change-this')

# Database Settings
DATABASE = {
    'host': os.getenv('DB_HOST', 'localhost'),
    'user': os.getenv('DB_USER', 'postgres'),
    'password': os.getenv('DB_PASSWORD', 'password'),
    'database': os.getenv('DB_NAME', 'aml_crm'),
    'port': os.getenv('DB_PORT', '5432')
}

# Session Settings
SESSION_COOKIE_SECURE = True  # Only send over HTTPS
SESSION_COOKIE_HTTPONLY = True  # JavaScript cannot access
SESSION_COOKIE_SAMESITE = 'Lax'  # Prevent CSRF attacks
PERMANENT_SESSION_LIFETIME = 3600  # 1 hour timeout

# Temporary test user (CHANGE THIS!)
TEST_USER = {
    'email': 'admin@aml-crm.com',
    'password': 'TempPass123!',  # This will be hashed in database
    'name': 'Administrator',
    'role': 'admin'
}
