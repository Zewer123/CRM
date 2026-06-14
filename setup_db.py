import sqlite3
from werkzeug.security import generate_password_hash
import os
import json

DB = 'aml_crm.db'

def setup_database():
    """Create database with complete schema for 47 company fields + 23 UBO fields"""
    
    if os.path.exists(DB):
        os.remove(DB)
        print("✅ Old database removed")
    
    conn = sqlite3.connect(DB)
    conn.execute('PRAGMA foreign_keys = ON')
    cursor = conn.cursor()
    
    # Users table
    cursor.execute('''
        CREATE TABLE users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            name TEXT NOT NULL,
            role TEXT DEFAULT 'staff',
            is_active INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Companies table (47 fields from Excel)
    cursor.execute('''
        CREATE TABLE companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            
            -- Basic Info
            ac_code TEXT UNIQUE NOT NULL,
            client_name TEXT NOT NULL,
            customer_name TEXT,
            ac_opening_date DATE,
            
            -- Status & Type
            ac_status TEXT DEFAULT 'Active',
            active_till_year TEXT,
            nature TEXT,
            type_of_client TEXT,
            name_of_freezone TEXT,
            mode_of_ac TEXT,
            
            -- Location & Contact
            country_of_incorporation TEXT,
            region TEXT,
            address TEXT,
            telephone TEXT,
            mobile TEXT,
            whatsapp_link TEXT,
            email_id TEXT,
            
            -- Address Proof / Ejari
            address_proof_type TEXT,
            address_proof_expiry DATE,
            address_proof_days_left INTEGER,
            address_proof_status TEXT,
            
            -- KYC
            kyc_status TEXT,
            
            -- Trade License
            trade_license_no TEXT,
            issuing_authority TEXT,
            legal_type TEXT,
            incorporation_date DATE,
            trade_license_expiry DATE,
            trade_license_days_left INTEGER,
            trade_license_valid TEXT,
            
            -- Tax & VAT
            tax_no_trn TEXT,
            vat_cert TEXT,
            vat_declaration TEXT,
            vat_declaration_date DATE,
            
            -- Beneficial Owners
            num_beneficial_owners INTEGER DEFAULT 0,
            
            -- Compliance Docs
            moa TEXT,
            pep TEXT,
            undertaking TEXT,
            source_of_fund TEXT,
            software_updation TEXT,
            
            -- Compliance Status
            doc_status TEXT DEFAULT 'Incomplete',
            screening_date DATE,
            screening_tool_registered TEXT,
            risk_status TEXT DEFAULT 'Unspecified',
            verified_by TEXT,
            verified_date DATE,
            
            -- Notes
            followup_details TEXT,
            crowe_feedback TEXT,
            zewer_comments TEXT,
            
            -- Metadata
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            
            FOREIGN KEY(created_by) REFERENCES users(id)
        )
    ''')
    
    # UBOs table (23 fields from Excel)
    cursor.execute('''
        CREATE TABLE ubos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            
            -- Company Link
            ac_code TEXT,
            client_name TEXT,
            
            -- UBO Info
            position TEXT,
            share_percentage DECIMAL(5,2),
            person_name TEXT NOT NULL,
            nationality TEXT,
            residential_status TEXT,
            group_of_companies TEXT,
            
            -- Passport
            passport_no TEXT,
            passport_expiry DATE,
            passport_days_left INTEGER,
            passport_status TEXT,
            
            -- Emirates ID
            emirates_id TEXT,
            emirates_id_expiry DATE,
            emirates_id_days_left INTEGER,
            emirates_id_status TEXT,
            
            -- Compliance
            doc_status TEXT DEFAULT 'Incomplete',
            verified_by TEXT,
            verified_date DATE,
            followup_details TEXT,
            
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            
            FOREIGN KEY(company_id) REFERENCES companies(id) ON DELETE CASCADE
        )
    ''')
    
    # Document files table
    cursor.execute('''
        CREATE TABLE document_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER,
            ubo_id INTEGER,
            doc_type TEXT NOT NULL,
            file_name TEXT,
            file_path TEXT,
            uploaded_by INTEGER,
            uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            
            FOREIGN KEY(company_id) REFERENCES companies(id) ON DELETE CASCADE,
            FOREIGN KEY(ubo_id) REFERENCES ubos(id) ON DELETE CASCADE,
            FOREIGN KEY(uploaded_by) REFERENCES users(id)
        )
    ''')
    
    # Followup notes table
    cursor.execute('''
        CREATE TABLE followup_notes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER,
            ubo_id INTEGER,
            note_type TEXT DEFAULT 'general',
            note_text TEXT NOT NULL,
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            
            FOREIGN KEY(company_id) REFERENCES companies(id) ON DELETE CASCADE,
            FOREIGN KEY(ubo_id) REFERENCES ubos(id) ON DELETE CASCADE,
            FOREIGN KEY(created_by) REFERENCES users(id)
        )
    ''')
    
    # Dropdowns table (for dynamic management in settings)
    cursor.execute('''
        CREATE TABLE dropdowns (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            field_name TEXT NOT NULL,
            value TEXT NOT NULL,
            is_active INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            
            UNIQUE(field_name, value)
        )
    ''')
    
    # Insert admin user
    admin_password = generate_password_hash('Admin@123')
    cursor.execute('''
        INSERT INTO users (email, password_hash, name, role)
        VALUES (?, ?, ?, ?)
    ''', ('admin@zewer.ae', admin_password, 'Administrator', 'admin'))
    
    # Insert compliance user
    compliance_password = generate_password_hash('Compliance@123')
    cursor.execute('''
        INSERT INTO users (email, password_hash, name, role)
        VALUES (?, ?, ?, ?)
    ''', ('compliance@zewer.ae', compliance_password, 'Compliance Officer', 'compliance'))
    
    # Insert dropdown values from Excel
    dropdowns_data = {
        'AC STATUS': ['Active', 'Inactive'],
        'NATURE': ['Individual', 'Legal entity'],
        'TYPE OF CLIENT': ['MainLand', 'Free Zone', 'Abroad', 'International Corporate'],
        'MODE OF AC': ['Supplier', 'Customer', 'Bullion', 'Refinery', 'Logistics Co'],
        'RISK STATUS': ['High', 'Medium', 'Low', 'Unspecified'],
        'DOC STATUS': ['Completed', 'Incompleted'],
        'KYC STATUS': ['Yes', 'No', 'Pending'],
        'ADDRESS PROOF TYPE': ['Ejari', 'Utility Bill', 'Other'],
        'LEGAL TYPE': ['Individual', 'Partnership', 'LLC', 'Corporation', 'Trust', 'Foundation'],
        'VAT CERT': ['Yes', 'No'],
        'VAT DECLARATION': ['Yes', 'No'],
        'COUNTRY': ['UAE', 'Saudi Arabia', 'Kuwait', 'Qatar', 'Bahrain', 'Oman', 'Other'],
        'REGION': ['Dubai', 'Abu Dhabi', 'Sharjah', 'Ajman', 'Umm Al Quwain', 'Ras Al Khaimah', 'Fujairah'],
        'FREEZONE': ['Jafza', 'DMCC', 'DAFZA', 'RAKEZ', 'ICAD', 'Other']
    }
    
    for field, values in dropdowns_data.items():
        for value in values:
            cursor.execute('''
                INSERT INTO dropdowns (field_name, value, is_active)
                VALUES (?, ?, 1)
            ''', (field, value))
    
    conn.commit()
    conn.close()
    
    print("✅ Database created successfully!")
    print("✅ Admin user: admin@zewer.ae / Admin@123")
    print("✅ Compliance user: compliance@zewer.ae / Compliance@123")
    print("✅ All 47 company fields configured")
    print("✅ All 23 UBO fields configured")
    print("✅ Dropdowns initialized from Excel data")

if __name__ == '__main__':
    setup_database()
