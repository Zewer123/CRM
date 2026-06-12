# 🚀 AML CRM - SETUP GUIDE FOR NON-CODERS

This guide explains EVERYTHING in plain English. You don't need to understand coding!

---

## ✅ BEFORE YOU START - WHAT YOU NEED

Check you have these:
- ✓ GitHub account (you have this)
- ✓ Railway account (you have this)
- ✓ PostgreSQL (you'll install this)
- ✓ Python (you'll install this)
- ✓ These 9 files (I created them for you)

---

## PHASE 1️⃣: SETUP ON YOUR COMPUTER (Local Testing)

### Step 1: Install Python & PostgreSQL

**Windows:**
1. Go to https://www.python.org/downloads/
2. Click "Download Python 3.11"
3. Run installer
4. ✓ CHECK "Add Python to PATH"
5. Click Install

**For PostgreSQL:**
1. Go to https://www.postgresql.org/download/windows/
2. Download "EDB Installer"
3. Run installer
4. Remember the password you set (you'll need it!)
5. Port = 5432 (default)

**Mac:**
```
# Install using Homebrew
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
brew install python
brew install postgresql
```

**Linux (Ubuntu):**
```bash
sudo apt-get update
sudo apt-get install python3 python3-pip postgresql postgresql-contrib
```

---

### Step 2: Create Project Folder

Open your computer's file explorer:

```
C:\Users\YourName\   (Windows)
or
/Users/YourName/     (Mac)
or
/home/YourName/      (Linux)
```

Create a new folder called: **aml-crm-system**

Inside it, create these folders:
```
aml-crm-system/
├── templates/      (for HTML files)
└── static/         (for CSS and JavaScript files)
```

---

### Step 3: Copy All Files Into Folders

I've created 9 files for you. Copy each one to the right place:

**In main folder (aml-crm-system/):**
- app.py
- config.py
- requirements.txt
- database.sql
- .env.example
- .gitignore

**In templates/ folder:**
- login.html
- dashboard.html

**In static/ folder:**
- style.css
- script.js

Your folder should look like:
```
aml-crm-system/
├── app.py
├── config.py
├── requirements.txt
├── database.sql
├── .env.example
├── .gitignore
├── templates/
│   ├── login.html
│   └── dashboard.html
└── static/
    ├── style.css
    └── script.js
```

---

### Step 4: Create .env File (SECRET KEYS)

1. In the **aml-crm-system** folder, find **.env.example**
2. **Make a copy** of it
3. **Rename the copy** to **.env** (remove "example")
4. **Open .env** with Notepad (right-click → Open With)
5. Change these values:

```
FLASK_ENV=development
SECRET_KEY=my-super-secret-key-12345-change-this

DB_HOST=localhost
DB_PORT=5432
DB_USER=postgres
DB_PASSWORD=YOUR-POSTGRES-PASSWORD-HERE    ← Put the password you set during PostgreSQL install
DB_NAME=aml_crm

PORT=5000
```

6. **Save the file**

⚠️ **IMPORTANT:** Never share this .env file! It has your password!

---

### Step 5: Create Database

Open PostgreSQL Command Line (pgAdmin or psql):

**Windows (pgAdmin):**
- Search for "pgAdmin" in Start Menu
- Click "Tools" → "Query Tool"
- Copy all code from **database.sql**
- Paste into Query Tool
- Click "Execute" (▶️ button)

**Mac/Linux (Terminal):**
```bash
psql -U postgres

# Then type this:
CREATE DATABASE aml_crm;

# Exit with:
\q
```

Then run database.sql:
```bash
psql -U postgres -d aml_crm -f database.sql
```

---

### Step 6: Install Python Packages

**Windows (Command Prompt):**
```
cd C:\Users\YourName\aml-crm-system
pip install -r requirements.txt
```

**Mac/Linux (Terminal):**
```bash
cd ~/aml-crm-system
pip3 install -r requirements.txt
```

This will download and install:
- Flask (makes website)
- psycopg2 (talks to PostgreSQL)
- python-dotenv (reads .env)

---

### Step 7: Run The App Locally

**Windows:**
```
python app.py
```

**Mac/Linux:**
```bash
python3 app.py
```

You should see:
```
* Running on http://127.0.0.1:5000
```

✅ **Success!** Open your browser and go to:
```
http://localhost:5000
```

You should see the **login page**!

---

### Step 8: Test Login

Try logging in with:
```
Email:    admin@aml-crm.com
Password: TempPass123!
```

If it works, you'll see the **Dashboard** page with "Welcome to AML CRM" 🎉

---

## PHASE 2️⃣: UPLOAD TO GITHUB

### Step 1: Open Command Line

**Windows:** Press `Win + R`, type `cmd`, hit Enter

**Mac/Linux:** Open Terminal

### Step 2: Go to Your Project

```bash
cd C:\Users\YourName\aml-crm-system    (Windows)
cd ~/aml-crm-system                     (Mac/Linux)
```

### Step 3: Initialize Git

```bash
git init
```

### Step 4: Add All Files

```bash
git add .
```

### Step 5: Create First Commit

```bash
git commit -m "Initial commit - Login page"
```

### Step 6: Add Your GitHub Repository

Replace `YOUR-USERNAME` with your GitHub username:
Replace `YOUR-REPO-NAME` with your repository name:

```bash
git remote add origin https://github.com/YOUR-USERNAME/YOUR-REPO-NAME.git
```

### Step 7: Push to GitHub

```bash
git branch -M main
git push -u origin main
```

Enter your GitHub username and password (or use token).

✅ **Done!** Your code is now on GitHub!

---

## PHASE 3️⃣: DEPLOY TO RAILWAY

### Step 1: Go to Railway

1. Open https://railway.app
2. Click "Login" (with GitHub)
3. Click "New Project"
4. Click "Deploy from GitHub"

### Step 2: Connect GitHub

1. Select your GitHub repository
2. Railway will auto-detect it's Python
3. Click "Deploy"

### Step 3: Add PostgreSQL

In Railway dashboard:
1. Click "Add Service"
2. Select "PostgreSQL"
3. Railway creates database automatically

### Step 4: Set Environment Variables

In Railway:
1. Click "Variables"
2. Add these:

```
FLASK_ENV=production
SECRET_KEY=your-new-secret-key-12345
```

Railway automatically provides:
- `DATABASE_URL` (from PostgreSQL service)

### Step 5: Initialize Database on Railway

In Railway Console:
```bash
psql $DATABASE_URL < database.sql
```

### Step 6: Deploy!

Click "Deploy" button. Railway will:
1. Install Python packages
2. Start the Flask app
3. Give you a URL like: https://aml-crm-system-production.up.railway.app

✅ **LIVE!** Your app is now online!

---

## 🧪 TESTING YOUR LOGIN

### Local (Your Computer)
Open: http://localhost:5000

### Online (Railway)
Open: https://aml-crm-system-production.up.railway.app

Use these credentials:
```
Email:    admin@aml-crm.com
Password: TempPass123!
```

---

## 📝 FILE EXPLANATIONS (PLAIN ENGLISH)

| File | What It Does |
|------|-------------|
| **app.py** | The brain - handles login, database connection |
| **config.py** | Settings file - database location, secrets |
| **login.html** | What users see - email box, password box |
| **dashboard.html** | Welcome page after login |
| **style.css** | Pretty colors, fonts, buttons |
| **script.js** | Form interactions - click login button |
| **database.sql** | Creates filing cabinets (tables) in PostgreSQL |
| **requirements.txt** | Shopping list for Python packages |
| **.env** | Secret passwords (NEVER share!) |
| **.gitignore** | Tell GitHub what NOT to upload |

---

## ⚠️ TROUBLESHOOTING

### "Can't connect to database"
- Check PostgreSQL is running
- Check DB_PASSWORD in .env is correct
- Check database name is "aml_crm"

### "Port 5000 already in use"
```bash
# Change port in .env
PORT=5001
```

### "Login doesn't work"
- Check database.sql ran successfully
- Check PostgreSQL has the users table
- Check email/password are correct

### "GitHub push fails"
```bash
# Check if git is installed
git --version

# If not: https://git-scm.com/download
```

---

## 🚀 NEXT STEPS

1. **Test everything locally first**
2. **Push to GitHub**
3. **Deploy to Railway**
4. **Test the live URL**
5. **Create admin account** (optional - add more users)
6. **Next feature:** Company Master (adding companies)

---

## 💡 QUICK COMMANDS REFERENCE

```bash
# Start app locally
python app.py  (Windows)
python3 app.py (Mac/Linux)

# Install packages
pip install -r requirements.txt

# Upload to GitHub
git add .
git commit -m "Your message"
git push

# Connect to PostgreSQL
psql -U postgres

# Check if port is in use
netstat -ano | findstr :5000  (Windows)
lsof -i :5000  (Mac/Linux)
```

---

## 🎯 COMPLETION CHECKLIST

- [ ] Python installed
- [ ] PostgreSQL installed
- [ ] 9 files created
- [ ] .env file created with password
- [ ] Database created
- [ ] Login page works locally
- [ ] Files pushed to GitHub
- [ ] Railway deployed
- [ ] Live URL works
- [ ] Login works with test account

---

**You did it! 🎉 You now have a working login system!**

Next: We'll add the Company Master form.

Questions? Check the error messages - they usually explain the problem!
