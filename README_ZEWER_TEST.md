# Zewer AML CRM — Local Test Guide

A short guide to run the system on a Windows PC for testing.

## 1. One-time setup
1. Install **Python 3.11+** from https://www.python.org/downloads/
   - On the first install screen, **tick "Add Python to PATH"**.
2. Copy the whole `CRM-live` folder onto the PC (e.g. to the Desktop).

## 2. Start the system
- Double-click **`START.bat`**.
- The first run takes a minute (it sets things up). A black window will open and stay open — that's normal.
- Open a browser and go to: **http://localhost:8000**

## 3. Log in
| Role | Email | Password |
|------|-------|----------|
| Admin | `admin@zewer.ae` | `Admin@123` |
| Compliance | `compliance@zewer.ae` | `Compliance@123` |

**Change these passwords immediately** under Admin Panel → Users.

## 4. Stop the system
- Close the black window (or press `Ctrl + C` in it).

## Where is the data stored?
- All data lives in the file **`data\aml_crm.db`** inside this folder.
- **To back up:** copy that one file somewhere safe (a USB drive or another folder).
- **To restore:** put the backed-up file back as `data\aml_crm.db`.

## Notes for this test phase
- Document **upload to cloud is off** in the local version. Saving documents to a
  local drive is a planned next step once you're happy with the rest.
- This runs on the built-in test server — fine for one office PC on the local network.
  Before exposing it to the internet, see `SECURITY_GUIDE.md`.
