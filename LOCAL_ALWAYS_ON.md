# Zewer CRM — Always-On Local Install (office PC)

This makes the CRM run automatically in the background on one office PC, so staff
just open a browser. It uses **waitress** (a proper Windows web server) instead of
the test launcher, and connects to the **same live database** as the cloud site —
so it's a second window onto the real data, not a separate copy.

## One-time setup
1. Copy the `CRM-live` folder onto the office "server" PC.
2. Install **Python 3.11+** (tick *Add Python to PATH*).
3. Get your **DATABASE_URL** from Railway → your service → **Variables** tab.
4. Double-click **`INSTALL_LOCAL_SERVICE.bat`** and paste the DATABASE_URL when asked.
   - It creates the environment, installs components, stores the DB link + a secret
     key, registers a Windows auto-start task, and starts the service.
   - If it says it couldn't register the task, right-click → **Run as administrator**.

## Daily use
- Nothing to start — it launches automatically at logon and runs hidden.
- On the server PC: **http://localhost:8000**
- From other office PCs: **http://SERVER-PC-IP:8000** (find the IP with `ipconfig`).

## Managing it
- **Logs:** `logs\local_service.log`
- **Stop / remove auto-start:** run `UNINSTALL_LOCAL_SERVICE.bat`
- **Update:** replace the folder contents, re-run `INSTALL_LOCAL_SERVICE.bat`.

> Data lives in the Railway database, so uninstalling or reinstalling never loses it.
> Uploaded documents also go to cloud storage; set a **Local Document Storage** path
> in Admin → Settings to keep local copies too.
