# Zewer AML CRM — Protection & Security Guide

Plain-language guide to **not losing data**, **not breaking the system**, and
**keeping it safe from misuse** (both insiders and outside attackers).

Work top-down: Tier 1 matters most.

---

## TIER 1 — Don't lose the data (most important)

A small business is hurt far more by *lost data* than by hackers.

1. **Back up the database file regularly.**
   - All data is in one file: `data\aml_crm.db`.
   - Copy it to a USB drive / another PC / cloud drive **daily or weekly**.
   - Keep several dated copies (e.g. `aml_crm_2026-06-25.db`). One backup that's
     also corrupted helps nobody.
2. **Before ANY change to the system, back up first.**
   - Zip the whole `CRM-live` folder (the database + code together).
3. **Test that a backup actually opens** once in a while — restore it on a spare PC.

> On the cloud version (Railway/PostgreSQL), turn on automatic database backups
> in the Railway dashboard. Same idea, different button.

---

## TIER 2 — Don't break it (change discipline)

Most "hacks" in small systems are actually *self-inflicted breakage*.

1. **One change at a time, then test.** Don't stack ten edits and hope.
2. **Always validate before deploying.** From the `CRM-live` folder:
   ```
   python -m py_compile app.py
   ```
   If that prints nothing, the Python is syntactically OK. If it errors, **stop**
   and fix before going further.
3. **Keep the working backup zip** from Tier 1 so you can always roll back.
4. **Never edit files directly on the live/production server.** Change locally,
   test, then deploy.
5. **Use Git** (the cloud copy already does). Every change is a commit you can undo.

---

## TIER 3 — Control who can do what (internal threats)

Insiders (staff, ex-staff) are the most common real risk for an office system.

1. **Change the default passwords immediately.**
   `admin@zewer.ae / Admin@123` and `compliance@zewer.ae / Compliance@123` are
   public knowledge — change them under **Admin Panel → Users** on day one.
2. **Give each person their own login.** Never share one account. The
   **Login History** page (Admin menu) only helps if logins are individual.
3. **Least privilege.** Give staff the *staff* role, not admin. Only one or two
   people should be admins.
4. **Remove leavers immediately.** When someone leaves, deactivate or delete
   their user the same day.
5. **Use the built-in audit tools:**
   - **Login History** (Admin menu) — watch for failed-login spikes or logins at
     odd hours / unknown devices.
   - **Submitted DPMSR reports are locked** to admins — staff can't quietly alter
     a filed report.
   - **Action password** protects sensitive company edits/deletes.
6. **Strong passwords** — at least 12 characters, not reused from email/banking.
7. **Lock the PC** when unattended (Windows key + L). The app auto-logs-out after
   8 hours, but an unlocked screen bypasses everything.

---

## TIER 4 — Outside attackers (only critical once it's on the internet)

While it runs only on the office PC / office network, external risk is low.
**Before** you ever expose it to the internet, do ALL of the following:

1. **Set a real secret key.** The app ships with a default
   (`zewer-aml-secret-2026`). Set your own private value via the `SECRET_KEY`
   environment variable. This key signs login sessions — if it's the public
   default, sessions can be forged.
2. **Don't use the built-in test server in production.** `START.bat` runs Flask's
   development server — fine for one office PC, **not** for the internet. Use a
   real server (gunicorn on Linux / waitress on Windows) behind **HTTPS**.
   The cloud (Railway) version already does this.
3. **Always use HTTPS** (the padlock). Passwords sent over plain `http` can be
   read on the network.
4. **Put it behind a firewall / VPN.** Ideally the system is only reachable from
   the office or via VPN, never open to the whole internet.
5. **Keep components updated.** Every few months:
   ```
   pip install -r requirements.txt --upgrade
   ```
   then test. Outdated libraries are the #1 external entry point.
6. **Limit login attempts** (rate-limiting) to slow password-guessing — ask your
   developer to add this before public exposure.

---

## What's already built in (your baseline)

You are not starting from zero. The system already:

- **Hashes all passwords** (never stored in plain text).
- **Uses parameterised database queries** — strong protection against the most
  common attack (SQL injection).
- **Marks session cookies HttpOnly + SameSite** — protects against common
  cookie theft / cross-site tricks.
- **Auto-logs-out** after 8 hours.
- **Records every login attempt** (success + failure, with IP and device).
- **Enforces roles on the server**, not just by hiding buttons (e.g. a staff
  member cannot delete a submitted DPMSR report even via a crafted request).

---

## Quick monthly checklist

- [ ] Database backed up and a copy stored off the PC
- [ ] Verified one backup actually opens
- [ ] No accounts for people who left
- [ ] Reviewed Login History for anything odd
- [ ] Everyone on their own account, admins kept to a minimum
- [ ] (If online) HTTPS working, libraries updated, secret key is private
