# Zewer AML CRM — Build & Operations Manual

The one document to reach for when something breaks or you need to remember how
this was built. Written for a non-developer owner.

---

## 1. What this is
A web-based **Anti-Money-Laundering (AML) compliance CRM** for a UAE jewellery
business (a "DPMS" — Dealer in Precious Metals & Stones, regulated under UAE AML
law / goAML). It replaces an Excel workflow: companies, individuals, beneficial
owners (UBOs), KYC/documents, risk assessments, DPMSR (goAML) transaction
tracking, alerts, tasks, and reports.

- **Live site:** https://crm-production-213f.up.railway.app
- **Source repo:** GitHub `Zewer123/CRM` (branch `master` is what's live)
- **The live code folder on your PC:** `CRM-live/` (ignore the other duplicate copies)

---

## 2. Architecture & tech stack
| Layer | Technology |
|---|---|
| Language / framework | Python 3.11+, **Flask** (one main file: `app.py`) |
| UI | Server-rendered **Jinja2** HTML templates + one CSS file (`static/style.css`) — dark charcoal + gold theme |
| Cloud database | **PostgreSQL** (hosted on Railway) |
| Local database | **SQLite** (a single file, `data/aml_crm.db`) — used only for local dev/testing |
| File storage | **Cloudinary** (uploaded documents) + optional local-folder copies |
| Cloud hosting | **Railway** (auto-deploys on git push), served by **gunicorn** |
| Local always-on hosting | **waitress** (via `run_local_service.py`) — gunicorn doesn't run on Windows |
| Excel export/import | **openpyxl** |

It is a **single-tenant** app: one deployment = one business's data.

---

## 3. The three ways it runs
1. **Production (cloud):** Railway runs `app.py` via gunicorn against the Railway
   Postgres DB. This is the live system everyone uses.
2. **Local always-on (office PC):** `run_local_service.py` (waitress) runs on an
   office PC and connects to the **same** Railway Postgres DB — a second window
   onto the *same live data*, giving a local disk for document copies + scheduled
   backups. Auto-starts at login (see `LOCAL_ALWAYS_ON.md`).
3. **Local developer test:** `python wsgi.py` with no `DATABASE_URL` → uses the
   local SQLite file. For trying changes safely, isolated from production.

> ⚠️ **The #1 safety rule for testing:** make sure `DATABASE_URL` is EMPTY before
> running a local test, or you'll be editing production. On this machine it has
> been set to prod — clear it in your test terminal (`$env:DATABASE_URL=""`).

---

## 4. Configuration (environment variables)
These are set in **Railway → your service → Variables** (for cloud) and via
`setx` on the office PC (for the local service). **Never** put them in the code.

| Variable | What it does | Required? |
|---|---|---|
| `SECRET_KEY` | Signs login sessions. **App refuses to boot without it.** | **Yes** |
| `DATABASE_URL` | Postgres connection string. If empty → local SQLite. | Cloud: yes |
| `CLOUDINARY_CLOUD_NAME` / `CLOUDINARY_API_KEY` / `CLOUDINARY_API_SECRET` | Document uploads | For docs |
| `SQLITE_PATH` | Override where the local SQLite file lives | Optional |
| `LOCAL_SERVICE` | Set to `1` by `run_local_service.py` → enables the daily backup scheduler | Local only |
| `LOCAL_PORT` | Port for the local service (default 8000) | Optional |

---

## 5. How to ship a change (deploy workflow)
1. Edit files in `CRM-live/`.
2. **Validate:** `python -m py_compile app.py` (silent = OK).
3. Commit: `git add -A && git commit -m "..."`.
4. **Push to master** (this auto-deploys to Railway in ~1–2 min):
   `git push origin <branch>:master`
   - The push needs your GitHub credentials. If a push hangs, it's waiting on
     auth — run it in *your own* terminal, not an automated one.
5. **Verify the deploy actually landed** (see §7).

---

## 6. Database & migrations — the important gotcha
- Schema changes are applied automatically on startup by `_run_migrations()` and
  `_pg_ensure_columns()` in `app.py`. New columns use `safe_alter(...)`; new
  tables are `CREATE TABLE IF NOT EXISTS`.
- **The trap that caused a past outage:** on PostgreSQL, if one migration
  statement fails, the whole transaction "aborts" and every *following* statement
  is silently skipped → a new column never gets created → every page that uses it
  returns **Internal Server Error**. The fix (already in the code):
  `_pg_ensure_columns()` adds new columns on a **fresh, autocommit** connection,
  which can't be poisoned this way. **When you add a new column in future, add it
  there too**, and always test against Postgres, not just SQLite.

---

## 7. Verifying a deploy
- **`/healthz`** (https://crm-production-213f.up.railway.app/healthz) returns
  `{"status":"ok"}` only if every expected DB column/table exists. Curl it after
  any deploy that changed the schema — no login needed.
- ⚠️ **`/healthz` is NOT proof the new *code* deployed** — it only checks the DB
  schema, which is the same whether old or new code is running. To confirm the
  actual deploy landed, check something code-specific (e.g. a new page/feature you
  just added is visible), or compare the live commit SHA on GitHub to what you
  pushed.

---

## 8. WHEN SOMETHING BREAKS — recovery playbook
**Golden rule: restore service first, diagnose second. Rolling back is safe.**

| Symptom | Likely cause | Fix |
|---|---|---|
| Every page = "Internal Server Error" right after a deploy | A migration didn't apply (missing column) — §6 | Check Railway **Deploy Logs** for `column "x" does not exist`. Then **roll back** (below). |
| App won't start at all | `SECRET_KEY` not set | Set it in Railway → Variables. |
| Local service won't start | `DATABASE_URL` not set on the PC | `setx DATABASE_URL "<value>"`, reopen terminal. |
| Login always fails | Wrong credentials / rate-limited (10 fails/15 min → temporary block) | Wait 15 min or check Login History. |
| A specific page errors but others work | A bug in that page's query/template | Read the Railway log traceback — it names the file + line. |

**To roll back to the last working version (restores the site in ~2 min):**
```
git revert --no-edit <bad-commit-sha>
git push origin <branch>:master
```
Find the bad SHA with `git log --oneline`. This creates a new commit that undoes
the change and redeploys the previous working state — non-destructive.

**Where to read errors:** Railway → your service → **Deployments → (latest) →
Deploy Logs**. The Python traceback there tells you the exact file and line.

---

## 9. Backups & restore
- **Manual:** Admin → Settings → **Download Full Backup** → one Excel of everything.
- **Scheduled (local install only):** Admin → Settings → **Scheduled Auto-Backup**
  → set a folder + daily time; the office PC writes a full Excel there every day.
- **Restore:** Admin → Settings → **Restore from Backup** → upload the Excel. It
  only *adds* missing records; it never overwrites existing ones or restores users/passwords.
- **Also back up the Railway database** itself from the Railway dashboard.
- **Rule: keep several dated backups off the main machine, and test that one opens.**

---

## 10. Security — what's built in, and what you must do
Built in: hashed passwords, parameterised SQL (blocks injection), mandatory
`SECRET_KEY`, XSS escaping, login rate-limiting, security headers, HttpOnly/Secure
cookies, 8-hour auto-logout, 15 MB upload cap, generic error messages, an action
password for edit/delete, edit/delete restricted to admin, per-role permissions,
and a login-history audit trail.

**You must still:**
1. 🔴 **Change the default passwords** (`admin@zewer.ae / Admin@123`,
   `compliance@zewer.ae / Compliance@123`) — they're in the source code until you do.
2. Keep `SECRET_KEY` and `DATABASE_URL` secret (never paste them in chats/screenshots).
3. Rotate the DB password if it's ever exposed.
4. Give each staff member their own login (never share); remove leavers same-day.

---

## 11. Module map (where things live in the app)
Companies · Individuals (Clients) · UBOs/Authorized Persons · DPMSR Report (goAML
transaction tracking) · Risk Assessment (now driven by an admin-editable
questionnaire — Settings → Risk Questionnaire) · Country Risk Scores · Alerts
(document expiry) · Health Check · Tasks + Regular Tasks · Zewer Docs (internal
documents) · Reports · Admin Panel (users, roles/permissions, dropdowns, action
password, backup, local paths) · Login History.

---

## 12. Key files
| File | What it is |
|---|---|
| `app.py` | The entire backend — routes, DB, migrations, logic (large single file) |
| `templates/*.html` | All pages (Jinja2). `base.html` = shared shell/nav |
| `static/style.css` | All styling |
| `wsgi.py` | Entry point (`python wsgi.py` for local dev) |
| `run_local_service.py` | Always-on local server (waitress) + backup scheduler |
| `requirements.txt` | Python dependencies |
| `Dockerfile` / `nixpacks.toml` | How Railway builds it |
| `INSTALL_LOCAL_SERVICE.bat` | One-time office-PC setup (auto-start) |
| `LOCAL_ALWAYS_ON.md` | Office-PC install guide |
| `SECURITY_GUIDE.md` | Plain-language security guide |

---

## 13. Known deferred items (not blockers, fix when convenient)
- `calculate_risk_score()` averages only *answered* questions, so a half-filled
  assessment can score "High" off one answer — consider requiring all questions.
- Minor dead code: `api_send_risk_email` stub; a KYC "expired" branch that never
  fires; company `pep` vs client `pep`/`pep_status` naming can disagree.
- Repo has duplicate old copies of the app — delete all but `CRM-live/`.

---

*Last updated: 2026-07-12. Keep this file with the code so it's always at hand.*
