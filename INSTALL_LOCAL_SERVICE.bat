@echo off
REM ============================================================
REM  Zewer CRM - install as an ALWAYS-ON local service (Windows)
REM  Run this ONCE on the office PC. It:
REM    1. creates a private Python environment + installs components
REM    2. stores the database connection + secret key
REM    3. registers a scheduled task so the CRM starts automatically
REM       (hidden, no window) every time you log in
REM  After this, staff just open a browser to this PC.
REM ============================================================
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo(
echo ================================================================
echo   Zewer CRM - Always-On Local Service Setup
echo ================================================================
echo(

REM --- 1. Python present? ---
where python >nul 2>nul
if errorlevel 1 (
  echo  Python is not installed.
  echo  Install Python 3.11+ from https://www.python.org/downloads/
  echo  During install TICK "Add Python to PATH", then run this file again.
  pause
  exit /b 1
)

REM --- 2. Environment + components ---
if not exist ".venv" (
  echo Creating environment ^(first run only^)...
  python -m venv .venv
)
call ".venv\Scripts\activate.bat"
echo Installing / updating components...
pip install -q -r requirements.txt
if errorlevel 1 ( echo Component install failed. & pause & exit /b 1 )

REM --- 3. DATABASE_URL (points at the live Railway database) ---
if "%DATABASE_URL%"=="" (
  echo(
  echo  Paste the DATABASE_URL value from your Railway project
  echo  ^(Railway - your service - Variables tab - DATABASE_URL^).
  set /p DBURL="DATABASE_URL: "
  if "!DBURL!"=="" ( echo No value entered - aborting. & pause & exit /b 1 )
  setx DATABASE_URL "!DBURL!" >nul
  set "DATABASE_URL=!DBURL!"
) else (
  echo DATABASE_URL already set - keeping existing value.
)

REM --- 4. SECRET_KEY (generate one if not already set) ---
if "%SECRET_KEY%"=="" (
  for /f %%i in ('".venv\Scripts\python.exe" -c "import secrets;print(secrets.token_hex(32))"') do set "SKEY=%%i"
  setx SECRET_KEY "!SKEY!" >nul
  set "SECRET_KEY=!SKEY!"
  echo Generated a new SECRET_KEY for this PC.
) else (
  echo SECRET_KEY already set - keeping existing value.
)

REM --- 5. Auto-start at logon (hidden background service) ---
echo(
echo Registering auto-start task...
schtasks /Create /TN "ZewerCRM" /F /SC ONLOGON /RL HIGHEST ^
  /TR "\"%~dp0.venv\Scripts\pythonw.exe\" \"%~dp0run_local_service.py\"" >nul
if errorlevel 1 (
  echo  Could not register the scheduled task automatically.
  echo  Right-click this file and choose "Run as administrator", then retry.
  pause
  exit /b 1
)

REM --- 6. Start it now for this session ---
echo Starting Zewer CRM now...
start "" ".venv\Scripts\pythonw.exe" "run_local_service.py"

echo(
echo ================================================================
echo   Done! Zewer CRM is now running and will start automatically
echo   whenever this PC is logged in.
echo(
echo   Open a browser to:   http://localhost:8000
echo   From other office PCs: http://THIS-PC-IP:8000
echo(
echo   Logs:      logs\local_service.log
echo   To remove: run UNINSTALL_LOCAL_SERVICE.bat
echo ================================================================
pause
