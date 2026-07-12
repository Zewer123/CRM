@echo off
REM ============================================================
REM  Zewer AML CRM - one-click local launcher (Windows)
REM ============================================================
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo.
  echo  Python is not installed.
  echo  Please install Python 3.11 or newer from https://www.python.org/downloads/
  echo  During install, TICK "Add Python to PATH", then run this file again.
  echo.
  pause
  exit /b 1
)

if not exist ".venv" (
  echo Creating a local environment ^(first run only^)...
  python -m venv .venv
)

call .venv\Scripts\activate.bat

if not exist ".venv\.installed" (
  echo Installing required components ^(first run only^)...
  pip install -q -r requirements.txt
  echo done> ".venv\.installed"
)

echo.
echo ================================================================
echo   Zewer AML CRM is starting...
echo.
echo   A browser will open in a few seconds. If it does not, open:
echo       http://localhost:8000
echo   First login:  admin@zewer.ae   /   Admin@123
echo   ^(Change this password immediately in Admin - Settings - Users^)
echo.
echo   Keep this window open while using the system.
echo   To stop, close this window or press Ctrl+C.
echo ================================================================
echo.

REM Open the browser a few seconds after the server starts.
start "Zewer CRM" cmd /c "timeout /t 4 /nobreak >nul & start http://localhost:8000"

python run_local.py
pause
