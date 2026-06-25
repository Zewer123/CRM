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
echo   Open your web browser to:   http://localhost:8000
echo   Login:  admin@zewer.ae   /   Admin@123
echo.
echo   Keep this window open while using the system.
echo   To stop, close this window or press Ctrl+C.
echo ================================================================
echo.

set PORT=8000
python wsgi.py
pause
