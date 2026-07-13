@echo off
setlocal enabledelayedexpansion
REM ============================================================
REM  Zewer AML CRM - one-click local launcher (Windows)
REM  Uses Python 3 explicitly, ignoring any old Python 2 on the PC.
REM ============================================================
cd /d "%~dp0"

REM --- Find a real Python 3, ignoring old Python 2 (e.g. BioTime) on PATH ---
set "PY="
py -3 --version >nul 2>nul
if not errorlevel 1 (
  set "PY=py -3"
) else (
  for /f "tokens=2" %%v in ('python --version 2^>^&1') do set "PYVER=%%v"
  echo !PYVER! | findstr /b "3." >nul 2>nul
  if not errorlevel 1 set "PY=python"
)

if not defined PY (
  echo.
  echo  Python 3 was not found on this computer.
  echo  ^(This PC may have an old Python 2 from other software, which will NOT work.^)
  echo.
  echo  Please install Python 3.11 or newer from:
  echo       https://www.python.org/downloads/
  echo  On the FIRST screen, TICK "Add python.exe to PATH", finish the
  echo  install, then run this file again.
  echo.
  pause
  exit /b 1
)

echo  Using Python 3:
%PY% --version
echo.

REM --- Create the isolated environment with Python 3 (first run only) ---
if not exist ".venv\Scripts\python.exe" (
  echo  Creating a local environment ^(first run only^)...
  %PY% -m venv .venv
)
if not exist ".venv\Scripts\python.exe" (
  echo.
  echo  ERROR: could not create the local environment. Make sure Python 3
  echo  installed correctly, then run this file again.
  pause
  exit /b 1
)

REM --- Install required components (first run only) ---
if not exist ".venv\.installed" (
  echo  Installing required components ^(first run only, ~2-3 minutes^)...
  ".venv\Scripts\python.exe" -m pip install --upgrade pip >nul 2>nul
  ".venv\Scripts\python.exe" -m pip install -q -r requirements.txt
  if errorlevel 1 (
    echo.
    echo  ERROR: could not install components. Check your internet connection
    echo  and run this file again.
    pause
    exit /b 1
  )
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

".venv\Scripts\python.exe" run_local.py
pause
