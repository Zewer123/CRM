@echo off
REM Removes the Zewer CRM always-on auto-start task and stops the running service.
echo Stopping Zewer CRM local service...
taskkill /F /IM pythonw.exe >nul 2>nul
echo Removing auto-start task...
schtasks /Delete /TN "ZewerCRM" /F >nul 2>nul
echo Done. Zewer CRM will no longer start automatically.
echo (Your data is untouched - it lives in the Railway database.)
pause
