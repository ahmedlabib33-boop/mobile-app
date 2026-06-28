@echo off
cd /d "%~dp0"
echo Starting Excel and CSV GitHub sync watcher for ahmedlabib33-boop/mobile-app...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\watch_excel_and_push_to_mobile_app.ps1"
