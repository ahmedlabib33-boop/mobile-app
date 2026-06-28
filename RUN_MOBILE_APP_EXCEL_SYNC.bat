@echo off
cd /d "%~dp0"
echo Starting Excel and CSV GitHub sync watcher for ahmedlabib33-boop/mobile-app...
echo Credentials are read only from PIH_MOBILE_APP_GITHUB_TOKEN or PIH_MOBILE_APP_GH_TOKEN.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\watch_excel_and_push_to_mobile_app.ps1"
