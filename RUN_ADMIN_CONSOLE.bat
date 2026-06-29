@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PIH_ADMIN_PORT=18756"
if not "%~1"=="" set "PIH_ADMIN_PORT=%~1"

echo Starting Project Intelligence Hub Admin Console
echo Local admin URL: http://localhost:%PIH_ADMIN_PORT%
echo Main dashboard remains separate from this admin host.
echo.

python -m streamlit run admin_app.py --server.port=%PIH_ADMIN_PORT% --server.headless=false --browser.gatherUsageStats=false
exit /b %ERRORLEVEL%
