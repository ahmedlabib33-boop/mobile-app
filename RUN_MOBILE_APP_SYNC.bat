@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "MODE=Watch"
set "INTERVAL_SECONDS=30"
if not "%~1"=="" set "MODE=%~1"
if not "%~2"=="" set "INTERVAL_SECONDS=%~2"

echo Project Intelligence Hub mobile-app no-Git synchronization
echo Target repository: ahmedlabib33-boop/mobile-app
echo Mode: %MODE%  Interval: %INTERVAL_SECONDS% second(s)
echo Credentials are read only from PIH_MOBILE_APP_GITHUB_TOKEN or PIH_MOBILE_APP_GH_TOKEN.
echo Create a new repository token first, then run this file.
echo.

if /I "%MODE%"=="Watch" (
    echo Running immediate one-time sync before starting the watcher...
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\github_no_git_sync.ps1" -Mode Once -IntervalSeconds %INTERVAL_SECONDS%
    if errorlevel 1 exit /b %ERRORLEVEL%
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\github_no_git_sync.ps1" -Mode "%MODE%" -IntervalSeconds %INTERVAL_SECONDS%
exit /b %ERRORLEVEL%
