@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "MODE=Watch"
set "INTERVAL_SECONDS=30"
if not "%~1"=="" set "MODE=%~1"
if not "%~2"=="" set "INTERVAL_SECONDS=%~2"

echo Project Intelligence Hub mobile-app full-workspace no-Git synchronization
echo Mode: %MODE%  Interval: %INTERVAL_SECONDS% second(s)
echo Target and deletion policy are controlled by tools\github_sync_config.json
echo Target repository: ahmedlabib33-boop/mobile-app
echo Credentials are read only from PIH_MOBILE_APP_GITHUB_TOKEN or PIH_MOBILE_APP_GH_TOKEN.
echo Codespaces user secrets do not authenticate this local Windows process.
echo.

if /I "%MODE%"=="Watch" (
    echo Running immediate one-time sync before starting the 30-second watcher...
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\github_no_git_sync.ps1" -Mode Once -IntervalSeconds %INTERVAL_SECONDS%
    if errorlevel 1 exit /b %ERRORLEVEL%
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\github_no_git_sync.ps1" -Mode "%MODE%" -IntervalSeconds %INTERVAL_SECONDS%
exit /b %ERRORLEVEL%
