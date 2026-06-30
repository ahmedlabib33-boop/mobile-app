@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo Project Intelligence Hub full-workspace 30-second no-Git watcher
echo This watches code, app files, and data files including CSV changes.
echo Target repository: ahmedlabib33-boop/mobile-app
echo.

call "%~dp0RUN_FULL_PROJECT_NO_GIT_SYNC.bat" Watch 30
exit /b %ERRORLEVEL%
