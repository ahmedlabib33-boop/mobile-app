@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PORT=%PIH_MOBILE_APP_PORT%"
if "%PORT%"=="" set "PORT=18755"
set "PY=%cd%\.venv\Scripts\python.exe"

if not exist "%PY%" (
  echo [ERROR] Missing virtual environment interpreter: %PY%
  pause
  exit /b 1
)

for /f "tokens=5" %%P in ('netstat -ano ^| findstr /r /c:":%PORT% .*LISTENING"') do (
  taskkill /PID %%P /F >nul 2>nul
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%PY%' -ArgumentList @('-m','streamlit','run','dashboard.py','--server.address','127.0.0.1','--server.port','%PORT%') -WorkingDirectory '%cd%' -WindowStyle Hidden"

timeout /t 8 /nobreak >nul
powershell -NoProfile -Command "Start-Process 'http://127.0.0.1:%PORT%'"
exit /b 0
