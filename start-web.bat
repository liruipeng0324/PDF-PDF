@echo off
setlocal

set "ROOT=%~dp0"
set "PYTHON=C:\Users\Administrator\AppData\Local\Python\bin\python.exe"

if not exist "%PYTHON%" (
  set "PYTHON=python"
)

for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":8765" ^| findstr "LISTENING"') do (
  taskkill /PID %%P /F >nul 2>nul
)

start "" /min "%PYTHON%" "%ROOT%web-server.py" --host 127.0.0.1 --port 8765
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:8765/"

endlocal
