@echo off
setlocal

set "ROOT=%~dp0"
set "PYTHON=%ROOT%tools\Python313\python.exe"

if not exist "%PYTHON%" (
  set "PYTHON=C:\Users\Administrator\AppData\Local\Python\bin\python.exe"
)

if not exist "%PYTHON%" (
  where python >nul 2>nul
  if errorlevel 1 (
    echo 未找到 Python。请先运行 setup-dependencies.ps1 检查依赖。
    pause
    exit /b 1
  )
  set "PYTHON=python"
)

for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":8765" ^| findstr "LISTENING"') do (
  taskkill /PID %%P /F >nul 2>nul
)

start "" /min "%PYTHON%" "%ROOT%web-server.py" --host 127.0.0.1 --port 8765
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:8765/"

endlocal
