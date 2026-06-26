@echo off
setlocal

set "ROOT=%~dp0"
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%PS%" (
  echo Windows PowerShell was not found.
  pause
  exit /b 1
)

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%ROOT%start-web.ps1"
if errorlevel 1 (
  echo.
  echo start-web.ps1 failed. Check web-server.log in this folder.
  pause
  exit /b 1
)

endlocal
