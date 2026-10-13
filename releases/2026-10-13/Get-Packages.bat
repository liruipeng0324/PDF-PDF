@echo off
setlocal
cd /d "%~dp0"
echo 1. Manual OCR package
echo 2. Full package
echo 3. Both packages
choice /C 123 /M "Select"
set "PROFILE=manual-ocr"
if errorlevel 3 (set "PROFILE=all") else if errorlevel 2 (set "PROFILE=full")
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Get-Packages.ps1" -Profile %PROFILE%
set "RESULT=%ERRORLEVEL%"
echo Exit code: %RESULT%
pause
exit /b %RESULT%
