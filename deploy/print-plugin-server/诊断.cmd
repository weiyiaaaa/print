@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "server\diagnose.ps1" %*
echo.
pause
