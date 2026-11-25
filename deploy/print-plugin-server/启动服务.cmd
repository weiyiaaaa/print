@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "server\start.ps1" %*
echo.
pause
