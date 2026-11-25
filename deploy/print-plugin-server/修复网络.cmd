@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "server\fix-net.ps1" %*
echo.
pause
