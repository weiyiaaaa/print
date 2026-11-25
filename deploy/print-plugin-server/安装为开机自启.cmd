@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "server\install-task.ps1" %*
echo.
pause
