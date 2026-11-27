@echo off
cd /d "%~dp0"
rem First time only: check that this PC can run the tool (read only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0system\Test-Environment.ps1" %*
pause
