@echo off
cd /d "%~dp0"
rem Rename files for real
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Rename-ScannedDocs.ps1" -Apply %*
pause
