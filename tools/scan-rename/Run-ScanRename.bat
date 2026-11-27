@echo off
cd /d "%~dp0"
rem Preview only (no files are changed)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Rename-ScannedDocs.ps1" %*
pause
