@echo off
cd /d "%~dp0"
rem Check invoice mails (read only). The result page opens automatically.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0system\Check-InvoiceMails.ps1" %*
pause
