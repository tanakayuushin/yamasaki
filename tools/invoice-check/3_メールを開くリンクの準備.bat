@echo off
cd /d "%~dp0"
rem First time only: enable the "Open in Outlook" links on the result page (current user only, no admin needed)
rem To remove: run with -Remove
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0system\Register-MailLink.ps1" %*
pause
