@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0gerar_dashboard_html.ps1" %*
exit /b %ERRORLEVEL%
