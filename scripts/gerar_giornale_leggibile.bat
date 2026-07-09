@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0gerar_giornale_leggibile.ps1" %*
exit /b %ERRORLEVEL%
