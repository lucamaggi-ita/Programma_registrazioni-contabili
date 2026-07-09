@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0reprocessar_importacao_real.ps1" %*
exit /b %ERRORLEVEL%
