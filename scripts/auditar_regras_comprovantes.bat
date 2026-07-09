@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0auditar_regras_comprovantes.ps1" %*
exit /b %ERRORLEVEL%
