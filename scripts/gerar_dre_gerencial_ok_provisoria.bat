@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0gerar_dre_gerencial_ok_provisoria.ps1" %*
exit /b %ERRORLEVEL%
