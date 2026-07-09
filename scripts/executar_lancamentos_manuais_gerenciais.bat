@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0executar_lancamentos_manuais_gerenciais.ps1" %*
exit /b %ERRORLEVEL%
