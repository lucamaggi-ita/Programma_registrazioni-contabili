@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0gerar_pdf_dre.ps1" %*
exit /b %ERRORLEVEL%
