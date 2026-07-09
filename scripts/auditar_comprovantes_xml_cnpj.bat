@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0auditar_comprovantes_xml_cnpj.ps1" %*
exit /b %ERRORLEVEL%
