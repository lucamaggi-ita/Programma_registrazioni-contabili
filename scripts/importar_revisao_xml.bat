@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0importar_revisao_xml.ps1" %*
exit /b %ERRORLEVEL%
