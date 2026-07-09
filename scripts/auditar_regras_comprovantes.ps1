#requires -Version 5.1
<#
    Launcher read-only: audita config/regras_classificacao_comprovantes.yaml
    senza modificare alcun file. Wrapper diretto di:
      .\.venv\Scripts\python.exe -m fantozzi.cli auditar-regras-comprovantes
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location -Path $RepoRoot

$PythonExe = Join-Path $RepoRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -Path $PythonExe -PathType Leaf)) {
    Write-Host "[ERRO] Python del virtualenv non trovato: $PythonExe" -ForegroundColor Red
    Write-Host "[ERRO] Preparare l'ambiente (.venv) prima di eseguire questo launcher." -ForegroundColor Red
    exit 1
}

Write-Host "[INFO] Repo root: $RepoRoot"
Write-Host "[INFO] Comando: auditar-regras-comprovantes (read-only, non modifica YAML)"

& $PythonExe -m fantozzi.cli auditar-regras-comprovantes
$ExitCode = $LASTEXITCODE

if ($ExitCode -eq 0) {
    Write-Host "[OK] Audit concluso." -ForegroundColor Green
}
else {
    Write-Host "[ERRO] Audit terminato con codice $ExitCode." -ForegroundColor Red
}

exit $ExitCode
