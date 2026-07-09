#requires -Version 5.1
<#
    Launcher read-only: audita 2 livello comprovantes x XML/NF-e per CNPJ,
    scrivendo solo dentro -OutputDir. Wrapper parametrico di:
      .\.venv\Scripts\python.exe -m fantozzi.cli auditar-comprovantes-xml-cnpj

    -ComprovantesCsv e obbligatorio: non usa un default storico che potrebbe
    essere stantio. -OutputDir, se omesso, viene generato con timestamp per
    evitare di sovrascrivere audit precedenti.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ComprovantesCsv,

    [Parameter(Mandatory = $false)]
    [string]$XmlRoot = "Database_Dev\XML_fatture_acquisto",

    [Parameter(Mandatory = $false)]
    [string]$OutputDir,

    [Parameter(Mandatory = $false)]
    [int]$DiasTolerancia = 7,

    [switch]$SomenteRevisar
)

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

if (-not (Test-Path -Path $ComprovantesCsv -PathType Leaf)) {
    Write-Host "[ERRO] --comprovantes-csv non trovato: $ComprovantesCsv" -ForegroundColor Red
    Write-Host "[ERRO] Percorso interpretato relativo al repo root: $RepoRoot" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path -Path $XmlRoot -PathType Container)) {
    Write-Host "[ERRO] --xml-root non trovato: $XmlRoot" -ForegroundColor Red
    exit 1
}

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $OutputDir = "output\audit_xml_comprovantes_launcher_$Timestamp"
    Write-Host "[INFO] -OutputDir non specificato: uso $OutputDir"
}

if (Test-Path -Path $OutputDir) {
    Write-Host "[ERRO] -OutputDir gia esistente, non sovrascrivo silenziosamente: $OutputDir" -ForegroundColor Red
    Write-Host "[ERRO] Indicare -OutputDir diverso oppure rimuovere manualmente la cartella esistente." -ForegroundColor Red
    exit 1
}

Write-Host "[INFO] Repo root: $RepoRoot"
Write-Host "[INFO] Comando: auditar-comprovantes-xml-cnpj (read-only, scrive solo in -OutputDir)"
Write-Host "[INFO] Comprovantes CSV: $ComprovantesCsv"
Write-Host "[INFO] XML root: $XmlRoot"
Write-Host "[INFO] Dias tolerancia: $DiasTolerancia"
Write-Host "[INFO] Output dir: $OutputDir"

$PythonArgs = @(
    "-m", "fantozzi.cli", "auditar-comprovantes-xml-cnpj",
    "--comprovantes-csv", $ComprovantesCsv,
    "--xml-root", $XmlRoot,
    "--output-dir", $OutputDir,
    "--dias-tolerancia", $DiasTolerancia
)

if ($SomenteRevisar) {
    $PythonArgs += "--somente-revisar"
}

& $PythonExe @PythonArgs
$ExitCode = $LASTEXITCODE

if ($ExitCode -eq 0) {
    Write-Host "[OK] Audit concluso. Output: $OutputDir" -ForegroundColor Green
}
else {
    Write-Host "[ERRO] Audit terminato con codice $ExitCode." -ForegroundColor Red
}

exit $ExitCode
