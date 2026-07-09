#requires -Version 5.1
<#
    Launcher operativo: valida registrazioni manuali, genera CSV tecnico ed
    esegue la pipeline gerencial provvisoria tramite:
      .\.venv\Scripts\python.exe -m fantozzi.cli executar-lancamentos-manuais-gerenciais

    Guardrail: non esegue Python senza -ConfirmarExecucao. Prima della
    esecuzione confermata blocca se esistono modifiche tracked/staged, salvo
    -SkipGitCheck.
#>

[CmdletBinding()]
param(
    [switch]$ConfirmarExecucao,
    [switch]$SkipGitCheck,
    [string]$TemplatePath = "Database_Dev\Registrazioni_manuali\registrazioni_manuali_base.xlsx",
    [string]$CsvTecnico = "Database_Dev\Registrazioni_manuali\registrazioni_manuali_tecnico.csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-RepoPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $RepoRoot $Path))
}

function Test-PathEquals {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Left,
        [Parameter(Mandatory = $true)]
        [string]$Right
    )

    return [string]::Equals(
        [System.IO.Path]::GetFullPath($Left).TrimEnd('\'),
        [System.IO.Path]::GetFullPath($Right).TrimEnd('\'),
        [System.StringComparison]::OrdinalIgnoreCase
    )
}

function Invoke-GitGuard {
    if ($SkipGitCheck) {
        Write-Host "[AVISO] Git check ignorato per richiesta esplicita: -SkipGitCheck" -ForegroundColor Yellow
        return
    }

    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        Write-Host "[ERRO] git non disponibile. Launcher bloccato." -ForegroundColor Red
        Write-Host "[ERRO] Rieseguire in ambiente Git valido oppure usare -SkipGitCheck solo se autorizzato." -ForegroundColor Red
        exit 1
    }

    $TrackedDiff = @(& git diff --name-status 2>&1)
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERRO] git diff --name-status fallito. Launcher bloccato." -ForegroundColor Red
        $TrackedDiff | ForEach-Object { Write-Host $_ }
        exit 1
    }

    $StagedDiff = @(& git diff --cached --name-status 2>&1)
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERRO] git diff --cached --name-status fallito. Launcher bloccato." -ForegroundColor Red
        $StagedDiff | ForEach-Object { Write-Host $_ }
        exit 1
    }

    if (($TrackedDiff.Count -gt 0) -or ($StagedDiff.Count -gt 0)) {
        Write-Host "[ERRO] Existem modificacoes tracked/staged. Launcher bloqueado." -ForegroundColor Red
        if ($TrackedDiff.Count -gt 0) {
            Write-Host "[ERRO] Tracked modificati:"
            $TrackedDiff | ForEach-Object { Write-Host "  $_" }
        }
        if ($StagedDiff.Count -gt 0) {
            Write-Host "[ERRO] Staged modificati:"
            $StagedDiff | ForEach-Object { Write-Host "  $_" }
        }
        Write-Host "[ERRO] Chiudere o classificare le modifiche prima della run operativa." -ForegroundColor Red
        exit 1
    }

    $Porcelain = @(& git status --porcelain=v1 --untracked-files=all 2>&1)
    if ($LASTEXITCODE -eq 0) {
        $Untracked = @($Porcelain | Where-Object { $_.StartsWith("?? ") })
        if ($Untracked.Count -gt 0) {
            Write-Host "[AVISO] Untracked presenti (non bloccanti): $($Untracked.Count)" -ForegroundColor Yellow
            $Untracked | Select-Object -First 10 | ForEach-Object { Write-Host "  $_" }
            if ($Untracked.Count -gt 10) {
                Write-Host "  ... altri untracked non mostrati"
            }
        }
    }
}

$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location -Path $RepoRoot

$PythonExe = Join-Path $RepoRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -Path $PythonExe -PathType Leaf)) {
    Write-Host "[ERRO] Python del virtualenv non trovato: $PythonExe" -ForegroundColor Red
    Write-Host "[ERRO] Preparare l'ambiente (.venv) prima di eseguire questo launcher." -ForegroundColor Red
    exit 1
}

$OfficialTemplate = Join-Path $RepoRoot "Database_Dev\Registrazioni_manuali\registrazioni_manuali_base.xlsx"
$LegacyTemplate = Join-Path $RepoRoot "Database_Dev\_templates\lancamentos_manuais_gerenciais_TEMPLATE.xlsx"
$ResolvedTemplate = Resolve-RepoPath -Path $TemplatePath
$ResolvedCsv = Resolve-RepoPath -Path $CsvTecnico

if (Test-PathEquals -Left $ResolvedTemplate -Right $LegacyTemplate) {
    Write-Host "[ERRO] Template legacy non operativo richiesto: $TemplatePath" -ForegroundColor Red
    Write-Host "[ERRO] Usare il template ufficiale: Database_Dev\Registrazioni_manuali\registrazioni_manuali_base.xlsx" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path -Path $OfficialTemplate -PathType Leaf)) {
    Write-Host "[ERRO] Template ufficiale non trovato: $OfficialTemplate" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path -Path $ResolvedTemplate -PathType Leaf)) {
    Write-Host "[ERRO] Template richiesto non trovato: $TemplatePath" -ForegroundColor Red
    Write-Host "[ERRO] Percorso risolto: $ResolvedTemplate" -ForegroundColor Red
    exit 1
}

$CsvDir = Split-Path -Parent $ResolvedCsv
if (-not (Test-Path -Path $CsvDir -PathType Container)) {
    Write-Host "[ERRO] Cartella del CSV tecnico non trovata: $CsvDir" -ForegroundColor Red
    exit 1
}

Write-Host "[INFO] Repo root: $RepoRoot"
Write-Host "[INFO] Template ufficiale: $OfficialTemplate"
Write-Host "[INFO] Template richiesto: $ResolvedTemplate"
Write-Host "[INFO] CSV tecnico: $ResolvedCsv"
Write-Host "[INFO] Comando: executar-lancamentos-manuais-gerenciais"
Write-Host "[AVISO] Questo comando puo generare output contabili e aggiornare file tecnici della run manuale." -ForegroundColor Yellow

if (-not $ConfirmarExecucao) {
    Write-Host "[AVISO] Esecuzione non confermata. Python NON e stato avviato." -ForegroundColor Yellow
    Write-Host "[INFO] Per eseguire realmente: scripts\executar_lancamentos_manuais_gerenciais.ps1 -ConfirmarExecucao"
    Write-Host "[INFO] Exit code 2 = non eseguito per mancanza conferma."
    exit 2
}

Invoke-GitGuard

$PythonArgs = @(
    "-m", "fantozzi.cli", "executar-lancamentos-manuais-gerenciais",
    "--template", $ResolvedTemplate,
    "--saida-csv", $ResolvedCsv
)

& $PythonExe @PythonArgs
$ExitCode = $LASTEXITCODE

if ($ExitCode -eq 0) {
    Write-Host "[OK] Lancamentos manuais conclusi." -ForegroundColor Green
}
else {
    Write-Host "[ERRO] Lancamentos manuais terminati con codice $ExitCode." -ForegroundColor Red
}

exit $ExitCode
