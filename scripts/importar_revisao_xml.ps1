#requires -Version 5.1
<#
    Launcher operativo per:
      .\.venv\Scripts\python.exe -m fantozzi.cli importar-revisao-xml

    Guardrail: senza -ConfirmarExecucao stampa il riepilogo e non avvia
    Python. Con conferma blocca prima su diff tracked/staged, salvo
    -SkipGitCheck, verifica il comando via --help e poi passa gli argomenti
    residui alla CLI.
#>

[CmdletBinding()]
param(
    [switch]$ConfirmarExecucao,
    [switch]$SkipGitCheck,

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$CliArgs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$CliCommand = "importar-revisao-xml"
$LauncherPath = "scripts\importar_revisao_xml.ps1"

function Format-DisplayArgs {
    param([string[]]$ArgsToShow)

    if (-not $ArgsToShow -or $ArgsToShow.Count -eq 0) {
        return ""
    }

    $Formatted = $ArgsToShow | ForEach-Object {
        if ($_ -match '\s') {
            '"' + ($_ -replace '"', '\"') + '"'
        }
        else {
            $_
        }
    }

    return " " + ($Formatted -join " ")
}

function Invoke-GitGuard {
    if ($SkipGitCheck) {
        Write-Host "[AVISO] Git check ignorato per richiesta esplicita: -SkipGitCheck" -ForegroundColor Yellow
        return
    }

    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        Write-Host "[ERRO] git non disponibile. Launcher bloccato." -ForegroundColor Red
        Write-Host "[ERRO] Usare -SkipGitCheck solo come deroga esplicita e consapevole." -ForegroundColor Red
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
}

function Test-CliCommandAvailable {
    $HelpOutput = @(& $PythonExe -m fantozzi.cli $CliCommand --help 2>&1)
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERRO] Comando CLI non disponibile: $CliCommand" -ForegroundColor Red
        $HelpOutput | ForEach-Object { Write-Host $_ }
        exit 1
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

$CliArgs = @($CliArgs)
$DisplayArgs = Format-DisplayArgs -ArgsToShow $CliArgs

Write-Host "[INFO] Repo root: $RepoRoot"
Write-Host "[INFO] Comando: $CliCommand"
Write-Host "[INFO] Python: .\.venv\Scripts\python.exe -m fantozzi.cli $CliCommand$DisplayArgs"
Write-Host "[AVISO] Questo comando puo generare planilha decisoria XML e manifest operativi." -ForegroundColor Yellow

if (-not $ConfirmarExecucao) {
    Write-Host "[AVISO] Esecuzione non confermata. Python NON e stato avviato." -ForegroundColor Yellow
    Write-Host "[INFO] Per eseguire realmente: $LauncherPath -ConfirmarExecucao <argomenti CLI>"
    Write-Host "[INFO] Exit code 2 = non eseguito per mancanza conferma."
    exit 2
}

Invoke-GitGuard
Test-CliCommandAvailable

$PythonArgs = @("-m", "fantozzi.cli", $CliCommand) + $CliArgs
& $PythonExe @PythonArgs
$ExitCode = $LASTEXITCODE

if ($ExitCode -eq 0) {
    Write-Host "[OK] Comando concluso: $CliCommand" -ForegroundColor Green
}
else {
    Write-Host "[ERRO] Comando terminato con codice ${ExitCode}: $CliCommand" -ForegroundColor Red
}

exit $ExitCode
