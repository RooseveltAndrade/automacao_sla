param(
    [string]$TaskName = "AutomacaoSlaMensal_10h"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$runScript = Join-Path $root "scripts\run_main.ps1"

$taskCmd = "powershell -NoProfile -ExecutionPolicy Bypass -File `"$runScript`""

schtasks /Create /F /TN $TaskName /SC MONTHLY /D 5 /ST 10:00 /TR $taskCmd

Write-Host "Tarefa criada: $TaskName"
