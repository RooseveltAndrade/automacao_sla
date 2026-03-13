$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$venvPy = Join-Path $root ".venv\Scripts\python.exe"
$logDir = Join-Path $root "logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $logDir ("main_" + $timestamp + ".log")

& $venvPy (Join-Path $root "main.py") *>> $logFile
