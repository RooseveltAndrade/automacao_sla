$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$venvPy = Join-Path $root ".venv\Scripts\python.exe"
$logDir = Join-Path $root "logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $logDir ("main_" + $timestamp + ".log")

Push-Location $root
try {
	$pythonScript = Join-Path $root "main.py"
	$cmdCommand = 'cd /d "{0}" & "{1}" "{2}" >> "{3}" 2>&1' -f $root, $venvPy, $pythonScript, $logFile
	cmd.exe /d /c $cmdCommand
	exit $LASTEXITCODE
}
catch {
	$_ | Out-File -FilePath $logFile -Append
	throw
}
finally {
	Pop-Location
}
