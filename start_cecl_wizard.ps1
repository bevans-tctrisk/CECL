$ErrorActionPreference = "SilentlyContinue"

$projectRoot = "Z:\Shared\TCT Files\CECL - CM Files"

# Use the local C: virtual environment for fast startup.  Importing 60+
# packages from a network share (.venv on Z:) added 15-30s and occasional
# PermissionErrors.  Fall back to the network venv if the C: one is missing.
$localVenvPython = "C:\cecl-venv\Scripts\python.exe"
$networkVenvPython = Join-Path $projectRoot ".venv\Scripts\python.exe"
if (Test-Path $localVenvPython) {
    $pythonExe = $localVenvPython
} else {
    $pythonExe = $networkVenvPython
}

# Keep Python bytecode cache off the network share too.
$env:PYTHONPYCACHEPREFIX = "C:\cecl-cache"
if (-not (Test-Path $env:PYTHONPYCACHEPREFIX)) {
    New-Item -ItemType Directory -Path $env:PYTHONPYCACHEPREFIX -Force | Out-Null
}

$appUrl = "http://127.0.0.1:5000/setup/step/identity"

Set-Location $projectRoot

$alreadyRunning = $false
try {
    $listener = Get-NetTCPConnection -LocalPort 5000 -State Listen | Select-Object -First 1
    if ($listener) {
        $alreadyRunning = $true
    }
} catch {
    $alreadyRunning = $false
}

if (-not $alreadyRunning) {
    Start-Process -FilePath $pythonExe -ArgumentList "run_ui.py" -WorkingDirectory $projectRoot
}

Start-Process $appUrl
