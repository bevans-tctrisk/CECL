$ErrorActionPreference = "SilentlyContinue"

# Project root is the local clone on C:.  The repo used to live on Egnyte
# (Z:\Shared\TCT Files\CECL - CM Files) but git operations and Python imports
# off the network share were the dominant latency source.  GitHub is the
# source of truth; the C: clone is the working copy.
$projectRoot = "C:\Dev\CECL"

# Use the local C: virtual environment for fast startup.  Importing 60+
# packages from a network share added 15-30s and occasional PermissionErrors.
$localVenvPython = "C:\cecl-venv\Scripts\python.exe"
if (Test-Path $localVenvPython) {
    $pythonExe = $localVenvPython
} else {
    throw "C:\cecl-venv not found. Run: python -m venv C:\cecl-venv; C:\cecl-venv\Scripts\pip install -r '$projectRoot\requirements.txt'"
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
