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

# Code lives on C: (this clone) but analyst data (wizard_drafts/, Raw_Uploads/,
# Generated_Reports/, client_configs/, admin_defaults.yaml) stays on Egnyte
# so it's shared and backed up.  cecl_ui/app.py reads CECL_WORKSPACE_ROOT to
# decouple code location from data location.
$env:CECL_WORKSPACE_ROOT = "Z:\Shared\TCT Files\CECL - CM Files"

$appUrl = "http://127.0.0.1:5000/setup/step/identity"

Set-Location $projectRoot

# Always restart cleanly: if a previous Flask server is still bound to
# port 5000, kill it first so code changes actually take effect.  Otherwise
# the shortcut silently re-opens the browser against the stale server and
# new edits never appear to load.
try {
    $existingPids = Get-NetTCPConnection -LocalPort 5000 -State Listen -ErrorAction Stop |
        Select-Object -ExpandProperty OwningProcess -Unique
    foreach ($procId in $existingPids) {
        Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
    }
    if ($existingPids) {
        Start-Sleep -Milliseconds 500
    }
} catch {
    # No listener on port 5000 -- nothing to clean up.
}

# Also sweep any orphan run_ui.py python processes (e.g. debug-mode child
# from Flask's reloader that didn't release the port cleanly).
Get-CimInstance Win32_Process -Filter "Name='python.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -match 'run_ui\.py' } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }

Start-Process -FilePath $pythonExe -ArgumentList "run_ui.py" -WorkingDirectory $projectRoot

# Give Flask a couple seconds to bind the port before opening the browser.
Start-Sleep -Seconds 2
Start-Process $appUrl
