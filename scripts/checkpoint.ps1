<#
.SYNOPSIS
  Dev checkpoint: stage tracked changes, verify staged Python compiles, commit, push.

.DESCRIPTION
  Keeps work backed up in small, regular increments so it never piles up into a
  big uncommitted mess. Stages tracked modifications/deletions (NOT untracked
  files, so investigation scaffolding / caches are never swept in); pass -Add to
  include specific new files. Refuses to commit if any staged .py fails to
  compile. Pushes to the current branch's upstream unless -NoPush.

.EXAMPLE
  ./scripts/checkpoint.ps1 -Message "PDF migration: Impr Deter page from data"

.EXAMPLE
  ./scripts/checkpoint.ps1 -Message "add helper" -Add cecl_report_web/new_thing.py
#>
param(
    [Parameter(Mandatory = $true)][string]$Message,
    [string[]]$Add = @(),   # extra (e.g. new/untracked) paths to include
    [switch]$NoPush
)
$ErrorActionPreference = 'Stop'
Set-Location (git rev-parse --show-toplevel)

git add -u
foreach ($p in $Add) { git add -- $p }

$staged = git diff --cached --name-only
if (-not $staged) { Write-Host "checkpoint: nothing staged, skipping."; exit 0 }

# Safety gate: never commit Python that doesn't even compile.
$pyStaged = @($staged | Where-Object { $_ -like '*.py' })
if ($pyStaged.Count -gt 0) {
    python -m py_compile @pyStaged
    if ($LASTEXITCODE -ne 0) {
        git reset -q
        Write-Error "checkpoint: staged Python failed to compile - aborting (unstaged)."
        exit 1
    }
}

Write-Host "checkpoint: committing $($staged.Count) file(s)..."
git commit -q -m $Message
if (-not $NoPush) {
    git push -q
    Write-Host "checkpoint: pushed."
}
git log -1 --oneline
