# One-shot push helper for the CECL repo.
#
# Before running for the first time:
#   1. Make sure https://github.com/bevans-tctrisk/CECL exists on GitHub
#      (create it as an EMPTY repo - no README, no .gitignore, no license).
#   2. Open a new PowerShell window so PortableGit is on PATH (it was
#      added to your USER PATH during initial setup).
#
# Then just run:
#   .\push_to_github.ps1
#
# On the very first push, Git Credential Manager will pop a browser
# window so you can sign in to GitHub. Credentials are cached after
# that, so subsequent pushes are silent.

$ErrorActionPreference = "Stop"
$ProgressPreference    = "SilentlyContinue"

# Make sure PortableGit is reachable even if this shell predates the
# PATH change.
$gitDir = Join-Path $env:LOCALAPPDATA "PortableGit\cmd"
if (Test-Path $gitDir) { $env:Path = "$env:Path;$gitDir" }

Set-Location -Path (Split-Path -Parent $MyInvocation.MyCommand.Path)

Write-Host ""
Write-Host "Repo state:" -ForegroundColor Cyan
git status --short
Write-Host ""
git log --oneline -5
Write-Host ""

$branch = git branch --show-current
Write-Host "Pushing branch '$branch' to origin..." -ForegroundColor Cyan
git push -u origin $branch
