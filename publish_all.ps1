param(
    [string]$Message = "",
    [switch]$PushTags = $true
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path ".git")) {
    throw "No .git folder found. Run this script from the repository root."
}

if ([string]::IsNullOrWhiteSpace($Message)) {
    $Message = "Sync all project files $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
}

Write-Host "Staging all repository changes..."
git add -A

$status = git status --porcelain
if (-not [string]::IsNullOrWhiteSpace($status)) {
    Write-Host "Creating commit..."
    git commit -m $Message
}
else {
    Write-Host "No file changes detected. Skipping commit."
}

Write-Host "Pushing main..."
git push origin main

if ($PushTags) {
    Write-Host "Pushing tags..."
    git push origin --tags
}

Write-Host "Publish complete."
git status --short --branch
