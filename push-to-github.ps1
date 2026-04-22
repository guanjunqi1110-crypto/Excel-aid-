# Run this in PowerShell from this folder AFTER you create an empty repo on GitHub.
# Usage: .\push-to-github.ps1 -GitHubUrl "https://github.com/YOUR_USERNAME/excel-readiness-ai-coach.git"
param(
    [Parameter(Mandatory = $true)]
    [string] $GitHubUrl
)

$ErrorActionPreference = "Stop"
$here = $PSScriptRoot
Set-Location $here

function Find-Git {
    $candidates = @(
        "git",
        "C:\Program Files\Git\bin\git.exe",
        "C:\Program Files (x86)\Git\bin\git.exe"
    )
    foreach ($c in $candidates) {
        if ($c -eq "git") {
            $g = Get-Command git -ErrorAction SilentlyContinue
            if ($g) { return $g.Source }
        }
        elseif (Test-Path $c) { return $c }
    }
    return $null
}

$git = Find-Git
if (-not $git) {
    Write-Host "Git not found. Install from: https://git-scm.com/download/win"
    Write-Host "Then reopen PowerShell and run this script again."
    exit 1
}

Write-Host "Using: $git"

# One-time identity (change if you want)
& $git config user.email 2>$null
if ($LASTEXITCODE -ne 0) {
    & $git config --global user.email "your.email@university.edu"
    & $git config --global user.name "Your Name"
}

if (-not (Test-Path ".git")) {
    & $git init
    & $git branch -M main
}

& $git add app.py requirements.txt .gitignore Excel_Readiness_AI_Coach_Content_Pack.xlsx README.md push-to-github.ps1
if (Test-Path ".streamlit\config.toml") { & $git add .streamlit\config.toml }
& $git status
& $git commit -m "Excel Readiness AI Coach - initial commit" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Nothing to commit or commit failed (maybe already committed)."
}

$hasRemote = & $git remote 2>$null
if ($hasRemote -match "origin") {
    & $git remote set-url origin $GitHubUrl
} else {
    & $git remote add origin $GitHubUrl
}

Write-Host "Pushing to $GitHubUrl ..."
& $git push -u origin main
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "If the push failed: create an EMPTY repo on GitHub (no README), then run this script again."
    Write-Host "Or use: & '$git' push -u origin main"
    exit 1
}
Write-Host "Done. Next: open https://streamlit.io/cloud and deploy this repository (main file: app.py)."
