# SFEPM Takeoff Wizard - GitHub Deploy Script
# Called automatically by build.bat after a successful build.
# Requires github_token.txt in the same folder.

param(
    [string]$Script,
    [string]$Version
)

$ErrorActionPreference = "Stop"

try {
    # Read token
    $tokenFile = Join-Path $PSScriptRoot "github_token.txt"
    if (-not (Test-Path $tokenFile)) {
        Write-Host "  github_token.txt not found - skipping deploy."
        exit 0
    }
    $token = (Get-Content $tokenFile -Raw).Trim()

    # Check git is available
    git --version 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  git not found - skipping deploy. Install Git from git-scm.com"
        exit 0
    }

    # Clone repo to temp folder
    $repoUrl = "https://$token@github.com/thecampbellm2/sfepm-wizard.git"
    $tempDir = Join-Path $env:TEMP "sfepm_deploy_$(Get-Random)"

    Write-Host "  Cloning repo..."
    git clone $repoUrl $tempDir --quiet 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "git clone failed" }

    # Copy new files into cloned repo
    Write-Host "  Copying files..."
    Copy-Item (Join-Path $PSScriptRoot "dist\SFEPM_Takeoff_Wizard.exe") (Join-Path $tempDir "SFEPM_Takeoff_Wizard.exe") -Force
    Copy-Item (Join-Path $PSScriptRoot $Script)         (Join-Path $tempDir $Script)         -Force
    Copy-Item (Join-Path $PSScriptRoot "version.txt")   (Join-Path $tempDir "version.txt")   -Force
    Copy-Item (Join-Path $PSScriptRoot "changelog.txt") (Join-Path $tempDir "changelog.txt") -Force
    Copy-Item (Join-Path $PSScriptRoot "build.bat")     (Join-Path $tempDir "build.bat")     -Force
    Copy-Item (Join-Path $PSScriptRoot "deploy.ps1")    (Join-Path $tempDir "deploy.ps1")    -Force

    # Remove old script versions from repo
    Get-ChildItem (Join-Path $tempDir "SFEPM_Takeoff_Wizard_v*.py") |
        Where-Object { $_.Name -ne $Script } |
        Remove-Item -Force

    # Commit and push
    Write-Host "  Committing and pushing v$Version..."
    Set-Location $tempDir
    git config user.email "build@sfepm.local"
    git config user.name  "SFEPM Build"
    git add -A
    git commit -m "Release v$Version" --quiet
    git push origin main --quiet 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "git push failed" }

    Write-Host "  GitHub deploy complete - v$Version is live."

} catch {
    Write-Host "  Deploy failed: $_"
    Write-Host "  Upload files to GitHub manually."
} finally {
    if ($tempDir -and (Test-Path $tempDir)) {
        Set-Location $env:TEMP
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}
