# build.ps1 - Build helpers for PSA_DIAG
# Usage: Open a PowerShell prompt in the repository root and run:
#   .\build.ps1
# This script builds the updater in --onedir mode, deploys the entire folder to tools\updater,
# removes any old single-file updater, then builds the main application and copies it to C:\INSTALL.

$ErrorActionPreference = 'Stop'

# Resolve repository root (script location)
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
Push-Location $ScriptRoot

try {
    Write-Host "[1/4] Building updater (onedir)..." -ForegroundColor Cyan
    pyinstaller --onedir --windowed --name updater updater.py --noconfirm
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller updater build failed with exit code $LASTEXITCODE"
    }

    Write-Host "[2/4] Deploying updater to tools\updater..." -ForegroundColor Cyan
    $dest = Join-Path $ScriptRoot 'tools\updater'
    if (Test-Path $dest) {
        Write-Host "Removing existing $dest" -ForegroundColor Yellow
        Remove-Item -Recurse -Force $dest
    }
    New-Item -ItemType Directory -Path $dest | Out-Null
    Copy-Item -Path (Join-Path $ScriptRoot 'dist\updater\*') -Destination $dest -Recurse -Force

    # Optionally remove old single-file updater to avoid confusion
    $single = Join-Path $ScriptRoot 'tools\updater.exe'
    if (Test-Path $single) {
        Write-Host "Removing legacy single-file updater: $single" -ForegroundColor Yellow
        Remove-Item -Force $single -ErrorAction SilentlyContinue
    }

    Write-Host "[3/4] Building main application (PSA_DIAG)..." -ForegroundColor Cyan
    pyinstaller PSA_DIAG.spec
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller main build failed with exit code $LASTEXITCODE"
    }

    Write-Host "[4/4] Deploying PSA_DIAG.exe to C:\\INSTALL..." -ForegroundColor Cyan
    $installTarget = 'C:\\INSTALL'
    if (-not (Test-Path $installTarget)) { New-Item -ItemType Directory -Path $installTarget | Out-Null }
    Copy-Item -Path (Join-Path $ScriptRoot 'dist\PSA_DIAG.exe') -Destination (Join-Path $installTarget 'PSA_DIAG.exe') -Force

    Write-Host "Build and deploy complete." -ForegroundColor Green
} catch {
    Write-Error "Build failed: $_"
    exit 1
} finally {
    Pop-Location
}
