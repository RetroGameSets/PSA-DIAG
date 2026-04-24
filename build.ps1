# build.ps1 - Build helpers for PSA_DIAG
# Usage: Open a PowerShell prompt in the repository root and run:
#   .\build.ps1
# This script builds the main application and copies it to C:\INSTALL.

$ErrorActionPreference = 'Stop'

# Resolve repository root (script location)
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
Push-Location $ScriptRoot

try {
    Write-Host "[1/2] Building main application (PSA_DIAG)..." -ForegroundColor Cyan
    pyinstaller PSA_DIAG.spec
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller main build failed with exit code $LASTEXITCODE"
    }

    Write-Host "[2/2] Deploying PSA_DIAG.exe to C:\\INSTALL..." -ForegroundColor Cyan
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
