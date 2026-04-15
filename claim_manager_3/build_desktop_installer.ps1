param(
    [string]$Version = "3.0.1"
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$portableBuilder = Join-Path $PSScriptRoot "build_desktop_release.ps1"
$portableRoot = Join-Path $projectRoot "Releases\desktop_portable\Claim Manager 3 Portable"
$installerRoot = Join-Path $projectRoot "Releases\desktop_installer"
$issPath = Join-Path $PSScriptRoot "ClaimManager3Setup.iss"
$innoRoot = Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6"
$isccPath = Join-Path $innoRoot "ISCC.exe"

if (-not (Test-Path $isccPath)) {
    throw "Inno Setup compiler not found at $isccPath"
}

if (Test-Path $installerRoot) {
    Remove-Item -LiteralPath $installerRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $installerRoot | Out-Null

Write-Host "Building portable desktop package first..."
powershell -ExecutionPolicy Bypass -File $portableBuilder

if (-not (Test-Path $portableRoot)) {
    throw "Portable desktop package was not created at $portableRoot"
}

Write-Host "Building installer with Inno Setup..."
& $isccPath "/DAppVersion=$Version" "/DSourcePortableDir=$portableRoot" "/DOutputDir=$installerRoot" $issPath
if ($LASTEXITCODE -ne 0) {
    throw "Inno Setup failed while building the installer."
}

$installerPath = Join-Path $installerRoot "Claim-Manager-3-Setup-$Version.exe"
if (-not (Test-Path $installerPath)) {
    throw "Installer was expected at $installerPath but was not found."
}

Write-Host ""
Write-Host "Desktop installer ready:"
Write-Host "  $installerPath"
