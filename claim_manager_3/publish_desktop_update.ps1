param(
    [string]$Version = "3.0.1",
    [string]$Repo = "fmarin27/ia-app",
    [string]$ReleaseTag = "",
    [string]$ReleaseTitle = ""
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

if (-not $ReleaseTag) {
    $ReleaseTag = "desktop-v$Version"
}
if (-not $ReleaseTitle) {
    $ReleaseTitle = "Claim Manager 3 Desktop $Version"
}

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$installerBuilder = Join-Path $PSScriptRoot "build_desktop_installer.ps1"
$manifestPath = Join-Path $PSScriptRoot "desktop_update\latest.json"
$installerPath = Join-Path $projectRoot "Releases\desktop_installer\Claim-Manager-3-Setup-$Version.exe"
$desktopFeedDir = Join-Path $projectRoot "Releases\desktop_update_feed"
$targetCommit = (git -C $projectRoot rev-parse HEAD).Trim()

Write-Host "Building desktop installer..."
powershell -ExecutionPolicy Bypass -File $installerBuilder -Version $Version

if (-not (Test-Path $installerPath)) {
    throw "Installer was expected at $installerPath but was not found."
}

& gh release view $ReleaseTag --repo $Repo > $null 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "Creating GitHub release $ReleaseTag..."
    gh release create $ReleaseTag $installerPath --repo $Repo --title $ReleaseTitle --notes "Claim Manager 3 desktop installer $Version" --target $targetCommit --prerelease
} else {
    Write-Host "Uploading installer to existing release $ReleaseTag..."
    gh release upload $ReleaseTag $installerPath --repo $Repo --clobber
}

$installerFileName = Split-Path $installerPath -Leaf
$installerUrl = "https://github.com/fmarin27/ia-app/releases/download/$ReleaseTag/$installerFileName"
$manifest = [ordered]@{
    version = $Version
    installer_file = $installerFileName
    published_at = (Get-Date).ToString("o")
    notes = "Desktop installer update $Version"
}
$manifest | ConvertTo-Json | Set-Content -Path $manifestPath -Encoding UTF8
if (Test-Path $desktopFeedDir) {
    Remove-Item -LiteralPath $desktopFeedDir -Recurse -Force
}
New-Item -ItemType Directory -Path $desktopFeedDir | Out-Null
Copy-Item -LiteralPath $installerPath -Destination (Join-Path $desktopFeedDir $installerFileName) -Force
Copy-Item -LiteralPath $manifestPath -Destination (Join-Path $desktopFeedDir "latest.json") -Force

Write-Host ""
Write-Host "Desktop update manifest refreshed:"
Write-Host "  $manifestPath"
Write-Host "Desktop update feed:"
Write-Host "  $desktopFeedDir"
Write-Host "GitHub backup installer URL:"
Write-Host "  $installerUrl"
