$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$claimManagerDir = (Resolve-Path $PSScriptRoot).Path
$releaseRoot = Join-Path $projectRoot "Releases\desktop_portable"
$buildRoot = Join-Path $projectRoot "build\desktop_portable"
$desktopUpdateBranch = "desktop-updates"
$desktopUpdateManifestUrl = "https://api.github.com/repos/fmarin27/ia-app/contents/claim_manager_3/desktop_update/latest.json?ref=$desktopUpdateBranch"

if (Test-Path $releaseRoot) {
    Remove-Item -LiteralPath $releaseRoot -Recurse -Force
}
if (Test-Path $buildRoot) {
    Remove-Item -LiteralPath $buildRoot -Recurse -Force
}

New-Item -ItemType Directory -Path $releaseRoot | Out-Null
New-Item -ItemType Directory -Path $buildRoot | Out-Null

$commonArgs = @(
    "--noconfirm",
    "--clean",
    "--windowed",
    "--onedir",
    "--distpath", $releaseRoot,
    "--workpath", $buildRoot,
    "--specpath", $buildRoot,
    "--paths", $claimManagerDir,
    "--add-data", "$claimManagerDir\ocr;claim_manager_3\ocr",
    "--add-data", "$claimManagerDir\refresh_scan_helper.py;claim_manager_3",
    "--add-data", "$claimManagerDir\autosource_page.png;claim_manager_3",
    "--add-data", "$claimManagerDir\README.md;claim_manager_3",
    "--add-data", "$claimManagerDir\_vendor;claim_manager_3\_vendor",
    "--collect-submodules", "docx",
    "--collect-submodules", "openpyxl",
    "--collect-submodules", "pypdf",
    "--collect-submodules", "reportlab"
)

Write-Host "Building Claim Manager 3 desktop..."
python -m PyInstaller @commonArgs --contents-directory "runtime-home" --name "Claim Manager 3" "$claimManagerDir\app.py"
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed while building Claim Manager 3."
}

Write-Host "Building Claim Manager 3 Office desktop..."
python -m PyInstaller @commonArgs --contents-directory "runtime-office" --name "Claim Manager 3 Office" "$claimManagerDir\officeversion.py"
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed while building Claim Manager 3 Office."
}

$portableRoot = Join-Path $releaseRoot "Claim Manager 3 Portable"
New-Item -ItemType Directory -Path $portableRoot | Out-Null
New-Item -ItemType Directory -Path (Join-Path $portableRoot "PENDING CLAIMS") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $portableRoot "Closed Claims") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $portableRoot "Claim Tools") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $portableRoot "Payroll") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $portableRoot "Office Updates") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $portableRoot "working_sheets") | Out-Null

'{}' | Set-Content -Path (Join-Path $portableRoot "claims_data.json") -Encoding UTF8
@{
    enabled = $true
    channel = "stable"
    manifest_url = $desktopUpdateManifestUrl
} | ConvertTo-Json | Set-Content -Path (Join-Path $portableRoot "desktop_update_config.json") -Encoding UTF8

Copy-Item -Path (Join-Path $releaseRoot "Claim Manager 3") -Destination $portableRoot -Recurse -Force
Copy-Item -Path (Join-Path $releaseRoot "Claim Manager 3 Office") -Destination $portableRoot -Recurse -Force

@'
@echo off
setlocal
set "APP_DIR=%~dp0"
set "CLAIM_MANAGER_HOME=%APP_DIR%"
pushd "%APP_DIR%"
"%APP_DIR%Claim Manager 3\Claim Manager 3.exe"
set "EXIT_CODE=%ERRORLEVEL%"
popd
if not "%EXIT_CODE%"=="0" (
  echo.
  echo Claim Manager 3 closed with exit code %EXIT_CODE%.
  pause
)
exit /b %EXIT_CODE%
'@ | Set-Content -Path (Join-Path $portableRoot "Start Claim Manager 3.bat") -Encoding ASCII

@'
@echo off
setlocal
set "APP_DIR=%~dp0"
set "CLAIM_MANAGER_HOME=%APP_DIR%"
pushd "%APP_DIR%"
"%APP_DIR%Claim Manager 3 Office\Claim Manager 3 Office.exe"
set "EXIT_CODE=%ERRORLEVEL%"
popd
if not "%EXIT_CODE%"=="0" (
  echo.
  echo Claim Manager 3 Office closed with exit code %EXIT_CODE%.
  pause
)
exit /b %EXIT_CODE%
'@ | Set-Content -Path (Join-Path $portableRoot "Start Claim Manager 3 Office.bat") -Encoding ASCII

@'
Claim Manager 3 Portable

What this package includes
- Claim Manager 3.exe
- Claim Manager 3 Office.exe
- packaged runtime folders for both desktop variants
- claims_data.json
- desktop_update_config.json
- default folders for PENDING CLAIMS, Closed Claims, Claim Tools, Payroll, Office Updates, and working_sheets

Suggested Saturday install
1. Copy this whole folder to the supervisor PC, for example C:\Claim Manager 3 Portable
2. Put the real claim folders or shortcuts in the included PENDING CLAIMS and Closed Claims folders, or update the watched folders inside Settings after launch.
3. Launch Start Claim Manager 3.bat
4. Confirm email settings, watched folders, and route home address
5. Test with one real claim
'@ | Set-Content -Path (Join-Path $portableRoot "README.txt") -Encoding ASCII

Compress-Archive -Path (Join-Path $portableRoot "*") -DestinationPath (Join-Path $releaseRoot "Claim-Manager-3-Portable.zip") -Force

Write-Host ""
Write-Host "Portable desktop package ready:"
Write-Host "  $portableRoot"
Write-Host "Zip package:"
Write-Host "  $(Join-Path $releaseRoot 'Claim-Manager-3-Portable.zip')"
