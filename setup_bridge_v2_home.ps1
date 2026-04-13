$ErrorActionPreference = 'Stop'

$newPath = 'C:\Users\ferna\Syncthing\codex-bridge-v2'
$newId = 'codex-bridge-v2'
$cfg = Join-Path $env:LOCALAPPDATA 'Syncthing\config.xml'
$cfgBackup = Join-Path $env:LOCALAPPDATA ('config.backup-bridgev2-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.xml')

Copy-Item $cfg $cfgBackup -Force

if (-not (Test-Path $newPath)) {
    New-Item -ItemType Directory -Force -Path $newPath | Out-Null
}
foreach ($name in '.stfolder', 'commands', 'messages', 'state', 'tasks') {
    New-Item -ItemType Directory -Force -Path (Join-Path $newPath $name) | Out-Null
}

[xml]$xml = Get-Content $cfg
$existing = @($xml.configuration.folder) | Where-Object {
    $_.id -like 'codex-bridge*' -or
    $_.path -like 'C:\Users\ferna\Syncthing\codex-bridge*' -or
    $_.path -eq 'C:\Users\ferna\Desktop\codex-bridge'
}
foreach ($node in $existing) {
    [void]$xml.configuration.RemoveChild($node)
}

$folder = $xml.CreateElement('folder')
$folder.SetAttribute('id', $newId)
$folder.SetAttribute('label', 'codex-bridge')
$folder.SetAttribute('path', $newPath)
$folder.SetAttribute('type', 'sendreceive')
$folder.SetAttribute('rescanIntervalS', '60')
$folder.SetAttribute('fsWatcherEnabled', 'true')
$folder.SetAttribute('fsWatcherDelayS', '10')
$folder.SetAttribute('fsWatcherTimeoutS', '0')
$folder.SetAttribute('ignorePerms', 'false')
$folder.SetAttribute('autoNormalize', 'true')

$filesystemType = $xml.CreateElement('filesystemType')
$filesystemType.InnerText = 'basic'
[void]$folder.AppendChild($filesystemType)

foreach ($id in @(
    'RRS4MOX-DRVTPLE-VKITJOO-WKEUQH2-R5I6OFT-3YEZDJF-JVCXFUJ-MP3LNQR',
    '5RVVHDU-FRTTH66-DHYM25J-JDVOOTE-OAFXL6M-YGEDD3H-NQEFFPK-V5TMXQ5'
)) {
    $dev = $xml.CreateElement('device')
    $dev.SetAttribute('id', $id)
    $dev.SetAttribute('introducedBy', '')
    $enc = $xml.CreateElement('encryptionPassword')
    $enc.InnerText = ''
    [void]$dev.AppendChild($enc)
    [void]$folder.AppendChild($dev)
}

$min = $xml.CreateElement('minDiskFree')
$min.SetAttribute('unit', '%')
$min.InnerText = '1'
[void]$folder.AppendChild($min)

$versioning = $xml.CreateElement('versioning')
foreach ($pair in @{'cleanupIntervalS'='3600'; 'fsPath'=''; 'fsType'='basic'}.GetEnumerator()) {
    $n = $xml.CreateElement($pair.Key)
    $n.InnerText = $pair.Value
    [void]$versioning.AppendChild($n)
}
[void]$folder.AppendChild($versioning)

foreach ($pair in @{
    'copiers'='0'
    'pullerMaxPendingKiB'='0'
    'hashers'='0'
    'order'='random'
    'ignoreDelete'='false'
    'scanProgressIntervalS'='0'
    'pullerPauseS'='0'
    'pullerDelayS'='1'
    'maxConflicts'='10'
    'disableSparseFiles'='false'
    'paused'='false'
    'markerName'='.stfolder'
    'copyOwnershipFromParent'='false'
    'modTimeWindowS'='0'
    'maxConcurrentWrites'='16'
    'disableFsync'='false'
    'blockPullOrder'='standard'
    'copyRangeMethod'='standard'
    'caseSensitiveFS'='false'
    'junctionsAsDirs'='false'
    'syncOwnership'='false'
    'sendOwnership'='false'
    'syncXattrs'='false'
    'sendXattrs'='false'
}.GetEnumerator()) {
    $n = $xml.CreateElement($pair.Key)
    $n.InnerText = $pair.Value
    [void]$folder.AppendChild($n)
}

$xattrFilter = $xml.CreateElement('xattrFilter')
foreach ($pair in @{'maxSingleEntrySize'='1024'; 'maxTotalSize'='4096'}.GetEnumerator()) {
    $n = $xml.CreateElement($pair.Key)
    $n.InnerText = $pair.Value
    [void]$xattrFilter.AppendChild($n)
}
[void]$folder.AppendChild($xattrFilter)

[void]$xml.configuration.AppendChild($folder)
$xml.Save($cfg)

Get-Process syncthing -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

$exe = 'C:\Users\ferna\AppData\Local\Programs\Syncthing\syncthing.exe'
if (Test-Path $exe) {
    Start-Process $exe
}

Start-Sleep -Seconds 8

foreach ($bridgeCfg in @(
    'C:\Users\ferna\Desktop\IA APP\codex_bridge_config.json',
    'C:\Users\ferna\Downloads\ia-app\codex_bridge_config.json'
)) {
    if (Test-Path $bridgeCfg) {
        $json = Get-Content $bridgeCfg -Raw | ConvertFrom-Json
        $json.bridge_root = $newPath
        $json | ConvertTo-Json -Depth 10 | Set-Content $bridgeCfg -Encoding utf8
    }
}

Write-Output ('BRIDGE_ROOT=' + $newPath)
Write-Output ('FOLDER_ID=' + $newId)
