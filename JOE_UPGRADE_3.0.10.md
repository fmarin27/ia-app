# Joe Upgrade 3.0.10 / Mobile 1.0.6

## Release artifacts

- Desktop installer: https://github.com/fmarin27/ia-app/releases/download/desktop-v3.0.10/Claim-Manager-3-Setup-3.0.10.exe
- Joe Android install page: https://expo.dev/accounts/fmarin27/projects/claim-manager-mobile/builds/30b0b161-3d7c-4a0d-a0cd-4b8300c27361
- Local Joe APK copy: `C:\Users\ferna\Desktop\IA APP\Releases\mobile_joe\Claim-Manager-Mobile-Joe-1.0.6-build-30b0b161.apk`
- Source branch: `codex/joes-apps`

## What Joe should do

1. Close Claim Manager 3 if it is open.
2. Install `Claim-Manager-3-Setup-3.0.10.exe` over the existing install. Do not uninstall first.
3. Start Claim Manager 3 and confirm Settings shows desktop version `3.0.10`.
4. Install the Joe Android build from the Expo link above, or from the APK if it is copied to his phone.
5. In the mobile app, confirm the footer says app version `1.0.6`, channel `joe`, and release `joe-2026.05.28.1`.

## Joe Codex checklist

Run these from PowerShell on Joe's PC.

```powershell
git fetch origin
git switch codex/joes-apps
git pull --ff-only
python -m py_compile claim_manager_3/app.py claim_manager_3/data_access.py mobile_api/server.py
```

Confirm the desktop install exists:

```powershell
$installRoot = Join-Path $env:ProgramFiles "Claim Manager 3"
Test-Path (Join-Path $installRoot "Claim Manager 3\Claim Manager 3.exe")
Test-Path (Join-Path $env:LOCALAPPDATA "Claim Manager 3\claims_data.json")
```

Start the Joe mobile API from the updated source checkout:

```powershell
$repo = "C:\Users\Joe\Desktop\IA APP"
$env:CLAIM_MANAGER_HOME = Join-Path $env:LOCALAPPDATA "Claim Manager 3"
$env:CLAIM_MANAGER_HOME_ROOT = $repo
$env:CLAIM_MANAGER_MOBILE_API_PORT = "8012"
python -u "$repo\mobile_api\server.py"
```

Verify the local API:

```powershell
Invoke-RestMethod "http://127.0.0.1:8012/api/mobile/health"
```

The Joe mobile app is pinned to:

```text
https://joe-api.luxuryimportsusa.shop
```

That hostname must be routed to Joe's PC tunnel. If public health fails with Cloudflare `1033`, the tunnel is not connected:

```powershell
Invoke-RestMethod "https://joe-api.luxuryimportsusa.shop/api/mobile/health"
cloudflared tunnel info claim-manager-mobile-joe
```

Once the tunnel is connected, the public health check should return:

```json
{"ok": true}
```

## Desktop updater check

Joe's desktop updater should see:

```powershell
Invoke-RestMethod "https://api2.luxuryimportsusa.shop/api/desktop-update/latest.json"
```

Expected version: `3.0.10`.
