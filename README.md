# Claims Manager

Local desktop app for tracking claim folders, open claims, and closed claims.

## SFTP Transfer Tool

This project also includes a small SFTP desktop utility for moving folders between PCs with the Windows OpenSSH client that already ships with Windows.

Run it with:

```powershell
python sftp_transfer_tool.py
```

Or use:

```powershell
start_sftp_transfer_tool.bat
```

What it does:

- Saves SFTP connection profiles in `sftp_profiles.json`
- Generates an `ed25519` SSH keypair for you in `sftp_keys/`
- Uploads a local folder to a remote folder over SFTP
- Downloads a remote folder back to your PC
- Shows the exact `sftp.exe` command and batch commands it runs so you can reuse the same flow in the claims manager app later

For tomorrow's office-PC move:

1. Enable **OpenSSH Server** on the new office PC.
2. Copy the generated `.pub` key into that PC user's `authorized_keys`.
3. Use this tool from home to upload your folders into a folder on the office PC.

The tool uses key-based auth on purpose because it is the cleanest and safest way to automate SFTP from Python without adding extra third-party packages.

## What it does

- Watches one or more folders you choose.
- Treats each file or subfolder inside those watched folders as a claim entry.
- Shows tabs for all claims, open claims, and closed claims.
- Lets you add manual claims that are not tied to a folder yet.
- Saves notes, status, and assignment details in `claims_data.json`.

## Run it

```powershell
python claims_manager.py
```

## How to use it

1. Click `Add Folder` and choose a folder that contains your claim folders or files.
2. Click a claim to review it on the right.
3. Add notes and save them.
4. Use `Mark Open` or `Mark Closed` to manage status.
5. Click `Refresh Scan` anytime after folder contents change.

## Stored data

- App file: `claims_manager.py`
- Local data: `claims_data.json`

The JSON file is created automatically the first time you use the app.
