# IA APP

Current project layout for the home PC, office PC, and mobile companion app.

## Live Apps

- Desktop app:
  - Home launcher: `C:\Users\ferna\Desktop\IA APP\claim_manager_3\Start Claim Manager.bat`
  - Office launcher: `C:\Users\ferna\Desktop\IA APP\claim_manager_3\Start Office Claim Manager.bat`
  - Shared desktop code: `C:\Users\ferna\Desktop\IA APP\claim_manager_3`
- Mobile API:
  - `C:\Users\ferna\Desktop\IA APP\mobile_api\server.py`
- Expo mobile app:
  - `C:\Users\ferna\Desktop\IA APP\mobile_app`

`claims_manager.py` is now legacy and is no longer the active desktop app.

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

## Codex Bridge

This project also includes a bridge app for coordinating work between the home PC and office PC through a synced folder such as Syncthing.

Run it with:

```powershell
python codex_bridge.py
```

Or use the default launcher:

```powershell
start_codex_bridge.bat
```

The default launcher now opens the newer Qt version with a nicer UI.

You can also run it directly:

```powershell
start_codex_bridge_qt.bat
```

What it does:

- Creates a shared chat timeline across both PCs
- Tracks pending command requests and asks for approval before running them
- Executes approved PowerShell commands on the receiving PC and writes the result back to the bridge
- Publishes local project folders and local focus notes so both sides can stay on the same page
- Stores local setup in `codex_bridge_config.json`

Typical setup:

1. Pick a shared bridge folder that Syncthing keeps in sync on both PCs.
2. Run the bridge app on both PCs.
3. Set one app's local node to `home` and the other to `office`.
4. Add the project root folders each PC should publish.
5. Send chat messages or command requests through the shared bridge UI.

The bridge is intentionally approval-gated for command execution. A command request can be broad, but the receiving PC still asks you before it runs.

## Shared Data

- Local claim data: `C:\Users\ferna\Desktop\IA APP\claims_data.json`
- The office launcher uses the same PySide6 app with office-specific sync behavior in `claim_manager_3\data_access.py`.
