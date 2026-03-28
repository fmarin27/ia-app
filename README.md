# Claims Manager

Local desktop app for tracking claim folders, open claims, and closed claims.

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
