# Claims Web App

First browser version of the claims manager.

## Run

```powershell
python web_app\server.py
```

Then open:

```text
http://127.0.0.1:8010
```

## What it does

- Reads the same `claims_data.json` used by the desktop app
- Shows all/open/closed claims
- Supports search
- Shows claim details in a side panel

## Notes

- This is a starting point, not a replacement for the desktop app yet
- It uses only Python standard library code, so no extra install is needed
