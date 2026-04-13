# Claim Manager 3.0

This is the live PySide6 desktop app for IA APP.

Current scope:
- Open claims
- Closed claims
- Claim detail panel
- Reports
- Office update
- Route planner
- Claim tools
- Body shops
- Insurance companies
- Open assignment sheet
- Mitchell and AppTrak helper workflows

## Run

```powershell
cd "C:\Users\ferna\Desktop\IA APP\claim_manager_3"
pip install -r requirements.txt
python app.py
```

Or use the launcher:

```powershell
.\Start Claim Manager.bat
```

For the office-connected desktop mode:

```powershell
.\Start Office Claim Manager.bat
```

## Notes

- Home and office use the same `app.py`.
- The office version changes mode through `officeversion.py`.
- Shared claim data is stored in `C:\Users\ferna\Desktop\IA APP\claims_data.json`.
