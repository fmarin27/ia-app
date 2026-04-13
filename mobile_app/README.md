# Claim Manager Mobile

Android and iPhone companion app for the local claims system.

Current scope:
- Open claims list
- Closed claims list
- Claim detail view
- Office update view
- Route planner view
- Claim tools library view
- Body shops and insurance companies reference views
- Open assignment sheet PDF
- Open any file from the claim folder
- Save a note back to the local app
- Upload photos from the phone directly into the claim folder on the PC

## Folders

- Mobile API:
  - `C:\Users\ferna\Desktop\IA APP\mobile_api\server.py`
- Mobile app:
  - `C:\Users\ferna\Desktop\IA APP\mobile_app`
- EAS config:
  - `C:\Users\ferna\Desktop\IA APP\mobile_app\eas.json`

## Run the API

```powershell
& 'C:\Program Files\Python311\python.exe' 'C:\Users\ferna\Desktop\IA APP\mobile_api\server.py'
```

This starts the API on:

```text
http://127.0.0.1:8011
```

For Android on the same Wi-Fi, use your PC's LAN IP in the app instead of `127.0.0.1`.

## Install mobile dependencies

From `C:\Users\ferna\Desktop\IA APP\mobile_app`:

```powershell
npm install
```

## Start the mobile app

```powershell
npm run android
```

Or:

```powershell
npm run ios
```

Or:

```powershell
npm run start
```

Then open it with Expo Go on Android or iPhone.

## Expo / EAS

- This app is the live Expo companion app.
- `app.json` includes the current Expo Updates project URL.
- `mobile_app\eas.json` is the active EAS build config.

## Notes

- The app fetches live data from the local API on your PC.
- Use your PC's LAN/Wi-Fi IP in the app instead of `127.0.0.1`.
- `Open Assign Sheet` and file links come from the claim folder on the PC.
- Photo upload saves a new image directly into the selected claim folder.
- Saving notes writes back into the local claim data so the desktop app can see them.
