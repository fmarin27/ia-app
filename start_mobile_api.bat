@echo off
cd /d "C:\Users\ferna\Desktop\IA APP"
title Claim Manager Mobile API
echo Starting Claim Manager Mobile API...
echo.
echo Local test:
echo   http://127.0.0.1:8011/api/mobile/health
echo.
echo Phone test:
echo   http://192.168.1.134:8011/api/mobile/health
echo.
"C:\Program Files\Python311\python.exe" -u "C:\Users\ferna\Desktop\IA APP\mobile_api\server.py"
echo.
echo Mobile API stopped. Press any key to close this window.
pause >nul
