@echo off
cd /d "C:\Users\ferna\Desktop\IA APP"
title Claim Manager Mobile API V2
echo Starting Claim Manager Mobile API V2...
echo.
echo Local test:
echo   http://127.0.0.1:8012/api/mobile/health
echo.
echo Public phone test:
echo   https://api2.luxuryimportsusa.shop/api/mobile/health
echo.
set CLAIM_MANAGER_MOBILE_API_PORT=8012
"C:\Program Files\Python311\python.exe" -u "C:\Users\ferna\Desktop\IA APP\mobile_api\server.py"
echo.
echo Mobile API stopped. Press any key to close this window.
pause >nul
