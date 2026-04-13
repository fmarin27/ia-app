@echo off
cd /d "C:\Users\ferna\Desktop\IA APP"
title Claim Manager Remote Access

echo Starting Claim Manager mobile API...
start "Claim Manager Mobile API" cmd /k ""C:\Program Files\Python311\python.exe" -u "C:\Users\ferna\Desktop\IA APP\mobile_api\server.py""

echo Starting Cloudflare tunnel...
start "Claim Manager Cloudflare Tunnel" cmd /k ""C:\Users\ferna\Desktop\IA APP\run_mobile_cloudflare_tunnel.bat""

echo.
echo Public mobile API should become available at:
echo   https://api.luxuryimportsusa.shop/api/mobile/health
echo.
echo Keep both windows open while using the phone app remotely.
echo.
pause
