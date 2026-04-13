@echo off
cd /d "C:\Users\ferna\Desktop\IA APP"
title Claim Manager Remote Access V2

echo Starting Claim Manager mobile API V2...
start "Claim Manager Mobile API V2" cmd /k ""C:\Users\ferna\Desktop\IA APP\start_mobile_api_v2.bat""

echo Starting Cloudflare tunnel V2...
start "Claim Manager Cloudflare Tunnel V2" cmd /k ""C:\Users\ferna\Desktop\IA APP\run_mobile_cloudflare_tunnel_v2.bat""

echo.
echo Public mobile API should become available at:
echo   https://api2.luxuryimportsusa.shop/api/mobile/health
echo.
echo Keep both windows open while using the phone app remotely.
echo.
pause
