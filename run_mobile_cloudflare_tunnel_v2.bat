@echo off
cd /d "C:\Users\ferna\Desktop\IA APP"
set TUNNEL_ORIGIN_CERT=%USERPROFILE%\.cloudflared\cert.pem
"C:\Program Files (x86)\cloudflared\cloudflared.exe" tunnel --config "C:\Users\ferna\Desktop\IA APP\cloudflared-mobile-v2.yml" run claim-manager-mobile-api-v2
