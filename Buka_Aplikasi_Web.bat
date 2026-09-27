@echo off
title SISTEM OTOMASI LAB KLINIS - WEB DASHBOARD
color 0A

echo ========================================================
echo       MEMBUKA WEB DASHBOARD OTOMASI LAB KLINIS
echo ========================================================
echo.
echo  [1/2] Menjalankan server aplikasi di latar belakang...
start /B python -m uvicorn app:app --host 127.0.0.1 --port 8000

echo  [2/2] Membuka antarmuka di browser Anda...
timeout /t 2 /nobreak >nul
start http://localhost:8000

echo.
echo ========================================================
echo  Dashboard telah dibuka di browser!
echo  Untuk menutup aplikasi, Anda dapat menutup jendela ini.
echo ========================================================
echo.
pause
