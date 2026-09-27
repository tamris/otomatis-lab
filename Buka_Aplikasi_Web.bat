@echo off
title SISTEM OTOMASI LAB KLINIS - WEB DASHBOARD
color 0A

echo ========================================================
echo       MEMBUKA WEB DASHBOARD OTOMASI LAB KLINIS
echo ========================================================
echo.

:: Bersihkan server lama jika masih ada yang berjalan di port 8000
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8000 ^| findstr LISTENING') do (
    echo  [INFO] Menutup instance server sebelumnya (PID: %%a)...
    taskkill /F /PID %%a >nul 2>&1
)

echo  [1/2] Menjalankan server aplikasi terbaru...
start /B python -m uvicorn app:app --host 127.0.0.1 --port 8000 --reload

echo  [2/2] Membuka antarmuka di browser Anda...
timeout /t 2 /nobreak >nul
start http://localhost:8000

echo.
echo ========================================================
echo  Dashboard telah aktif di browser: http://localhost:8000
echo  Biarkan jendela ini tetap terbuka selama aplikasi digunakan.
echo ========================================================
echo.
pause
