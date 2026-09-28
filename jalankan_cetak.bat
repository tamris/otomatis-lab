@echo off
chcp 65001 >nul
title OTOMASI LABORATORIUM - MODE CETAK ALL-IN-ONE
color 0B
cd /d "%~dp0"
cls
echo ============================================================
echo    MEMULAI PROSES GENERATE WORD ^& PDF ALL-IN-ONE
echo ============================================================
echo.
python otomasi_lab.py --generate %*
echo.
echo ============================================================
echo    PROSES SELESAI!
echo    Seluruh dokumen Word ^& PDF telah selesai dibuat.
echo ============================================================
echo.
pause
