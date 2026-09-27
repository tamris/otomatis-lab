@echo off
chcp 65001 >nul
title OTOMASI LABORATORIUM - MODE EKSTRAKSI FOTO
color 0A
cd /d "%~dp0"
cls
echo ============================================================
echo    MEMULAI PROSES EKSTRAKSI FOTO LABORATORIUM
echo ============================================================
echo.
python otomasi_lab.py --extract
echo.
echo ============================================================
echo    PROSES SELESAI!
echo    Silakan periksa data di Excel sebelum mencetak dokumen.
echo ============================================================
echo.
pause
