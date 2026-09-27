@echo off
chcp 65001 >nul
title SISTEM OTOMASI LAB - TOOL PEMBERSIH & RESET
color 0C

:MENU
cls
echo ============================================================
echo       SISTEM OTOMASI LAB - TOOL PEMBERSIH ^& RESET
echo ============================================================
echo.
echo  Pilih opsi pembersihan / reset:
echo.
echo  [1] Reset ^& Bersihkan Puskesmas Tertentu
echo      (Menghapus Rekap Excel ^& Folder Pasien PKM tersebut,
echo       lalu mengembalikan foto dari arsip ke foto_masuk)
echo.
echo  [2] Sterilkan File Template Master (Template Exel.xlsx)
echo      (Memastikan Template Exel.xlsx murni kosong / header only)
echo.
echo  [3] Cek Status Sistem (Health Check)
echo      (Melihat jumlah foto masuk, arsip, dan rekap yang ada)
echo.
echo  [0] Keluar
echo.
echo ============================================================
set /p opt="Pilih nomor [0-3]: "

if "%opt%"=="1" goto RESET_PKM
if "%opt%"=="2" goto CLEAN_TEMPLATE
if "%opt%"=="3" goto HEALTH_CHECK
if "%opt%"=="0" exit /b
goto MENU

:RESET_PKM
cls
echo ============================================================
echo              RESET DATA PER PUSKESMAS
echo ============================================================
echo.
echo  Puskesmas yang tersedia saat ini:
dir /b /ad HASIL_PUSKESMAS 2>nul
echo.
set /p pkm="Ketik Nama Puskesmas yang ingin direset (misal PUSKESMAS LEBAKSIU): "
if "%pkm%"=="" goto MENU
echo.
echo Sedang mereset %pkm%...
python reset_tool.py --reset-pkm "%pkm%"
echo.
pause
goto MENU

:CLEAN_TEMPLATE
cls
echo.
echo Sedang mensterilkan Template Exel.xlsx...
python reset_tool.py --clean-template
echo.
pause
goto MENU

:HEALTH_CHECK
cls
python reset_tool.py --health
pause
goto MENU
