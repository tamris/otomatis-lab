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

rem Ambil tanggal default terakhir dari config.txt
python -c "import sys, otomasi_lab as o; c=o.load_config(); open(sys.argv[1], 'w', encoding='utf-8').write(f'set \"DEF_EXAM={c.get(\"TANGGAL_EXAM\", \"\")}\"\nset \"DEF_SURAT={c.get(\"TANGGAL_SURAT\", \"\")}\"\n')" "%TEMP%\lab_cfg.bat"
call "%TEMP%\lab_cfg.bat"
del "%TEMP%\lab_cfg.bat" 2>nul

echo ------------------------------------------------------------
echo  PENGATURAN TANGGAL PEMERIKSAAN ^& SURAT
echo  (Ketik tanggal baru, atau tekan ENTER jika ingin pakai default)
echo ------------------------------------------------------------
set "INPUT_EXAM="
set /p "INPUT_EXAM= Tanggal Exam  [%DEF_EXAM%]: "
if "%INPUT_EXAM%"=="" set "INPUT_EXAM=%DEF_EXAM%"

set "INPUT_SURAT="
set /p "INPUT_SURAT= Tanggal Surat [%DEF_SURAT%]: "
if "%INPUT_SURAT%"=="" set "INPUT_SURAT=%DEF_SURAT%"

echo.
echo  [+] Tanggal Exam  : %INPUT_EXAM%
echo  [+] Tanggal Surat : %INPUT_SURAT%
echo ------------------------------------------------------------
echo.

python otomasi_lab.py --extract --tgl-exam "%INPUT_EXAM%" --tgl-surat "%INPUT_SURAT%" %*
echo.
echo ============================================================
echo    PROSES SELESAI!
echo    Silakan periksa data di Excel sebelum mencetak dokumen.
echo ============================================================
echo.
pause
