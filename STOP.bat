@echo off
title Stop Excel Agent
color 0E
cls
echo ========================================
echo   STOP EXCEL AGENT
echo ========================================
echo.
echo [INFO] Menghentikan server port 3000...
taskkill /F /FI "WINDOWTITLE eq node*" >nul 2>&1
taskkill /F /IM node.exe >nul 2>&1
echo [OK] Server dihentikan.
timeout /t 1 /nobreak >nul
echo.
choice /C YN /M "Tutup Excel juga"
if %errorlevel% equ 1 (
    taskkill /F /IM EXCEL.EXE >nul 2>&1
    echo [OK] Excel ditutup.
) else (
    echo [INFO] Excel dibiarkan terbuka.
)
echo.
echo ----------------------------------------
echo [SELESAI] Semua proses dihentikan.
echo ----------------------------------------
pause
