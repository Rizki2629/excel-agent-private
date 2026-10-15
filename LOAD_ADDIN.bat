@echo off
title Load Excel Agent Add-in
color 0B
cls
echo ========================================
echo   EXCEL AGENT ADD-IN INSTALLER
echo ========================================
echo.

cd /d "%~dp0"

:: Check if Excel is installed
where excel >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Excel not found!
    pause
    exit /b 1
)

echo [INFO] Installing Excel Agent add-in...
echo.
echo Please make sure:
echo   1. Server is already running on port 3000
echo   2. You trust this add-in in Excel
echo.
echo ----------------------------------------
echo Registering add-in and opening Excel...
echo ----------------------------------------

call npx office-addin-debugging start manifest.xml desktop --app excel --no-debug
if exist "C:\Users\acer\agent_excel\ADMINISTRSAI KELAS.xlsm" (
    excel "C:\Users\acer\agent_excel\ADMINISTRSAI KELAS.xlsm"
)

echo.
echo [SUCCESS] Add-in registered!
echo.
echo Cara membuka:
echo   1. Buka tab HOME (Beranda) di Ribbon Excel
echo   2. Cari grup "Excel Agent" di sebelah kanan
echo   3. Klik tombol "Buka Asisten AI"
echo.
pause