@echo off
title Excel Agent - Master Launcher (One Click)
color 0D
cls
echo ========================================
echo   EXCEL AGENT MASTER LAUNCHER
echo         ONE-CLICK SOLUTION
echo ========================================
echo.

cd /d "%~dp0"

:: Check if running as admin
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [WARNING] This script should be run as Administrator for best results!
    echo          Right-click and "Run as administrator" recommended.
    echo.
)

echo Step 1 of 4: Stopping all processes...
taskkill /F /IM EXCEL.EXE >nul 2>&1
taskkill /F /FI "WINDOWTITLE eq node*" >nul 2>&1
timeout /t 1 /nobreak >nul
echo [OK] Old processes stopped.
echo.

echo Step 2 of 4: Clearing Office cache...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\" (
    del /q /s "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*" >nul 2>&1
    echo [OK] Cache cleared.
) else (
    echo [INFO] No cache to clear.
)
echo.

echo Step 3 of 4: Starting server on port 3000...
start "" cmd /k "node server.js"
echo [OK] Server starting...
echo.

timeout /t 3 /nobreak >nul

echo Step 4 of 4: Sideloading add-in and opening Excel...
call npx office-addin-debugging start manifest.xml desktop --app excel --no-debug
if exist "C:\Users\acer\agent_excel\ADMINISTRSAI KELAS.xlsm" (
    excel "C:\Users\acer\agent_excel\ADMINISTRSAI KELAS.xlsm"
)
echo [OK] Add-in registered and Excel opened.
echo.

echo ----------------------------------------
echo 🎉 EVERYTHING IS NOW RUNNING!
echo ----------------------------------------
echo.
echo Your workbook has been opened in Excel.
echo Untuk membuka panel:
echo   1. Buka tab HOME (Beranda) di Ribbon Excel
echo   2. Cari grup "Excel Agent" di sebelah kanan
echo   3. Klik tombol "Buka Asisten AI"
echo.
echo Server is running at: https://localhost:3000/taskpane.html
echo.
echo To STOP: Close all windows or press CTRL+C on server console
echo.
pause