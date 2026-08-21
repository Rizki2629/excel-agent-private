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
echo Opening Excel to load add-in...
echo ----------------------------------------

:: Open Excel with the workbook
excel "C:\Users\acer\agent_excel\ADMINISTRSAI KELAS.xlsm"

echo.
echo [SUCCESS] Excel opened!
echo.
echo Next steps:
echo   1. Go to ADD-INS tab in Excel ribbon
echo   2. Click "Excel Agent" button
echo   3. If prompted, click "Allow" or "Trust"
echo.
pause