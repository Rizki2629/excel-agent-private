@echo off
title Stop Excel Agent
color 0E
cls
echo ========================================
echo   STOP EXCEL AGENT
echo ========================================
echo.

:: Stop server
echo [INFO] Stopping server on port 3000...
taskkill /F /FI "WINDOWTITLE eq node*" >nul 2>&1
if %errorlevel% equ 0 (
    echo [OK] Server stopped.
) else (
    echo [INFO] Server was not running.
)

:: Stop Node processes
echo [INFO] Stopping all Node.js processes...
taskkill /F /IM node.exe >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] No Node processes to stop.
)

timeout /t 1 /nobreak >nul

:: Close Excel
echo [INFO] Closing Excel (optional)...
choice /C YN /M "Do you want to close Excel as well?"
if %errorlevel% equ 1 (
    taskkill /F /IM EXCEL.EXE >nul 2>&1
    echo [OK] Excel closed.
) else (
    echo [INFO] Excel left open.
)

echo.
echo ----------------------------------------
echo ✅ ALL PROCESSES STOPPED!
echo ----------------------------------------
pause