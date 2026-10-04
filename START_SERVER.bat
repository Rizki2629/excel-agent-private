@echo off
title Excel Agent Server
color 0A
cls
echo ========================================
echo   EXCEL AGENT SERVER STARTER
echo ========================================
echo.
echo Starting server on port 3000...
echo.

cd /d "%~dp0"

:: Check if Node.js is installed
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Node.js not found!
    echo Please install Node.js 18+ from https://nodejs.org/
    pause
    exit /b 1
)

echo [OK] Node.js detected: 
node --version

echo.
echo [INFO] Server berjalan di https://localhost:3000
echo [INFO] Add-in sudah terpasang permanen di Tab HOME (Beranda) Excel
echo [INFO] Tekan CTRL+C untuk menghentikan server
echo.
echo ----------------------------------------
node server.js

pause