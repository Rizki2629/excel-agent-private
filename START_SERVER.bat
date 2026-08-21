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
echo [INFO] Server will run at:
echo         https://localhost:3000/taskpane.html
echo.
echo [INFO] Press CTRL+C to stop the server
echo.
echo ----------------------------------------
echo Opening browser in 5 seconds...
echo ----------------------------------------
timeout /t 5 /nobreak >nul

start https://localhost:3000/taskpane.html

node server.js

pause