@echo off
title Excel Agent - Master Menu
color 0F
cls
:MENU
echo ========================================
echo   EXCEL AGENT - CONTROL MENU
echo ========================================
echo.
echo [1] 🚀 LAUNCH ALL (Recommended for first time)
echo [2] 🖥️  Start Server Only
echo [3] 📁 Load Add-in to Excel
echo [4] 🗑️  Clear All Cache
echo [5] ⏹️  Stop Everything
echo [6] 🔍 Check Status
echo [Q] Quit
echo.
echo ----------------------------------------
echo Current working directory: %CD%
echo ----------------------------------------
echo.

set /p choice="Select option [1-6 or Q]: "
if "%choice%"=="1" goto LAUNCH_ALL
if "%choice%"=="2" goto START_SERVER
if "%choice%"=="3" goto LOAD_ADDIN
if "%choice%"=="4" goto CLEAR_CACHE
if "%choice%"=="5" goto STOP_ALL
if "%choice%"=="6" goto CHECK_STATUS
if /i "%choice%"=="Q" goto QUIT
if /i "%choice%"=="q" goto QUIT

echo [ERROR] Invalid option! Press any key to retry...
pause >nul
cls
goto MENU

:LAUNCH_ALL
cls
echo ========================================
echo   LAUNCHING EVERYTHING...
echo ========================================
call START_ALL.bat
goto MENU

:START_SERVER
cls
echo Starting server...
start "" cmd /k "cd /d %~dp0 && node server.js"
echo [OK] Server starting on port 3000
timeout /t 2 /nobreak >nul
goto MENU

:LOAD_ADDIN
cls
echo Loading add-in...
excel "C:\Users\acer\agent_excel\ADMINISTRSAI KELAS.xlsm"
echo [OK] Excel opened
timeout /t 2 /nobreak >nul
goto MENU

:CLEAR_CACHE
cls
echo Clearing cache...
call CLEAR_CACHE.bat
goto MENU

:STOP_ALL
cls
echo Stopping everything...
call STOP_ALL.bat
goto MENU

:CHECK_STATUS
cls
echo ========================================
echo   CURRENT STATUS
echo ========================================
echo.
echo Checking server on port 3000...
netstat -ano | findstr :3000 >nul
if %errorlevel% equ 0 (
    echo [✅] Server is RUNNING
) else (
    echo [❌] Server is NOT running
)
echo.
echo Checking Excel processes...
tasklist | findstr excel.exe >nul
if %errorlevel% equ 0 (
    echo [✅] Excel is RUNNING
    tasklist | findstr excel.exe
) else (
    echo [❌] Excel is NOT running
)
echo.
echo Checking Node.js processes...
tasklist | findstr node.exe >nul
if %errorlevel% equ 0 (
    echo [✅] Node.js processes are running
    tasklist | findstr node.exe
) else (
    echo [❌] No Node.js processes found
)
echo.
echo Config file location: %APPDATA%\ExcelAgent\config.json
echo Scratch workdir: %TEMP%\excelagent-opencode\
echo.
pause
goto MENU

:QUIT
cls
echo Thank you for using Excel Agent!
echo Repository: https://github.com/Rizki2629/excel-agent-private
echo.
exit

:: Labels and commands that might be needed
:start_all
cls
call LAUNCH_ALL.bat
goto MENU
