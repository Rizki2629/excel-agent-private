@echo off
title Clear Excel Agent Cache
color 0C
cls
echo ========================================
echo   EXCEL AGENT CACHE CLEARER
echo ========================================
echo.
echo This will:
echo   1. Stop all Excel processes
echo   2. Clear Office WEF cache
echo   3. Clear Chrome cache (if used)
echo.

:: Stop Excel
echo [INFO] Stopping Excel processes...
taskkill /F /IM EXCEL.EXE >nul 2>&1
if %errorlevel% equ 0 (
    echo [OK] All Excel instances closed.
) else (
    echo [INFO] No Excel processes were running.
)
timeout /t 2 /nobreak >nul

:: Clear Office WEF cache
echo [INFO] Clearing Office WEF cache...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\" (
    del /q /s "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*" >nul 2>&1
    echo [OK] Office WEF cache cleared.
) else (
    echo [INFO] WEF cache directory not found.
)
timeout /t 1 /nobreak >nul

:: Clear Chrome cache (optional - if you see taskpane in Chrome)
echo [INFO] Clearing Chrome cache (if exists)...
if exist "%LOCALAPPDATA%\Google\Chrome\User Data\Default\Cache\" (
    del /q /s "%LOCALAPPDATA%\Google\Chrome\User Data\Default\Cache\*" >nul 2>&1
    echo [OK] Chrome cache cleared.
) else (
    echo [INFO] Chrome cache not found or not needed.
)
timeout /t 1 /nobreak >nul

:: Clear browser session storage for localhost:3000
echo [INFO] Clearing session storage for localhost:3000...
del /q /s "%LOCALAPPDATA%\Google\Chrome\User Data\Default\Synced\Local Storage\https_localhost.3000_0.localstorage" >nul 2>&1

echo.
echo ----------------------------------------
echo [SUCCESS] Cache cleared successfully!
echo ----------------------------------------
echo.
echo Next step: Open the START_SERVER.bat file
echo to restart everything fresh.
echo.
pause