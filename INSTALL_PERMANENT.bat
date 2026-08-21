@echo off
title Install Excel Agent Add-in Permanen
color 0A
cls
echo ========================================
echo   PERMANENT INSTALLATION SCRIPT
echo ========================================
echo.
echo This will register the add-in so it loads automatically!
echo.

cd /d "%~dp0"

:: Check admin rights
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [WARNING] Please run as Administrator for permanent installation!
    echo          Right-click this file and select "Run as administrator".
    echo.
    choice /C YN /M "Continue without admin privileges?"
    if %errorlevel% equ 2 goto ADMIN_REQUIRED
)

echo [INFO] Registering add-in via manifest...
echo.
echo This requires Excel to be installed.
echo If you see errors, make sure Excel is installed first.
echo.

pause

:: Try to open Excel with catalog command
echo [INFO] Opening Excel to install catalog...
start "" excel.exe /catalog "%CD%\catalog\manifest.xml"

echo.
echo ----------------------------------------
echo [SUCCESS] Installation process started!
echo ----------------------------------------
echo.
echo Next steps in Excel:
echo   1. Close any open workbooks
echo   2. Go to File → Options → Add-ins
echo   3. Select "COM Add-ins" from dropdown
echo   4. Click "Go..."
echo   5. Find and check "Excel Agent"
echo   6. Click OK
echo.
echo OR manually browse to manifest.xml
echo.
pause
exit

:ADMIN_REQUIRED
echo [ERROR] Admin rights required for permanent installation.
echo.
echo Right-click this file and select "Run as administrator".
pause
exit /b 1
