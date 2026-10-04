@echo off
title Excel Agent - START
color 0A
cd /d "%~dp0"

:MENU
cls
echo ========================================
echo   EXCEL AGENT
echo ========================================
echo.
echo [1] Jalankan Semua (server + buka Excel)
echo [2] Server saja
echo [3] Load Add-in ke Excel
echo [4] Bersihkan cache
echo [5] Cek status
echo [Q] Keluar
echo.
set /p choice="Pilih [1-5/Q]: "
if "%choice%"=="1" goto LAUNCH_ALL
if "%choice%"=="2" goto START_SERVER
if "%choice%"=="3" goto LOAD_ADDIN
if "%choice%"=="4" goto CLEAR_CACHE
if "%choice%"=="5" goto CHECK_STATUS
if /i "%choice%"=="Q" exit /b 0
echo [ERROR] Pilihan tidak valid.
pause >nul
goto MENU

:LAUNCH_ALL
cls
echo --- [1/4] Menghentikan proses lama...
taskkill /F /IM EXCEL.EXE >nul 2>&1
taskkill /F /FI "WINDOWTITLE eq node*" >nul 2>&1
timeout /t 1 /nobreak >nul
echo [OK] Proses lama dihentikan.
echo.
echo --- [2/4] Membersihkan cache Office...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\" (
    del /q /s "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*" >nul 2>&1
    echo [OK] Cache dibersihkan.
) else (
    echo [INFO] Tidak ada cache.
)
echo.
echo --- [3/4] Menjalankan server port 3000...
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Node.js tidak ditemukan! Jalankan SETUP.bat dulu.
    pause
    goto MENU
)
start "" cmd /k "cd /d %~dp0 && node server.js"
echo [OK] Server dijalankan.
timeout /t 3 /nobreak >nul
echo.
echo --- [4/4] Mendaftarkan add-in dan membuka Excel...
call npx office-addin-debugging start manifest.xml desktop --app excel --no-debug >nul 2>&1
if exist "C:\Users\acer\agent_excel\ADMINISTRSAI KELAS.xlsm" (
    start "" excel "C:\Users\acer\agent_excel\ADMINISTRSAI KELAS.xlsm"
)
echo [OK] Selesai.
echo.
echo ----------------------------------------
echo  Server: https://localhost:3000
echo  Di Excel: tab HOME -^> grup "Excel Agent" -^> "Buka Asisten AI"
echo ----------------------------------------
pause
goto MENU

:START_SERVER
cls
echo ========================================
echo   SERVER SAJA
echo ========================================
echo.
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Node.js tidak ditemukan! Jalankan SETUP.bat dulu.
    pause
    goto MENU
)
echo [INFO] Server berjalan di https://localhost:3000
echo [INFO] Tekan CTRL+C untuk menghentikan server
echo ----------------------------------------
node server.js
pause
goto MENU

:LOAD_ADDIN
cls
echo ========================================
echo   LOAD ADD-IN KE EXCEL
echo ========================================
echo.
where excel >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Excel tidak ditemukan!
    pause
    goto MENU
)
echo [INFO] Pastikan server sudah berjalan (opsi [2]).
echo.
call npx office-addin-debugging start manifest.xml desktop --app excel --no-debug
if exist "C:\Users\acer\agent_excel\ADMINISTRSAI KELAS.xlsm" (
    start "" excel "C:\Users\acer\agent_excel\ADMINISTRSAI KELAS.xlsm"
)
echo.
echo [OK] Add-in dimuat. Di Excel: tab HOME -^> "Buka Asisten AI".
pause
goto MENU

:CLEAR_CACHE
cls
echo ========================================
echo   BERSIHKAN CACHE
echo ========================================
echo.
echo [INFO] Menghentikan Excel...
taskkill /F /IM EXCEL.EXE >nul 2>&1
timeout /t 2 /nobreak >nul
echo [INFO] Membersihkan cache Office WEF...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\" (
    del /q /s "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*" >nul 2>&1
    echo [OK] Cache Office dibersihkan.
) else (
    echo [INFO] Folder cache Office tidak ditemukan.
)
if exist "%LOCALAPPDATA%\Google\Chrome\User Data\Default\Cache\" (
    del /q /s "%LOCALAPPDATA%\Google\Chrome\User Data\Default\Cache\*" >nul 2>&1
    echo [OK] Cache Chrome dibersihkan.
)
echo.
echo [OK] Selesai. Jalankan opsi [1] untuk mulai fresh.
pause
goto MENU

:CHECK_STATUS
cls
echo ========================================
echo   STATUS
echo ========================================
echo.
netstat -ano | findstr :3000 >nul
if %errorlevel% equ 0 (
    echo [OK] Server port 3000: JALAN
) else (
    echo [--] Server port 3000: MATI
)
tasklist | findstr /i excel.exe >nul
if %errorlevel% equ 0 (
    echo [OK] Excel: JALAN
) else (
    echo [--] Excel: MATI
)
tasklist | findstr /i node.exe >nul
if %errorlevel% equ 0 (
    echo [OK] Node.js: JALAN
) else (
    echo [--] Node.js: MATI
)
echo.
pause
goto MENU
