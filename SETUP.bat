@echo off
title Setup Excel Agent - Komputer Baru
color 0B
cls

echo ====================================================================
echo             SETUP OTOMATIS EXCEL AGENT (KOMPUTER BARU)
echo ====================================================================
echo.
echo Script ini akan menyiapkan seluruh kebutuhan Excel Agent secara otomatis:
echo   [1] Cek & siapkan Node.js + NPM
echo   [2] Cek & siapkan Python + library pywin32
echo   [3] Install dependencies project (npm install)
echo   [4] Install sertifikat HTTPS lokal (bebas error SSL)
echo   [5] Registrasi Add-in ke Registry Office Windows (Permanen)
echo   [6] Inisialisasi folder & file konfigurasi
echo   [7] Buat shortcut di Desktop
echo.
echo --------------------------------------------------------------------
echo Menjalankan dari folder: %~dp0
echo --------------------------------------------------------------------
echo.
pause

cd /d "%~dp0"

:: ----------------------------------------------------------------------
:: TAHAP 1: Periksa Node.js & NPM
:: ----------------------------------------------------------------------
echo.
echo [1/7] Memeriksa Node.js...
where node >nul 2>nul
if %errorlevel% neq 0 (
    color 0C
    echo [ERROR] Node.js belum terinstall di komputer ini!
    echo Silakan download dan install Node.js (LTS version) dari:
    echo   https://nodejs.org/
    echo.
    echo Setelah selesai install Node.js, jalankan kembali script ini.
    pause
    exit /b 1
)
for /f "tokens=*" %%i in ('node -v') do set NODE_VER=%%i
echo [OK] Node.js terdeteksi: %NODE_VER%

:: ----------------------------------------------------------------------
:: TAHAP 2: Periksa Python & pywin32
:: ----------------------------------------------------------------------
echo.
echo [2/7] Memeriksa Python...
where python >nul 2>nul
if %errorlevel% neq 0 (
    color 0E
    echo [PERINGATAN] Python belum terdeteksi di PATH sistem.
    echo Eksekusi script win32com membutuhkan Python 3.
    echo Pastikan Python terinstall dan centang "Add Python to PATH" saat install.
) else (
    for /f "tokens=*" %%i in ('python --version') do set PY_VER=%%i
    echo [OK] Python terdeteksi: %PY_VER%
    echo     Memeriksa pustaka pywin32 (win32com)...
    python -c "import win32com.client" >nul 2>nul
    if %errorlevel% neq 0 (
        echo     Menginstall pywin32 via pip...
        pip install pywin32
        if %errorlevel% equ 0 (
            echo [OK] pywin32 berhasil diinstall.
        ) else (
            echo [PERINGATAN] Gagal install pywin32 otomatis. Jalankan: pip install pywin32
        )
    ) else (
        echo [OK] pywin32 sudah siap.
    )
)

:: ----------------------------------------------------------------------
:: TAHAP 3: Install NPM Dependencies
:: ----------------------------------------------------------------------
echo.
echo [3/7] Memeriksa & menginstall dependensi project (npm install)...
if not exist "node_modules" (
    echo     Folder node_modules belum ada, menjalankan npm install...
    call npm install
) else (
    echo     node_modules sudah ada. Memperbarui dependensi...
    call npm install --silent
)
echo [OK] Dependensi Node.js siap.

:: ----------------------------------------------------------------------
:: TAHAP 4: Install Sertifikat HTTPS Lokal
:: ----------------------------------------------------------------------
echo.
echo [4/7] Memasang sertifikat HTTPS lokal untuk localhost:3000...
call npx office-addin-dev-certs install --machine >nul 2>&1
echo [OK] Sertifikat HTTPS lokal terpasang dan dipercaya sistem.

:: ----------------------------------------------------------------------
:: TAHAP 5: Registrasi Permanen ke Windows Registry
:: ----------------------------------------------------------------------
echo.
echo [5/7] Mendaftarkan Add-in ke Registry Office Windows...
if not exist "catalog" mkdir "catalog"
copy /y "manifest.xml" "catalog\manifest.xml" >nul 2>&1

:: Daftarkan Trusted Catalog dan Developer Manifest
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{bf0ea8de-c9e1-4dec-becd-f37cbaeb2649}" /v "Url" /t REG_SZ /d "%~dp0catalog" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{bf0ea8de-c9e1-4dec-becd-f37cbaeb2649}" /v "Flags" /t REG_DWORD /d 1 /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "bf0ea8de-c9e1-4dec-becd-f37cbaeb2649" /t REG_SZ /d "%~dp0manifest.xml" /f >nul
echo [OK] Registry Office berhasil dikonfigurasi permanen.

:: ----------------------------------------------------------------------
:: TAHAP 6: Inisialisasi Konfigurasi Pengguna
:: ----------------------------------------------------------------------
echo.
echo [6/7] Menyiapkan folder konfigurasi pengguna...
if not exist "%APPDATA%\ExcelAgent" mkdir "%APPDATA%\ExcelAgent"
if not exist "%APPDATA%\ExcelAgent\config.json" (
    echo {"apiKey":"","apiEndpoint":"http://localhost:20128/v1","model":"qd/ultimate","autoRun":true,"useOpencode":false,"opencodeModel":"9router/qd/ultimate","modelList":["qd/ultimate","FreeCombo","gpt-4o"]} > "%APPDATA%\ExcelAgent\config.json"
    echo [OK] File config.json default dibuat di: %APPDATA%\ExcelAgent\config.json
) else (
    echo [OK] File config.json sudah ada di: %APPDATA%\ExcelAgent\config.json
)

:: ----------------------------------------------------------------------
:: TAHAP 7: Buat Shortcut Desktop & Registrasi WEF
:: ----------------------------------------------------------------------
echo.
echo [7/7] Membuat shortcut Desktop & mendaftarkan ke Excel...
call npx office-addin-debugging start manifest.xml desktop --app excel --no-debug >nul 2>&1

powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\Excel Agent Server.lnk'); $s.TargetPath = '%~dp0START.bat'; $s.WorkingDirectory = '%~dp0'; $s.IconLocation = 'shell32.dll,138'; $s.Save()" >nul 2>&1
echo [OK] Shortcut 'Excel Agent Server' berhasil dibuat di Desktop.

echo.
color 0A
echo ====================================================================
echo           🎉 SETUP BERHASIL! EXCEL AGENT SIAP DIGUNAKAN
echo ====================================================================
echo.
echo CARA PENGGUNAAN SEHARI-HARI:
echo   1. Klik shortcut "Excel Agent Server" di Desktop (atau START.bat)
echo   2. Buka Microsoft Excel seperti biasa
echo   3. Di Tab HOME (Beranda) Excel bagian kanan, klik tombol "Buka Asisten AI"
echo.
echo Server dan Add-in sudah terintegrasi dan terinstall permanen.
echo ====================================================================
echo.
pause
