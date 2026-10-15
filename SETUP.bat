@echo off
title Setup Excel Agent - Komputer Baru
color 0B
cls

echo ====================================================================
echo             SETUP OTOMATIS EXCEL AGENT (KOMPUTER BARU)
echo ====================================================================
echo.
echo Script ini akan menyiapkan seluruh kebutuhan Excel Agent secara otomatis:
echo   [1] Cek & siapkan Node.js + NPM (versi minimum dicek)
echo   [2] Cek & siapkan Python + library pywin32 (versi terkunci)
echo   [3] Install dependencies project (npm ci - deterministik)
echo   [4] Install sertifikat HTTPS lokal (bebas error SSL)
echo   [5] Registrasi Add-in ke Registry Office Windows (Permanen)
echo   [6] Inisialisasi folder & file konfigurasi (termasuk profil identitas)
echo   [7] Buat shortcut di Desktop
echo   [8] Verifikasi instalasi (self-test)
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
echo [1/8] Memeriksa Node.js...
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
:: Versi minimum v18 agar hasil sama di semua komputer
for /f "tokens=1 delims=." %%a in ("%NODE_VER:~1%") do set NODE_MAJOR=%%a
if %NODE_MAJOR% lss 18 (
    color 0E
    echo [PERINGATAN] Node.js %NODE_VER% terlalu lama. Dibutuhkan v18 atau lebih baru.
    echo Silakan update dari https://nodejs.org/ lalu jalankan ulang script ini.
    pause
    exit /b 1
)

:: ----------------------------------------------------------------------
:: TAHAP 2: Periksa Python & pywin32
:: ----------------------------------------------------------------------
echo.
echo [2/8] Memeriksa Python...
where python >nul 2>nul
if %errorlevel% neq 0 (
    color 0E
    echo [PERINGATAN] Python belum terdeteksi di PATH sistem.
    echo Eksekusi script win32com membutuhkan Python 3.10+.
    echo Pastikan Python terinstall dan centang "Add Python to PATH" saat install.
) else (
    for /f "tokens=*" %%i in ('python --version') do set PY_VER=%%i
    echo [OK] Python terdeteksi: %PY_VER%
    python -c "import sys; exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>nul
    if %errorlevel% neq 0 (
        color 0E
        echo [PERINGATAN] Python %PY_VER% terlalu lama. Dibutuhkan Python 3.10 atau lebih baru.
    )
    echo     Memeriksa pustaka pywin32 (win32com)...
    python -c "import win32com.client" >nul 2>nul
    if %errorlevel% neq 0 (
        echo     Menginstall pywin32 versi terkunci via pip...
        pip install -r requirements.txt
        if %errorlevel% equ 0 (
            echo [OK] pywin32 berhasil diinstall.
        ) else (
            echo [PERINGATAN] Gagal install pywin32 otomatis. Jalankan: pip install -r requirements.txt
        )
    ) else (
        echo [OK] pywin32 sudah siap.
    )
)

:: ----------------------------------------------------------------------
:: TAHAP 3: Install NPM Dependencies
:: ----------------------------------------------------------------------
echo.
echo [3/8] Menginstall dependensi project (npm ci - versi terkunci)...
echo     Menggunakan package-lock.json agar versi SAMA di semua komputer.
call npm ci
if %errorlevel% neq 0 (
    color 0C
    echo [ERROR] npm ci gagal. Cek koneksi internet lalu jalankan ulang SETUP.bat.
    pause
    exit /b 1
)
echo [OK] Dependensi Node.js siap (versi terkunci).

:: ----------------------------------------------------------------------
:: TAHAP 4: Install Sertifikat HTTPS Lokal
:: ----------------------------------------------------------------------
echo.
echo [4/8] Memasang sertifikat HTTPS lokal untuk localhost:3000...
call npx office-addin-dev-certs install --machine >nul 2>&1
echo [OK] Sertifikat HTTPS lokal terpasang dan dipercaya sistem.

:: ----------------------------------------------------------------------
:: TAHAP 5: Registrasi Permanen ke Windows Registry
:: ----------------------------------------------------------------------
echo.
echo [5/8] Mendaftarkan Add-in ke Registry Office Windows...
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
echo [6/8] Menyiapkan folder konfigurasi pengguna...
if not exist "%APPDATA%\ExcelAgent" mkdir "%APPDATA%\ExcelAgent"
if not exist "%APPDATA%\ExcelAgent\config.json" (
    echo {"apiKey":"","apiEndpoint":"http://localhost:20128/v1","model":"qd/ultimate","autoRun":true,"useOpencode":false,"opencodeModel":"9router/qd/ultimate","modelList":["qd/ultimate","FreeCombo","gpt-4o"]} > "%APPDATA%\ExcelAgent\config.json"
    echo [OK] File config.json default dibuat di: %APPDATA%\ExcelAgent\config.json
) else (
    echo [OK] File config.json sudah ada di: %APPDATA%\ExcelAgent\config.json
)
:: Profil identitas: agar model mengenali pemilik di semua komputer.
:: Tidak ditimpa bila sudah ada (bisa diubah lewat Pengaturan di taskpane).
if not exist "%APPDATA%\ExcelAgent\profil.json" (
    echo {"nama":"Muhammad Rizki Pratama, S.Pd","peran":"Guru kelas IV, SDN Grogol Utara 09","zonaWaktu":"Asia/Jakarta (WIB)","bahasa":"Bahasa Indonesia, santai tapi jelas","gayaJawab":"Singkat, langsung ke hasil. Tidak bertele-tele."} > "%APPDATA%\ExcelAgent\profil.json"
    echo [OK] File profil.json default dibuat di: %APPDATA%\ExcelAgent\profil.json
) else (
    echo [OK] File profil.json sudah ada, tidak diubah.
)

:: ----------------------------------------------------------------------
:: TAHAP 7: Buat Shortcut Desktop & Registrasi WEF
:: ----------------------------------------------------------------------
echo.
echo [7/8] Membuat shortcut Desktop & mendaftarkan ke Excel...
call npx office-addin-debugging start manifest.xml desktop --app excel --no-debug >nul 2>&1

powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\Excel Agent Server.lnk'); $s.TargetPath = '%~dp0START.bat'; $s.WorkingDirectory = '%~dp0'; $s.IconLocation = 'shell32.dll,138'; $s.Save()" >nul 2>&1
echo [OK] Shortcut 'Excel Agent Server' berhasil dibuat di Desktop.

:: ----------------------------------------------------------------------
:: TAHAP 8: Verifikasi instalasi (self-test)
:: ----------------------------------------------------------------------
echo.
echo [8/8] Verifikasi instalasi (self-test)...
set TEST_OK=1

node --check server.js >nul 2>&1
set CHK=%errorlevel%
if %CHK% equ 0 echo [OK] server.js valid.
if %CHK% neq 0 echo [GAGAL] server.js rusak.
if %CHK% neq 0 set TEST_OK=0

node --check taskpane/taskpane.js >nul 2>&1
set CHK=%errorlevel%
if %CHK% equ 0 echo [OK] taskpane.js valid.
if %CHK% neq 0 echo [GAGAL] taskpane.js rusak.
if %CHK% neq 0 set TEST_OK=0

node --check hermes_providers.js >nul 2>&1
set CHK=%errorlevel%
if %CHK% equ 0 echo [OK] hermes_providers.js valid.
if %CHK% neq 0 echo [GAGAL] hermes_providers.js rusak.
if %CHK% neq 0 set TEST_OK=0

if exist "manifest.xml" (
    echo [OK] manifest.xml ada.
) else (
    echo [GAGAL] manifest.xml hilang.
    set TEST_OK=0
)

python -c "import win32com.client" >nul 2>&1
set CHK=%errorlevel%
if %CHK% equ 0 echo [OK] pywin32 bisa diimport.
if %CHK% neq 0 echo [GAGAL] pywin32 tidak bisa diimport. Jalankan: pip install -r requirements.txt
if %CHK% neq 0 set TEST_OK=0

if %TEST_OK% equ 1 goto setupsuccess
color 0C
echo.
echo ====================================================================
echo        SETUP SELESAI TAPI ADA YANG GAGAL - CEK TANDA [GAGAL] DI ATAS
echo ====================================================================
echo Perbaiki yang gagal, lalu jalankan ulang SETUP.bat.
echo.
pause
exit /b 1

:setupsuccess
echo.
color 0A
echo ====================================================================
echo           SETUP BERHASIL! EXCEL AGENT SIAP DIGUNAKAN
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
