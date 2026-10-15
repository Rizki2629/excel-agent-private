@echo off
title Install Excel Agent - Permanen di Excel
color 0A
cls
echo ===================================================
echo   INSTALASI PERMANEN EXCEL AGENT KE EXCEL DESKTOP
echo ===================================================
echo.
echo Script ini akan:
echo   1. Mendaftarkan manifest Excel Agent ke Registry Office (Trusted Catalog)
echo   2. Mendaftarkan developer manifest ke Office WEF
echo   3. Memasang sertifikat HTTPS localhost
echo.

cd /d "%~dp0"

echo [1/3] Memasang sertifikat HTTPS lokal...
call npx office-addin-dev-certs install --machine >nul 2>&1
echo [OK] Sertifikat HTTPS terpasang.

echo [2/3] Mendaftarkan Add-in ke Windows Registry...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{bf0ea8de-c9e1-4dec-becd-f37cbaeb2649}" /v "Url" /t REG_SZ /d "%~dp0catalog" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{bf0ea8de-c9e1-4dec-becd-f37cbaeb2649}" /v "Flags" /t REG_DWORD /d 1 /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "bf0ea8de-c9e1-4dec-becd-f37cbaeb2649" /t REG_SZ /d "%~dp0manifest.xml" /f >nul
echo [OK] Registry Office berhasil didaftarkan.

echo [3/3] Mendaftarkan Add-in ke Office WEF cache...
call npx office-addin-debugging start manifest.xml desktop --app excel --no-debug >nul 2>&1
echo [OK] Registrasi selesai.

echo.
echo ===================================================
echo [BERHASIL] Excel Agent telah TERINSTALL PERMANEN!
echo ===================================================
echo.
echo Mulai sekarang:
echo   1. Anda cukup jalankan server (START_SERVER.bat / node server.js)
echo   2. Buka Excel kapan saja
echo   3. Tombol "Buka Asisten AI" sudah SELALU ADA di Tab HOME (Beranda) Excel
echo.
pause
