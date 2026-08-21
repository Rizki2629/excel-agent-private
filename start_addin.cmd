@echo off
setlocal
cd /d "%~dp0"

netstat -ano | findstr /R /C:":3000 .*LISTENING" >nul
if errorlevel 1 (
    schtasks /run /tn "ExcelAgentServer" >nul 2>&1
)

for /l %%I in (1,1,15) do (
    curl.exe --insecure --silent --fail --max-time 2 "https://localhost:3000/taskpane.html" >nul 2>&1
    if not errorlevel 1 goto server_ready
    ping.exe 127.0.0.1 -n 2 >nul
)

echo Excel Agent server gagal aktif di https://localhost:3000.
exit /b 1

:server_ready

call npx office-addin-debugging start manifest.xml desktop --app excel --no-debug
