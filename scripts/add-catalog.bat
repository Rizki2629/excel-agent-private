@echo off
echo Adding Excel Agent to Trusted Add-in Catalogs...
echo.

:: Create registry key for trusted catalog
reg add "HKCU\Software\Microsoft\Office\16.0\excel\options" /v "NetworkCatalogs" /t REG_SZ /d "https://localhost:3000" /f

echo.
echo Registry updated. Please:
echo 1. Restart Excel
echo 2. Go to Insert ^> Office Add-ins ^> My Add-ins
echo 3. You should see "Shared Folder" tab
echo 4. Click Add next to Excel Agent
echo.
pause
