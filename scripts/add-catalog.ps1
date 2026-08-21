# Excel Agent - Add Trusted Catalog
# Run as Administrator

Write-Host "========================================" -ForegroundColor Green
Write-Host "Excel Agent - Add-in Catalog Setup" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# Check if running as admin
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

if (-not $isAdmin) {
    Write-Host "Please run as Administrator!" -ForegroundColor Red
    Write-Host "Right-click PowerShell > Run as Administrator" -ForegroundColor Yellow
    pause
    exit
}

# Add network catalog
$catalogUrl = "https://localhost:3000"
$regPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"

Write-Host "Adding catalog: $catalogUrl" -ForegroundColor Yellow

try {
    # Check if NetworkCatalogs exists
    if (Test-Path "$regPath\NetworkCatalogs") {
        # Add to existing
        $catalogs = Get-ItemProperty -Path $regPath -Name "NetworkCatalogs" -ErrorAction SilentlyContinue
        if ($catalogs.NetworkCatalogs -notlike "*$catalogUrl*") {
            Set-ItemProperty -Path $regPath -Name "NetworkCatalogs" -Value "$($catalogs.NetworkCatalogs);$catalogUrl"
        }
    } else {
        New-ItemProperty -Path $regPath -Name "NetworkCatalogs" -Value $catalogUrl -PropertyType String -Force
    }
    
    Write-Host "Catalog added successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "1. Restart Excel" -ForegroundColor White
    Write-Host "2. Go to Insert > Office Add-ins > My Add-ins" -ForegroundColor White
    Write-Host "3. Click 'Shared Folder' tab" -ForegroundColor White
    Write-Host "4. Click Add next to Excel Agent" -ForegroundColor White
    
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
}

Write-Host ""
pause
