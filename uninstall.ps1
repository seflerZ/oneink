# OneInk - Uninstall Script
# Unregisters the COM AddIn and removes files installed via .\deploy.ps1 -Mode Production
#
# Usage:
#   .\uninstall.ps1 -Platform x64

param(
    [ValidateSet("x86", "x64", "arm64")]
    [string]$Platform = "x64"
)

. "$PSScriptRoot\config.ps1"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "OneInk - Uninstall ($Platform)" -ForegroundColor Cyan
Write-Host "========================================"
Write-Host

$InstallPath = $Global:InstallPath

# 1. Unregister COM
Write-Host "[1/3] Unregistering COM AddIn..." -ForegroundColor Yellow
$RegAsm = if ($Platform -eq "x86") { $Global:RegAsmX86 } elseif ($Platform -eq "arm64") { $Global:RegAsmARM64 } else { $Global:RegAsmX64 }
$addinDll = Join-Path $InstallPath "OneInk.dll"
if (Test-Path $addinDll) {
    Set-Location $InstallPath
    & $RegAsm /u $addinDll
    Set-Location $PSScriptRoot
    Write-Host "  regasm /u done." -ForegroundColor Green
} else {
    Write-Host "  DLL not found, skipping regasm /u." -ForegroundColor Yellow
}

# 2. Remove registry entries
Write-Host "[2/3] Removing registry entries..." -ForegroundColor Yellow
$regPaths = @(
    "HKLM:\SOFTWARE\Classes\CLSID\$Global:AddInCLSID",
    "HKLM:\SOFTWARE\Classes\AppID\$Global:AddInAppID",
    "HKCU:\SOFTWARE\Microsoft\Office\OneNote\AddIns\OneInk.AddIn"
)
foreach ($path in $regPaths) {
    if (Test-Path $path) {
        Remove-Item $path -Recurse -Force
        Write-Host "  Removed: $path" -ForegroundColor Green
    } else {
        Write-Host "  Not found: $path" -ForegroundColor Gray
    }
}

# 3. Remove install directory
Write-Host "[3/3] Removing install directory..." -ForegroundColor Yellow
if (Test-Path $InstallPath) {
    Remove-Item $InstallPath -Recurse -Force
    Write-Host "  Removed: $InstallPath" -ForegroundColor Green
} else {
    Write-Host "  Not found: $InstallPath" -ForegroundColor Gray
}

Write-Host
Write-Host "Restart OneNote to confirm the add-in is removed." -ForegroundColor Cyan
