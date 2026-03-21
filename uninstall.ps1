# OneInk - Uninstall Script
# Unregisters the COM AddIn and removes files

param(
    [ValidateSet("x86", "x64", "arm64")]
    [string]$Platform = "x64"
)

. "$PSScriptRoot\config.ps1"

$InstallPath = if ($Platform -eq "x86") {
    "C:\Program Files\OneInk"
} elseif ($Platform -eq "x64") {
    $Global:InstallPathX64
} else {
    $Global:InstallPathARM64
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "OneInk - Uninstall ($Platform)" -ForegroundColor Cyan
Write-Host "========================================"
Write-Host

# 1. Unregister COM
Write-Host "[1/3] Unregistering COM AddIn..." -ForegroundColor Yellow
$RegAsm = if ($Platform -eq "x86") { $Global:RegAsmX86 } else { $Global:RegAsmX64 }
$addinDll = Join-Path $InstallPath "OneInk.dll"
if (Test-Path $addinDll) {
    Set-Location $InstallPath
    & $RegAsm /u $addinDll
    Write-Host "  regasm /u done." -ForegroundColor Green
} else {
    Write-Host "  DLL not found at $addinDll, skipping regasm /u." -ForegroundColor Yellow
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
        Write-Host "  Not found (already clean): $path" -ForegroundColor Gray
    }
}

# 3. Remove install directory
Write-Host "[3/3] Removing install directory..." -ForegroundColor Yellow
if (Test-Path $InstallPath) {
    Remove-Item $InstallPath -Recurse -Force
    Write-Host "  Removed: $InstallPath" -ForegroundColor Green
} else {
    Write-Host "  Not found (already clean): $InstallPath" -ForegroundColor Gray
}

Write-Host
Write-Host "========================================" -ForegroundColor Green
Write-Host "[OK] Uninstall complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host
Write-Host "Restart OneNote to confirm the add-in is removed." -ForegroundColor Yellow
