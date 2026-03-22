# OneInk - ARM64 COM Registration Script
# Bypasses RegAsm by directly writing registry entries
# Run as Administrator after deploying files

$ErrorActionPreference = "Stop"

$AddInCLSID = "{E1F2A3B4-CF2D-409B-B65A-BDBACB9F21DC}"
$AddInAppID = "{E1F2A3B4-CF2D-409B-B65A-BDBACB9F21DC}"
$InstallPath = "C:\Program Files\OneInk-arm64"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "OneInk - ARM64 Registration" -ForegroundColor Cyan
Write-Host "========================================"
Write-Host

# 1. Check admin
Write-Host "[1/3] Checking Administrator privileges..." -ForegroundColor Yellow
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "[ERROR] Please run as Administrator." -ForegroundColor Red
    exit 1
}
Write-Host "  OK" -ForegroundColor Green

# 2. Clean old registry entries
Write-Host "[2/3] Cleaning old registry entries..." -ForegroundColor Yellow
$hklmPaths = @(
    "HKLM:\SOFTWARE\Classes\CLSID\$AddInCLSID",
    "HKLM:\SOFTWARE\Classes\AppID\$AddInAppID"
)
foreach ($path in $hklmPaths) {
    if (Test-Path $path) { Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue }
}

# 3. Write registry entries directly (bypass RegAsm)
Write-Host "[3/3] Writing registry entries..." -ForegroundColor Yellow

# AppID with DllSurrogate
New-Item -Path "HKLM:\SOFTWARE\Classes\AppID\$AddInAppID" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\AppID\$AddInAppID" -Name DllSurrogate -Value ""

# CLSID
New-Item -Path "HKLM:\SOFTWARE\Classes\CLSID\$AddInCLSID" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\CLSID\$AddInCLSID" -Name AppID -Value $AddInAppID
Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\CLSID\$AddInCLSID" -Name "(Default)" -Value "OneInk.AddIn"

# InprocServer32
$inprocPath = "HKLM:\SOFTWARE\Classes\CLSID\$AddInCLSID\InprocServer32"
New-Item -Path $inprocPath -Force | Out-Null
Set-ItemProperty -Path $inprocPath -Name "(Default)" -Value "mscoree.dll"
Set-ItemProperty -Path $inprocPath -Name ThreadingModel -Value "Both"
Set-ItemProperty -Path $inprocPath -Name CodeBase -Value "file:///$($InstallPath.Replace('\', '/'))/OneInk.dll"

# AddIn entry
$addinRegPath = "HKCU:\SOFTWARE\Microsoft\Office\OneNote\AddIns\OneInk.AddIn"
if (-not (Test-Path $addinRegPath)) { New-Item -Path $addinRegPath -Force | Out-Null }
Set-ItemProperty -Path $addinRegPath -Name LoadBehavior -Value 3 -Type DWord
Set-ItemProperty -Path $addinRegPath -Name FriendlyName -Value "OneInk"
Set-ItemProperty -Path $addinRegPath -Name Description -Value "OneInk - OneNote Ink Operations COM AddIn"

Write-Host "  CLSID registered" -ForegroundColor Green
Write-Host "  AppID registered" -ForegroundColor Green
Write-Host "  AddIn entry registered" -ForegroundColor Green

# 4. Verify
Write-Host
Write-Host "[4/4] Verification..." -ForegroundColor Yellow
$codeBase = (Get-ItemProperty "HKLM:\SOFTWARE\Classes\CLSID\$AddInCLSID\InprocServer32" 'CodeBase' -ErrorAction SilentlyContinue).CodeBase
$loadBehavior = (Get-ItemProperty $addinRegPath 'LoadBehavior' -ErrorAction SilentlyContinue).LoadBehavior
Write-Host "  CodeBase    : $codeBase" -ForegroundColor $(if ($codeBase) { "Green" } else { "Red" })
Write-Host "  LoadBehavior: $loadBehavior" -ForegroundColor $(if ($loadBehavior -eq 3) { "Green" } else { "Red" })

Write-Host
Write-Host "========================================" -ForegroundColor Green
Write-Host "Registration complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host
Write-Host "Please restart OneNote to load the add-in." -ForegroundColor Cyan
