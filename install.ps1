# OneInk - Install Script (no build)
# Installs to Program Files and registers COM (requires admin)
#
# Usage:
#   .\install.ps1                       # uses default DLL: C:\Program Files\OneInk\OneInk.dll
#   .\install.ps1 -DllPath "path\to\OneInk.dll"

param(
    [string]$DllPath = "$env:ProgramFiles\OneInk\OneInk.dll",

    [ValidateSet("x86", "x64", "arm64")]
    [string]$Platform = "x64"
)

. "$PSScriptRoot\config.ps1"

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "OneInk Install ($Platform)" -ForegroundColor Cyan
Write-Host "========================================"
Write-Host

# Source DLL - default to existing install path, can override via parameter
$SourceDll = $DllPath
if (-not (Test-Path $SourceDll)) {
    Write-Host "[ERROR] DLL not found: $SourceDll" -ForegroundColor Red
    Write-Host "Please specify -DllPath to your OneInk.dll" -ForegroundColor Yellow
    exit 1
}

$OutputPath = $Global:InstallPath

# 1. Copy to Program Files
Write-Host "[1/3] Copying to $OutputPath ..." -ForegroundColor Yellow
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}
$srcBin = Split-Path $SourceDll
robocopy $srcBin $OutputPath /MIR /R:3 /W:1 | Out-Null
# Set Everyone read permissions
cmd /c "icacls `"$OutputPath`" /grant Everyone:(OI)(CI)RX /T" 2>$null
Write-Host "[OK]" -ForegroundColor Green

# 2. Register COM
Write-Host "[2/3] Registering COM AddIn..." -ForegroundColor Yellow

# Clean HKLM before regasm
$hklmPaths = @(
    "HKLM:\SOFTWARE\Classes\CLSID\$Global:AddInCLSID",
    "HKLM:\SOFTWARE\Classes\AppID\$Global:AddInAppID"
)
foreach ($path in $hklmPaths) {
    if (Test-Path $path) { Remove-Item $path -Recurse -Force }
}

$RegAsm = if ($Platform -eq "x86") { $Global:RegAsmX86 } elseif ($Platform -eq "arm64") { $Global:RegAsmARM64 } else { $Global:RegAsmX64 }
Set-Location $OutputPath
& $RegAsm /codebase /tlb "$OutputPath\OneInk.dll"
if ($LASTEXITCODE -ne 0) { Write-Host "[WARNING] regasm exited with code $LASTEXITCODE" -ForegroundColor Yellow }

# HKLM AppID + DllSurrogate
New-Item -Path "HKLM:\SOFTWARE\Classes\AppID\$Global:AddInAppID" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\AppID\$Global:AddInAppID" -Name DllSurrogate -Value ""
Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\CLSID\$Global:AddInCLSID" -Name AppID -Value $Global:AddInAppID

# HKLM InprocServer32 with all required entries
$inprocPath = "HKLM:\SOFTWARE\Classes\CLSID\$Global:AddInCLSID\InprocServer32"
New-Item -Path $inprocPath -Force | Out-Null
Set-ItemProperty -Path $inprocPath -Name "(Default)" -Value "mscoree.dll"
Set-ItemProperty -Path $inprocPath -Name ThreadingModel -Value "Both"
Set-ItemProperty -Path $inprocPath -Name Class -Value "OneInk.AddIn"
Set-ItemProperty -Path $inprocPath -Name Assembly -Value "OneInk, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
Set-ItemProperty -Path $inprocPath -Name RuntimeVersion -Value "v4.0.30319"
Set-ItemProperty -Path $inprocPath -Name CodeBase -Value "file:///$($OutputPath.Replace('\', '/'))/OneInk.dll"

# HKCU AddIn entry (must be last - regasm resets LoadBehavior)
$addinRegPath = "HKCU:\SOFTWARE\Microsoft\Office\OneNote\AddIns\OneInk.AddIn"
New-Item -Path $addinRegPath -Force | Out-Null
Set-ItemProperty -Path $addinRegPath -Name LoadBehavior -Value 3 -Type DWord
Set-ItemProperty -Path $addinRegPath -Name FriendlyName -Value "OneInk"
Set-ItemProperty -Path $addinRegPath -Name Description -Value "OneInk - OneNote Ink Operations COM AddIn"
Set-ItemProperty -Path $addinRegPath -Name CommandLineSafe -Value 1 -Type DWord
Write-Host "[OK]" -ForegroundColor Green

# 3. Verify
Write-Host "[3/3] Verification..." -ForegroundColor Yellow
Get-ChildItem $OutputPath -Recurse -File | Select-Object Name, @{N="SizeKB";E={[Math]::Round($_.Length/1KB,1)}} | Format-Table -AutoSize
Write-Host "DLL: $OutputPath\OneInk.dll" -ForegroundColor Green
Write-Host "[OK] Installation complete!" -ForegroundColor Green

Write-Host
Write-Host "Restart OneNote to load the add-in." -ForegroundColor Cyan
