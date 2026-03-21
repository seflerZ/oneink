# OneInk - Install Script
# Copies built output from local bin to Program Files and registers the COM AddIn
# Usage: install.ps1 [x86|x64|arm64]

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("x86", "x64", "arm64")]
    [string]$Platform
)

. "$PSScriptRoot\config.ps1"

$SrcDir = Join-Path $Global:ProjectRoot "OneInk\bin\$Platform\Release"
$DstDir = if ($Platform -eq "x86") { "C:\Program Files\OneInk" } elseif ($Platform -eq "x64") { $Global:InstallPathX64 } else { $Global:InstallPathARM64 }
$RegAsm = if ($Platform -eq "x86") { $Global:RegAsmX86 } else { $Global:RegAsmX64 }

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "OneInk - Install ($Platform)" -ForegroundColor Cyan
Write-Host "========================================"
Write-Host

# 1. Check source
if (-not (Test-Path (Join-Path $SrcDir "OneInk.dll"))) {
    Write-Host "[ERROR] OneInk.dll not found at $SrcDir" -ForegroundColor Red
    Write-Host "Please run build-$Platform.ps1 first." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# 2. Copy to Program Files
Write-Host "[1/3] Copying to Program Files..." -ForegroundColor Yellow
if (-not (Test-Path $DstDir)) {
    New-Item -ItemType Directory -Path $DstDir -Force | Out-Null
}
Copy-Item -Path "$SrcDir\*" -Destination $DstDir -Recurse -Force
Write-Host "[OK] Files copied." -ForegroundColor Green

# 3. Set folder permissions
Write-Host "[2/3] Setting folder permissions..." -ForegroundColor Yellow
cmd /c "icacls `"$DstDir`" /grant Everyone:(OI)(CI)RX /T"
Write-Host "[OK] Permissions set." -ForegroundColor Green

# 4. Register COM AddIn
Write-Host "[3/3] Registering COM AddIn..." -ForegroundColor Yellow

# Clean HKLM registry entries BEFORE regasm
$hklmPaths = @(
    "HKLM:\SOFTWARE\Classes\CLSID\$Global:AddInCLSID",
    "HKLM:\SOFTWARE\Classes\AppID\$Global:AddInAppID"
)
foreach ($path in $hklmPaths) {
    if (Test-Path $path) { Remove-Item $path -Recurse -Force }
}

Set-Location $DstDir
& $RegAsm /codebase "OneInk.dll"

# HKLM AppID + DllSurrogate
New-Item -Path "HKLM:\SOFTWARE\Classes\AppID\$Global:AddInAppID" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\AppID\$Global:AddInAppID" -Name DllSurrogate -Value ""
Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\CLSID\$Global:AddInCLSID" -Name AppID -Value $Global:AddInAppID

# HKLM InprocServer32
$inprocPath = "HKLM:\SOFTWARE\Classes\CLSID\$Global:AddInCLSID\InprocServer32"
New-Item -Path $inprocPath -Force | Out-Null
Set-ItemProperty -Path $inprocPath -Name "(Default)" -Value "mscoree.dll"
Set-ItemProperty -Path $inprocPath -Name ThreadingModel -Value "Both"
Set-ItemProperty -Path $inprocPath -Name CodeBase -Value "file:///$($DstDir.Replace('\', '/'))/OneInk.dll"

# HKCU OneNote AddIn (MUST be last — regasm resets LoadBehavior to 2)
$addinRegPath = "HKCU:\SOFTWARE\Microsoft\Office\OneNote\AddIns\OneInk.AddIn"
New-Item -Path $addinRegPath -Force | Out-Null
Set-ItemProperty -Path $addinRegPath -Name LoadBehavior -Value 3 -Type DWord
Set-ItemProperty -Path $addinRegPath -Name FriendlyName -Value "OneInk"
Set-ItemProperty -Path $addinRegPath -Name Description -Value "OneInk - OneNote Ink Operations COM AddIn"
Set-ItemProperty -Path $addinRegPath -Name CommandLineSafe -Value 1 -Type DWord

$lb = (Get-ItemProperty -Path $addinRegPath -Name LoadBehavior).LoadBehavior
Write-Host "[OK] HKCU LoadBehavior = $lb" -ForegroundColor Green

Write-Host
Write-Host "========================================" -ForegroundColor Green
Write-Host "[OK] Install completed!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "1. Close OneNote completely (check Task Manager)" -ForegroundColor White
Write-Host "2. Restart OneNote" -ForegroundColor White
Write-Host "3. Check the OneInk tab" -ForegroundColor White
Write-Host
Read-Host "Press Enter to exit"
