# OneInk - x64 Deploy Script
# Builds, copies to Program Files, and registers the COM AddIn

. "$PSScriptRoot\config.ps1"

$Platform = "x64"
$OutputPath = $Global:InstallPathX64

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "OneInk - x64 Deployment" -ForegroundColor Cyan
Write-Host "========================================"
Write-Host

# 1. Build
Write-Host "[1/4] Building OneInk (x64)..." -ForegroundColor Yellow
& $Global:MSBuildPath $Global:ProjectFile /p:Configuration=Release /p:Platform=$Platform /p:OutputPath="$OutputPath\" /t:Rebuild
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] Build failed!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Host "[OK] Build completed." -ForegroundColor Green
Write-Host

# 2. Set folder permissions
Write-Host "[2/4] Setting folder permissions..." -ForegroundColor Yellow
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}
# Use cmd /c to avoid PowerShell interpreting (OI)(CI) flags
cmd /c "icacls `"$OutputPath`" /grant Everyone:(OI)(CI)RX /T"
Write-Host "[OK] Permissions set." -ForegroundColor Green
Write-Host

# 3. Register COM AddIn
Write-Host "[3/4] Registering COM AddIn..." -ForegroundColor Yellow

# Clean HKLM registry entries BEFORE regasm
$hklmPaths = @(
    "HKLM:\SOFTWARE\Classes\CLSID\$Global:AddInCLSID",
    "HKLM:\SOFTWARE\Classes\AppID\$Global:AddInAppID"
)
foreach ($path in $hklmPaths) {
    if (Test-Path $path) { Remove-Item $path -Recurse -Force }
}

$addinDll = Join-Path $OutputPath "OneInk.dll"
Set-Location $OutputPath
& $Global:RegAsmX64 /codebase $addinDll

# HKLM AppID + DllSurrogate
New-Item -Path "HKLM:\SOFTWARE\Classes\AppID\$Global:AddInAppID" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\AppID\$Global:AddInAppID" -Name DllSurrogate -Value ""
Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\CLSID\$Global:AddInCLSID" -Name AppID -Value $Global:AddInAppID

# HKLM InprocServer32
$inprocPath = "HKLM:\SOFTWARE\Classes\CLSID\$Global:AddInCLSID\InprocServer32"
New-Item -Path $inprocPath -Force | Out-Null
Set-ItemProperty -Path $inprocPath -Name "(Default)" -Value "mscoree.dll"
Set-ItemProperty -Path $inprocPath -Name ThreadingModel -Value "Both"
Set-ItemProperty -Path $inprocPath -Name CodeBase -Value "file:///$($OutputPath.Replace('\', '/'))/OneInk.dll"

# HKCU OneNote AddIn (MUST be last — regasm resets LoadBehavior to 2)
$addinRegPath = "HKCU:\SOFTWARE\Microsoft\Office\OneNote\AddIns\OneInk.AddIn"
New-Item -Path $addinRegPath -Force | Out-Null
Set-ItemProperty -Path $addinRegPath -Name LoadBehavior -Value 3 -Type DWord
Set-ItemProperty -Path $addinRegPath -Name FriendlyName -Value "OneInk"
Set-ItemProperty -Path $addinRegPath -Name Description -Value "OneInk - OneNote Ink Operations COM AddIn"
Set-ItemProperty -Path $addinRegPath -Name CommandLineSafe -Value 1 -Type DWord

$lb = (Get-ItemProperty -Path $addinRegPath -Name LoadBehavior).LoadBehavior
Write-Host "[OK] HKCU LoadBehavior = $lb" -ForegroundColor Green

# 4. Verify install
Write-Host "[4/4] Verifying installation..." -ForegroundColor Yellow
$installedDll = Join-Path $OutputPath "OneInk.dll"
$installedRes = Join-Path $OutputPath "Resources\Logo.png"
if (Test-Path $installedDll) {
    Write-Host "  DLL: $installedDll ($([Math]::Round((Get-Item $installedDll).Length / 1KB, 1)) KB)" -ForegroundColor Green
} else {
    Write-Host "  DLL missing!" -ForegroundColor Red
}
$files = Get-ChildItem $OutputPath -Recurse -File
foreach ($f in $files) {
    $rel = $f.FullName.Substring($OutputPath.Length).TrimStart('\')
    Write-Host "  $rel ($([Math]::Round($f.Length / 1KB, 1)) KB)" -ForegroundColor Green
}

Write-Host
Write-Host "========================================" -ForegroundColor Green
Write-Host "[OK] Deployment completed!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "1. Close OneNote completely (check Task Manager)" -ForegroundColor White
Write-Host "2. Restart OneNote" -ForegroundColor White
Write-Host "3. Check the OneInk tab" -ForegroundColor White
Write-Host
Read-Host "Press Enter to exit"
