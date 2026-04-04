# OneInk - Deploy Script
# Production: copies to Program Files and registers (requires admin)
# Dev: registers from build output directory (no admin required)
#
# Usage:
#   Deploy to Program Files (production):
#     .\deploy.ps1 -Mode Production -Platform x64
#   Register from build dir (development):
#     .\deploy.ps1 -Mode Dev

param(
    [ValidateSet("Production", "Dev")]
    [string]$Mode = "Dev",

    [ValidateSet("x86", "x64", "arm64")]
    [string]$Platform = "x64"
)

. "$PSScriptRoot\config.ps1"

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "OneInk Deploy - $Mode ($Platform)" -ForegroundColor Cyan
Write-Host "========================================"
Write-Host

if ($Mode -eq "Production") {
    # ===================== Production =====================
    # 1. Build
    Write-Host "[1/4] Building OneInk..." -ForegroundColor Yellow
    & $Global:MSBuildPath $Global:ProjectFile /p:Configuration=Release /p:Platform=$Platform /t:Rebuild /v:m | Out-Null
    if ($LASTEXITCODE -ne 0) { Write-Host "[ERROR] Build failed" -ForegroundColor Red; exit 1 }
    Write-Host "[OK] Build completed" -ForegroundColor Green

    # 2. Copy to Program Files
    $OutputPath = $Global:InstallPath
    Write-Host "[2/4] Copying to $OutputPath ..." -ForegroundColor Yellow
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
    $srcBin = Join-Path (Split-Path $Global:ProjectFile) "bin\$Platform\Release"
    robocopy $srcBin $OutputPath /MIR /R:3 /W:1
    # Set Everyone read permissions
    cmd /c "icacls `"$OutputPath`" /grant Everyone:(OI)(CI)RX /T"
    Write-Host "[OK]" -ForegroundColor Green

    # 3. Register COM AddIn
    Write-Host "[3/4] Registering COM AddIn..." -ForegroundColor Yellow
    # Clean HKLM before regasm
    $hklmPaths = @(
        "HKLM:\SOFTWARE\Classes\CLSID\$Global:AddInCLSID",
        "HKLM:\SOFTWARE\Classes\AppID\$Global:AddInAppID"
    )
    foreach ($path in $hklmPaths) { if (Test-Path $path) { Remove-Item $path -Recurse -Force } }

    $RegAsm = if ($Platform -eq "x86") { $Global:RegAsmX86 } elseif ($Platform -eq "arm64") { $Global:RegAsmARM64 } else { $Global:RegAsmX64 }
    Set-Location $OutputPath
    & $RegAsm /codebase /tlb "$OutputPath\OneInk.dll"
    if ($LASTEXITCODE -ne 0) { Write-Host "[WARNING] regasm exited with code $LASTEXITCODE" -ForegroundColor Yellow }

    # HKLM AppID + DllSurrogate
    New-Item -Path "HKLM:\SOFTWARE\Classes\AppID\$Global:AddInAppID" -Force | Out-Null
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\AppID\$Global:AddInAppID" -Name DllSurrogate -Value ""
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\CLSID\$Global:AddInCLSID" -Name AppID -Value $Global:AddInAppID

    # Fix InprocServer32 with all required entries for .NET COM
    $inprocPath = "HKLM:\SOFTWARE\Classes\CLSID\$Global:AddInCLSID\InprocServer32"
    New-Item -Path $inprocPath -Force | Out-Null
    Set-ItemProperty -Path $inprocPath -Name "(Default)" -Value "mscoree.dll"
    Set-ItemProperty -Path $inprocPath -Name ThreadingModel -Value "Both"
    Set-ItemProperty -Path $inprocPath -Name Class -Value "OneInk.AddIn"
    Set-ItemProperty -Path $inprocPath -Name Assembly -Value "OneInk, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
    Set-ItemProperty -Path $inprocPath -Name RuntimeVersion -Value "v4.0.30319"
    Set-ItemProperty -Path $inprocPath -Name CodeBase -Value "file:///$($OutputPath.Replace('\', '/'))/OneInk.dll"

    # HKCU AddIn (must be last — regasm resets LoadBehavior)
    $addinRegPath = "HKCU:\SOFTWARE\Microsoft\Office\OneNote\AddIns\OneInk.AddIn"
    New-Item -Path $addinRegPath -Force | Out-Null
    Set-ItemProperty -Path $addinRegPath -Name LoadBehavior -Value 3 -Type DWord
    Set-ItemProperty -Path $addinRegPath -Name FriendlyName -Value "OneInk"
    Set-ItemProperty -Path $addinRegPath -Name Description -Value "OneInk - OneNote Ink Operations COM AddIn"
    Set-ItemProperty -Path $addinRegPath -Name CommandLineSafe -Value 1 -Type DWord
    Write-Host "[OK]" -ForegroundColor Green

    # 4. Verify
    Write-Host "[4/4] Verification..." -ForegroundColor Yellow
    Get-ChildItem $OutputPath -Recurse -File | Select-Object Name, @{N="SizeKB";E={[Math]::Round($_.Length/1KB,1)}} | Format-Table -AutoSize
    Write-Host "DLL: $OutputPath\OneInk.dll" -ForegroundColor Green
    Write-Host "[OK] Production deployment complete!" -ForegroundColor Green

} else {
    # ===================== Dev =====================
    # Build output dir: OneInk\bin\x64\Release\
    Write-Host "[1/4] Building OneInk..." -ForegroundColor Yellow
    & $Global:MSBuildPath $Global:ProjectFile /p:Configuration=Release /p:Platform=$Platform /t:Rebuild /v:m | Out-Null
    if ($LASTEXITCODE -ne 0) { Write-Host "[ERROR] Build failed" -ForegroundColor Red; exit 1 }
    Write-Host "[OK] Build completed" -ForegroundColor Green

    $BuildDir = Join-Path (Split-Path $Global:ProjectFile) "bin\$Platform\Release"

    # Register DLL
    Write-Host "[2/4] Registering DLL: $BuildDir\OneInk.dll" -ForegroundColor Yellow

    $RegAsm = if ($Platform -eq "x86") { $Global:RegAsmX86 } elseif ($Platform -eq "arm64") { $Global:RegAsmARM64 } else { $Global:RegAsmX64 }
    $dllPath = "$BuildDir\OneInk.dll"

    # Clean HKLM CLSID/AppID (best effort - may need admin)
    try {
        $hklmClsId = "HKLM:\SOFTWARE\Classes\CLSID\$Global:AddInCLSID"
        $hklmAppId = "HKLM:\SOFTWARE\Classes\AppID\$Global:AddInAppID"
        if (Test-Path $hklmClsId) { Remove-Item $hklmClsId -Recurse -Force -EA SilentlyContinue }
        if (Test-Path $hklmAppId) { Remove-Item $hklmAppId -Recurse -Force -EA SilentlyContinue }
    } catch { }

    # Run regasm to register HKCR entries (requires admin)
    # First unregister any existing registration
    & $RegAsm /unregister $dllPath 2>$null
    $regasmScript = "$env:TEMP\OneInk_regasm_$PID.ps1"
    @"
`$dll = "$dllPath"
`$regasm = "$RegAsm"
& `$regasm /codebase /tlb `$dll
"@ | Out-File -FilePath $regasmScript -Encoding UTF8
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell.exe"
    $psi.Arguments = "-ExecutionPolicy Bypass -NoProfile -File `"$regasmScript`""
    $psi.Verb = "RunAs"
    $psi.UseShellExecute = $true
    [System.Diagnostics.Process]::Start($psi) | Out-Null
    # Wait for elevated process to finish
    Start-Sleep -Seconds 3
    Remove-Item $regasmScript -Force -EA SilentlyContinue

    # HKCU: AppID with DllSurrogate (fallback)
    New-Item -Path "HKCU:\SOFTWARE\Classes\AppID\$Global:AddInAppID" -Force | Out-Null
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Classes\AppID\$Global:AddInAppID" -Name DllSurrogate -Value "" -ErrorAction SilentlyContinue

    # HKCU: CLSID with InprocServer32 (fallback)
    $clsidPath = "HKCU:\SOFTWARE\Classes\CLSID\$Global:AddInCLSID"
    New-Item -Path $clsidPath -Force | Out-Null
    Set-ItemProperty -Path $clsidPath -Name AppID -Value $Global:AddInAppID -ErrorAction SilentlyContinue
    $inprocPath = "$clsidPath\InprocServer32"
    New-Item -Path $inprocPath -Force | Out-Null
    Set-ItemProperty -Path $inprocPath -Name "(Default)" -Value "mscoree.dll" -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $inprocPath -Name ThreadingModel -Value "Both" -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $inprocPath -Name CodeBase -Value "file:///$($BuildDir.Replace('\', '/'))/OneInk.dll" -ErrorAction SilentlyContinue

    # HKCU AddIn entry (LoadBehavior=3)
    $addinRegPath = "HKCU:\SOFTWARE\Microsoft\Office\OneNote\AddIns\OneInk.AddIn"
    if (-not (Test-Path $addinRegPath)) { New-Item -Path $addinRegPath -Force | Out-Null }
    Set-ItemProperty -Path $addinRegPath -Name LoadBehavior -Value 3 -Type DWord -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $addinRegPath -Name FriendlyName -Value "OneInk" -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $addinRegPath -Name Description -Value "OneInk - OneNote Ink Operations" -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $addinRegPath -Name CommandLineSafe -Value 1 -Type DWord -ErrorAction SilentlyContinue

    Write-Host "[OK]" -ForegroundColor Green

    Write-Host "[3/4] Verification..." -ForegroundColor Yellow
    $hkcrInproc = (Get-ItemProperty "HKCR:\CLSID\$Global:AddInCLSID\InprocServer32" -EA SilentlyContinue)
    $hkcuLB = (Get-ItemProperty $addinRegPath 'LoadBehavior' -EA SilentlyContinue).LoadBehavior
    if ($hkcrInproc.Class) {
        Write-Host "  HKCR InprocServer32: Class=$(($hkcrInproc.Class))" -ForegroundColor Green
    } else {
        Write-Host "  HKCR InprocServer32: Class=NOT SET (may need admin)" -ForegroundColor Yellow
    }
    Write-Host "  HKCU LoadBehavior: $hkcuLB" -ForegroundColor $(if ($hkcuLB -eq 3) { "Green" } else { "Red" })
    Write-Host "[OK] Dev deployment complete!" -ForegroundColor Green

    Write-Host "[4/4] Note: If HKCR is incomplete, restart PowerShell as Admin and run:" -ForegroundColor Yellow
    Write-Host "  regasm `"$dllPath`" /codebase /tlb" -ForegroundColor Cyan
}

Write-Host
Write-Host "Restart OneNote to load the add-in." -ForegroundColor Cyan
