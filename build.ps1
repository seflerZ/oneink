# OneInk - Build Script
# Builds the project to the bin directory

param(
    [ValidateSet("x86", "x64", "arm64")]
    [string]$Platform = "x64",

    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release"
)

. "$PSScriptRoot\config.ps1"

Write-Host "Building OneInk ($Platform) - $Configuration..." -ForegroundColor Cyan
& $Global:MSBuildPath $Global:ProjectFile /p:Configuration=$Configuration /p:Platform=$Platform /t:Rebuild /v:m
if ($LASTEXITCODE -ne 0) { exit 1 }
Write-Host "OK: OneInk\bin\$Platform\$Configuration\OneInk.dll" -ForegroundColor Green
