# OneInk - Build Script (x64)
# Builds to local bin directory: OneInk\bin\x64\Release\

. "$PSScriptRoot\config.ps1"

param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release"
)

Write-Host "Building OneInk (x64) - $Configuration..." -ForegroundColor Cyan
& $Global:MSBuildPath $Global:ProjectFile /p:Configuration=$Configuration /p:Platform=x64 /t:Rebuild
if ($LASTEXITCODE -ne 0) { exit 1 }
Write-Host "Build OK: OneInk\bin\x64\$Configuration\OneInk.dll" -ForegroundColor Green
