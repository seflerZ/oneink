# OneInk - Build Script (ARM64)
# Builds to local bin directory: OneInk\bin\ARM64\Debug\

. "$PSScriptRoot\config.ps1"

Write-Host "Building OneInk (ARM64)..." -ForegroundColor Cyan
& $Global:MSBuildPath $Global:ProjectFile /p:Configuration=Debug /p:Platform=ARM64 /t:Rebuild
