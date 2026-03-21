# OneInk - Configuration
# Centralized configuration for all build and deployment scripts

# Visual Studio / MSBuild
$Global:VsWhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$Global:MSBuildPath = if (Test-Path $VsWhere) {
    & $VsWhere -latest -products * -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe | Select-Object -First 1
} else {
    "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
}

# Project paths
$Global:ProjectRoot = $PSScriptRoot
$Global:ProjectFile = Join-Path $PSScriptRoot "OneInk\OneInk.csproj"

# .NET Framework regasm paths
$Global:RegAsmX86 = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe"
$Global:RegAsmX64 = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"

# Install destinations
$Global:InstallPathX64 = "C:\Program Files\OneInk-x64"
$Global:InstallPathARM64 = "C:\Program Files\OneInk-arm64"

# COM AddIn CLSID
$Global:AddInCLSID = "{E1F2A3B4-CF2D-409B-B65A-BDBACB9F21DC}"
$Global:AddInAppID = "{E1F2A3B4-CF2D-409B-B65A-BDBACB9F21DC}"
