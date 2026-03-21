# OneInk - OneNote Ink Operations COM AddIn

OneInk is a minimal COM AddIn for Microsoft OneNote that provides ink manipulation tools.

## Features

- **Clear All Ink**: Remove all ink strokes from the current page
- **Delete Ink by Color**: Select a color and delete all strokes of that color (accurate color detection via Microsoft.Ink API)

## Requirements

- Windows 10/11
- Microsoft OneNote (Office 2016 or later, x64 or ARM64)
- .NET Framework 4.8
- Visual Studio 2022 (for building)
- PowerShell 5.1+ (for deployment scripts)

## Building

```powershell
# x64 (Release)
.\build-x64.ps1

# x64 (Debug)
.\build-x64.ps1 Debug

# ARM64
.\build-arm64.ps1
```

Output: `OneInk\bin\x64\Release\` (default) or `OneInk\bin\x64\Debug\`

## Deployment Scripts

| Script | Description |
|--------|-------------|
| `deploy-x64.ps1` | Build Release + copy to Program Files + register (x64) |
| `deploy-arm64.ps1` | Build Release + copy to Program Files + register (ARM64) |
| `install.ps1 x64` | Copy pre-built Release output to Program Files + register |
| `install.ps1 arm64` | Copy pre-built Release output to Program Files + register |
| `uninstall.ps1 x64` | Unregister COM and remove install directory |

> Note: Deploy scripts require administrator privileges (UAC prompt will appear).

### Workflow

1. **Develop**: Modify code → `.\build-x64.ps1` (compiles Release to `OneInk\bin\x64\Release\`)
2. **Deploy**: `.\deploy-x64.ps1` (full build + install + register, requires admin)

Or for already-built output:
```powershell
# After building with build-x64.ps1
.\install.ps1 x64
```

### Configuration

All paths are centralized in `config.ps1`. Edit this file if MSBuild or other tool paths change.

## Uninstallation

```powershell
.\uninstall.ps1 x64
```

Or manually:
```
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /u "C:\Program Files\OneInk-x64\OneInk.dll"
Remove-Item "C:\Program Files\OneInk-x64" -Recurse
```

## Usage

After installation, open OneNote. A **OneInk** tab appears in the ribbon with two buttons:
- **Clear All Ink**: Removes all ink strokes from the current page
- **Delete by Color**: Opens a dialog listing detected ink colors; select one to delete all strokes of that color

## Project Structure

```
OneInk/
├── OneInk.sln              # Visual Studio solution
├── OneInk/                 # Main AddIn project
│   ├── AddIn.cs            # COM AddIn entry point + ribbon callbacks
│   ├── CCOMStreamWrapper.cs # IStream wrapper for COM
│   ├── ColorSelectionDialog.cs # Ink color selection dialog
│   ├── InkColorExtractor.cs  # ISF color extraction via Microsoft.Ink
│   ├── Strings.cs           # i18n (Chinese/English)
│   ├── ribbon.xml          # Ribbon UI definition
│   └── Properties/
├── config.ps1               # Centralized configuration
├── build-x64.ps1           # Build script (x64)
├── build-arm64.ps1         # Build script (ARM64)
├── deploy-x64.ps1          # Deploy script (x64)
├── deploy-arm64.ps1        # Deploy script (ARM64)
├── install.ps1             # Install from pre-built output
└── Setup/                  # Installer project
    └── Setup.vdproj
```

## Development Notes

- The AddIn uses the OneNote Interop API (`Microsoft.Office.Interop.OneNote`)
- Ribbon UI is defined in `ribbon.xml` and loaded via `IRibbonExtensibility`
- Ink operations work with OneNote's XML page format

## License

MIT License - See LICENSE file for details.

## Acknowledgments

Based on the VanillaAddIn template from Microsoft.
