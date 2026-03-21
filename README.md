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
.\build.ps1 -Platform x64 -Configuration Release
```

Output: `OneInk\bin\x64\Release\`

## Deployment

### Development Mode (recommended for active development)

```powershell
.\deploy.ps1 -Mode Dev -Platform x64
```

- Registers the add-in from the build output directory (`OneInk\bin\x64\Release\`)
- No admin required
- After code changes: rebuild, then re-run this script, then restart OneNote

### Production Mode (for end-user installation)

```powershell
.\deploy.ps1 -Mode Production -Platform x64
```

- Builds Release, copies to `C:\Program Files\OneInk-x64`
- Registers COM AddIn (requires administrator privileges)
- Sets HKLM registry entries + HKCU LoadBehavior

## Uninstallation

```powershell
.\uninstall.ps1 -Platform x64
```

> Note: Only for production installations. Dev mode registration is cleaned up automatically by the system when the build directory changes.

## Configuration

All paths are centralized in `config.ps1`. Edit this file if MSBuild or other tool paths change.

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
│   ├── BitmapExtensions.cs # Bitmap → IStream extension
│   ├── ReadOnlyStream.cs   # IStream COM wrapper
│   ├── ColorSelectionDialog.cs # Ink color selection dialog
│   ├── InkColorExtractor.cs  # ISF color extraction via Microsoft.Ink
│   ├── Strings.cs           # i18n (Chinese/English)
│   ├── Resources/           # Ribbon icons
│   └── Properties/
│       └── Resources.resx   # Ribbon XML + strings
├── config.ps1               # Centralized configuration
├── build.ps1               # Build script
├── deploy.ps1             # Deployment script (Dev + Production modes)
├── uninstall.ps1           # Production uninstall script
└── Setup/                  # Installer project
    └── Setup.vdproj
```

## Development Notes

- Ribbon icons are loaded via the `loadImage` callback, returning `IStream` (PNG data)
- The add-in uses the OneNote Interop API (`Microsoft.Office.Interop.OneNote`)
- Ribbon UI is defined in `ribbon.xml` and loaded via `IRibbonExtensibility`
- Ink operations work with OneNote's XML page format

## License

MIT License - See LICENSE file for details.

## Acknowledgments

Based on the VanillaAddIn template from Microsoft.
