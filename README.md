# OneInk - OneNote Ink Operations COM AddIn

OneInk is a minimal COM AddIn for Microsoft OneNote that provides ink manipulation tools.

## Features

- **Clear All Ink**: Remove ink strokes from the current page — if ink is selected (lasso), only selected ink is removed
- **Delete Ink by Color**: Select colors (multi-select supported) and delete strokes of those colors — if ink is selected (lasso), only selected ink colors are shown and deleted
- **To Dashed Lines**: Convert ink strokes to dashed/dotted lines — supports three density presets (dense, medium, sparse); if ink is selected (lasso), only selected ink is converted
- **Smooth Ink**: Smooth hand-drawn strokes into cleaner lines — supports two modes:
  - **Smooth Curve (曲线平滑)**: Chaikin's algorithm for flowing Bezier curves
  - **Smooth Polyline (折线平滑)**: Ramer-Douglas-Peucker algorithm for simplified straight-line segments
  - If ink is selected (lasso), only selected ink is smoothed
- **Align Ink**: Align multiple selected ink strokes by their edges — supports two modes:
  - **Align Top (顶边对齐)**: Aligns all ink strokes to the highest (topmost) stroke's top edge
  - **Align Bottom (底边对齐)**: Aligns all ink strokes to the lowest (bottommost) stroke's bottom edge
  - If ink is selected (lasso), only selected ink is aligned

## Installing the release pacakge
Just unzip the release file and execute the install.ps1 script as Administrator.

## Requirements for Development

- Windows 10/11
- Microsoft OneNote (Office 2016 or later, **x64**)
- .NET Framework 4.8
- Visual Studio 2022 (for building)
- PowerShell 5.1+ (for deployment scripts)

> **Note for ARM64 Windows users**: Even on ARM64 Windows, Office is typically installed as **x64** (for compatibility). If your Office is x64, follow the x64 installation instructions below — do NOT use the arm64 platform option.

## Building

```powershell
.\build.ps1 -Platform x64 -Configuration Release
```

Output: `OneInk\bin\x64\Release\`

## Deployment

### Development Mode (recommended for active development)

```powershell
.\deploy.ps1 -Mode Dev
```

- Registers the add-in from the build output directory (`OneInk\bin\x64\Release\`)
- No admin required
- After code changes: rebuild, then re-run this script, then restart OneNote

### Production Mode (for end-user installation)

```powershell
.\deploy.ps1 -Mode Production
```

- Builds Release, copies to `C:\Program Files\OneInk`
- Registers COM AddIn (requires administrator privileges)
- Sets HKLM registry entries + HKCU LoadBehavior

## Uninstallation

```powershell
.\uninstall.ps1
```

> Note: Only for production installations. Dev mode uses HKCU registration (no admin), cleaned up automatically by rebuilding with different paths.

## Configuration

All paths are centralized in `config.ps1`. Edit this file if MSBuild or other tool paths change.

## Usage

After installation, open OneNote. A **OneInk** tab appears in the ribbon with tools:

- **Clear All Ink**: Removes ink strokes from the current page — if ink is selected (lasso selection), only selected ink is removed
- **Delete by Color**: Opens a dialog listing detected ink colors on the page — if ink is selected (lasso selection), only selected ink colors are shown; check multiple colors to delete them all at once
- **To Dashed Lines** (split button):
  - Main button label shows current density (e.g., `转为虚线（中等）`)
  - Click the dropdown arrow to select density: **密集** (dense), **中等** (medium), **稀疏** (sparse)
  - Each density has its own icon; clicking a menu item updates the button label and icon
  - Click the main button to convert ink with the selected density
  - If ink is selected (lasso selection), only selected ink is converted
- **Smooth Ink** (split button):
  - Main button label shows current mode (e.g., `平滑至（曲线）`)
  - Click the dropdown arrow to select mode: **曲线** (curve smoothing) or **折线** (polyline simplification)
  - Each mode has its own icon; clicking a menu item updates the button label and icon
  - Click the main button to smooth ink with the selected mode
  - Curve smoothing uses Chaikin's algorithm to create flowing Bezier curves
  - Polyline smoothing uses Ramer-Douglas-Peucker algorithm to simplify strokes to straight segments
  - If ink is selected (lasso selection), only selected ink is smoothed
- **Align Ink** (split button):
  - Main button label shows current mode (e.g., `对齐（顶边对齐）`)
  - Click the dropdown arrow to select mode: **顶边对齐** (align top) or **底边对齐** (align bottom)
  - Each mode has its own icon; clicking a menu item updates the button label and icon
  - Click the main button to align ink with the selected mode
  - Uses intelligent clustering to group strokes that form logical shapes (e.g., hand-drawn cube)
  - Align Top: each cluster aligns to the highest stroke's top edge within that cluster
  - Align Bottom: each cluster aligns to the lowest stroke's bottom edge within that cluster
  - Strokes far apart horizontally are treated as separate clusters for better alignment

## Project Structure

```
OneInk/
├── OneInk.sln              # Visual Studio solution
├── OneInk/                 # Main AddIn project
│   ├── AddIn.cs            # COM AddIn entry point + ribbon callbacks
│   ├── BitmapExtensions.cs # Bitmap → IStream extension
│   ├── ReadOnlyStream.cs   # IStream COM wrapper
│   ├── ColorSelectionDialog.cs # Ink color selection dialog (multi-select)
│   ├── InkColorExtractor.cs  # ISF color extraction via Microsoft.Ink
│   ├── InkDashedConverter.cs # Ink conversion: dashed lines + smoothing + alignment via Microsoft.Ink
│   ├── Strings.cs           # i18n (Chinese/English)
│   ├── Resources/           # Ribbon icons
│   └── Properties/
│       └── Resources.resx   # Ribbon XML + embedded strings
├── config.ps1               # Centralized configuration
├── build.ps1               # Build script
├── deploy.ps1              # Deployment script (Dev + Production modes)
├── uninstall.ps1           # Production uninstall script
└── docs/                   # Development notes and learnings
```

## Development Notes

- Ribbon icons are loaded via the `loadImage` callback, returning `IStream` (PNG data)
- Dynamic ribbon images (e.g., density-dependent icons) use `getImage` callback with `IPictureDisp` return type
- The add-in uses the OneNote Interop API (`Microsoft.Office.Interop.OneNote`)
- Ribbon UI is defined in `ribbon.xml` and loaded via `IRibbonExtensibility`
- Selection detection: `piBinaryDataSelection` returns `selected="all"` on selected `InkDrawing` elements; `piBinaryData` provides ISF stroke data — use both via objectID matching
- Ink operations work with OneNote's XML page format and ISF (Ink Serialized Format) via `Microsoft.Ink`
- `splitButton` is used for combined button + menu UI; menu items use `onAction` callbacks
- Ribbon `dropDown` `onAction` is known to not fire reliably in OneNote — use separate `button` elements or `splitButton` with `menu` instead
- **Smooth Curve**: Chaikin's corner-cutting algorithm (3 iterations) produces C^1 continuous smooth curves
- **Smooth Polyline**: Ramer-Douglas-Peucker algorithm (epsilon=500 HIMETRIC ≈ 12.7mm) simplifies strokes to straight segments
- **Align Ink**: Hierarchical clustering (single-linkage) groups strokes by position; distance threshold is 30 HIMETRIC; uses Euclidean distance with X normalized to page-level scale (`sqrt((dx/100)^2 + dy^2)`)

## License

MIT License - See LICENSE file for details.

## Acknowledgments

Based on the VanillaAddIn template from Microsoft.
