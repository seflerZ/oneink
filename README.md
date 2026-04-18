# OneInk - OneNote Ink Operations

A COM AddIn for Microsoft OneNote that provides ink (handwriting/drawing) manipulation tools.

![OneInk Logo](docs/Logo.png)

---

[中文说明](#中文说明) | [English](#english)

---

## ✨ Features

| Feature | Description |
|---------|-------------|
| **Clear All Ink** | Remove all ink from the current page — selected ink only if lasso selection is active |
| **Delete by Color** | Select colors (multi-select) and delete strokes of those colors — shows/deletes only selected ink when lasso selection is active |
| **To Dashed Lines** | Convert ink strokes to dashed/dotted lines — three density presets (dense/medium/sparse) |
| **Smooth Ink** | Smooth hand-drawn strokes — supports curve smoothing (Chaikin algorithm) and polyline simplification (Ramer-Douglas-Peucker) |
| **Align Ink** | Align multiple ink strokes by their edges — top/bottom/left/right alignment |
| **Partial Selection Handling** | When lasso-selecting merged ink containers, auto-detects and offers to split them |

## 📥 Installation

**Requirements:**
- Windows 10/11
- Microsoft OneNote (Office 2016 or later, **64-bit desktop version only**)
- .NET Framework 4.8 (usually pre-installed)

**Steps:**

1. Download `OneInk-vX.X.X-win64.zip` from [Releases](https://github.com/seflerZ/oneink/releases)
2. Unzip to any location
3. Right-click `install.ps1` → **Run with PowerShell** (as Administrator)
4. Press `R` to confirm and wait for installation
5. Restart OneNote — the **OneInk** tab will appear in the ribbon

**Uninstall:** Run `uninstall.ps1` (from the same zip) as Administrator.

---

## ⚠️ Known Limitations

- ❌ Does **not** support OneNote UWP (Microsoft Store version)
- ❌ Does **not** support 32-bit Office
- ❌ Does **not** support Mac/iOS/Android versions

**How to check:** OneNote → File → Account → About OneNote — must show "64-bit" desktop edition.

## 🛠️ Build from Source

**Requirements:**
- Visual Studio 2022
- .NET Framework 4.8 SDK
- PowerShell 5.1+

```powershell
# Build
.\build.ps1 -Platform x64 -Configuration Release

# Dev deployment (no admin required)
.\deploy.ps1 -Mode Dev

# Production deployment
.\deploy.ps1 -Mode Production
```

## 📁 Project Structure

```
OneInk/
├── OneInk/                    # Source code
│   ├── AddIn.cs              # COM entry point + ribbon callbacks
│   ├── InkColorExtractor.cs   # Ink color extraction
│   ├── InkDashedConverter.cs  # Dashed/smooth/align conversion
│   ├── ColorSelectionDialog.cs# Color selection dialog
│   └── Properties/
│       └── Resources.resx     # Ribbon XML + string resources
├── libs/                      # Pre-built Interop DLLs
├── docs/                      # Development notes
├── install.ps1                # Installation script
├── uninstall.ps1              # Uninstallation script
└── LICENSE                    # MIT License
```

## 📝 Technical Notes

- Built on `IDTExtensibility2` + `IRibbonExtensibility` COM interfaces
- Ink operations via `Microsoft.Ink` library and ISF (Ink Serialized Format)
- Ribbon UI defined in `Properties/Resources.resx` (ribbon.xml)
- Selection detection: `piBinaryDataSelection` + `piBinaryData` with `objectID` matching

See [`docs/`](docs/) for development documentation.

## 📄 License

MIT License — see [LICENSE](LICENSE).

## 🙏 Acknowledgments

Based on Microsoft's VanillaAddIn template.

---

## 中文说明

# OneInk - OneNote 墨迹工具箱

一款为 Microsoft OneNote 打造的 COM 插件，提供墨迹（手写/画图）操作工具。

## ✨ 功能

| 功能 | 说明 |
|------|------|
| **清除全部墨迹** | 删除当前页面所有墨迹 — 选中时仅删除选中部分 |
| **按颜色删除** | 选择颜色（支持多选），删除该颜色的墨迹 — 选中时仅显示/删除选中部分 |
| **转为虚线** | 将墨迹转为虚线/点线 — 支持三种密度（密集/中等/稀疏） |
| **平滑墨迹** | 将手绘线条平滑化 — 支持曲线平滑（Chaikin算法）和折线简化（Ramer-Douglas-Peucker） |
| **对齐墨迹** | 将多个墨迹按边缘对齐 — 支持顶边/底边/左边/右边对齐 |
| **局部选择处理** | 当套索选中的是合并的墨迹容器时，自动检测并询问是否拆分 |

## 📥 安装

**要求：**
- Windows 10/11
- Microsoft OneNote（Office 2016 或更新版本，**必须是 64 位桌面版**）
- .NET Framework 4.8（通常系统已内置）

**安装步骤：**

1. 下载 `OneInk-vX.X.X-win64.zip`（见 [Releases](https://github.com/seflerZ/oneink/releases)）
2. 解压到任意位置
3. 右键 `install.ps1` → **使用 PowerShell 运行**（或右键 → "用管理员身份运行 PowerShell"）
4. 输入 `R` 确认，等待安装完成
5. 重启 OneNote，插件自动加载

> 安装后 OneNote 功能区会出现 **OneInk** 标签页。

**卸载：** 解压包里的 `uninstall.ps1` 同样以管理员身份运行即可。

## 🔧 系统要求详解

- ⚠️ **不支持 OneNote UWP 版**（Windows 10/11 自带的那个应用商店版）
- ⚠️ **不支持 32 位 Office**
- ⚠️ **不支持 Mac/iOS/Android 版 OneNote**

确认方法：OneNote → 文件 → 账户 → 关于 OneNote，查看是否为 "64 位" 桌面版。

## 🛠️ 从源码编译

**环境要求：**
- Visual Studio 2022
- .NET Framework 4.8 SDK
- PowerShell 5.1+

```powershell
# 编译
.\build.ps1 -Platform x64 -Configuration Release

# 开发模式部署（无需管理员）
.\deploy.ps1 -Mode Dev

# 生产模式部署
.\deploy.ps1 -Mode Production
```

## 📝 技术笔记

- 基于 `IDTExtensibility2` + `IRibbonExtensibility` COM 接口
- 墨迹操作通过 `Microsoft.Ink` 库处理 ISF（墨迹序列化格式）
- Ribbon UI 在 `Properties/Resources.resx` 的 `ribbon.xml` 中定义
- 选择检测：`piBinaryDataSelection` + `piBinaryData` 配合 `objectID` 匹配

详细开发文档见 [`docs/`](docs/) 目录。

## 📄 协议

MIT License - 详见 [LICENSE](LICENSE) 文件。

## 🙏 致谢

基于 Microsoft 的 VanillaAddIn 模板开发。
