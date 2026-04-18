# OneInk - OneNote 墨迹工具箱

一款为 Microsoft OneNote 打造的 COM 插件，提供墨迹（手写/画图）操作工具。

![OneInk Logo](docs/Logo.png)

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

1. 下载 `OneInk-v1.0.0-win64.zip`（见 [Releases](https://github.com/seflerZ/oneink/releases)）
2. 解压到任意位置
3. 右键 `install.ps1` → **使用 PowerShell 运行**（或右键 → "用管理员身份运行 PowerShell"）
4. 输入 `R` 确认，等待安装完成
5. 重启 OneNote，插件自动加载

> 安装后 OneNote 功能区会出现 **OneInk** 标签页。

**卸载：** 解压包里的 `uninstall.ps1` 同样以管理员身份运行即可。

---

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

## 📁 项目结构

```
OneInk/
├── OneInk/                    # 源码
│   ├── AddIn.cs              # COM 入口 + Ribbon 回调
│   ├── InkColorExtractor.cs   # 墨迹颜色提取
│   ├── InkDashedConverter.cs # 虚线/平滑/对齐转换
│   ├── ColorSelectionDialog.cs# 颜色选择对话框
│   └── Properties/
│       └── Resources.resx    # Ribbon XML + 字符串资源
├── docs/                     # 开发文档
├── install.ps1               # 安装脚本
├── uninstall.ps1             # 卸载脚本
└── LICENSE                  # MIT 协议
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
