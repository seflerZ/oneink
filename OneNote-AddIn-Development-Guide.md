# OneNote COM AddIn 开发完整指南

> 基于 OneInk 项目开发经验总结 - 2026 年 3 月 18 日

---

## 📋 目录

1. [环境要求](#环境要求)
2. [项目配置](#项目配置)
3. [代码规范](#代码规范)
4. [编译配置](#编译配置)
5. [注册表配置](#注册表配置)
6. [COM 注册步骤](#com 注册步骤)
7. [常见问题排查](#常见问题排查)
8. [调试技巧](#调试技巧)

---

## 环境要求

### 必需软件
- **Visual Studio 2022** (带 .NET Framework 4.8 SDK)
- **OneNote 桌面版** (Office 365/2016/2019/2021)
- **.NET Framework 4.8** (不能用 .NET Core/.NET 5+)

### 验证 OneNote 版本
```
OneNote → File → Account → About OneNote
```
必须显示 **32-bit** 或 **64-bit** 桌面版，不能是 UWP 版或 Web 版。

### 验证 Office 架构
```powershell
# 查看 Office 架构
reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v Platform
```
输出 `x64` 或 `x86`，决定你需要编译的 DLL 架构。

---

## 项目配置

### 1. 创建项目
```
Visual Studio → Create Project → Class Library (.NET Framework)
Framework: .NET Framework 4.8
```

### 2. 项目文件配置 (YourProject.csproj)

```xml
<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup>
    <TargetFrameworkVersion>v4.8</TargetFrameworkVersion>
    <Configuration Condition=" '$(Configuration)' == '' ">Debug</Configuration>
    <Platform Condition=" '$(Platform)' == '' ">x64</Platform>
    <OutputType>Library</OutputType>
    <RootNamespace>YourNamespace</RootNamespace>
    <AssemblyName>YourAddIn</AssemblyName>
  </PropertyGroup>
  
  <!-- x64 配置 -->
  <PropertyGroup Condition="'$(Configuration)|$(Platform)' == 'Release|x64'">
    <OutputPath>bin\x64\Release\</OutputPath>
    <PlatformTarget>x64</PlatformTarget>
    <RegisterForComInterop>false</RegisterForComInterop>
  </PropertyGroup>
  
  <!-- x86 配置 -->
  <PropertyGroup Condition="'$(Configuration)|$(Platform)' == 'Release|x86'">
    <OutputPath>bin\x86\Release\</OutputPath>
    <PlatformTarget>x86</PlatformTarget>
    <RegisterForComInterop>false</RegisterForComInterop>
  </PropertyGroup>
  
  <ItemGroup>
    <!-- COM 引用 - EmbedInteropTypes 必须为 False -->
    <COMReference Include="Extensibility">
      <Guid>{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}</Guid>
      <VersionMajor>7</VersionMajor>
      <VersionMinor>0</VersionMinor>
      <Lcid>0</Lcid>
      <WrapperTool>tlbimp</WrapperTool>
      <EmbedInteropTypes>False</EmbedInteropTypes>
    </COMReference>
    
    <COMReference Include="Microsoft.Office.Core">
      <Guid>{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}</Guid>
      <VersionMajor>2</VersionMajor>
      <VersionMinor>8</VersionMinor>
      <Lcid>0</Lcid>
      <WrapperTool>tlbimp</WrapperTool>
      <EmbedInteropTypes>False</EmbedInteropTypes>
    </COMReference>
    
    <COMReference Include="Microsoft.Office.Interop.OneNote">
      <Guid>{0EA692EE-BB50-4E3C-AEF0-356D91732725}</Guid>
      <VersionMajor>1</VersionMajor>
      <VersionMinor>1</VersionMinor>
      <Lcid>0</Lcid>
      <WrapperTool>tlbimp</WrapperTool>
      <EmbedInteropTypes>False</EmbedInteropTypes>
    </COMReference>
  </ItemGroup>
</Project>
```

### 3. AssemblyInfo.cs 配置

```csharp
using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// 程序集信息
[assembly: AssemblyTitle("YourAddIn")]
[assembly: AssemblyDescription("Your AddIn Description")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("YourAddIn")]
[assembly: AssemblyCopyright("Copyright © 2026")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// ⚠️ 关键：ComVisible 必须为 true
[assembly: ComVisible(true)]

// 推荐：CLSCompliant 设为 true
[assembly: CLSCompliant(true)]

// GUID - 用于类型库 ID
[assembly: Guid("YOUR-GUID-HERE")]

// 版本信息
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]
```

---

## 代码规范

### 1. AddIn 主类

```csharp
using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using ClassInterfaceType = System.Runtime.InteropServices.ClassInterfaceType;

namespace YourNamespace
{
    // ⚠️ 关键特性配置
    [ComVisible(true)]
    [Guid("YOUR-GUID-HERE")]
    [ProgId("YourAddIn.AddIn")]
    [ClassInterface(ClassInterfaceType.None)] // 必须设为 None
    public class AddIn : IDTExtensibility2, IRibbonExtensibility
    {
        // ⚠️ 必须有无参构造函数！
        public AddIn()
        {
            // 初始化代码
        }

        #region IDTExtensibility2 实现

        public void OnConnection(object Application,
                                 ext_ConnectMode ConnectMode,
                                 object AddInInst,
                                 ref Array custom)
        {
            // OneNote 启动时调用
            // 保存 Application 对象引用
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode,
                                    ref Array custom)
        {
            // OneNote 关闭时调用
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        #endregion

        #region IRibbonExtensibility 实现

        public string GetCustomUI(string RibbonID)
        {
            // 返回 Ribbon XML
            return Properties.Resources.ribbon;
        }

        #endregion
    }
}
```

### 2. Ribbon XML (ribbon.xml)

```xml
<?xml version="1.0" encoding="utf-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" 
          loadImage="GetImage">
    <ribbon>
        <tabs>
            <!-- 添加到 Home 标签 -->
            <tab idMso="TabHome">
                <group id="groupYourAddIn" label="YourAddIn">
                    <button id="buttonYourAction" 
                            label="Do Something" 
                            size="large" 
                            screentip="Click to do something"
                            onAction="YourButtonClicked" 
                            image="Logo.png"/>
                </group>
            </tab>
            
            <!-- 或创建独立标签 -->
            <tab id="tabYourAddIn" label="YourAddIn">
                <group id="groupYourTools" label="Tools">
                    <button id="buttonTool1" 
                            label="Tool 1" 
                            size="large" 
                            onAction="Tool1Clicked"/>
                </group>
            </tab>
        </tabs>
    </ribbon>
</customUI>
```

---

## 编译配置

### 1. 编译项目

```powershell
# 使用 MSBuild 编译
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" `
  YourProject.sln `
  /p:Configuration=Release `
  /p:Platform=x64 `
  /t:Rebuild
```

### 2. 输出目录

推荐输出到：`C:\Program Files\YourAddIn\`

```powershell
# 创建目录
New-Item -ItemType Directory -Path "C:\Program Files\YourAddIn" -Force

# 复制文件
Copy-Item "bin\x64\Release\*.dll" -Destination "C:\Program Files\YourAddIn\" -Force

# 设置权限
$acl = Get-Acl "C:\Program Files\YourAddIn"
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    "Everyone", 
    "ReadAndExecute", 
    "ContainerInherit,ObjectInherit", 
    "None", 
    "Allow"
)
$acl.SetAccessRule($rule)
Set-Acl "C:\Program Files\YourAddIn" $acl
```

---

## 注册表配置

### ⚠️ 最关键部分

OneNote COM AddIn 需要以下注册表配置才能正常工作：

### 1. DllSurrogate 配置 (最关键！)

```registry
[HKEY_CLASSES_ROOT\AppID\{YOUR-GUID}]
"DllSurrogate"=""

[HKEY_CLASSES_ROOT\CLSID\{YOUR-GUID}]
"AppID"="{YOUR-GUID}"
```

**没有 DllSurrogate，OneNote 会静默拒绝加载 AddIn！**

### 2. OneNote AddIn 注册表

```registry
[HKEY_CURRENT_USER\Software\Microsoft\Office\OneNote\AddIns\YourAddIn.AddIn]
"LoadBehavior"=dword:00000003
"FriendlyName"="YourAddIn"
"Description"="Your AddIn Description"
"CommandLineSafe"=dword:00000001
```

**注意：路径中不能包含版本号（如 16.0）！**

### 3. COM 注册表

```registry
[HKEY_CLASSES_ROOT\CLSID\{YOUR-GUID}]
@="YourAddIn.AddIn"

[HKEY_CLASSES_ROOT\CLSID\{YOUR-GUID}\InprocServer32]
@="mscoree.dll"
"ThreadingModel"="Both"
"Class"="YourAddIn.AddIn"
"Assembly"="YourAddIn, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
"RuntimeVersion"="v4.0.30319"
"CodeBase"="file:///C:/Program Files/YourAddIn/YourAddIn.DLL"
```

### 4. 完整注册表脚本示例

```registry
Windows Registry Editor Version 5.00

; DllSurrogate 配置
[HKEY_CLASSES_ROOT\AppID\{E1F2A3B4-CF2D-409B-B65A-BDBACB9F21DC}]
"DllSurrogate"=""

[HKEY_CLASSES_ROOT\CLSID\{E1F2A3B4-CF2D-409B-B65A-BDBACB9F21DC}]
"AppID"="{E1F2A3B4-CF2D-409B-B65A-BDBACB9F21DC}"

; OneNote AddIn 配置
[HKEY_CURRENT_USER\Software\Microsoft\Office\OneNote\AddIns\OneInk.AddIn]
"LoadBehavior"=dword:00000003
"FriendlyName"="OneInk"
"Description"="OneInk - OneNote Ink Operations COM AddIn"
"CommandLineSafe"=dword:00000001
```

---

## COM 注册步骤

### 1. 使用正确的 regasm

```powershell
# 64 位 Office - 使用 Framework64
$regasm = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"

# 32 位 Office - 使用 Framework
$regasm = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe"
```

### 2. 完整注册流程

```powershell
$dllPath = "C:\Program Files\YourAddIn\YourAddIn.dll"

# 1. 反注册旧版本
& $regasm /unregister $dllPath

# 2. 清理 Wow6432Node (32 位残留)
Remove-Item -Path "HKCR:\Wow6432Node\CLSID\{YOUR-GUID}" -Recurse -Force -ErrorAction SilentlyContinue

# 3. 注册新版本（带 /codebase 和 /tlb 参数）
& $regasm /codebase /tlb $dllPath

# 4. 添加 DllSurrogate
New-Item -Path "HKCR:\AppID\{YOUR-GUID}" -Force | Out-Null
New-ItemProperty -Path "HKCR:\AppID\{YOUR-GUID}" -Name "DllSurrogate" -Value "" -PropertyType String -Force

# 5. 关联 CLSID 到 AppID
New-ItemProperty -Path "HKCR:\CLSID\{YOUR-GUID}" -Name "AppID" -Value "{YOUR-GUID}" -PropertyType String -Force

# 6. 设置 OneNote AddIn 注册表
$addinPath = "HKCU:\Software\Microsoft\Office\OneNote\AddIns\YourAddIn.AddIn"
New-Item -Path $addinPath -Force | Out-Null
New-ItemProperty -Path $addinPath -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force
New-ItemProperty -Path $addinPath -Name "FriendlyName" -Value "YourAddIn" -PropertyType String -Force
New-ItemProperty -Path $addinPath -Name "Description" -Value "Description" -PropertyType String -Force
New-ItemProperty -Path $addinPath -Name "CommandLineSafe" -Value 1 -PropertyType DWord -Force
```

### 3. 验证注册

```powershell
# 检查 CLSID 注册
reg query "HKCR\CLSID\{YOUR-GUID}" /s

# 检查 CodeBase
reg query "HKCR\CLSID\{YOUR-GUID}\InprocServer32" /v CodeBase

# 检查 OneNote AddIn 注册表
reg query "HKCU\Software\Microsoft\Office\OneNote\AddIns\YourAddIn.AddIn" /s

# 测试 COM 对象创建
$com = New-Object -ComObject "YourAddIn.AddIn"
```

---

## 常见问题排查

### 问题 1: AddIn 不显示在 OneNote 中

**可能原因：**
1. ❌ 缺少 DllSurrogate 配置
2. ❌ 注册表路径包含版本号（如 16.0）
3. ❌ 使用了 32 位 regasm 注册 64 位 DLL
4. ❌ DLL 在用户目录下

**解决方案：**
```powershell
# 1. 添加 DllSurrogate
New-Item -Path "HKCR:\AppID\{YOUR-GUID}" -Force
New-ItemProperty -Path "HKCR:\AppID\{YOUR-GUID}" -Name "DllSurrogate" -Value ""

# 2. 使用正确的注册表路径
# HKCU\Software\Microsoft\Office\OneNote\AddIns\YourAddIn.AddIn

# 3. 使用 64 位 regasm
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe

# 4. 移动 DLL 到 Program Files
C:\Program Files\YourAddIn\
```

### 问题 2: LoadBehavior 从 3 变成 2

**原因：** OneNote 尝试加载但失败了

**排查步骤：**
1. 检查 Fusion 日志：`C:\FusionLogs\`
2. 检查事件查看器：Windows Logs → Application → .NET Runtime
3. 验证 DLL 签名
4. 检查依赖项是否完整

### 问题 3: 提示"不是有效的 Office 加载项"

**可能原因：**
1. ❌ DLL 没有代码签名
2. ❌ BlockUnmanagedAddins 策略启用
3. ❌ .NET 框架版本不匹配

**解决方案：**
```powershell
# 1. 创建自签名证书
$cert = New-SelfSignedCertificate `
  -Type CodeSigning `
  -Subject "CN=YourAddIn Dev" `
  -CertStoreLocation "Cert:\CurrentUser\My"

# 2. 签名 DLL
& "C:\Program Files (x86)\Windows Kits\10\App Certification Kit\signtool.exe" `
  sign /a /fd SHA256 /tr http://timestamp.digicert.com `
  "C:\Program Files\YourAddIn\YourAddIn.dll"

# 3. 检查 BlockUnmanagedAddins
reg query "HKCU\Software\Policies\Microsoft\Office\16.0\OneNote\Resiliency" /v BlockUnmanagedAddins
# 如果值为 1，删除它
```

### 问题 4: COM 对象能创建但 OneNote 不加载

**检查清单：**
- [ ] DllSurrogate 是否配置
- [ ] AppID 是否关联到 CLSID
- [ ] 注册表路径是否正确（无版本号）
- [ ] CommandLineSafe 是否设为 1
- [ ] DLL 是否在 Program Files
- [ ] 文件夹权限是否正确

---

## 调试技巧

### 1. 启用 Fusion 日志

```powershell
# 启用 Fusion 日志
reg add "HKLM\SOFTWARE\Microsoft\Fusion" /v EnableLog /t REG_DWORD /d 1 /f
reg add "HKLM\SOFTWARE\Microsoft\Fusion" /v ForceLog /t REG_DWORD /d 1 /f
reg add "HKLM\SOFTWARE\Microsoft\Fusion" /v LogPath /t REG_SZ /d "C:\FusionLogs" /f

# 查看日志
Get-ChildItem "C:\FusionLogs" -Recurse -Filter "*.htm" | Sort-Object LastWriteTime -Descending | Select-Object -First 5
```

### 2. 添加日志到 AddIn

```csharp
private static void Log(string message)
{
    var logPath = Path.Combine(Path.GetTempPath(), "YourAddIn.log");
    var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
    File.AppendAllText(logPath, $"[{timestamp}] {message}\n");
}

public AddIn()
{
    Log("=== AddIn Constructor called ===");
    Log($"Assembly: {Assembly.GetExecutingAssembly().Location}");
}
```

### 3. 测试 COM 对象

```powershell
# 测试 COM 对象创建
try {
    $com = New-Object -ComObject "YourAddIn.AddIn"
    Write-Host "✓ COM object created" -ForegroundColor Green
    
    # 测试方法调用
    $ribbon = $com.GetCustomUI("Test")
    Write-Host "✓ GetCustomUI returned $($ribbon.Length) chars" -ForegroundColor Green
} catch {
    Write-Host "✗ COM creation failed: $_" -ForegroundColor Red
}
```

### 4. 检查注册表

```powershell
# 导出注册表验证
reg export "HKCR\CLSID\{YOUR-GUID}" "C:\temp\clsid.reg"
Get-Content "C:\temp\clsid.reg" | Select-String "CodeBase|InprocServer32"
```

---

## 快速参考清单

### 编译前
- [ ] 目标框架：.NET Framework 4.8
- [ ] 平台目标：x64 或 x86（匹配 Office）
- [ ] ComVisible：true（Assembly 和 Class 级别）
- [ ] CLSCompliant：true
- [ ] ClassInterface：ClassInterfaceType.None

### 编译后
- [ ] DLL 输出到 Program Files
- [ ] 设置文件夹权限（Everyone: Read & Execute）
- [ ] 代码签名（可选但推荐）
- [ ] 复制所有依赖 DLL

### 注册时
- [ ] 使用正确的 regasm（Framework64 或 Framework）
- [ ] 使用 /codebase 参数
- [ ] 使用 /tlb 参数（推荐）
- [ ] 添加 DllSurrogate
- [ ] 关联 AppID
- [ ] 设置 OneNote 注册表路径（无版本号）
- [ ] 设置 LoadBehavior = 3
- [ ] 设置 CommandLineSafe = 1

### 测试前
- [ ] 完全关闭 OneNote（任务管理器）
- [ ] 清理 Wow6432Node 残留
- [ ] 验证 COM 对象创建
- [ ] 验证注册表配置

---

## 附录：完整注册表脚本

```registry
Windows Registry Editor Version 5.00

; DllSurrogate 配置
[HKEY_CLASSES_ROOT\AppID\{E1F2A3B4-CF2D-409B-B65A-BDBACB9F21DC}]
"DllSurrogate"=""

[HKEY_CLASSES_ROOT\CLSID\{E1F2A3B4-CF2D-409B-B65A-BDBACB9F21DC}]
"AppID"="{E1F2A3B4-CF2D-409B-B65A-BDBACB9F21DC}"

[HKEY_CLASSES_ROOT\CLSID\{E1F2A3B4-CF2D-409B-B65A-BDBACB9F21DC}\InprocServer32]
@="mscoree.dll"
"ThreadingModel"="Both"
"Class"="OneInk.AddIn"
"Assembly"="OneInk, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
"RuntimeVersion"="v4.0.30319"
"CodeBase"="file:///C:/Program Files/OneInk/OneInk.DLL"

; OneNote AddIn 配置
[HKEY_CURRENT_USER\Software\Microsoft\Office\OneNote\AddIns\OneInk.AddIn]
"LoadBehavior"=dword:00000003
"FriendlyName"="OneInk"
"Description"="OneInk - OneNote Ink Operations COM AddIn"
"CommandLineSafe"=dword:00000001
```

---

## 参考资料

- [Microsoft Office COM Add-ins](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Creating a COM Add-in](https://docs.microsoft.com/en-us/visualstudio/extensibility/creating-a-com-add-in)
- [Regasm.exe (Assembly Registration Tool)](https://docs.microsoft.com/en-us/dotnet/framework/tools/regasm-exe-assembly-registration-tool)
- [OneNote Add-in Development](https://docs.microsoft.com/en-us/office/dev/add-ins/onenote/onenote-add-ins)

---

**文档版本**: 1.0  
**更新日期**: 2026-03-18  
**基于项目**: OneInk, NoteHighlight2016, OnenoteAddin, VanillaAddIn
