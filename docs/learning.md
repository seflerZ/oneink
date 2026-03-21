# OneInk 开发经验总结

## 1. PowerShell 脚本中 param 块必须放在最前面

PowerShell 脚本中，`param()` 块**必须**放在文件的最前面（注释除外），在任何其他语句（包括 dot-source）之前。

```powershell
# ❌ 错误：param 在 dot-source 之后
. "$PSScriptRoot\config.ps1"
param(
    [string]$Platform = "x64"
)

# ✅ 正确：param 在最前面
param(
    [string]$Platform = "x64"
)
. "$PSScriptRoot\config.ps1"
```

`param` 出现在其他语句之后会报错：`The assignment expression is not valid`。

## 2. HKCU vs HKLM 注册——Dev 模式无需 admin

COM AddIn 注册既可以写 HKLM（需要管理员权限），也可以写 HKCU（无需管理员）。

### HKLM（需要 admin）
```powershell
# HKLM 注册需要管理员权限
New-Item -Path "HKLM:\SOFTWARE\Classes\CLSID\..."
Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\CLSID\..." -Name CodeBase -Value "..."
```

### HKCU（无需 admin）——推荐用于 Dev 模式
```powershell
# HKCU 注册不需要管理员权限，可以日常开发中使用
New-Item -Path "HKCU:\SOFTWARE\Classes\CLSID\..." -Force
Set-ItemProperty -Path "HKCU:\SOFTWARE\Classes\CLSID\..." -Name CodeBase -Value "..."
```

完整的手动 HKCU 注册结构（无需 regasm）：
```powershell
$clsid = "{E1F2A3B4-CF2D-409B-B65A-BDBACB9F21DC}"
$appId = "{E1F2A3B4-CF2D-409B-B65A-BDBACB9F21DC}"
$buildDir = "C:\path\to\bin\x64\Release"

# AppID with DllSurrogate
New-Item -Path "HKCU:\SOFTWARE\Classes\AppID\$appId" -Force
Set-ItemProperty -Path "HKCU:\SOFTWARE\Classes\AppID\$appId" -Name DllSurrogate -Value ""

# CLSID
$clsidPath = "HKCU:\SOFTWARE\Classes\CLSID\$clsid"
New-Item -Path $clsidPath -Force
Set-ItemProperty -Path $clsidPath -Name AppID -Value $appId

# InprocServer32
$inprocPath = "$clsidPath\InprocServer32"
New-Item -Path $inprocPath -Force
Set-ItemProperty -Path $inprocPath -Name "(Default)" -Value "mscoree.dll"
Set-ItemProperty -Path $inprocPath -Name ThreadingModel -Value "Both"
Set-ItemProperty -Path $inprocPath -Name CodeBase -Value "file:///$($buildDir.Replace('\', '/'))/OneInk.dll"

# HKCU AddIn entry
$addinPath = "HKCU:\SOFTWARE\Microsoft\Office\OneNote\AddIns\OneInk.AddIn"
New-Item -Path $addinPath -Force
Set-ItemProperty -Path $addinPath -Name LoadBehavior -Value 3 -Type DWord
```

**注意**：HKLM 注册不会被 HKCU 覆盖。如果之前用 admin 注册过 HKLM，再次用 HKCU 注册不会生效，因为 COM 会优先查找 HKLM。

## 3. OneNote GetPageContent API 选择检测

### 问题
需要检测 OneNote 页面中被 lasso 选中的墨迹，以便只在选中区域内操作。

### 关键发现

**`piBinaryData` 不包含 `selected` 属性**

`GetPageContent(pageId, out xml, PageInfo.piBinaryData)` 返回的 XML 中，InkDrawing 元素**没有** `selected` 属性，即使这些墨迹被选中。

**`piBinaryDataSelection` 包含 `selected` 属性，但没有 ISF 数据**

`GetPageContent(pageId, out xml, PageInfo.piBinaryDataSelection)` 返回的 XML 中，选中的 InkDrawing 元素有 `selected="all"` 属性。但这个 API **不包含** ISF（Ink Serialized Format）二进制数据（`<Data>` 元素为空或不存在）。

### 解决方案：两步法

1. 用 `piBinaryDataSelection` 获取选中墨迹的 objectID（通过 `selected="all"` 属性筛选）
2. 用 `piBinaryData` 获取完整页面数据（包含 ISF 数据），通过 objectID 匹配

```csharp
// Step 1: 获取选择元数据
OneNoteApplication.GetPageContent(pageId, out string xmlSelection,
    Microsoft.Office.Interop.OneNote.PageInfo.piBinaryDataSelection);

XDocument docSel = XDocument.Parse(xmlSelection);
var selectedObjectIds = new HashSet<string>(
    docSel.Descendants(ns + "InkDrawing")
          .Where(e => e.Attribute("selected")?.Value == "all")
          .Select(e => e.Attribute("objectID")?.Value ?? "")
          .Where(id => !string.IsNullOrEmpty(id))
);
bool hasSelection = selectedObjectIds.Count > 0;

// Step 2: 获取完整 ISF 数据
OneNoteApplication.GetPageContent(pageId, out string xml,
    Microsoft.Office.Interop.OneNote.PageInfo.piBinaryData);

XDocument doc = XDocument.Parse(xml);
var inkElements = doc.Descendants(ns + "InkDrawing").ToList();

foreach (var ink in inkElements)
{
    string objectId = ink.Attribute("objectID")?.Value ?? "";
    // 如果有选择，只处理选中的墨迹
    if (hasSelection && !selectedObjectIds.Contains(objectId))
        continue;
    // 处理墨迹...
}
```

## 4. Ribbon loadImage 回调返回类型

OneNote Ribbon `loadImage` 回调的返回类型必须是 `IStream`，**不是** `IPictureDisp`。

```csharp
// ✅ 正确：返回 IStream
public IStream GetImage(string imageName)
{
    using (var bitmap = new Bitmap(imagePath))
        return bitmap.GetReadOnlyStream();
}

// ❌ 错误：IPictureDisp 在 OneNote 中无法正常工作
public IPictureDisp GetImage(string imageName) { ... }
```

`IStream` 的实现可以参考 [OneMore 项目的 ReadOnlyStream.cs](https://github.com/oneNote-dev/OneMore)。

## 5. Program Files 权限问题

部署到 `C:\Program Files\` 目录时，文件由 TrustedInstaller 所有，普通用户无法写入或删除。

解决方案：Dev 模式注册到项目生成目录（无需 admin），Production 模式才部署到 Program Files（需要 admin 一次）。

## 6. Build Configuration 保持一致

确保 build 使用的配置与注册指向的目录一致：

- `build.ps1 -Configuration Debug` → 输出到 `bin\x64\Debug`
- `build.ps1 -Configuration Release` → 输出到 `bin\x64\Release`
- `deploy.ps1 -Mode Dev` → 从 Release 目录注册（`bin\x64\Release`）

Dev 模式开发建议始终使用 **Release 配置**（与生产环境一致）。
