# OneNote COM API 性能优化经验总结

> 基于 OneInk 项目开发经验 - 2026 年 3 月

---

## 核心问题

OneNote COM API 的 `GetPageContent` 方法有多种模式，性能差异巨大：

| 模式 | 值 | 性能 | 用途 |
|------|-----|------|------|
| `piBasic` | 0 | ~20ms | 仅结构，无二进制数据 |
| `piBinaryData` | 1 | ~20秒 | 完整页面+墨迹数据 |
| `piSelection` | 2 | ~20ms | 结构+选中标记 (`selected="all"`) |
| `piBinaryDataSelection` | 3 | ~23秒 | 完整+选中标记+二进制数据 |

---

## 正确获取墨迹数据的流程

### API 关系图

```
GetPageContent(piSelection)     → 获取 objectId + selected="all" 标记
GetPageContent(piBasic)        → 获取 objectId + Position + Size（无 Data）
GetBinaryPageContent(objectId)  → 获取单个墨迹的 ISF 二进制数据
UpdatePageContent(partialXml)   → 部分更新（只需包含修改的对象）
```

### InkDrawing XML 结构

```xml
<InkDrawing objectID="{id}" lastModifiedTime="{time}">
  <Position/>
  <Size/>
  <Data>base64_encoded_isf_data</Data>  <!-- 只有 piBinaryData 才有 -->
</InkDrawing>
```

### 关键发现

#### 1. piSelection 返回全量结构但带选中标记

`piSelection` 返回页面的完整 XML 结构，会在选中的 `InkDrawing` 元素上标记 `selected="all"` 属性。

```xml
<InkDrawing objectID="{...}" selected="all">
  <Position/>
  <Size/>
  <!-- 没有 <Data> -->
</InkDrawing>
```

#### 2. GetBinaryPageContent 逐个获取墨迹数据

根据官方文档，`GetBinaryPageContent(pageId, callbackId)` 可以获取单个墨迹的二进制数据。其中 `callbackId` 就是 `InkDrawing` 的 `objectID`。

```csharp
// 获取单个墨迹的二进制数据
OneNoteApplication.GetBinaryPageContent(pageId, objectId, out string isfBase64);
```

#### 3. UpdatePageContent 支持部分更新

不需要传入完整页面 XML，可以只传入修改的对象。

```xml
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="{pageId}">
  <one:InkDrawing objectID="{objectId}" lastModifiedTime="{timestamp}">
    <one:Position/>
    <one:Size/>
    <one:Data>converted_base64_data</one:Data>
  </one:InkDrawing>
</one:Page>
```

---

## 通用优化模式

### 获取选中墨迹 ID（快）

```csharp
OneNoteApplication.GetPageContent(pageId, out string xml,
    Microsoft.Office.Interop.OneNote.PageInfo.piSelection);
var selectedIds = xml.Descendants(ns + "InkDrawing")
    .Where(e => e.Attribute("selected")?.Value == "all")
    .Select(e => e.Attribute("objectID")?.Value)
    .Where(id => !string.IsNullOrEmpty(id))
    .ToHashSet();
```

### 获取所有墨迹 ID（快）

```csharp
OneNoteApplication.GetPageContent(pageId, out string xml,
    Microsoft.Office.Interop.OneNote.PageInfo.piBasic);
var allInkIds = xml.Descendants(ns + "InkDrawing")
    .Select(e => e.Attribute("objectID")?.Value)
    .Where(id => !string.IsNullOrEmpty(id))
    .ToList();
```

### 逐个获取墨迹 ISF 数据

```csharp
foreach (var objectId in selectedIds)
{
    OneNoteApplication.GetBinaryPageContent(pageId, objectId, out string isfBase64);
    // 处理 isfBase64...
}
```

### 删除墨迹（极快）

```csharp
OneNoteApplication.DeletePageContent(pageId, objectId); // ~1ms/条
```

---

## 优化案例

### ClearInk 清除墨迹

**优化后（~40ms）**

```csharp
// Step 1: piSelection 获取选中状态（~20ms）
OneNoteApplication.GetPageContent(pageId, out string xmlSel, PageInfo.piSelection);
var selectedIds = ParseSelectedIds(xmlSel);

// Step 2: piBasic 获取所有墨迹 ID（~20ms）
OneNoteApplication.GetPageContent(pageId, out string xmlBasic, PageInfo.piBasic);
var allIds = ParseAllInkIds(xmlBasic);

// Step 3: DeletePageContent 删除（~1ms/条）
foreach (var id in allIds)
{
    if (HasSelection && !selectedIds.Contains(id)) continue;
    OneNoteApplication.DeletePageContent(pageId, id);
}
```

**性能**：~43秒 → ~40ms

---

### ToDashed 转换虚线

**优化后（~200ms）**

```csharp
// Step 1: piSelection 获取选中 objectIds（~20ms）
OneNoteApplication.GetPageContent(pageId, out string xmlSel, PageInfo.piSelection);
var selectedIds = ParseSelectedIds(xmlSel);

// Step 2: piBasic 获取页面结构（~20ms）
OneNoteApplication.GetPageContent(pageId, out string xmlBasic, PageInfo.piBasic);
var selectedInks = ParseInkElements(xmlBasic, selectedIds);

// Step 3: GetBinaryPageContent 逐个获取并转换
var modifiedInks = new List<XElement>();
foreach (var ink in selectedInks)
{
    string objectId = ink.Attribute("objectID").Value;
    OneNoteApplication.GetBinaryPageContent(pageId, objectId, out string isfBase64);
    string dashedBase64 = ConvertToDashed(isfBase64);

    modifiedInks.Add(new XElement(ns + "InkDrawing",
        new XAttribute("objectID", objectId),
        new XAttribute("lastModifiedTime", ink.Attribute("lastModifiedTime")?.Value),
        ink.Element(ns + "Position"),
        ink.Element(ns + "Size"),
        new XElement(ns + "Data", dashedBase64)
    ));
}

// Step 4: UpdatePageContent 部分更新
var pageXml = BuildPartialPageXml(pageId, modifiedInks);
OneNoteApplication.UpdatePageContent(pageXml);
```

**性能**：~40秒 → ~200ms

---

### SelectInkColor 按颜色删除墨迹

**有选中时（~100-200ms）**

```csharp
// Step 1: piSelection 获取选中状态（~20ms）
// Step 2: piBasic 获取所有墨迹（~20ms）
// Step 3: 只检查选中的墨迹，逐个 GetBinaryPageContent
// Step 4: DeletePageContent 删除匹配的颜色
```

**无选中时**：仍需要检查所有墨迹（API 限制）

---

## 教训总结

1. **不要只看 API 名字猜测行为** - `piSelection` 并不是"只返回选中内容"，而是"返回全量+标记"

2. **GetBinaryPageContent 是关键** - callbackId 就是 objectId，可以逐个获取墨迹数据

3. **UpdatePageContent 支持部分更新** - 不需要传入完整页面

4. **XML 命名空间必须正确** - 使用 XDocument/XElement 构建

5. **测试是验证假设的唯一方式** - 必须看实际输出

6. **根据场景选择策略**：
   - 删除：只需要 ID，piBasic + piSelection
   - 转换/颜色：需要 ISF 数据，用 GetBinaryPageContent 逐个获取

---

## 相关代码参考

- `OneInk/AddIn.cs` - `ClearInkButtonClicked`、`ExecuteToDashed`、`SelectInkColorButtonClicked`
- `docs/How-to-extract-color-from-ink.md` - 墨迹颜色提取经验
