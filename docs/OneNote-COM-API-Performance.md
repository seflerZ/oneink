# OneNote COM API 性能优化经验总结

> 基于 OneInk 项目开发经验 - 2026 年 3 月 22 日

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

## 关键发现

### 1. piSelection 返回全量结构但带选中标记

`piSelection` 会返回页面的完整 XML 结构（不是只有选中的内容），但会在选中的 `InkDrawing` 元素上标记 `selected="all"` 属性。

```xml
<InkDrawing objectID="{...}" selected="all">
  <!-- 内容 -->
</InkDrawing>
```

### 2. piSelection 不返回二进制数据

即使 `selected="all"`，`piSelection` 返回的 `InkDrawing` 也不包含 `<Data>` 子元素。

```xml
<!-- piSelection 返回的结构 -->
<InkDrawing objectID="{...}" selected="all">
  <Position/>
  <Size/>
  <!-- 没有 <Data> -->
</InkDrawing>
```

### 3. 只有 piBinaryData 返回墨迹二进制数据

```xml
<!-- piBinaryData 返回的结构 -->
<InkDrawing objectID="{...}">
  <Position/>
  <Size/>
  <Data>base64_encoded_ink_data</Data>
</InkDrawing>
```

### 4. 删除墨迹不需要二进制数据

删除墨迹只需要 `objectId`，可以直接用 `DeletePageContent(objectId)` API。这个 API 是单条删除，非常快（~1ms/条）。

---

## 优化案例：ClearInk 清除墨迹

### 优化前（慢 ~43秒）

```csharp
// Step 1: 获取选中信息（慢 ~23秒）
OneNoteApplication.GetPageContent(pageId, out string xml,
    Microsoft.Office.Interop.OneNote.PageInfo.piBinaryDataSelection);

// Step 2: 获取完整页面（慢 ~20秒）
OneNoteApplication.GetPageContent(pageId, out string xml,
    Microsoft.Office.Interop.OneNote.PageInfo.piBinaryData);

// Step 3: 删除
foreach (var ink in inkElements)
{
    OneNoteApplication.DeletePageContent(pageId, objectId);
}
```

### 优化后（快 ~40ms）

```csharp
// Step 1: 获取选中状态（快 ~20ms）
OneNoteApplication.GetPageContent(pageId, out string xmlSelection,
    Microsoft.Office.Interop.OneNote.PageInfo.piSelection);
var selectedIds = ParseSelectedObjectIds(xmlSelection);

// Step 2: 获取所有墨迹 ID（快 ~20ms）
OneNoteApplication.GetPageContent(pageId, out string xmlBasic,
    Microsoft.Office.Interop.OneNote.PageInfo.piBasic);
var allInkIds = ParseAllInkObjectIds(xmlBasic);

// Step 3: 删除（快 ~1ms/条）
foreach (var objectId in allInkIds)
{
    if (HasSelection && !selectedIds.Contains(objectId))
        continue;
    OneNoteApplication.DeletePageContent(pageId, objectId);
}
```

### 性能对比

| 步骤 | 优化前 | 优化后 |
|------|--------|--------|
| API 调用 | piBinaryDataSelection (~23秒) | piSelection (~20ms) |
| 获取页面 | piBinaryData (~20秒) | piBasic (~20ms) |
| 删除 | ~1ms/条 | ~1ms/条 |
| **总计** | **~43秒** | **~40ms** |

---

## 教训总结

1. **不要只看 API 名字猜测行为** - `piSelection` 并不是"只返回选中内容"，而是"返回全量+标记"

2. **根据实际需要选择 API** - 如果只需要 objectId，就不需要调用返回二进制数据的 API

3. **性能问题往往在 API 调用策略** - 删除操作本身很快，慢的是获取数据的 API

4. **测试是验证假设的唯一方式** - 不能假设某个 API 返回什么结构，必须看实际输出

---

## 相关代码参考

- `OneInk/AddIn.cs` - `ClearInkButtonClicked` 方法实现
- `docs/How-to-extract-color-from-ink.md` - 墨迹颜色提取经验
