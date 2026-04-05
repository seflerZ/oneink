# Partial Selection 处理方案

## 问题背景

当用户在 OneNote 中复制多个独立绘制的墨迹对象时，OneNote 会将它们**合并成一个 InkDrawing**（包含多个 stroke）。选择这个合并后的 InkDrawing 时，OneNote 在 XML 中标记为 `selected="partial"`，但返回的 ISF 数据始终是**完整的所有 stroke**，而不是用户实际选中的子集。

## 根本原因

### OneNote API 限制

1. **`GetPageContent` with `piBinaryDataSelection`**：返回 `selected="partial"` 标记，但 ISF 数据是完整的
2. **无笔画级选择信息**：API 不提供用户实际选中了哪些 stroke 的信息
3. **无选区边界信息**：无法获取用户框选的坐标范围

## 解决方案：拆分后重新选择

采用**拆分后让用户重新选择**的交互方案：

```
用户框选 merged InkDrawing 中的部分笔画
        ↓
点击操作按钮（转虚线/平滑/对齐等）
        ↓
检测到 selected="partial"
        ↓
弹出确认对话框："检测到只有一个墨迹容器，需要拆分后才能操作，是否拆分？"
        ↓
执行拆分：聚类 strokes → 创建多个独立 InkDrawing → 删除原 InkDrawing
        ↓
显示结果：
  - 拆分成功："拆分完成，请重新选择墨迹进行操作"
  - 无法拆分："拆分失败，容器不可再拆"
        ↓
用户重新选择想要操作的独立 InkDrawing
        ↓
再次点击操作按钮，正常执行
```

## 实现细节

### 1. 核心组件

| 组件 | 文件 | 职责 |
|------|------|------|
| `CheckAndHandlePartialSelection` | `AddIn.cs` | 检测 partial selection，协调拆分流程 |
| `HandlePartialSelection` | `AddIn.cs` | 执行实际的拆分操作 |
| `SplitInkDrawingByConnectivity` | `InkDashedConverter.cs` | 按连通性聚类 strokes |
| `ExtractStrokes` | `InkDashedConverter.cs` | 提取指定 strokes 生成新 ISF |

### 2. 关键算法

#### Stroke 聚类（连通性分析）

使用并查集（Union-Find）算法按空间连通性聚类：

```csharp
// 参数配置
SamplingInterval = 500;  // 采样间隔（HIMETRIC）
DistanceThreshold = 20;   // 连通距离阈值（HIMETRIC）

// 距离计算：页面坐标系下的欧氏距离
double dx = (x1 - x2) * scaleX;
double dy = (y1 - y2) * scaleY;
double distance = Math.Sqrt(dx * dx + dy * dy);
```

#### 坐标变换

ISF 内部坐标与页面坐标的转换：

```csharp
// 缩放因子计算
scaleX = pageSizeW / isfBoundsW;
scaleY = pageSizeH / isfBoundsH;

// 新位置计算
newPosX = pagePosX + (isfX - isfMinX) * scaleX;
newPosY = pagePosY + (isfY - isfMinY) * scaleY;

// 新大小计算
newSizeW = isfW * scaleX;
newSizeH = isfH * scaleY;
```

#### ObjectID 格式

OneNote 要求的 objectID 格式：

```csharp
// 格式：{GUID}{数字}{字母+数字}
// 示例：{548E953B-0931-4F1A-9C32-1F582198B627}{33}{B0}

string newId = "{" + Guid.NewGuid().ToString("B").ToUpperInvariant().Trim('{', '}') + "}{1}{A0}";
```

### 3. DrawingAttributes 保留

拆分过程中显式复制所有 DrawingAttributes 属性：

```csharp
var newAttr = new DrawingAttributes
{
    Width = attr.Width,
    Height = attr.Height,
    Color = attr.Color,
    FitToCurve = attr.FitToCurve,
    IgnorePressure = attr.IgnorePressure,
    AntiAliased = attr.AntiAliased,
    RasterOperation = attr.RasterOperation,
    Transparency = attr.Transparency
};
```

## 已知限制

### 1. 笔迹渲染差异

**现象**：拆分后的独立 InkDrawing 与原始合并容器中的笔迹在视觉上可能有细微差异（如宽度感知上的轻微变化）。

**原因**：
- OneNote 对同一容器内的多个笔画可能应用了特殊的渲染优化（如抗锯齿、统一缩放）
- 独立容器的渲染方式与合并容器略有不同
- 坐标缩放计算的精度限制（浮点数转 XML 字符串）

**验证**：
- ISF 数据中的 `DrawingAttributes.Width` 被正确保留（日志验证）
- 保存并重新加载后属性值一致
- 差异在 OneNote 渲染层，而非数据层

### 2. 无法自动完成操作

由于 API 不提供笔画级选择信息，**无法在拆分后自动对用户原本选中的笔画执行操作**。必须让用户重新选择。

### 3. 交互流程增加

相比直接操作，此方案增加了一个确认步骤和重新选择步骤，但这是 API 限制下的必要折中。

## 代码集成

### 在操作按钮中集成检测

所有操作按钮（转虚线、平滑、对齐、清除等）都需要在开头检测 partial selection：

```csharp
private void ExecuteToDashed()
{
    // 1. 获取选择信息
    OneNoteApplication.GetPageContent(pageId, out xmlSel, PageInfo.piBinaryDataSelection);
    
    // 2. 检测并处理 partial selection
    if (CheckAndHandlePartialSelection(docSel, pageId, selSettings, ns, "ToDashed"))
        return; // 已处理，等待用户重新选择
    
    // 3. 正常流程（此时都是 selected="all"）
    // ...
}
```

## 测试场景

1. **正常拆分**：选择包含多个连通区域的 merged InkDrawing → 确认拆分 → 成功拆分为多个独立对象
2. **无法拆分**：选择只包含单个连通区域的 InkDrawing → 提示"容器不可再拆"
3. **用户取消**：弹出确认对话框 → 选择"否" → 操作取消
4. **各操作集成**：转虚线、平滑、对齐、清除等操作都能正确触发拆分流程

## 结论

Partial selection 问题通过**拆分后重新选择**的方案已解决。该方案在 OneNote COM API 的限制下提供了可行的用户体验，尽管增加了交互步骤，但确保了操作的精确性和可预测性。
