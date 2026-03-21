# OneInk Ribbon 开发经验总结

## 1. Ribbon `loadImage` 回调返回类型

OneNote Ribbon `loadImage` 回调的返回类型必须是 `IStream`，**不是** `IPictureDisp`。

```csharp
// ✅ 正确：返回 IStream
public IStream GetImage(string imageName)
{
    using (var bitmap = new Bitmap(imagePath))
        return bitmap.GetReadOnlyStream();
}

// ❌ IPictureDisp 在 OneNote 中无法正常工作
public IPictureDisp GetImage(string imageName) { ... }
```

`IStream` 的实现参考 `ReadOnlyStream.cs`（来自 OneMore 项目）。

## 2. `getImage` 回调签名

Ribbon XML 中有两种方式使用 `getImage`：

### 方式一：`loadImage`（全局回调）
```xml
<customUI onLoad="OnRibbonLoad" loadImage="GetImage">
```
所有按钮的 `image` 属性直接指定图片名，回调签名：
```csharp
public IStream GetImage(string imageName) { ... }
```

### 方式二：`getImage`（控件级回调）
```xml
<button id="btn1" getImage="GetButtonImage" ... />
```
每个需要动态图片的控件单独指定，回调签名：
```csharp
public IStream GetImage(string imageName) { ... }
```

> 注意：`getImage` 回调的 `imageName` 参数就是 XML 中指定的名字字符串，不是控件 ID。

## 3. `splitButton` 控件组合

`splitButton` 将一个主按钮和一个下拉菜单组合在一起：

```xml
<splitButton id="splitToDashed" size="large">
    <button id="buttonToDashed" getLabel="GetLabel" getScreentip="GetScreentip"
            onAction="ToDashedInkButtonClicked" getImage="GetToDashedImage"/>
    <menu>
        <button id="menuItemDense" label="密集" onAction="OnMenuDenseClicked" image="ToDashedDense.png"/>
        <button id="menuItemMedium" label="中等" onAction="OnMenuMediumClicked" image="ToDashedMedium.png"/>
        <button id="menuItemSparse" label="稀疏" onAction="OnMenuSparseClicked" image="ToDashedSparse.png"/>
    </menu>
</splitButton>
```

- 点击主按钮触发 `onAction`
- 点击下拉菜单项也触发各自的 `onAction`
- **重要**：OneNote 中 `menu` 内的 `button` 的 `onAction` 回调是**可靠**的，与 `dropDown` 不同

## 4. `dropDown` 的 `onAction` 问题

OneNote 中 `dropDown` 控件的 `onAction` 回调**不会触发**：

```xml
<!-- ❌ onAction 在 OneNote 的 dropDown 上不工作 -->
<dropDown id="dropDownDensity"
          getSelectedItemIndex="GetDashedDensitySelectedIndex"
          onAction="OnDashedDensityChanged"
          getItemCount="GetDashedDensityItemCount"
          getItemLabel="GetDashedDensityItemLabel"/>
```

**解决方案**：用 `splitButton` + `menu` + `button` 代替 `dropDown`，每个菜单项按钮单独 `onAction`。

## 5. 动态标签更新（Invalidate）

当菜单项被点击后需要更新主按钮的标签，使用 `IRibbonUI.InvalidateControl`：

```csharp
private static IRibbonUI _ribbon;

public void OnRibbonLoad(IRibbonUI ribbon)
{
    _ribbon = ribbon;
}

// 在菜单项点击回调中
public void OnMenuDenseClicked(IRibbonControl control)
{
    _selectedDensityIndex = 0;
    if (_ribbon != null)
        _ribbon.InvalidateControl("buttonToDashed"); // 触发 GetLabel 重新渲染
}
```

必须在 Ribbon XML 中声明 `onLoad` 回调：
```xml
<customUI xmlns="..." loadImage="GetImage" onLoad="OnRibbonLoad">
```

## 6. `splitButton` 中 `getImage` 动态图标

主按钮的 `getImage` 回调根据当前选中密度返回不同图标：

```xml
<button id="buttonToDashed" getLabel="GetLabel" getScreentip="GetScreentip"
        onAction="ToDashedInkButtonClicked" getImage="GetToDashedImage"/>
```

```csharp
public IStream GetToDashedImage(string imageName)
{
    string densityImage = _selectedDensityIndex switch
    {
        0 => "ToDashedDense.png",
        2 => "ToDashedSparse.png",
        _ => "ToDashedMedium.png"
    };
    return GetImage(densityImage);
}
```

## 7. `box` 控件

`box` 可以将多个控件水平或垂直排列：

```xml
<box id="boxToDashed" boxStyle="horizontal">
    <button id="buttonToDashed" getLabel="GetLabel" .../>
    <dropDown id="dropDownDensity" .../>
</box>
```

`boxStyle` 可选 `horizontal` 或 `vertical`。注意 `dropDown` 的 `onAction` 在 OneNote 中不触发的问题。

## 8. menu 与 dropDown 的区别

| 特性 | `menu` | `dropDown` |
|------|--------|------------|
| OneNote onAction 触发 | ✅ 是 | ❌ 否 |
| 支持多级菜单 | ✅ | ❌ |
| 显示选中状态 | ❌ | ✅ |
| `getSelectedItemIndex` 回调 | N/A | 支持但不触发 |

## 9. C# 版本注意

项目使用 .NET Framework 4.8（C# 7.3），**不支持** switch 表达式（需要 C# 8.0+）：

```csharp
// ❌ C# 7.3 不支持
var x = index switch { 0 => "a", _ => "b" };

// ✅ 正确写法
string x;
switch (index)
{
    case 0: x = "a"; break;
    default: x = "b"; break;
}
```
