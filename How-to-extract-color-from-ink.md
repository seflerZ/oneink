# 墨迹颜色解析问题排查与解决方案

## 问题背景

OneNote 中绘制的墨迹（Ink）以 ISF（Ink Serialized Format）格式存储在页面 XML 的 `<Data>` 节点中。在尝试按颜色删除墨迹时，需要正确解析这些二进制数据以获取每个笔划的实际颜色值。

## 关键发现

### 1. ISF 数据的压缩方式

OneNote 返回的墨迹二进制数据是经过 **ZLIB 压缩**的原始 ISF 流。直接解析压缩后的乱码数据无法匹配到 ISF 的结构，自然无法读取颜色和笔宽信息。

### 2. 正确的解析流程

必须先对拿到的二进制数据进行 **ZLIB 解压**，得到原始的 ISF 流之后，再进行后续的解析，才能正确读取到 OneNote 自定义属性中的颜色和笔宽信息。

### 3. 最简解决方案：Microsoft.Ink.dll

Tablet PC SDK 提供的 `Microsoft.Ink.dll` 封装了完整的 ISF 解析逻辑（包括 ZLIB 解压）。使用 `Microsoft.Ink.Ink.Load(byte[])` 方法即可自动处理所有压缩/解析步骤，直接获取笔划颜色。

```csharp
var ink = new Microsoft.Ink.Ink();
ink.Load(isfData);  // 自动处理 ZLIB 解压
foreach (Microsoft.Ink.Stroke stroke in ink.Strokes)
{
    var color = stroke.DrawingAttributes.Color;
}
```

### 4. 为什么手动解析 ISF 困难

- ISF 格式本身的结构复杂，包含元数据、笔划点数据、属性数据等多个部分
- ZLIB 压缩层的存在使得手动解析前需要额外的解压步骤
- 笔划颜色存储在 DrawingAttributes 中，位置不固定
- OneNote 可能对标准 ISF 格式有定制化处理

## 排查过程

1. **初步尝试**：直接对 base64 解码后的数据进行字节级搜索，查找颜色值（如 #E71225 的 RGB/BGR 排列），未成功
2. **ZLIB 疑云**：尝试使用多种 ZLIB 解压库和方法，格式不匹配导致解压失败
3. **发现关键 API**：`Microsoft.Ink.dll` 中的 `Ink.Load()` 方法直接处理压缩的 ISF 数据，一次调用即可获得完整解析结果

## 经验总结

- 处理 Office/OneNote 的墨迹数据时，优先查找官方或 Tablet PC SDK 提供的 COM 组件，而非自行解析二进制格式
- ISF 数据通常是压缩的（常见为 ZLIB），直接解析是走不通的
- `Microsoft.Ink.dll` 是处理墨迹数据的标准方案，功能完善且经过验证

## 相关技术

- **Microsoft.Ink.dll**：Tablet PC Ink API，位置 `C:\Windows\assembly\GAC_64\Microsoft.Ink\6.1.0.0__31bf3856ad364e35\`
- **ISF (Ink Serialized Format)**：墨迹序列化格式，用于存储笔划、点数据、属性等
- **ZLIB**：压缩算法，用于压缩 ISF 数据
- **IDTExtensibility2**：Office COM Add-In 接口标准
- **IRibbonExtensibility**：Office 功能区（Ribbon）扩展接口
