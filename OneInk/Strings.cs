/*
 *  Strings - Localization for OneInk
 *  Supports: Chinese (zh-CN), English (en-US)
 */

using System;
using System.Collections.Generic;
using System.Globalization;

namespace OneInk
{
    internal static class Strings
    {
        private static readonly bool IsChinese =
            CultureInfo.CurrentUICulture.Name.StartsWith("zh", StringComparison.OrdinalIgnoreCase);

        // Dialog: Color Selection
        internal static string DialogTitle => IsChinese ? "选择墨迹颜色" : "Select Ink Color";
        internal static string DialogHeader => IsChinese
            ? "选择一个颜色，删除该颜色的所有墨迹："
            : "Select a color to delete all ink strokes of that color:";
        internal static string OkButton => IsChinese ? "删除选中颜色" : "Delete Selected Color";
        internal static string CancelButton => IsChinese ? "取消" : "Cancel";
        internal static string NoSelection => IsChinese ? "请先选择一个颜色。" : "Please select a color first.";
        internal static string NoSelectionTitle => IsChinese ? "未选择" : "No Selection";

        // MessageBox
        internal static string AppNotAvailable => IsChinese ? "OneNote 应用程序不可用。" : "OneNote application not available.";
        internal static string RetrieveFailed => IsChinese ? "无法获取页面内容。" : "Could not retrieve page content.";
        internal static string NoInkStrokes => IsChinese ? "当前页面没有墨迹。" : "No ink strokes found on the current page.";
        internal static string NoInkStrokesInSelection => IsChinese ? "所选区域内没有墨迹。" : "No ink strokes found in the selected area.";
        internal static string NoSelectionForDashed => IsChinese ? "请先选择要转换的墨迹（套索工具框选）" : "Please select ink strokes to convert (use lasso to select).";
        internal static string ErrorClear => IsChinese ? "清除墨迹时出错：{0}" : "Error clearing ink: {0}";
        internal static string ErrorDelete => IsChinese ? "按颜色删除墨迹时出错：{0}" : "Error deleting ink by color: {0}";
        internal static string ErrorDashed => IsChinese ? "转换为虚线墨迹时出错：{0}" : "Error converting to dashed ink: {0}";
        internal static string ErrorSmooth => IsChinese ? "平滑墨迹时出错：{0}" : "Error smoothing ink: {0}";
        internal static string DashedSuccess => IsChinese ? "已成功转换 {0} 个墨迹为虚线样式。" : "Converted {0} ink stroke(s) to dashed style.";

        // Ribbon
        internal static string RibbonTabLabel => IsChinese ? "OneInk" : "OneInk";
        internal static string RibbonGroupLabel => IsChinese ? "墨迹工具" : "Ink Tools";
        internal static string ButtonClearInkLabel => IsChinese ? "清除全部墨迹" : "Clear All Ink";
        internal static string ButtonDeleteByColorLabel => IsChinese ? "按颜色删除" : "Delete by Color";
        internal static string ButtonToDashedLabel => IsChinese ? "转为虚线" : "To Dashed";
        internal static string ButtonSmoothLabel => IsChinese ? "平滑至" : "Smooth";
        internal static string ButtonSmoothCurveLabel => IsChinese ? "曲线" : "Curve";
        internal static string ButtonSmoothPolyLabel => IsChinese ? "折线" : "Polyline";
        internal static string ButtonClearInkScreentip => IsChinese ? "删除所选区域（或整页）的所有墨迹" : "Remove ink strokes from selection or entire page";
        internal static string ButtonDeleteByColorScreentip => IsChinese ? "按所选颜色删除所选区域（或整页）的墨迹" : "Delete ink strokes by color from selection or entire page";
        internal static string ButtonToDashedScreentip => IsChinese ? "将墨迹转换为虚线样式" : "Convert ink strokes to dashed lines";
        internal static string ButtonSmoothCurveScreentip => IsChinese ? "将墨迹转换为平滑的贝塞尔曲线" : "Convert ink strokes to smooth Bezier curves";
        internal static string ButtonSmoothPolyScreentip => IsChinese ? "将墨迹转换为保留转角的折线" : "Convert ink strokes to simplified polylines preserving corners";
        internal static string ButtonDenseScreentip => IsChinese ? "密集虚线" : "Dense dashed lines";
        internal static string ButtonSparseScreentip => IsChinese ? "稀疏虚线" : "Sparse dashed lines";

        // Ribbon: Dashed Density
        internal static string DashedDense => IsChinese ? "密集" : "Dense";
        internal static string DashedMedium => IsChinese ? "中等" : "Medium";
        internal static string DashedSparse => IsChinese ? "稀疏" : "Sparse";
    }
}
