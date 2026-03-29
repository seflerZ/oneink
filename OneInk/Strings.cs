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
        internal static readonly bool IsChinese =
            CultureInfo.CurrentUICulture.Name.StartsWith("zh", StringComparison.OrdinalIgnoreCase);

        // Dialog: Color Selection
        internal static string DialogTitle => IsChinese ? "选择墨迹颜色" : "Select Ink Color";
        internal static string DialogHeader => IsChinese
            ? "勾选要删除的颜色（可多选）："
            : "Check colors to delete (multi-select):";
        internal static string OkButton => IsChinese ? "删除所选颜色" : "Delete Selected Colors";
        internal static string CancelButton => IsChinese ? "取消" : "Cancel";
        internal static string NoSelection => IsChinese ? "请先选择至少一个颜色。" : "Please select at least one color.";
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
        internal static string ErrorAlign => IsChinese ? "对齐墨迹时出错：{0}" : "Error aligning ink: {0}";
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
        internal static string ButtonAlignLabel => IsChinese ? "对齐" : "Align";
        internal static string ButtonAlignTopLabel => IsChinese ? "顶边对齐" : "Align Top";
        internal static string ButtonAlignBottomLabel => IsChinese ? "底边对齐" : "Align Bottom";
        internal static string ButtonAlignLeftLabel => IsChinese ? "左边对齐" : "Align Left";
        internal static string ButtonAlignRightLabel => IsChinese ? "右边对齐" : "Align Right";
        internal static string ButtonClearInkScreentip => IsChinese ? "删除所选区域（或整页）的所有墨迹" : "Remove ink strokes from selection or entire page";
        internal static string ButtonDeleteByColorScreentip => IsChinese ? "按所选颜色删除所选区域（或整页）的墨迹" : "Delete ink strokes by color from selection or entire page";
        internal static string ButtonToDashedScreentip => IsChinese ? "将墨迹转换为虚线样式" : "Convert ink strokes to dashed lines";
        internal static string ButtonSmoothCurveScreentip => IsChinese ? "将墨迹转换为平滑的贝塞尔曲线" : "Convert ink strokes to smooth Bezier curves";
        internal static string ButtonSmoothPolyScreentip => IsChinese ? "将墨迹转换为保留转角的折线" : "Convert ink strokes to simplified polylines preserving corners";
        internal static string ButtonAlignTopScreentip => IsChinese ? "将所选墨迹顶边对齐到第一个所选墨迹" : "Align top edges of selected ink to the first selected ink";
        internal static string ButtonAlignBottomScreentip => IsChinese ? "将所选墨迹底边对齐到第一个所选墨迹" : "Align bottom edges of selected ink to the first selected ink";
        internal static string ButtonAlignLeftScreentip => IsChinese ? "将所选墨迹左边对齐到第一个所选墨迹" : "Align left edges of selected ink to the first selected ink";
        internal static string ButtonAlignRightScreentip => IsChinese ? "将所选墨迹右边对齐到第一个所选墨迹" : "Align right edges of selected ink to the first selected ink";
        internal static string ButtonDenseScreentip => IsChinese ? "密集虚线" : "Dense dashed lines";
        internal static string ButtonSparseScreentip => IsChinese ? "稀疏虚线" : "Sparse dashed lines";

        // Ribbon: Dashed Density
        internal static string DashedDense => IsChinese ? "密集" : "Dense";
        internal static string DashedMedium => IsChinese ? "中等" : "Medium";
        internal static string DashedSparse => IsChinese ? "稀疏" : "Sparse";
    }
}
