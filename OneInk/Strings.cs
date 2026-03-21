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
        internal static string ErrorClear => IsChinese ? "清除墨迹时出错：{0}" : "Error clearing ink: {0}";
        internal static string ErrorDelete => IsChinese ? "按颜色删除墨迹时出错：{0}" : "Error deleting ink by color: {0}";
    }
}
