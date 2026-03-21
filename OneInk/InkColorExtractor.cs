/*
 *  InkColorExtractor - Extracts ink colors from OneNote ISF data.
 *  Uses Microsoft.Ink.Ink class to load and parse ISF.
 */

using System;
using System.Collections.Generic;
using System.IO;

namespace OneInk
{
    public static class InkColorExtractor
    {
        private static readonly string LogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Temp", "OneInk.log");

        private static void Log(string msg)
        {
            try
            {
                File.AppendAllText(LogPath,
                    $"[{DateTime.Now:HH:mm:ss.fff}] [ISF] {msg}" + Environment.NewLine);
            }
            catch { }
        }

        public static List<string> ExtractInkColors(string base64Data)
        {
            var colors = new List<string>();
            if (string.IsNullOrEmpty(base64Data))
                return colors;

            try
            {
                byte[] raw = Convert.FromBase64String(base64Data);

                // Strategy 1: Use Microsoft.Ink.Ink to parse ISF
                // Ink class handles all compression/decompression internally
                var inkColors = TryLoadWithMicrosoftInk(raw);
                if (inkColors != null && inkColors.Count > 0)
                    return inkColors;

                // Strategy 2: Strip 0x00 prefix and try again
                if (raw.Length > 0 && raw[0] == 0x00)
                {
                    var stripped = new byte[raw.Length - 1];
                    Array.Copy(raw, 1, stripped, 0, stripped.Length);
                    inkColors = TryLoadWithMicrosoftInk(stripped);
                    if (inkColors != null && inkColors.Count > 0)
                        return inkColors;
                }

                colors.Add("#000000");
            }
            catch (Exception ex)
            {
                Log($"EXCEPTION: {ex.Message}");
                colors.Add("#000000");
            }

            return colors;
        }

        /// <summary>
        /// Uses Microsoft.Ink.Ink class to load ISF data and extract stroke colors.
        /// The Ink class handles compression internally.
        /// </summary>
        private static List<string> TryLoadWithMicrosoftInk(byte[] isfData)
        {
            try
            {
                // Create a new Ink object
                var ink = new Microsoft.Ink.Ink();

                // Load ISF data - Ink.Load handles all compression internally
                ink.Load(isfData);

                var strokes = ink.Strokes;

                if (strokes.Count == 0)
                {
                    ink.Dispose();
                    return null;
                }

                var colors = new List<string>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (Microsoft.Ink.Stroke stroke in strokes)
                {
                    var color = stroke.DrawingAttributes.Color;
                    string hex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";

                    if (seen.Add(hex))
                        colors.Add(hex);
                }

                ink.Dispose();
                return colors.Count > 0 ? colors : null;
            }
            catch (Exception ex)
            {
                Log($"Microsoft.Ink load error: {ex.Message}");
                return null;
            }
        }
    }
}
