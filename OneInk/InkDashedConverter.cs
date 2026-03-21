/*
 *  InkDashedConverter - Converts ink strokes to dashed/dotted lines.
 *  Uses Microsoft.Ink.Ink to load ISF, extracts stroke geometry,
 *  modifies it by removing gap points, then creates new strokes
 *  using CreateStroke(Point[]) which produces standard ISF that OneNote accepts.
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Microsoft.Ink;

namespace OneInk
{
    public static class InkDashedConverter
    {
        public static float DefaultDashFraction { get; set; } = 0.05f;
        public static float DefaultGapFraction { get; set; } = 0.05f;

        public static string ConvertToDashed(string base64Data)
        {
            return ConvertToDashed(base64Data, DefaultDashFraction, DefaultGapFraction);
        }

        public static string ConvertToDashed(string base64Data, float dashFraction, float gapFraction)
        {
            if (string.IsNullOrEmpty(base64Data))
                return null;

            byte[] raw;
            try
            {
                raw = Convert.FromBase64String(base64Data);
            }
            catch
            {
                return null;
            }

            try
            {
                var result = ConvertToDashedCore(raw, dashFraction, gapFraction);
                if (result != null)
                    return Convert.ToBase64String(result);
                return null;
            }
            catch
            {
                return null;
            }
        }

        private static byte[] ConvertToDashedCore(byte[] isfData, float dashFraction, float gapFraction)
        {
            byte[] data = isfData;

            // Try ZLIB decompression if the data appears compressed
            if (isfData.Length > 0 && isfData[0] == 0x00 && isfData.Length > 4)
            {
                byte second = isfData[1];
                if (second != 0x01)
                {
                    try { data = DecompressZlib(isfData); }
                    catch { }
                }
            }

            Ink srcInk = null;
            try
            {
                srcInk = new Ink();
                srcInk.Load(data);
            }
            catch
            {
                // Try stripping 0x00 prefix as fallback
                if (isfData.Length > 0 && isfData[0] == 0x00)
                {
                    var stripped = new byte[isfData.Length - 1];
                    Array.Copy(isfData, 1, stripped, 0, stripped.Length);
                    try
                    {
                        srcInk = new Ink();
                        srcInk.Load(stripped);
                    }
                    catch
                    {
                        if (srcInk != null) { srcInk.Dispose(); srcInk = null; }
                    }
                }
                if (srcInk == null)
                    return null;
            }

            if (srcInk.Strokes.Count == 0)
            {
                srcInk.Dispose();
                return null;
            }

            var strokesData = new List<StrokeGeometry>(srcInk.Strokes.Count);
            foreach (Stroke s in srcInk.Strokes)
            {
                strokesData.Add(ExtractStrokeGeometry(s));
            }
            srcInk.Dispose();

            var dstInk = new Ink();
            foreach (var sg in strokesData)
            {
                CreateDashedStroke(dstInk, sg, dashFraction, gapFraction);
            }

            byte[] result = dstInk.Save();
            dstInk.Dispose();
            return result;
        }

        private static byte[] DecompressZlib(byte[] data)
        {
            using (var input = new MemoryStream(data))
            using (var zlib = new System.IO.Compression.DeflateStream(input, System.IO.Compression.CompressionMode.Decompress))
            using (var output = new MemoryStream())
            {
                zlib.CopyTo(output);
                return output.ToArray();
            }
        }

        private class StrokeGeometry
        {
            public List<Point> Points { get; set; }
            public DrawingAttributes Attr { get; set; }
        }

        private static StrokeGeometry ExtractStrokeGeometry(Stroke stroke)
        {
            var sg = new StrokeGeometry();
            sg.Points = ResampleStroke(stroke.GetFlattenedBezierPoints(), 200);
            sg.Attr = stroke.DrawingAttributes.Clone();
            return sg;
        }

        private static List<Point> ResampleStroke(Point[] pts, int numPoints)
        {
            if (pts == null || pts.Length < 2)
                return new List<Point>(pts ?? new Point[0]);

            // Pre-compute cumulative lengths once
            double[] cumLen = new double[pts.Length];
            cumLen[0] = 0;
            double totalLen = 0;
            for (int i = 1; i < pts.Length; i++)
            {
                double dx = pts[i].X - pts[i - 1].X;
                double dy = pts[i].Y - pts[i - 1].Y;
                totalLen += Math.Sqrt(dx * dx + dy * dy);
                cumLen[i] = totalLen;
            }

            if (totalLen <= 0)
                return new List<Point>(pts);

            // Resample at equal geometric distances - O(n) using running pointer
            var result = new List<Point>(numPoints);
            result.Add(pts[0]);

            int ptr = 0;
            for (int i = 1; i < numPoints - 1; i++)
            {
                double targetDist = (totalLen * i) / (numPoints - 1);
                while (ptr < pts.Length - 1 && cumLen[ptr + 1] <= targetDist)
                    ptr++;
                if (ptr >= pts.Length - 1)
                {
                    result.Add(pts[pts.Length - 1]);
                    continue;
                }
                double segLen = cumLen[ptr + 1] - cumLen[ptr];
                double t = segLen > 0 ? (targetDist - cumLen[ptr]) / segLen : 0;
                int x = (int)Math.Round(pts[ptr].X + t * (pts[ptr + 1].X - pts[ptr].X));
                int y = (int)Math.Round(pts[ptr].Y + t * (pts[ptr + 1].Y - pts[ptr].Y));
                result.Add(new Point(x, y));
            }
            result.Add(pts[pts.Length - 1]);

            return result;
        }

        private static void CreateDashedStroke(Ink ink, StrokeGeometry sg,
            float dashFraction, float gapFraction)
        {
            var pts = sg.Points;
            if (pts.Count < 2)
                return;

            // Calculate cumulative path length
            var cumLen = new double[pts.Count];
            cumLen[0] = 0;
            double totalLen = 0;
            for (int i = 1; i < pts.Count; i++)
            {
                double dx = pts[i].X - pts[i - 1].X;
                double dy = pts[i].Y - pts[i - 1].Y;
                totalLen += Math.Sqrt(dx * dx + dy * dy);
                cumLen[i] = totalLen;
            }

            if (totalLen <= 0)
            {
                // Zero-length stroke: just create it as-is
                try
                {
                    var s = ink.CreateStroke(pts.ToArray());
                    s.DrawingAttributes = sg.Attr;
                    s.DrawingAttributes.FitToCurve = true;
                }
                catch { }
                return;
            }

            double dashLen = Math.Max(totalLen * dashFraction, 1.0);
            double gapLen = Math.Max(totalLen * gapFraction, 1.0);

            bool inDash = true;
            double pos = 0;

            while (pos < totalLen)
            {
                if (inDash)
                {
                    int startIdx = FindIdx(cumLen, pos);
                    double dashEnd = Math.Min(pos + dashLen, totalLen);
                    int endIdx = FindIdx(cumLen, dashEnd);

                    if (endIdx > startIdx && endIdx - startIdx >= 2)
                    {
                        var seg = new Point[endIdx - startIdx];
                        for (int i = startIdx; i < endIdx; i++)
                            seg[i - startIdx] = pts[i];
                        try
                        {
                            var s = ink.CreateStroke(seg);
                            s.DrawingAttributes = sg.Attr;
                            s.DrawingAttributes.FitToCurve = true;
                        }
                        catch { }
                    }

                    pos = dashEnd;
                    inDash = false;
                }
                else
                {
                    pos = Math.Min(pos + gapLen, totalLen);
                    inDash = true;
                }
            }
        }

        private static int FindIdx(double[] cumLen, double target)
        {
            int n = cumLen.Length;
            if (target <= cumLen[0]) return 0;
            if (target >= cumLen[n - 1]) return n - 1;
            int lo = 0, hi = n - 1;
            while (lo < hi)
            {
                int mid = (lo + hi + 1) / 2;
                if (cumLen[mid] <= target) lo = mid; else hi = mid - 1;
            }
            return lo;
        }
    }
}
