/*
 *  InkDashedConverter - Ink stroke manipulation via Microsoft.Ink API.
 *  Supports:
 *    - ConvertToDashed: Convert strokes to dashed/dotted lines
 *    - SmoothStroke: Smooth strokes (curve or polyline)
 *  Uses CreateStroke(Point[]) to produce standard ISF that OneNote accepts.
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using Microsoft.Ink;

namespace OneInk
{
    public static class InkDashedConverter
    {
        // Fixed dash/gap sizes in ink units (HIMETRIC: 1 unit = 0.001 inch)
        public const int DenseDashGap = 120;
        public const int MediumDashGap = 250;
        public const int SparseDashGap = 500;

        public static string ConvertToDashed(string base64Data)
        {
            return ConvertToDashed(base64Data, MediumDashGap);
        }

        public static string ConvertToDashed(string base64Data, int dashGapSize)
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
                var result = ConvertToDashedCore(raw, dashGapSize);
                if (result != null)
                    return Convert.ToBase64String(result);
                return null;
            }
            catch
            {
                return null;
            }
        }

        private static byte[] ConvertToDashedCore(byte[] isfData, int dashGapSize)
        {
            Ink srcInk = LoadIsf(isfData, out _);
            if (srcInk == null)
                return null;

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
                CreateDashedStroke(dstInk, sg, dashGapSize);
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

        /// <summary>
        /// Load ISF data into an Ink object. Handles ZLIB compression and 0x00 prefix.
        /// </summary>
        private static Ink LoadIsf(byte[] isfData, out byte[] remaining)
        {
            remaining = isfData;
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

            Ink ink = null;
            try
            {
                ink = new Ink();
                ink.Load(data);
                remaining = data;
                return ink;
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
                        ink = new Ink();
                        ink.Load(stripped);
                        remaining = stripped;
                        return ink;
                    }
                    catch
                    {
                        if (ink != null) { ink.Dispose(); ink = null; }
                    }
                }
            }
            remaining = data;
            return ink;
        }

        public static string SmoothStroke(string base64Data, bool curveSmoothing)
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
                var result = SmoothStrokeCore(raw, curveSmoothing);
                if (result != null)
                    return Convert.ToBase64String(result);
                return null;
            }
            catch
            {
                return null;
            }
        }

        private static byte[] SmoothStrokeCore(byte[] isfData, bool curveSmoothing)
        {
            Ink srcInk = LoadIsf(isfData, out _);
            if (srcInk == null)
                return null;

            if (srcInk.Strokes.Count == 0)
            {
                srcInk.Dispose();
                return null;
            }

            var strokesData = new List<StrokeGeometry>(srcInk.Strokes.Count);
            foreach (Stroke s in srcInk.Strokes)
            {
                strokesData.Add(ExtractStrokeGeometryForSmooth(s, curveSmoothing));
            }
            srcInk.Dispose();

            var dstInk = new Ink();
            foreach (var sg in strokesData)
            {
                CreateSmoothedStroke(dstInk, sg, curveSmoothing);
            }

            byte[] result = dstInk.Save();
            dstInk.Dispose();
            return result;
        }

        private static StrokeGeometry ExtractStrokeGeometryForSmooth(Stroke stroke, bool curveSmoothing)
        {
            var sg = new StrokeGeometry();
            var rawPts = GetStrokePoints(stroke);

            // Resample by fixed distance interval (1000 HIMETRIC ≈ 1mm for more detail)
            var resampledPts = ResampleByDistance(rawPts, 1000);

            List<Point> smoothedPts;
            if (curveSmoothing)
            {
                // Chaikin's algorithm for curve smoothing
                smoothedPts = ChaikinsSmooth(resampledPts, 1);
            }
            else
            {
                // Detect corners first before RDP simplification
                var corners = new HashSet<int>();
                for (int i = 1; i < resampledPts.Count - 1; i++)
                {
                    double angle = CalculateAngle(resampledPts[i - 1], resampledPts[i], resampledPts[i + 1]);
                    double deviation = Math.Abs(180 - angle);
                    if (deviation >= 30) // Corner: angle deviates > 30° from straight
                    {
                        corners.Add(i);
                    }
                }

                // RDP simplification preserving corners
                smoothedPts = RDPSimplifyPreservingCorners(resampledPts, 500, corners);
                // Snap angles to multiples of 30 degrees
                smoothedPts = SnapAnglesTo30(smoothedPts);
            }

            sg.Points = smoothedPts;
            sg.Attr = stroke.DrawingAttributes.Clone();
            return sg;
        }

        private static List<Point> ResampleByDistance(List<Point> pts, int interval)
        {
            if (pts == null || pts.Count < 2)
                return pts;

            var result = new List<Point>(pts.Count);
            result.Add(pts[0]);

            double accumDist = 0;
            int lastIdx = 0;
            Point lastPt = pts[0];

            for (int i = 1; i < pts.Count; i++)
            {
                double dx = pts[i].X - lastPt.X;
                double dy = pts[i].Y - lastPt.Y;
                double dist = Math.Sqrt(dx * dx + dy * dy);
                accumDist += dist;

                while (accumDist >= interval)
                {
                    double t = (accumDist - interval) / dist;
                    int x = (int)Math.Round(pts[i].X - t * (pts[i].X - lastPt.X));
                    int y = (int)Math.Round(pts[i].Y - t * (pts[i].Y - lastPt.Y));
                    result.Add(new Point(x, y));

                    accumDist -= interval;
                }

                lastPt = pts[i];
            }

            // Always include the last point
            if (result.Count == 0 || result[result.Count - 1] != pts[pts.Count - 1])
                result.Add(pts[pts.Count - 1]);

            return result;
        }

        private static List<Point> ChaikinsSmooth(List<Point> pts, int iterations)
        {
            if (pts == null || pts.Count < 2 || iterations <= 0)
                return pts;

            // For 1 iteration: result has 2*n - 1 points, direct calculation avoids buffer swapping
            if (iterations == 1)
            {
                int n = pts.Count;
                var result = new List<Point>(n * 2 - 1);
                result.Add(pts[0]);

                for (int i = 0; i < n - 1; i++)
                {
                    int p0x = pts[i].X, p0y = pts[i].Y;
                    int p1x = pts[i + 1].X, p1y = pts[i + 1].Y;

                    // Q = 3/4 * P0 + 1/4 * P1
                    result.Add(new Point((3 * p0x + p1x) / 4, (3 * p0y + p1y) / 4));
                    // R = 1/4 * P0 + 3/4 * P1
                    result.Add(new Point((p0x + 3 * p1x) / 4, (p0y + 3 * p1y) / 4));
                }

                return result;
            }

            // For multiple iterations, use buffer swapping
            int count = pts.Count;
            Point[] arr = pts.ToArray();
            Point[] buffer = new Point[count * 2];

            for (int iter = 0; iter < iterations; iter++)
            {
                int j = 0;
                buffer[j++] = arr[0];

                for (int i = 0; i < count - 1; i++)
                {
                    int p0x = arr[i].X, p0y = arr[i].Y;
                    int p1x = arr[i + 1].X, p1y = arr[i + 1].Y;

                    buffer[j++] = new Point((3 * p0x + p1x) / 4, (3 * p0y + p1y) / 4);
                    buffer[j++] = new Point((p0x + 3 * p1x) / 4, (p0y + 3 * p1y) / 4);
                }

                buffer[j++] = arr[count - 1];

                count = j;
                var temp = arr;
                arr = buffer;
                buffer = temp;
            }

            var final = new List<Point>(count);
            for (int i = 0; i < count; i++)
                final.Add(arr[i]);
            return final;
        }

        private static List<Point> RDPSimplify(List<Point> pts, double epsilon)
        {
            if (pts == null || pts.Count < 3)
                return pts;

            // Iterative RDP with stack to avoid recursion overhead
            var keepFlags = new bool[pts.Count];
            keepFlags[0] = true;
            keepFlags[pts.Count - 1] = true;

            var stack = new Stack<(int start, int end)>();
            stack.Push((0, pts.Count - 1));

            while (stack.Count > 0)
            {
                var (start, end) = stack.Pop();
                if (end - start < 2)
                    continue;

                double maxDist = 0;
                int maxIndex = start;

                var first = pts[start];
                var last = pts[end];

                for (int i = start + 1; i < end; i++)
                {
                    double dist = PerpendicularDistance(pts[i], first, last);
                    if (dist > maxDist)
                    {
                        maxDist = dist;
                        maxIndex = i;
                    }
                }

                if (maxDist > epsilon)
                {
                    keepFlags[maxIndex] = true;
                    stack.Push((start, maxIndex));
                    stack.Push((maxIndex, end));
                }
            }

            // Build result from kept points
            var result = new List<Point>(pts.Count);
            for (int i = 0; i < pts.Count; i++)
            {
                if (keepFlags[i])
                    result.Add(pts[i]);
            }

            return result;
        }

        // RDP simplification that preserves corner points
        private static List<Point> RDPSimplifyPreservingCorners(List<Point> pts, double epsilon, HashSet<int> corners)
        {
            if (pts == null || pts.Count < 3)
                return pts;

            // Mark corner indices as kept
            var keepFlags = new bool[pts.Count];
            keepFlags[0] = true;
            keepFlags[pts.Count - 1] = true;
            foreach (int c in corners)
            {
                keepFlags[c] = true;
            }

            // Use iterative RDP on non-corner segments
            var stack = new Stack<(int start, int end)>();
            stack.Push((0, pts.Count - 1));

            while (stack.Count > 0)
            {
                var (start, end) = stack.Pop();
                if (end - start < 2)
                    continue;

                // Check if this segment contains any corners
                bool hasCorner = false;
                for (int i = start + 1; i < end; i++)
                {
                    if (keepFlags[i])
                    {
                        hasCorner = true;
                        break;
                    }
                }

                if (hasCorner)
                {
                    // Find the first and last corner in segment
                    int firstCorner = -1, lastCorner = -1;
                    for (int i = start; i <= end; i++)
                    {
                        if (keepFlags[i])
                        {
                            if (firstCorner == -1) firstCorner = i;
                            lastCorner = i;
                        }
                    }

                    // Process sub-segment before first corner
                    if (firstCorner > start)
                        stack.Push((start, firstCorner));

                    // Process sub-segment after last corner
                    if (lastCorner < end)
                        stack.Push((lastCorner, end));

                    // Process sub-segments between consecutive corners
                    int prev = firstCorner;
                    for (int i = firstCorner + 1; i <= lastCorner; i++)
                    {
                        if (keepFlags[i] && i > prev)
                        {
                            if (i - prev > 1)
                                stack.Push((prev, i));
                            prev = i;
                        }
                    }
                }
                else
                {
                    // No corners in this segment, use standard RDP
                    double maxDist = 0;
                    int maxIndex = start;

                    var first = pts[start];
                    var last = pts[end];

                    for (int i = start + 1; i < end; i++)
                    {
                        double dist = PerpendicularDistance(pts[i], first, last);
                        if (dist > maxDist)
                        {
                            maxDist = dist;
                            maxIndex = i;
                        }
                    }

                    if (maxDist > epsilon)
                    {
                        keepFlags[maxIndex] = true;
                        stack.Push((start, maxIndex));
                        stack.Push((maxIndex, end));
                    }
                }
            }

            // Build result from kept points
            var result = new List<Point>(pts.Count);
            for (int i = 0; i < pts.Count; i++)
            {
                if (keepFlags[i])
                    result.Add(pts[i]);
            }

            return result;
        }

        // Snap all segment angles to 30-degree multiples
        private static List<Point> SnapAnglesTo30(List<Point> pts)
        {
            if (pts == null || pts.Count < 2)
                return pts;

            // Pre-process: merge very short segments (< 100 HIMETRIC ~2.5mm)
            var merged = new List<Point>(pts.Count);
            merged.Add(pts[0]);

            for (int i = 1; i < pts.Count - 1; i++)
            {
                double dx = pts[i].X - merged[merged.Count - 1].X;
                double dy = pts[i].Y - merged[merged.Count - 1].Y;
                double dist = Math.Sqrt(dx * dx + dy * dy);

                // Only add point if it's far enough from the last merged point
                if (dist >= 100 || i == pts.Count - 2)
                    merged.Add(pts[i]);
            }
            merged.Add(pts[pts.Count - 1]);

            // Now snap each segment
            var result = new List<Point>();
            result.Add(merged[0]);

            for (int i = 0; i < merged.Count - 1; i++)
            {
                double dx = merged[i + 1].X - merged[i].X;
                double dy = merged[i + 1].Y - merged[i].Y;
                double length = Math.Sqrt(dx * dx + dy * dy);

                if (length < 1)
                    continue;

                // Calculate angle
                double angleRad = Math.Atan2(-dy, dx);
                double angleDeg = angleRad * 180.0 / Math.PI;
                if (angleDeg < 0) angleDeg += 360;

                // Snap to nearest 30 degrees
                double snappedAngle = Math.Round(angleDeg / 30) * 30;
                if (snappedAngle >= 360) snappedAngle -= 360;
                if (snappedAngle < 0) snappedAngle += 360;

                // Calculate new endpoint
                double snappedRad = snappedAngle * Math.PI / 180.0;
                int newX = merged[i].X + (int)Math.Round(length * Math.Cos(snappedRad));
                int newY = merged[i].Y - (int)Math.Round(length * Math.Sin(snappedRad));

                result.Add(new Point(newX, newY));
            }

            return result;
        }

        // Snap a single segment's angle to nearest 30 degrees, append to result
        private static void SnapSegmentAngle(List<Point> pts, List<Point> result, int start, int end)
        {
            if (end <= start) return;

            double dx = pts[end].X - pts[start].X;
            double dy = pts[end].Y - pts[start].Y;
            double length = Math.Sqrt(dx * dx + dy * dy);

            if (length < 1)
                return;

            // Calculate angle
            double angleRad = Math.Atan2(-dy, dx); // Negative dy because Y increases downward
            double angleDeg = angleRad * 180.0 / Math.PI;
            if (angleDeg < 0) angleDeg += 360;

            // Snap to nearest 30 degrees
            double snappedAngle = Math.Round(angleDeg / 30) * 30;
            if (snappedAngle >= 360) snappedAngle -= 360;
            if (snappedAngle < 0) snappedAngle += 360;

            // Calculate new endpoint
            double snappedRad = snappedAngle * Math.PI / 180.0;
            int newX = pts[start].X + (int)Math.Round(length * Math.Cos(snappedRad));
            int newY = pts[start].Y - (int)Math.Round(length * Math.Sin(snappedRad));

            result.Add(new Point(newX, newY));
        }

        private static double CalculateAngle(Point p1, Point p2, Point p3)
        {
            double v1x = p1.X - p2.X;
            double v1y = p1.Y - p2.Y;

            double v2x = p3.X - p2.X;
            double v2y = p3.Y - p2.Y;

            double dot = v1x * v2x + v1y * v2y;
            double len1 = Math.Sqrt(v1x * v1x + v1y * v1y);
            double len2 = Math.Sqrt(v2x * v2x + v2y * v2y);

            if (len1 < 1e-10 || len2 < 1e-10)
                return 180;

            double cosAngle = Math.Max(-1, Math.Min(1, dot / (len1 * len2)));
            double angleRad = Math.Acos(cosAngle);

            return angleRad * 180.0 / Math.PI;
        }

        private static double PerpendicularDistance(Point pt, Point lineStart, Point lineEnd)
        {
            double dx = lineEnd.X - lineStart.X;
            double dy = lineEnd.Y - lineStart.Y;

            if (dx == 0 && dy == 0)
            {
                double px = pt.X - lineStart.X;
                double py = pt.Y - lineStart.Y;
                return Math.Sqrt(px * px + py * py);
            }

            double len = Math.Sqrt(dx * dx + dy * dy);
            dx /= len;
            dy /= len;

            double vx = pt.X - lineStart.X;
            double vy = pt.Y - lineStart.Y;

            double cross = dx * vy - dy * vx;
            return Math.Abs(cross);
        }

        private static void CreateSmoothedStroke(Ink ink, StrokeGeometry sg, bool curveSmoothing)
        {
            var pts = sg.Points;
            if (pts.Count < 2)
                return;

            try
            {
                var s = ink.CreateStroke(pts.ToArray());
                s.DrawingAttributes = sg.Attr;
                // FitToCurve only makes sense for curve smoothing
                s.DrawingAttributes.FitToCurve = curveSmoothing;
            }
            catch { }
        }

        private class StrokeGeometry
        {
            public List<Point> Points { get; set; }
            public DrawingAttributes Attr { get; set; }
        }

        private static StrokeGeometry ExtractStrokeGeometry(Stroke stroke)
        {
            var sg = new StrokeGeometry();
            var rawPts = GetStrokePoints(stroke);
            // Shapes: few points, keep as-is. Normal strokes: resample to 500 points
            int numPoints = rawPts.Count <= 20 ? rawPts.Count : 500;
            sg.Points = numPoints == rawPts.Count
                ? rawPts
                : ResampleStroke(rawPts.ToArray(), numPoints);
            sg.Attr = stroke.DrawingAttributes.Clone();
            return sg;
        }

        private static List<Point> GetStrokePoints(Stroke stroke)
        {
            try
            {
                var points = stroke.GetPoints();
                if (points != null && points.Length >= 2)
                    return new List<Point>(points);
            }
            catch { }

            return new List<Point>(stroke.GetFlattenedBezierPoints());
        }

        private static List<Point> InterpolatePoints(List<Point> pts, int targetCount)
        {
            // Linearly interpolate between points to create more points
            // This preserves straight line segments (important for shapes)
            if (pts.Count < 2 || targetCount <= pts.Count)
                return pts;

            var result = new List<Point>(targetCount);
            result.Add(pts[0]);

            // Calculate segment lengths
            double[] segLen = new double[pts.Count - 1];
            double totalLen = 0;
            for (int i = 0; i < pts.Count - 1; i++)
            {
                double dx = pts[i + 1].X - pts[i].X;
                double dy = pts[i + 1].Y - pts[i].Y;
                segLen[i] = Math.Sqrt(dx * dx + dy * dy);
                totalLen += segLen[i];
            }

            if (totalLen <= 0)
                return pts;

            // Interpolate points at equal distances
            for (int i = 1; i < targetCount - 1; i++)
            {
                double targetDist = (totalLen * i) / (targetCount - 1);
                double acc = 0;
                for (int j = 0; j < pts.Count - 1; j++)
                {
                    if (acc + segLen[j] >= targetDist)
                    {
                        double t = segLen[j] > 0 ? (targetDist - acc) / segLen[j] : 0;
                        int x = (int)Math.Round(pts[j].X + t * (pts[j + 1].X - pts[j].X));
                        int y = (int)Math.Round(pts[j].Y + t * (pts[j + 1].Y - pts[j].Y));
                        result.Add(new Point(x, y));
                        break;
                    }
                    acc += segLen[j];
                }
            }

            result.Add(pts[pts.Count - 1]);
            return result;
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
            int dashGapSize)
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
                }
                catch { }
                return;
            }

            // If stroke is too short for the dash gap, skip dash conversion
            if (totalLen < dashGapSize * 1.5)
            {
                try
                {
                    var s = ink.CreateStroke(pts.ToArray());
                    s.DrawingAttributes = sg.Attr;
                }
                catch { }
                return;
            }

            // Shapes (rectangles, ellipses) have few points, interpolate for dash conversion
            if (pts.Count <= 20)
            {
                pts = InterpolatePoints(pts, 500);
                // Recalculate cumulative lengths with new points
                cumLen = new double[pts.Count];
                cumLen[0] = 0;
                totalLen = 0;
                for (int i = 1; i < pts.Count; i++)
                {
                    double dx = pts[i].X - pts[i - 1].X;
                    double dy = pts[i].Y - pts[i - 1].Y;
                    totalLen += Math.Sqrt(dx * dx + dy * dy);
                    cumLen[i] = totalLen;
                }
            }

            double dashLen = dashGapSize;
            double gapLen = dashGapSize;

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
