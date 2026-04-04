/*
 *  InkAlignmentCluster - Point-based connectivity clustering for ink alignment.
 *  Uses single-linkage Union-Find to group strokes that form logical shapes.
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Microsoft.Ink;

namespace OneInk
{
    [System.Runtime.InteropServices.ComVisible(false)]
    public static class InkAlignmentCluster
    {
        [System.Runtime.InteropServices.ComVisible(false)]
        public class StrokeClusterPoint
        {
            public double X { get; set; }
            public double Y { get; set; }
            public string ObjectId { get; set; }
            public Stroke Stroke { get; set; }
        }

        [System.Runtime.InteropServices.ComVisible(false)]
        public class StrokeCluster
        {
            public List<StrokeClusterPoint> Points { get; set; } = new List<StrokeClusterPoint>();
            public double GroupX { get; set; }
            public double GroupY { get; set; }
            public double GroupWidth { get; set; }
            public double GroupHeight { get; set; }
        }

        [System.Runtime.InteropServices.ComVisible(false)]
        public class SampledStrokeInfo
        {
            public string ObjectId { get; set; }
            public List<Point> Points { get; set; } = new List<Point>(); // ISF coordinates
            public double PageX { get; set; }
            public double PageY { get; set; }
        }

        internal static (int idx1, Point pt1, int idx2, Point pt2, double minDist) FindClosestPointPair(
            SampledStrokeInfo stroke1, SampledStrokeInfo stroke2)
        {
            int bestI = -1, bestJ = -1;
            double minDist = double.MaxValue;

            for (int i = 0; i < stroke1.Points.Count; i++)
            {
                double px1 = (double)stroke1.Points[i].X; // ISF X directly
                double py1 = (double)stroke1.Points[i].Y; // ISF Y directly

                for (int j = 0; j < stroke2.Points.Count; j++)
                {
                    double px2 = (double)stroke2.Points[j].X; // ISF X directly
                    double py2 = (double)stroke2.Points[j].Y; // ISF Y directly

                    double dx = px1 - px2;
                    double dy = py1 - py2;
                    double dist = Math.Sqrt(dx * dx + dy * dy);
                    if (dist < minDist)
                    {
                        minDist = dist;
                        bestI = i;
                        bestJ = j;
                    }
                }
            }

            return (bestI, stroke1.Points[bestI], bestJ, stroke2.Points[bestJ], minDist);
        }

        internal static List<SampledStrokeInfo> GetSampledStrokesForDebug(
            string[] isfBase64Array,
            string[] objectIds,
            double[] inkDrawingXPositions,
            double[] inkDrawingYPositions,
            double pointIntervalHimetric = 1000.0)
        {
            var result = new List<SampledStrokeInfo>();

            for (int i = 0; i < isfBase64Array.Length; i++)
            {
                string isfBase64 = isfBase64Array[i];
                string objectId = objectIds[i];
                double pageX = (inkDrawingXPositions != null && i < inkDrawingXPositions.Length) ? inkDrawingXPositions[i] : 0;
                double pageY = (inkDrawingYPositions != null && i < inkDrawingYPositions.Length) ? inkDrawingYPositions[i] : 0;

                if (string.IsNullOrEmpty(isfBase64))
                    continue;

                byte[] raw;
                try
                {
                    raw = Convert.FromBase64String(isfBase64);
                }
                catch
                {
                    continue;
                }

                Ink ink = null;
                try
                {
                    ink = InkDashedConverter.LoadIsf(raw, out _);
                    if (ink == null || ink.Strokes.Count == 0)
                        continue;

                    foreach (Stroke s in ink.Strokes)
                    {
                        var rawPts = InkDashedConverter.GetStrokePoints(s);
                        if (rawPts.Count < 2)
                            continue;

                        // Compute total arc length
                        double totalLen = 0;
                        for (int k = 1; k < rawPts.Count; k++)
                        {
                            double dx = (double)rawPts[k].X - (double)rawPts[k - 1].X;
                            double dy = (double)rawPts[k].Y - (double)rawPts[k - 1].Y;
                            totalLen += Math.Sqrt(dx * dx + dy * dy);
                        }

                        int numPoints = Math.Max(2, (int)Math.Ceiling(totalLen / pointIntervalHimetric) + 1);
                        var sampledPts = InkDashedConverter.ResampleStroke(rawPts.ToArray(), numPoints);

                        result.Add(new SampledStrokeInfo
                        {
                            ObjectId = objectId,
                            Points = sampledPts,
                            PageX = pageX,
                            PageY = pageY
                        });
                    }
                }
                finally
                {
                    if (ink != null) ink.Dispose();
                }
            }

            return result;
        }

        private struct StrokeInfo
        {
            public Stroke Stroke;
            public string ObjectId;
            public List<Point> SampledPoints; // ISF coordinates
            public double PageX; // page-level X of this InkDrawing
            public double PageY; // page-level Y of this InkDrawing
            public double PageW; // page-level Width of this InkDrawing
            public double PageH; // page-level Height of this InkDrawing
            public double IsfMinX; // ISF bounding box min X (for normalization)
            public double IsfMinY; // ISF bounding box min Y (for normalization)
            public double IsfMaxX; // ISF bounding box max X (for normalization)
            public double IsfMaxY; // ISF bounding box max Y (for normalization)
        }

        private class UnionFind
        {
            private int[] parent;
            private byte[] rank;

            public UnionFind(int n)
            {
                parent = new int[n];
                rank = new byte[n];
                for (int i = 0; i < n; i++)
                    parent[i] = i;
            }

            public int Find(int x)
            {
                if (parent[x] != x)
                    parent[x] = Find(parent[x]);
                return parent[x];
            }

            public void Union(int x, int y)
            {
                int px = Find(x);
                int py = Find(y);
                if (px == py) return;
                if (rank[px] < rank[py])
                    parent[px] = py;
                else if (rank[px] > rank[py])
                    parent[py] = px;
                else
                {
                    parent[py] = px;
                    rank[px]++;
                }
            }
        }

        /// <summary>
        /// Clusters strokes from multiple ISF ink drawings using single-linkage
        /// point-based connectivity. Strokes whose resampled points are within
        /// threshold distance of each other belong to the same cluster (shape).
        ///
        /// Algorithm: each stroke is resampled at fixed arc-length intervals.
        /// Any two strokes with a point pair within threshold distance are connected.
        /// Connectivity is transitive — if A connects to B and B connects to C,
        /// then A, B, C are all in the same cluster (Union-Find).
        /// </summary>
        /// <param name="isfBase64Array">Array of ISF data as base64 strings.</param>
        /// <param name="objectIds">ObjectId strings for each InkDrawing.</param>
        /// <param name="pointIntervalHimetric">
        /// Fixed arc-length interval for point sampling in HIMETRIC units.
        /// Stroke length / interval = number of sample points (min 2).
        /// Default 1000 (~25mm interval).
        /// </param>
        /// <param name="thresholdHimetric">
        /// Point-to-point distance threshold for connectivity in page-level units.
        /// Default 50 (suitable for both small rectangles and hand-drawn strokes).
        /// </param>
        /// <returns>List of StrokeCluster, one per connected component (shape).</returns>
        internal static List<StrokeCluster> ClusterStrokesByConnectivity(
            string[] isfBase64Array,
            string[] objectIds,
            double[] inkDrawingXPositions,
            double[] inkDrawingYPositions,
            double[] inkDrawingWidths,
            double[] inkDrawingHeights,
            double pointIntervalHimetric = 1000.0,
            double thresholdHimetric = 50.0)
        {
            if (isfBase64Array.Length == 0)
                return new List<StrokeCluster>();

            // Phase 1: Extract and resample all strokes
            var strokeInfos = new List<StrokeInfo>();

            for (int i = 0; i < isfBase64Array.Length; i++)
            {
                string isfBase64 = isfBase64Array[i];
                string objectId = objectIds[i];
                double pageX = (inkDrawingXPositions != null && i < inkDrawingXPositions.Length) ? inkDrawingXPositions[i] : 0;
                double pageY = (inkDrawingYPositions != null && i < inkDrawingYPositions.Length) ? inkDrawingYPositions[i] : 0;
                double pageW = (inkDrawingWidths != null && i < inkDrawingWidths.Length) ? inkDrawingWidths[i] : 0;
                double pageH = (inkDrawingHeights != null && i < inkDrawingHeights.Length) ? inkDrawingHeights[i] : 0;

                if (string.IsNullOrEmpty(isfBase64))
                    continue;

                byte[] raw;
                try
                {
                    raw = Convert.FromBase64String(isfBase64);
                }
                catch
                {
                    continue;
                }

                Ink ink = null;
                try
                {
                    ink = InkDashedConverter.LoadIsf(raw, out _);
                    if (ink == null || ink.Strokes.Count == 0)
                        continue;

                    foreach (Stroke s in ink.Strokes)
                    {
                        var rawPts = InkDashedConverter.GetStrokePoints(s);
                        if (rawPts.Count < 2)
                            continue;

                        // Compute total arc length
                        double totalLen = 0;
                        for (int k = 1; k < rawPts.Count; k++)
                        {
                            double dx = (double)rawPts[k].X - (double)rawPts[k - 1].X;
                            double dy = (double)rawPts[k].Y - (double)rawPts[k - 1].Y;
                            totalLen += Math.Sqrt(dx * dx + dy * dy);
                        }

                        // Calculate number of points based on fixed interval
                        // Minimum 2 points (start + end)
                        int numPoints = Math.Max(2, (int)Math.Ceiling(totalLen / pointIntervalHimetric) + 1);

                        var sampledPts = InkDashedConverter.ResampleStroke(rawPts.ToArray(), numPoints);

                        // Compute ISF bounding box for normalization
                        double minX = double.MaxValue, minY = double.MaxValue;
                        double maxX = double.MinValue, maxY = double.MinValue;
                        foreach (var pt in sampledPts)
                        {
                            if (pt.X < minX) minX = pt.X;
                            if (pt.Y < minY) minY = pt.Y;
                            if (pt.X > maxX) maxX = pt.X;
                            if (pt.Y > maxY) maxY = pt.Y;
                        }

                        strokeInfos.Add(new StrokeInfo
                        {
                            Stroke = s,
                            ObjectId = objectId,
                            SampledPoints = sampledPts,
                            PageX = pageX,
                            PageY = pageY,
                            PageW = pageW,
                            PageH = pageH,
                            IsfMinX = minX,
                            IsfMinY = minY,
                            IsfMaxX = maxX,
                            IsfMaxY = maxY
                        });
                    }
                }
                finally
                {
                    if (ink != null) ink.Dispose();
                }
            }

            if (strokeInfos.Count == 0)
                return new List<StrokeCluster>();

            var logPath = Path.Combine(Path.GetTempPath(), "OneInk.log");
            File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] ClusterStrokesByConnectivity: pointInterval={pointIntervalHimetric}, threshold={thresholdHimetric}\r\n");

            // Debug: log all strokes with ALL sampled points
            for (int i = 0; i < strokeInfos.Count; i++)
            {
                var info = strokeInfos[i];
                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}]   DEBUG stroke[{i}] ObjId={info.ObjectId} pts={info.SampledPoints.Count} PageX={info.PageX:F2} PageY={info.PageY:F2}\r\n");
                for (int p = 0; p < info.SampledPoints.Count; p++)
                {
                    var pt = info.SampledPoints[p];
                    File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}]     stroke[{i}] point[{p}]: ISF=({pt.X},{pt.Y})\r\n");
                }
            }

            // Phase 2: Precompute ISF-level bounding boxes for quick rejection
            // Use ISF coordinates directly (no page offset conversion)
            var strokeBounds = new (double MinX, double MinY, double MaxX, double MaxY)[strokeInfos.Count];
            for (int i = 0; i < strokeInfos.Count; i++)
            {
                var info = strokeInfos[i];
                double minX = double.MaxValue, minY = double.MaxValue;
                double maxX = double.MinValue, maxY = double.MinValue;
                foreach (var pt in info.SampledPoints)
                {
                    double px = (double)pt.X; // ISF X directly
                    double py = (double)pt.Y; // ISF Y directly
                    if (px < minX) minX = px;
                    if (py < minY) minY = py;
                    if (px > maxX) maxX = px;
                    if (py > maxY) maxY = py;
                }
                strokeBounds[i] = (minX, minY, maxX, maxY);
            }

            // Phase 3: Union-Find single-linkage connectivity
            var uf = new UnionFind(strokeInfos.Count);

            for (int i = 0; i < strokeInfos.Count; i++)
            {
                for (int j = i + 1; j < strokeInfos.Count; j++)
                {
                    // Compare strokes using page-level coordinates.
                    // Each InkDrawing maps its ISF bounding box to its Page Position+Size independently.
                    // scale = PageSize / ISFSize
                    double iPageX = strokeInfos[i].PageX;
                    double iPageY = strokeInfos[i].PageY;
                    double iPageW = strokeInfos[i].PageW;
                    double iPageH = strokeInfos[i].PageH;
                    double iIsfMinX = strokeInfos[i].IsfMinX;
                    double iIsfMinY = strokeInfos[i].IsfMinY;
                    double iIsfMaxX = strokeInfos[i].IsfMaxX;
                    double iIsfMaxY = strokeInfos[i].IsfMaxY;
                    double jPageX = strokeInfos[j].PageX;
                    double jPageY = strokeInfos[j].PageY;
                    double jPageW = strokeInfos[j].PageW;
                    double jPageH = strokeInfos[j].PageH;
                    double jIsfMinX = strokeInfos[j].IsfMinX;
                    double jIsfMinY = strokeInfos[j].IsfMinY;
                    double jIsfMaxX = strokeInfos[j].IsfMaxX;
                    double jIsfMaxY = strokeInfos[j].IsfMaxY;

                    double iScaleX = (iIsfMaxX > iIsfMinX) ? iPageW / (iIsfMaxX - iIsfMinX) : 0;
                    double iScaleY = (iIsfMaxY > iIsfMinY) ? iPageH / (iIsfMaxY - iIsfMinY) : 0;
                    double jScaleX = (jIsfMaxX > jIsfMinX) ? jPageW / (jIsfMaxX - jIsfMinX) : 0;
                    double jScaleY = (jIsfMaxY > jIsfMinY) ? jPageH / (jIsfMaxY - jIsfMinY) : 0;

                    double minDist = double.MaxValue;
                    int minPi = -1, minPj = -1;
                    double minPix = 0, minPiy = 0, minPjx = 0, minPjy = 0;
                    for (int pi = 0; pi < strokeInfos[i].SampledPoints.Count; pi++)
                    {
                        var piPt = strokeInfos[i].SampledPoints[pi];
                        double pix = ((double)piPt.X - iIsfMinX) * iScaleX + iPageX;
                        double piy = ((double)piPt.Y - iIsfMinY) * iScaleY + iPageY;
                        for (int pj = 0; pj < strokeInfos[j].SampledPoints.Count; pj++)
                        {
                            var pjPt = strokeInfos[j].SampledPoints[pj];
                            double pjx = ((double)pjPt.X - jIsfMinX) * jScaleX + jPageX;
                            double pjy = ((double)pjPt.Y - jIsfMinY) * jScaleY + jPageY;
                            double dx = pix - pjx;
                            double dy = piy - pjy;
                            double dist = Math.Sqrt(dx * dx + dy * dy);
                            if (dist < minDist) { minDist = dist; minPi = pi; minPj = pj; minPix = pix; minPiy = piy; minPjx = pjx; minPjy = pjy; }
                        }
                    }
                    // After checking ALL point pairs, compare minimum to threshold
                    if (minDist <= thresholdHimetric)
                    {
                        uf.Union(i, j);
                        File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}]   UNION stroke[{i}] + stroke[{j}]: minDist={minDist:F0} <= {thresholdHimetric}, closest point[{minPi}]=({minPix:F0},{minPiy:F0}) vs point[{minPj}]=({minPjx:F0},{minPjy:F0})\r\n");
                    }
                    else
                    {
                        File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}]   SKIP stroke[{i}] vs stroke[{j}]: minDist={minDist:F0} > {thresholdHimetric}, closest point[{minPi}]=({minPix:F0},{minPiy:F0}) vs point[{minPj}]=({minPjx:F0},{minPjy:F0})\r\n");
                    }
                }
            }

            // Phase 4: Build StrokeCluster list from Union-Find groups
            var groups = new Dictionary<int, List<int>>();
            for (int idx = 0; idx < strokeInfos.Count; idx++)
            {
                int root = uf.Find(idx);
                if (!groups.ContainsKey(root))
                    groups[root] = new List<int>();
                groups[root].Add(idx);
            }

            var result = new List<StrokeCluster>();
            foreach (var group in groups.Values)
            {
                var cluster = new StrokeCluster();
                foreach (int idx in group)
                {
                    var info = strokeInfos[idx];
                    foreach (var pt in info.SampledPoints)
                    {
                        cluster.Points.Add(new StrokeClusterPoint
                        {
                            X = pt.X,
                            Y = pt.Y,
                            ObjectId = info.ObjectId,
                            Stroke = info.Stroke
                        });
                    }
                }
                CalculateClusterBounds(cluster);
                result.Add(cluster);
            }

            return result;
        }

        private static void CalculateClusterBounds(StrokeCluster cluster)
        {
            if (cluster.Points.Count == 0) return;

            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;

            foreach (var pt in cluster.Points)
            {
                if (pt.X < minX) minX = pt.X;
                if (pt.Y < minY) minY = pt.Y;
                if (pt.X > maxX) maxX = pt.X;
                if (pt.Y > maxY) maxY = pt.Y;
            }

            cluster.GroupX = minX;
            cluster.GroupY = minY;
            cluster.GroupWidth = maxX - minX;
            cluster.GroupHeight = maxY - minY;
        }
    }
}
