# Alignment Feature - Key Learnings

## Overview
Implemented top and bottom alignment for ink strokes in OneNote, with intelligent clustering to keep related strokes together.

## Key Findings

### 1. OneNote Coordinate System
- **Position Y**: Page-level Y coordinate of InkDrawing element (small values, e.g., 150-600 HIMETRIC)
- **Size Height**: Height of the InkDrawing bounding box from the XML
- **ISF Internal Coordinates**: Internal stroke coordinates within the ISF data (large values, e.g., 5000-20000)

These are **different coordinate systems** and cannot be mixed!

### 2. Top Alignment
- Use **Position Y** directly as the top edge
- Align all strokes to the minimum Position Y

### 3. Bottom Alignment
- Bottom edge = **Position Y + Size Height**
- Use Size Height (not ISF internal bounding box height)
- ISF internal coordinates contain too much padding/margin and don't scale correctly with Position Y

### 4. Clustering Algorithm
- Used **Hierarchical Clustering (single-linkage)** instead of K-Means/DBSCAN/OPTICS
- Each InkDrawing is treated as an **atomic unit** - all strokes within stay together
- Distance threshold: 2500 HIMETRIC (~64mm) - clusters strokes that form logical shapes
- Distance calculation: Euclidean distance between bounding box centers

### 5. Implementation Details

```
ClusterInkDrawings:
- Groups strokes by InkDrawing first
- Calculates collective bounding box for each InkDrawing
- Uses hierarchical clustering to merge nearby InkDrawings
- Returns clusters with GroupY (top) and GroupMaxY (bottom for alignment)
```

### 6. Bottom Alignment Formula
```
referenceY = max(PositionY + SizeHeight)  // lowest bottom
yOffset = referenceY - (PositionY + SizeHeight)  // per stroke
newPositionY = PositionY + yOffset
```

## Files Modified
- `OneInk/AddIn.cs` - ExecuteAlign logic with clustering
- `OneInk/InkDashedConverter.cs` - ClusterInkDrawings method

## Common Mistakes

### Mistake 1: Mixing ISF and Page Coordinates
**Wrong**: `GroupY = bbox.Y + inkDrawingY` where bbox.Y is ISF internal (large) and inkDrawingY is page-level (small)

**Correct**: Use Position Y (page-level) directly for top alignment, and Position Y + Size Height for bottom alignment

### Mistake 2: Using ISF Bounding Box for Height
**Wrong**: Using ISF internal `maxY - minY` for bottom calculation
**Correct**: Use Size Height from XML for bottom alignment

### Mistake 3: Overlapping Clusters
**Problem**: When bounding boxes overlap, single-linkage distance returns 0, causing immediate merge
**Solution**: Always use center-point Euclidean distance, never return 0 for overlap

## Debugging Tips
- Use `Log()` to output Position Y, Size Height, GroupY, GroupMaxY for each stroke
- Check if offset values are reasonable (typically < 1000 HIMETRIC for small adjustments)
- If strokes fly off screen, likely mixing coordinate systems
