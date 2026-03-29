# Alignment Feature - Key Learnings

## Overview
Implemented top and bottom alignment for ink strokes in OneNote, with intelligent clustering to keep related strokes together.

## Key Findings

### 1. OneNote Coordinate System
- **Position Y**: Page-level Y coordinate of InkDrawing element (small values, e.g., 150-600 HIMETRIC)
- **Size Height**: Height of the InkDrawing bounding box from the XML
- **ISF Internal Coordinates**: Internal stroke coordinates within the ISF data (large values, e.g., 5000-20000)
- **X coordinate**: ISF internal (large values), different scale from page-level Y

These are **different coordinate systems** that cannot be mixed!

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
- **Fixed threshold: 30 HIMETRIC** (DPI-independent, physical units)
- **Distance calculation**: Euclidean distance with X normalized to page-level scale:
  ```
  distance = sqrt((dx/100)^2 + dy^2)
  ```
  - X is ISF internal (large), divided by ~100 to normalize to page-level scale
  - Y is page-level, used directly

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

### 7. Single InkDrawing with Multiple Strokes
When OneNote merges multiple strokes into one InkDrawing (e.g., drawn without lifting pen), each stroke's Y is tracked via `StrokeYs` list for proper clustering.

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

### Mistake 3: Double Counting Y in MaxY Calculation
**Problem**: In `CalculateMergedBounds`, using `bounds[idx].Y + bounds[idx].MaxY` where MaxY already equals `inkDrawingY + inkDrawingHeight`

**Correct**: Use `bounds[idx].MaxY` directly since it already contains the page-level bottom Y

### Mistake 4: Using Y-Only Distance
**Problem**: Only considering Y distance, ignoring X. Causes strokes far apart horizontally but close vertically to merge incorrectly.

**Correct**: Use Euclidean distance with X normalized: `sqrt((dx/100)^2 + dy^2)`

### Mistake 5: Dynamic Threshold Based on Gap
**Problem**: Using `maxGap * 0.6` or similar ratios fails across different DPI settings because the actual gaps vary.

**Correct**: Use fixed HIMETRIC threshold (30) since HIMETRIC is a physical unit, DPI-independent.

## Debugging Tips
- Use `Log()` to output Position Y, Size Height, GroupY, GroupMaxY for each stroke
- Check if offset values are reasonable (typically < 1000 HIMETRIC for small adjustments)
- If strokes fly off screen, likely mixing coordinate systems
- If all strokes merge into one cluster despite being far apart, check distance calculation
- Compare logs between computers to identify coordinate system differences
