# Canonical PDF Layout Contract

The PDF renderer owns page breaking and resolved geometry. A captured trace is an internal semantic record
of the same drawing pass, not a reverse-engineered interpretation of the finished PDF.

## Required trace fields

- Trace version and report-definition identity.
- Page width, height, margins, body top/bottom, header, and footer regions.
- Page-local ordered items with source item identity, type, z-order, and point coordinates.
- Text lines with resolved runs, font file/family/variant/size, line height, baseline, padding, and alignment.
- Tablix fragments with grid widths, rows, cells, spans, repeated-header state, logical continuations,
  resolved shared edges, backgrounds, and final fragment closure.
- Images and charts with source type and rendered point bounds.

## Precision

- Preserve point values until the target renderer converts units.
- Coalesce coincident edges at the renderer's established 0.25-point precision.
- Word conversion uses twentieths of a point and must distribute rounding so summed columns equal the
  traced table width.

## Certification

- Compare every page, not only a sample.
- Page count and dimensions are exact gates.
- Geometry tolerance is 0.5 point.
- Raster comparison is 144 DPI with a per-channel threshold of 16 and a per-page differing-pixel ratio
  no greater than 0.005.
- A trace recorder change must produce no visible PDF change from identical inputs.
