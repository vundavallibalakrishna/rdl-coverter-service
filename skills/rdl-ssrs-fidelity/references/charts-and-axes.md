# Charts and Axes Contract (SSRS)

Reproduce the SSRS chart engine's axis math and label layout. The plot area yields to labels — SSRS does
not clip axis labels or emit unformatted numbers.

## Value (numeric) axis

- Default scale is **0-based** and rounds to a readable maximum with a "nice" interval from the
  1 / 2 / 5 · 10ⁿ family, giving ~4–6 gridlines. Honor explicit axis `Minimum`/`Maximum`/`Interval` when
  set; honor `LogScale`.
- Tick values are multiples of the interval — format them so binary-float accumulation never leaks
  (`3 * 0.2` → render `0.6`, not `0.6000000000000001`). Use the axis `LabelFormat` when present; otherwise
  format to the interval's own decimal precision (drop trailing zeros: `1.0` → `1`).
- Stacked charts scale to the max stack total; percent-stacked scale to 100.

## Category axis and label auto-fit

- SSRS fits category labels without clipping, in this order: **shrink** font toward a floor, then **rotate**
  (commonly to an angle or vertical), then **stagger/wrap** to multiple lines. The plot rectangle is
  reduced to reserve whatever height/width the fitted labels need.
- Reproduce by reserving the labels' true rendered band (measure at the fitted font and slot width,
  accounting for wrapping) **before** sizing the plot. Never reserve a fixed gutter and let the outer clip
  rectangle crop a wrapped label — that is the canonical chart-label-crop defect.
- Respect `LabelsAutoFitDisabled` (keep the declared font, do not shrink) and any declared label rotation.

## Series, labels, legend

- Data-point labels (`UseValueAsLabel` or an expression) render with the value's format (same formatting
  contract as textboxes).
- Legend visibility/position/layout follow the RDL; wrap legend entries within the reserved legend box.
- `PointWidth` custom property controls bar/column thickness.

## No data

- When a series has no points, render the chart's `NoDataMessage` (SSRS default "No Data Available")
  centered in the plot — do not draw empty axes as if populated.

## Supported set and fail-closed

- Column, bar, line, area, pie/doughnut, scatter and the documented stacked variants are supported.
- `Aggregate()`, maps, gauges, KPIs, and uncatalogued chart types/features are fail-closed
  (`UNSUPPORTED_FEATURE`) — never approximate them.

## Renderer parity

Charts render once through the shared PDFKit chart pass for PDF and are rasterized from that same pass for
DOCX_EDITABLE (embedded picture) and DOCX_VISUAL. XLSX represents chart source as typed cells (no axis
drawing). An axis/label fix in the chart layer therefore fixes PDF and both DOCX modes together; note XLSX
non-applicability with its reason.

## Verification

Assert clean tick labels (no float noise) at a max that forces fractional intervals (e.g. max 1 → interval
0.2). Assert long category labels wrap and remain fully visible (measure the label band; confirm no crop in
the raster). Assert `NoDataMessage` for an empty series.
