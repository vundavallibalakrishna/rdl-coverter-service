# Border Resolution Contract (SSRS)

Borders in RDL are per-item and per-edge. Reproduce SSRS's visible result: each declared edge is drawn once,
at its exact coordinate; shared edges coincide into a single line; a visible **double** line is an engine
defect.

## Declaration model

- `Style/Border` sets all four sides; `TopBorder`/`BottomBorder`/`LeftBorder`/`RightBorder` override
  individual sides. `Style=None` on a side means that side draws nothing even if `Border` set a color.
- Each side carries `Style` (None/Solid/Dotted/Dashed/Double/…), `Color`, and `Width`.
- Expression-backed border properties resolve through the style helpers at render time; never compare or
  measure a raw `=expression` string.

## Shared edges (the key rule)

- Tablix cells have **no spacing**: adjacent cells share an edge at the **same** coordinate. Cell A's right
  border and cell B's left border are the same line; cell A's bottom and cell B (below) top are the same
  line. When both declare a border there, the two strokes **coincide** and read as a single line.
- Therefore SSRS shows a single grid line at a shared edge. Two parallel lines with a gap is **never**
  correct for a shared tablix edge — it means the engine computed the two cells' edges at different
  coordinates (a geometry/growth bug) or drew a redundant second stroke.
- A group/total separator that looks like a "double" (e.g. a black line plus a light-grey line) is the same
  failure: two edges that should be one coordinate landed apart, usually because a row/container grew and
  only one of the two edges tracked the growth.

## Style precedence on a single shared edge

- When two coincident edges declare **different** styles/colors/widths, resolve to **one** drawn line, not
  two overlapping strokes. Precedence: a heavier/darker declared border wins over `None`/lighter (Solid over
  None; a specified color over an unspecified/light default; greater width over lesser). The result must
  match what SSRS shows — validate against the oracle when precedence is ambiguous.
- Do not rely on draw-order (last stroke wins) to hide a redundant edge. That is fragile: a later change to
  order, rounding, or growth re-exposes the hidden stroke as a visible double. Resolve to a single edge
  explicitly.

## Drawing mechanics to preserve

- De-duplicate collinear edges: each grid line is drawn once. Merge collinear intervals that touch within
  the renderer's established coincidence precision (a named constant, ~0.25–0.75 pt) — but never fuse two
  edges that are intentionally distinct rows apart.
- Solid edges use square cap / miter join so corners close; dashed/dotted keep butt caps and honor the
  dash pattern and width floor. A `Line` report item honors its `Style` (dashed/dotted render dashed/dotted,
  not solid).
- `Double` borders are two strands; extend each strand at joins so inner corners close.
- Repeated tablix headers and split rows always close their fragment with the correct bottom/edge borders;
  a merged (row-span) cell redraws its shared borders per fragment.

## What counts as a declared edge

- A cell's grid edge is declared by the cell, or by the **textbox** that fills it — the common RDL idiom of
  a `Border=None` tablix whose every cell textbox carries the rule.
- An item that draws its own box at its own coordinates inside the cell does **not** declare a cell edge: a
  `Line` (whose rule is expressed through `Style.Border`, RDL giving it no stroke property), a nested
  `Tablix` (which strokes its own outer border), a `Rectangle` frame. Reading any of those as grid intent
  makes a borderless form/layout tablix look like a bordered grid.
- A synthesized fragment closure **reuses** a border the report declares. Inventing one where the report
  declares none draws a decoration SSRS never draws.

## Renderer parity

Border resolution is a shared semantic. Fix it in the tablix/border layer and reflect it in PDF strokes,
DOCX_EDITABLE cell borders (twips), and XLSX cell borders together — Word OOXML and Excel express borders
per cell, so a coincident-edge resolution must map to the correct single per-cell border on each side, not
a doubled pair.

## Verification

Intercept the actual draw calls (edge coordinate + color + width) for the table region and assert: one
stroke per visible grid line; shared edges at identical coordinates; group/total separators single; no
change under row growth. Confirm with a raster that no two parallel lines appear where SSRS shows one.
