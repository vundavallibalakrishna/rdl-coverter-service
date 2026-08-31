# Known Engine ↔ SSRS Deviations

Running log of places this engine is known to differ from SSRS, or where fidelity is unproven. Each entry
names the SSRS rule, the current behavior, and how to reconcile. Remove an entry only when it is fixed and
verified against a real SSRS oracle. Never let an entry justify a report-specific workaround.

## Open

- **Shape-chart radius is smaller than SSRS's.** With outside labels, SSRS sized a 335×229.5 pt pie chart's
  radius to 69.8 pt; this engine now produces 59.4 pt for the same chart. The callout reservation itself
  matches SSRS (`1.4·r + gap ≈ half of the plot's smaller dimension`, verified against two SSRS charts), so
  the difference is the *plot rectangle*: `drawChart` insets the chart by 8 pt horizontally and 6 pt
  vertically, adds another 6 pt above the plot, and reserves 8 pt between the plot and the legend, where
  SSRS reserves ≈0 and lets the legend sit flush. Fix: derive the chart content box and legend gutter from
  SSRS's own metrics, with a corpus diff over every chart type before landing (the insets are shared by
  bar/column/line/area/scatter). Ref: `charts-and-axes.md`.

- **`Format` = `d`/`g` use `Intl` short patterns, not the .NET ones.** `d` under en-US renders `5/15/26`
  where .NET/SSRS render `5/15/2026`; `formatNet` does not implement the single-letter *date* specifiers at
  all (`NUMERIC_HINT_RE` claims them first), so only the `formatValue` path handles them. Fix: implement the
  .NET standard date/time specifiers (`d D f F g G m M o R s t T u U y`) from the culture's own patterns.
  Ref: `value-formatting.md`.

- **No-format date order is not culture-reordered.** The no-format DateTime default is a fixed
  `dd/MM/yyyy HH:mm:ss` regardless of culture, so an en-US report shows `dd/MM` order for a *no-format*
  DateTime instead of `M/d`. Explicit formats and standard specifiers (`d`,`D`,`g`,`G`) already reorder by
  culture; only the bare default does not. Fix: derive the culture's general date/time pattern for the
  default. Ref: `value-formatting.md`.

- **Item-level `Language` override.** Only the report-level `Language` is honored; a per-textbox
  `Style/Language` override is not yet applied. Rare. Fix: resolve item Language over report Language in the
  style helpers. Ref: `value-formatting.md`.

- **Body coordinate model is partly flow-based.** SSRS positions body items absolutely by `Top`/`Left`;
  the top-level body band currently stacks items vertically (`gap = max(0, item.top − previousBottom)`),
  which is only correct for stacked layouts. Rectangle children and header/footer items already use true
  absolute positioning. Fix: absolute 2-D positioning for the body band; preserve output for purely stacked
  reports. Ref: `page-and-flow.md`.

- **Wide-matrix horizontal pagination.** SSRS splits a matrix wider than the printable body across
  horizontal page columns, repeating row headers. The engine draws the natural width and can overflow the
  right edge. Fix behind the advanced-materialization path. Ref: `page-and-flow.md`.

- **Body-scope `TotalPages`.** Correct in the header/footer pass; a body/tablix cell referencing
  `Globals!TotalPages` currently sees a provisional value. Fix: resolve final page count for body scope, or
  confirm no supported report needs it and document. Ref: `page-and-flow.md`.

- **Non-embedded images.** `Image.Source ∈ {External, Database}` currently renders nothing silently, which
  violates fail-closed. Fix: classify as `UNSUPPORTED_FEATURE` at analyze time (metadata-only policy still
  forbids fetching external resources). Ref: service security invariants.

- **Redundant coincident border strokes.** At shared tablix/group edges the engine can emit two coincident
  strokes (e.g. black + light-grey) and rely on draw-order to hide one. It renders as a single line today
  and is stable under row growth, but is fragile. Fix: resolve conflicting collinear coincident edges to a
  single precedence winner in the edge-flush layer. Ref: `border-resolution.md`.

## Verified fixed (kept for regression awareness)

- **A free-form tablix cell collapsed into one Excel cell value.** A cell whose CellContents Rectangle was
  flattened is a CANVAS: its children keep their own declared position and size. `pdf.js` places each one
  (`isFreeFormCell` / `drawCanvasRow`); `renderReportTablix` joined them with `cellText()` into a single
  merged cell, so a report whose whole body is one canvas cell put 2,468 characters — every heading and
  paragraph — into `A3:A4`, a range 0.14 characters wide. The prose was in the workbook and invisible in
  it. Fixed in `excel.js`: `isCanvasCell`/`isFreeFormCell` moved to `common.js` so both renderers agree
  what a canvas is; a free-form cell now writes no value, merge or region paint of its own and instead
  emits each child through `renderFreeformItem` at its own section coordinates, with SSRS vertical
  displacement (`nestedTopsFor` now resolves canvas peers too) and the row profile extended to reserve the
  canvas extent. Three supporting corrections were needed: `applyRegionBorder` no longer marks every cell
  of a region as styled with an all-undefined border object (which `copyPlannedTablix` then materialized
  as a merge per column); the row profile records the regions' MEASURED, displaced boundaries rather than
  their declared ones; and a child whose resolved rows still meet an occupied range steps below it, which
  is the same displacement rule applied where the point model and the row grid disagree by one interval.
  Verified through Excel COM on the client report: all ten numbered sections plus the prose land on their
  own rows at the RDL's declared widths (555 pt full-width, 140 pt for "1. INTRODUCTION"), alongside the
  13 pictures. Regression: `test/canvas-cell-chart.test.js`. Ref: `page-and-flow.md`.

- **A canvas-cell chart drew the whole report in Word, and nothing at all in Excel.** A chart on an SSRS
  List / canvas cell is scoped to its group instance — it sees that instance's rows, not the dataset.
  **Word**: the canonical trace recorded only the chart's geometry, so `pagedDocx` re-materialized it from
  the report-level datasets with an empty row scope; one chart carrying every category in the report
  replaced the per-instance ones (the client report showed a 6-slice pie where the PDF drew a single
  slice). Fixed by recording the resolved series on the chart's trace item (`chartData`, materialized in
  the PDF pass with the chart's own context) and consuming it in `pagedDocx`, so Word rasterizes exactly
  the chart the PDF drew — which is what the renderer-parity contract already required.
  **Excel**: only BODY-level charts were embedded, so a report that puts every chart inside a List cell
  produced no chart at all. Fixed in `excel.js`: a tablix cell's canvas media (`Chart`/`Image`) is
  collected while the grid is written, its left/right edges join the section column grid
  (`collectXBoundaries`) and its top/bottom the row grid, and it is anchored as a floating picture over
  that region in the CELL's scope. Verified through Excel COM on the client report: 12 chart shapes at
  459.8 x 156.8 pt against the RDL's declared 459.3 x 157.5. Only pictures are placed this way; a canvas
  cell's text is placed the same way (see the entry above).
  Regression: `test/canvas-cell-chart.test.js`.
  Ref: `charts-and-axes.md`.

- **"Continued from previous page" was never emitted for a child region.** The annotation
  (`pagination.continuationMarkers`) was drawn only by the top-level tablix's `startContinuationPage`, so a
  report whose tables are all CHILD regions — every table inside one body-level canvas — produced no marker
  at all even with the option on, while an ordinary report produced them normally. Fixed in `pdf.js`:
  `drawContinuationMarker` now takes the box and cursor to draw at (defaulting to the top-level tablix's own
  geometry and cursor) and returns its height, and `drawNestedTablixAcrossPages` labels each of its own
  continuation fragments over the child's box. Verified on the client report: 10 markers in the PDF, the
  same 10 in DOCX_EDITABLE through the layout trace, page count unchanged at 42; XLSX has no pagination and
  correctly has none. Regression: `test/nested-region-fragment-grid.test.js`, which also pins the option as
  opt-in. Ref: `page-and-flow.md`.

- **A child region's column grid re-flowed on its continuation pages.** Cell placement walks a data region's
  rows in order, carrying each row-span's occupancy forward. A page fragment of a child region is a SLICE of
  those rows, and both slicing loops recomputed placement from the slice alone — so every span that began in
  an earlier fragment was lost and the cells to the right of a spanned row header shifted left into the
  header's own columns. In the client report the per-audit findings, which belong in the last column, were
  drawn in the second column from the continuation page onwards, with the right-hand columns left empty and
  unruled — reported as "this tablix is still missing borders". Fixed in `pdf.js` by pinning the region's
  canonical row order before slicing and resolving placement once over it (`nestedPlacements`), so every
  fragment reads its own rows out of that map. Ref: `sizing-and-growth.md`.

- **A row header spanning into a continuation fragment vanished.** SSRS repeats it there, clipped to the
  rows the fragment shows; dropping it leaves the header columns empty and unruled for the rest of the
  region. Fixed in `pdf.js` (`carryOpenSpans`): each fragment re-attaches the still-open spans to its first
  row with the span clipped, and registers that row's placement so the carried cells keep the columns they
  hold in the whole region. PDF owns the fragmenting; DOCX_EDITABLE inherits both fixes through the layout
  trace (verified: the continuation Word row carries `1 | Test New | … | Findings`, fully bordered), and
  XLSX is unaffected because it has no pagination. Regression:
  `test/nested-region-fragment-grid.test.js`. Ref: `page-and-flow.md`.

- **Tablix borders landed on the wrong page (no borders here, a stray grid there).** Borders are collected
  as edges and stroked as merged runs when a page closes, but only the tablix's own `startContinuationPage`
  flushed them. Every page break taken by the OTHER pagination paths in `renderTablix` — the SSRS
  List/canvas reflow (`drawCanvasRow`) and the nested-region continuation (`drawNestedTablixAcrossPages`),
  both of which call `addPage` directly — therefore carried a page's worth of edges forward. A report whose
  whole body is a canvas cell (a 1×1 tablix holding a Rectangle with the report's items) rendered its tables
  with **no borders at all**, and the accumulated runs were painted, at the earlier pages' coordinates, over
  whatever content sat on the page where the flush finally landed. Fixed in `renderTablix` (`pdf.js`) by
  wrapping the injected page advance so a page can never close with edges pending. PDF-only: DOCX_EDITABLE
  and XLSX resolve borders per cell and always had them, so this also removed a PDF↔Word disagreement.
  Regression: `test/canvas-page-break-borders.test.js`. Ref: `border-resolution.md`.

- **Row-header band collapsed into one column.** Every `TablixHeader` is as wide as its declared `Size`, so
  a row-header cell spans as many leaf hierarchy columns as that size covers — the flat materializer already
  sized its headers this way. The advanced materializer only honored it for `static` units and for a
  single-descriptor full-width band, so a group's own header branch (shallower than the detail branch) got
  `colSpan: 1` plus a synthesized blank tail: a quarter band that SSRS draws across the whole table rendered
  as a narrow wrapped cell in the second column. Fixed in `materializeAdvancedRows` (`validation.js`) by
  deriving the spans from the declared sizes in every case, with the tail sized from what those spans leave
  uncovered; `headerSpans` no longer short-circuits a single descriptor to the whole grid, so a narrow
  one-column band stays one column. Shared layer — PDF, DOCX_EDITABLE and XLSX all inherit it.
  Regression: `test/tablix-row-header-band.test.js`. Ref: `sizing-and-growth.md`.

- **Running aggregates numbered the unsorted rows.** `RowNumber`/`RunningValue` accumulate in the order the
  data region processes rows, which the recursive walk records. Role alone did not separate the coarse unit
  from the fine ones: a group's own header band and the per-instance rows under it are both static leaves
  inside that group, so both are labelled `header`. The band sits at a shallower hierarchy depth and carries
  the WHOLE group scope in dataset arrival order, so it recorded first and reinstated exactly the arrival
  order a member `SortExpression` had just replaced — a `RunningValue(..., CountDistinct, "Group")` column
  numbered the sorted rows 2, 1, 3. Fixed by restricting the ordering pass to the deepest units of the
  ordering role. Shared layer. Regression: `test/tablix-row-header-band.test.js`.
  Ref: `page-and-flow.md`.

- **XLSX columns were ~5 px narrow each, so dates rendered as `#####`.** `excelWidthFromPoints` converted a
  point width with the widely-quoted `width = (pixels − 5) / maxDigitWidth`, subtracting Excel's per-cell
  inset. Excel does not add that inset back: it renders a stored width `w` at exactly `w * 7` device pixels
  (verified through Excel COM — stored 10 → 70 px, stored 20 → 140 px). Every grid column therefore came out
  5 px (3.75 pt) short, and an RDL column that the shared grid slices into N columns came out 5N px short.
  Text merely clipped so nobody noticed; a date or number is never clipped by Excel — it becomes `#####`, so
  a "Due Date" column declared at 53.17 pt arrived as 45 pt and its value disappeared. Fixed by converting
  `points → pixels / maxDigitWidth` with no inset; the inset stays where it belongs, in text measurement
  (`EXCEL_CELL_TEXT_INSET_PT`). Impact is whole-sheet: on the client report the detail sheet went from
  704 pt wide to 829 pt, against a printable report width of ~828 pt, and Excel now reports the Due Date
  merge at 53.25 pt against the RDL's 53.171 pt. PDF/DOCX are unaffected (they never used this conversion);
  the corpus PDFs are byte-identical. Regression: `test/excel-column-geometry.test.js`.
  Ref: `sizing-and-growth.md`.

- **Chart picture silently lost slices in Word/Excel (PDF looked fine).** A chart's strings are absolutely
  positioned, never flowed, but `fillText` called `doc.text(..., { width })` without a `height`. PDFKit
  treats a string that reaches the bottom of the page as overflowing body copy: its `LineWrapper` calls
  `continueOnNewPage()`, so that label **and everything drawn after it** move to a new page. `renderChartPng`
  draws each chart onto a one-page document exactly the size of the chart and rasterizes with
  `-singlefile`, so the spill was discarded — a pie whose bottom outside label touched the page edge lost
  its remaining slice and embedded as a half-circle in DOCX_EDITABLE and XLSX, while the same chart on a
  tall PDF report page was correct. (The same hazard could have injected a stray page into the report PDF
  itself.) Fixed in `fillText` (`chart.js`): an explicit `height` bounds the wrapper, which then clips at
  the chart edge — what SSRS does — instead of paginating. `renderChartPng` also asserts its document
  stayed one page and fails with `RENDER_FAILED` rather than embedding an incomplete chart. Reproduced from
  the client's own pre-fix DOCX (identical media SHA-1s) and still reachable under current geometry for
  short shape charts. Regression: `test/chart-image-single-page.test.js`. Ref: `charts-and-axes.md`.

- **Shape-chart empty points took a legend entry and a palette colour.** SSRS calls a data point whose value
  is `Nothing`/null an EMPTY POINT. On a shape chart (pie/doughnut) it draws no slice, takes **no legend
  entry**, and consumes **no palette colour**, so the points after it keep the colours they would have had
  without it. A zero is a real value and keeps both. The engine allocated `palette[categoryIndex]` and built
  the legend from every category, so one `=IIF(x > 0, Count(x), Nothing)` point shifted every later slice by
  one palette colour and added a phantom legend row. Fixed in `materializeChart` (`chartData.js`), the layer
  all four renderers share. Verified against a real SSRS 2019 export of the same report: legend and slice
  colours now match exactly. Ref: `charts-and-axes.md`.

- **Outside shape-chart labels had no callout elbow.** SSRS draws an outside pie/doughnut label as a
  callout: a radial stub off the slice edge, then a **horizontal elbow**, with the label starting past the
  elbow and vertically centred on it; and it shrinks the shape so the whole callout stays inside the plot.
  The engine drew a single short radial tick and placed the text beside it, with the side chosen by
  `Math.cos(middle) >= 0` — which flips on the ±1e-16 cosine of a vertical bisector, so two equal slices put
  their labels on opposite sides. Fixed in `drawPieChart` (`chart.js`): two-segment callout, a
  tolerance-based side test, and a radius budget that reserves the callout band. Measured from the SSRS
  raster: stub ≈ elbow ≈ 0.20·r, label gap ≈ 3.5 pt. Ref: `charts-and-axes.md`.

- **Shape-chart data labels wrapped inside a fixed-width box.** Inside labels were drawn into a hardcoded
  24 pt box and outside labels into a 40 pt box, so `1 (50.0%)` broke across two lines. The box is now the
  measured width of the text. Ref: `charts-and-axes.md`.

- **Unzoned DateTime values rendered the previous day.** An RDL DateTime is a wall-clock value: SSRS renders
  exactly what it was given and never converts time zones. Every formatter here reads a Date through its UTC
  accessors, but `new Date('2026-05-15T00:00:00')` — a date-time with no offset — is parsed by JavaScript as
  **local** time, so on any host east of UTC midnight became the previous day in UTC (`15/05/2026` rendered
  as `14/05/2026`) and the same inputs rendered differently on different servers, breaking determinism.
  Fixed with one shared parser, `parseDateValue` (`src/rdl/dateValue.js`): an ISO-like value with no `Z` or
  `±HH:MM` is built from UTC components; a value that names a real instant keeps the standard parse. Applied
  in `format.js` (`coerceDate`), `expression.js` (`toDate` and the no-format date path),
  `functions/shared.js` (`toDate`), `validation.js` (DateTime parameter validation and canonicalization),
  and `excel.js` (typed date cells). Regression: `test/datetime-wall-clock.test.js`, including a sweep over
  four host time zones so a UTC build machine cannot hide the defect. Ref: `value-formatting.md`.

- **Parent row holding a child data region was treated as atomic (large blank page tail).** SSRS breaks a
  page at the deepest boundary that can still fill it: a tablix row whose cell holds a child data region
  (nested tablix or bundled subreport) is *not* atomic — the break falls between the **child** region's own
  rows, so the current page is filled and the remainder continues on the next page. A whole-row move is
  correct only when KeepTogether is declared on the child region (`Subreport`/`Tablix`) or on the owning
  tablix member, and the row still fits a fresh page. The engine moved the row to a fresh page whenever it
  did not fit the remainder, and split the child only when the child grid exceeded an *entire* page — so a
  child region a little taller than what remained stranded up to a full page of white space (observed: 74%
  of a page blank, and one extra page in the document). Fixed in `renderTablix`/`drawRow` (`pdf.js`): the
  overflowing child region is now split at a child-row boundary in place, with the fresh-page move reserved
  for the KeepTogether/unsplittable cases; the same rule was applied to the List/canvas cell path
  (`drawCanvasRow`), which had the identical "move the whole region" behavior. `validation.js` now carries
  the invoking `Subreport`'s `KeepTogether` onto the materialized nested entry, since that — not the child
  report body's own tablix — is the property SSRS honors for subreport content. Shared-layer fix: PDF owns
  pagination, `DOCX_EDITABLE` and `DOCX_VISUAL` inherit it from the canonical trace/raster, and XLSX has no
  pagination. Regression: `test/nested-region-page-fill.test.js` (split + KeepTogether counterexamples for
  both nested tablix and bundled subreport, plus DOCX_EDITABLE and XLSX coverage). Ref: `page-and-flow.md`.

- **XLSX confined a merged cell's child data region to the row it starts in.** A cell that spans several
  tablix rows holds its child region in the whole BLOCK of rows it covers: the fixed layout draws the child
  from the block's top and lets it flow past the first row's bottom, so the child's rows and the spanned
  rows interleave. `renderReportTablix` (`excel.js`) gave every child edge to the profile of the row the
  cell starts in, which made that row as tall as the entire child grid and pushed every later row of the
  block below it — the worksheet form of the defect the fixed-layout renderer fixed by growing a merge's
  LAST spanned row. In a report whose group band pairs a merged subreport column with three detail rows,
  the second and third detail rows began under the whole child table instead of beside its second and third
  rows, so none of the row rules lined up across the two columns. Fixed in the Excel grid layer: child edges
  are now measured from the block top and distributed to whichever spanned row contains each one
  (`blockRowOffsets`/`blockPosition`), a row that only INHERITS such an edge is recorded so the
  grow-to-measured-height pass appends rather than overwrites it, and regions are placed against one flat
  map of physical row tops (`physicalRowAt`/`physicalRowBefore`) instead of a single row's profile.
  Format-specific by construction: PDF already laid the block out correctly and is the oracle here,
  `DOCX_EDITABLE`/`DOCX_VISUAL` inherit that PDF geometry, and only XLSX rebuilds the grid itself.
  Verified by comparing, for every pair of content blocks, the vertical relationship the canonical PDF gives
  them against the one the workbook gives them: the reported report went from 26 mismatched pairs to 0, and
  three unrelated corpus reports from 40, 36 and 37 to 0, 0 and 3. Regression:
  `test/excel-merged-cell-child-grid.test.js`. Ref: `sizing-and-growth.md`.

- **Continuation markers fired on every page break instead of on a real continuation.** The opt-in
  `pagination.continuationMarkers` label was emitted whenever a tablix crossed a page boundary with any
  row-span open, and unconditionally on every page break of a paginated child region. A grouped tablix
  therefore carried "Continued from previous page" on every page of the report even though each of those
  pages simply STARTED a fresh row (measured: a 244-page grouped fixture drew 243 of them; the bands also
  cost 35 pages of body height, so removing them returned the same data to 209 pages). The label is now
  decided from the one piece of measured state that means it: a physical row whose own content the break
  cut and which resumes on the page being started — the split-text loop and the child-region fragment loop
  in `drawRow`, which already know they are mid-row. An open row-span no longer qualifies (a group
  spanning the boundary is not a row that was cut), and neither does a merge that outgrew its rows: that
  break passes `rowContinuation: false`, keeping the fragment's open bottom edge without claiming a row
  continued. Text and on/off are deployment config (`continuation.rowLabel`); disabled, the geometry is
  identical to rendering without the request option. Shared-layer fix in the tablix pagination engine
  (`renderTablix`, `pdf.js`), so `DOCX_EDITABLE` and `DOCX_VISUAL` inherit it through the canonical
  trace; XLSX has no pagination and therefore no continuation. Regression:
  `test/tablix-continuation-labels.test.js`. Ref: `page-and-flow.md`.

- **A merge's extra height belonged to no row (grid missing under a grown group).** A merged (row-span) cell
  taller than the rows it spans grows its group; SSRS sizes a row to the tallest content that *ends* in it,
  so that growth belongs to the merge's LAST spanned row and every cell of that row is painted and ruled at
  the grown height. `renderTablix` (`pdf.js`) appended the difference after the last row instead, as a band
  no row owned: the merged cell drew its own box through it while every other column stopped at its natural
  height, so the band showed one tall cell flanked by columns with no borders at all and the row's fills
  stopped short of the row. Fixed in the shared PDF tablix layer — `closingMergeRequirement` hands the
  closing row the height `growGroupForClosingSpans` would otherwise append, so `drawRowContent` sizes,
  fills, and rules all of that row's cells together; the page-filling path is unchanged for a merge that
  cannot close on the current page. XLSX already implemented the same rule (`excel.js` grows
  `measuredHeights[endRow]` for a span it cannot contain), and `DOCX_EDITABLE`/`DOCX_VISUAL` inherit the
  corrected geometry through the canonical trace. Regression: `test/rowspan-merge-closing-row.test.js`
  (PDF geometry, the Word row that holds the closing detail, the Excel row height, plus the short-merge
  counterexample). Ref: `sizing-and-growth.md`, `border-resolution.md`.

- **A tablix row was split mid-text as soon as it overflowed the page remainder.** A tablix row is SSRS's
  indivisible pagination unit: a row that does not fit what is left of the page moves whole to the next one,
  and only a row that cannot fit a page at all is split. `drawRow` (`pdf.js`) moved the whole row only when
  KeepTogether was declared or the row carried no continuation-able text, and split every other overflowing
  row. Two visible consequences: a value SSRS keeps on one page was broken across two, and — because a cut
  inside a row is deliberately left open, drawing no closing rule (`tablix-split-row-open-edge.test.js`) —
  the table stopped flush on the printable body boundary, hard against the page footer's own rule, instead
  of closing at its last complete row with the gap that leaves. Fixed by moving any non-fitting row that
  fits a fresh page; the split path stays for a row taller than a page, and the child-data-region row keeps
  its own deeper break (entry above). PDF owns pagination, `DOCX_EDITABLE` and `DOCX_VISUAL` inherit it from
  the canonical trace/raster, and XLSX has no pagination. Regression:
  `test/tablix-row-page-atomicity.test.js` (whole-row move, closure clear of the footer band, and the
  oversized-row counterexample that must still split). Ref: `page-and-flow.md`.

- **Row-span group header duplicated across a page break (DOCX overlap).** A row-span (merged) group-header
  cell whose group crossed a page boundary could be emitted twice at one origin, so DOCX_EDITABLE failed
  with "Overlapping editable PDF regions" (native Word cells cannot overlap; PDF and XLSX tolerate it). Root
  cause: the trailing row of a row-span group is text-less (columns covered by the merged header, empty
  body); when it did not fit at a page boundary the text-split loop could not advance it, so it was silently
  dropped — its open row-span stayed open and the header's residual later closed against the *next* group's
  cursor, painting a duplicate. Fixed in the shared tablix layer (`renderTablix`, `pdf.js`): a non-fitting,
  unsplittable row now moves to a fresh page and draws in order instead of being dropped. Shared-layer (the
  duplicate was in the PDF trace both renderers consume), construct-driven (any row-span group header
  spilling a text-less trailing row across a page break), not report-specific. Verified against the real
  production payload (overlap gone, full 297-page DOCX renders) and the full corpus (zero page-count/size/
  status changes). Regression: `test/rowspan-header-page-break.test.js`. Ref: `page-and-flow.md`.


- **SSRS List / canvas cells (all renderers).** A grouped 1×1 tablix whose cell is a Rectangle free-form
  canvas (textboxes, lines, charts, images, nested tablixes) now renders. **PDF**: cell content types
  Line/Chart/Image are supported (`RENDERABLE_CELL_ITEMS`), the canvas is drawn item-by-item at each item's
  position per group instance, sized to its content extent (not the large design row height), reflowed
  across pages, with SSRS vertical displacement so a grown nested region pushes later items down instead of
  overlapping. **DOCX_EDITABLE** inherits it via the layout trace (canvas items are recorded during the PDF
  pass → native text + embedded chart pictures), and **DOCX_VISUAL** rasterizes the same PDF. **XLSX**
  renders the canvas text and tables through the shared materialization. The change is gated to cells that
  carry a Line/Chart/Image — content that was previously *refused* — so no flat-cell report can regress
  (proven by a full 34-report corpus diff: zero status/page-count changes). Regression:
  `test/list-canvas-cell.test.js` (PDF/DOCX/XLSX), `test/tablixCellContent.test.js`. **Residuals:** XLSX does
  not yet embed the chart/image *inside a canvas cell* (only body-level charts become images); a canvas
  textbox's CanGrow uses declared height for displacement (not re-measured); a single canvas item taller
  than a full page is drawn where it starts rather than split; DOCX_EDITABLE canvas layout still needs
  Microsoft Word for Windows certification. Ref: `page-and-flow.md`.



- **Report culture (`Language`).** Formatting now follows the report `Language`: date order, decimal/
  thousands separators, currency symbol, and month/day names all resolve to the declared culture
  (en-ZA → `R 1 234.50` / `2026/03/04`, de-DE → `1.234,50 €` / `4. März 2026`), while a report with no
  `Language` keeps the legacy defaults (en-US numbers/currency, en-GB dates). `formatValue`/`formatNet`
  take a culture, `globals.culture` carries the resolved locale to every renderer, and `C` currency uses the
  locale's currency. Residuals above: no-format default order and item-level Language.
  Regression: `test/culture-formatting.test.js`. Ref: `value-formatting.md`.

- **Excel honors explicit date formats.** An explicitly-formatted date cell now stays a live typed Excel
  date whose number format is translated from the RDL `Format` (`dd/MM/yyyy` → `dd/mm/yyyy`, `MMMM yyyy` →
  `mmmm yyyy`, `y` → `mmmm yyyy`, `G` → `dd/mm/yyyy hh:mm:ss`), so Excel shows what PDF/DOCX show instead of
  a date-only default; an untranslatable format writes the exact formatted string. `excelDateFormat` does
  the translation. Regression: `test/excel-date-format.test.js`. Ref: `value-formatting.md`.

- **No-format DateTime default.** SSRS renders a `DateTime` with no `Format` as general date/time (date and
  time, with seconds) — the value's default `ToString`. The engine emitted a bare date. Fixed: the no-format
  Date default is now general date/time (`dd/MM/yyyy HH:mm:ss`) in `formatValue`; the text path routes any
  Date-or-ISO-string value with no format through the formatter; XLSX writes a typed date with a
  general-date/time number format. Explicit formats and culture ordering remain per the open items above.
  Regression: `test/date-format-tokens.test.js`. Ref: `value-formatting.md`.

- **Container-growth double border.** A Rectangle wrapping a Tablix drew its border at the declared height
  while the tablix grew, stranding a second bottom line. Fixed by painting the container border/box at the
  rendered extent (and correcting the trace height so DOCX inherits it). Ref: `sizing-and-growth.md`.

- **Footer date dump / crop.** `=Globals!ExecutionTime` was coerced with JS `String(Date)`, producing a
  multi-line timezone string that overflowed the fixed footer. Fixed by routing no-format Date values
  through the .NET format engine. (See the open deviation above for the remaining date-vs-datetime default.)
  Ref: `value-formatting.md`.

- **Chart float axis label + cropped category labels.** Tick values leaked `0.6000000000000001`; long
  category labels wrapped and were clipped by the plot clip rectangle. Fixed by formatting tick values to
  the interval precision and reserving the measured category-label band before sizing the plot.
  Ref: `charts-and-axes.md`.

## How to use this log

Before changing a renderer, check whether the construct appears here. When you fix an open item, move it to
"Verified fixed" with the SSRS rule and the verification evidence. When you discover a new deviation during
RCA, add it here with the rule it violates — even if you are not fixing it in the same change.
