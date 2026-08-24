# Known Engine ↔ SSRS Deviations

Running log of places this engine is known to differ from SSRS, or where fidelity is unproven. Each entry
names the SSRS rule, the current behavior, and how to reconcile. Remove an entry only when it is fixed and
verified against a real SSRS oracle. Never let an entry justify a report-specific workaround.

## Open

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
