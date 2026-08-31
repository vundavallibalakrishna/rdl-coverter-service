# Positioning, Page, and Flow Contract (SSRS)

Reproduce SSRS's absolute coordinate model and its page/pagination rules.

## Coordinate model

- Report items are **absolutely positioned** by `Top`/`Left` within their container's coordinate space
  (body, rectangle, or cell). They are not a vertical flow: two items with the same `Top` and different
  `Left` sit **side by side**; items may **overlap**, ordered by `ZIndex`.
- Reproduce by placing each item at its resolved `left/top` within its container. A vertical-stacking
  approximation collapses genuine 2-D layouts and is only coincidentally correct for reports whose items
  happen to be stacked.
- Sizes convert from RDL units (in/cm/mm/pt) to points as the layout source of truth; keep point precision
  until the target renderer converts (Word uses twips; distribute rounding so summed columns equal the
  traced width).

## Page geometry

- `PageHeight`/`PageWidth` and the four margins define the physical page. The printable body width is
  `PageWidth − LeftMargin − RightMargin`; body height similarly less header/footer bands.
- **PageHeader / PageFooter** are fixed-height bands repeated per page (subject to `PrintOnFirstPage` /
  `PrintOnLastPage`). They do not grow; overflow is clipped (see sizing contract). `Globals!PageNumber` and
  `Globals!TotalPages` resolve per rendered page — a body/cell reference to `TotalPages` must resolve to the
  final count, not a provisional value.

## Pagination

- Page breaks occur in the growing body. Honor explicit breaks: group `PageBreak` (Start/End/Between),
  rectangle/tablix `PageBreak`, and `PageBreakAtEnd/Start`. A break must not emit a blank trailing page.
- `KeepTogether` is best-effort ("keep on one page if possible") and is emitted on nearly every textbox;
  do not treat it as an absolute veto that stops long content from ever crossing a page. Atomic
  keep-together is enforced where the unit is truly indivisible (a tablix row).
- A tablix **row** is that indivisible unit: a row that does not fit the space left on the page moves whole
  to the next page. Split a row across pages only when it cannot fit a page at all — a row split at a page
  boundary breaks a value SSRS keeps together and ends the fragment mid-row, where no closing rule is drawn,
  so the table stops flush on the printable body boundary against the page footer's own rule. A row whose
  cell holds a **child data region** is the documented exception: the break falls between the child's rows.
- Tablix `RepeatColumnHeaders`/`RepeatRowHeaders` repeat headers at the top of each page/fragment;
  `KeepWithGroup`/group headers stay with their data. Repeated headers must not add a blank page.
- **Wide content**: a tablix/matrix wider than the printable body is split across horizontal page columns,
  repeating row-header columns on each — SSRS paginates horizontally as well as vertically. Do not silently
  overflow the right edge.

## Subreports, groups, variables

- Subreports render only from caller-bundled, invocation-scoped content (metadata only otherwise); a
  DATA-mode XLSX subreport is fail-closed.
- Matrix/cross-tab column hierarchies (`TablixHeader`, `TablixCorner`), group header/footer subtotals, and
  recursive/parent (`Group/Parent`) row groups render per RDL. `Group/Variables` resolve as
  `Variables!Name.Value` in the current row scope.

## Determinism

Same RDL + parameters + datasets + fonts + runtime ⇒ identical output bytes. `Globals!ExecutionTime` is the
only intentional non-determinism; keep it isolated so nothing else varies per run.

## Verification

Two textboxes at equal `Top`, different `Left` render side by side (not stacked). A body/cell
`=Globals!TotalPages` resolves to the final page count. A wide matrix splits into horizontal page groups
with repeated row headers. Header/footer content that exceeds the band clips rather than growing the band.
