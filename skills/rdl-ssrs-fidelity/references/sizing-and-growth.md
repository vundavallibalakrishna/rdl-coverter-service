# Sizing and Growth Contract (SSRS)

SSRS lays out from declared sizes but resolves final geometry from content. Reproduce the growth chain and
paint every container at its **rendered** extent.

## Item growth

- `CanGrow` (default true for textboxes): the item grows **vertically** to fit its content; it never shrinks
  below its declared `Height` unless `CanShrink` is also set.
- `CanShrink`: the item may shrink below declared `Height` when content is shorter.
- Width is fixed (report items do not auto-grow horizontally); text wraps within the width, and wrapping is
  what drives vertical growth. An embedded line break in the value also adds lines.
- A rotated textbox (`WritingMode`) consumes its declared cross-axis extent; wrapping grows across the
  physical width, not down the page.

## Tablix rows and cells

- A tablix cell's height is `max(declared row height, measured content height)` using the same font metrics
  and rich-text run boundaries as the final draw.
- A row's height is the max over its cells (including the row-span contribution of merged cells).
- A grown row pushes every following row down; the tablix's total height is the sum of rendered row
  heights, not the sum of declared heights.

## Container envelope (the critical rule)

- A container — Rectangle, tablix cell holding nested items, the body — grows to **contain** its grown
  children. Its own `Height` is a minimum, not a cap.
- The container's **border and background are painted at the rendered (grown) extent**, computed *after* its
  children are laid out. Painting them at the declared height strands the bottom edge above the grown
  content and, where that edge is meant to coincide with a child's edge (e.g. a tablix's last-row border),
  produces a visible **double line**. This is the canonical container-growth defect.
- Coincident edges stay coincident through growth: when a container and its sole/last child share a bottom
  edge by design, the grown geometry must keep both at one coordinate (one visible line).

## Fixed vs growing regions

- **Page**: fixed (`PageHeight`/`PageWidth`, margins). Never grows.
- **PageHeader / PageFooter**: **fixed** `Height`. They do **not** grow; content taller than the band is
  **clipped**, not expanded. Therefore content that must fit a footer must be short — the correct fix for a
  clipped footer is almost always upstream (e.g. a value formatted to one line), not enlarging the band.
  Honor `PrintOnFirstPage` / `PrintOnLastPage`.
- **Body**: grows with content; page breaks occur here.

## Page fragmentation of containers

- A Rectangle (or any container) that carries a **visible border or background** and would span a page
  boundary cannot be safely fragmented (the extent would paint once at the wrong height). Fail closed
  (`UNSUPPORTED_FEATURE`) rather than draw a misleading partial box. A borderless, fill-less container may
  flow across pages.
- A split tablix segment is capped at the printable body boundary (`page height − bottom margin − footer
  height`); a split row never paints into the footer band.

## Verification

Grow a cell past its declared height (wrap or embedded newline) and assert the container's border/box moves
with it — no double edge, no clipped content — in PDF, in the trace-driven DOCX_EDITABLE, and in XLSX row
heights. DOCX_VISUAL inherits from the canonical PDF.
