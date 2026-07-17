import { pointsToTwips } from '../units.js';

// Single source of truth for tablix grid geometry, consumed by BOTH the PDF and DOCX renderers
// so the two outputs build the identical column grid. Points are the float source of truth;
// twip rounding happens only at the DOCX boundary (toDocxTwips).

export function gridColumnCount(item) {
  if (item.columns?.length) return item.columns.length;
  return Math.max(1, ...(item.rows || []).map((row) => (row.cells || []).reduce((sum, cell) => sum + (cell.colSpan || 1), 0)));
}

// Per-grid-column widths in points. Explicit RDL column widths (row-header + body columns) are
// scaled so their sum equals the tablix's declared width (item.width), the fixed-layout source of
// truth. When explicit widths are absent, item.width is split evenly across the derived column
// count. Both renderers call this, so neither can drift from the other.
export function resolveGridColumns(item) {
  const count = gridColumnCount(item);
  const source = item.columns?.length
    ? item.columns
    : Array.from({ length: count }, () => (item.width || count) / count);
  const positive = source.map((width) => (Number.isFinite(width) && width > 0 ? width : 0));
  const totalSource = positive.reduce((sum, width) => sum + width, 0) || positive.length;
  const scale = item.width > 0 && totalSource > 0 ? item.width / totalSource : 1;
  const columnsPt = positive.map((width) => width * scale);
  const totalPt = columnsPt.reduce((sum, width) => sum + width, 0);
  return { columnsPt, totalPt, gridColumnCount: columnsPt.length };
}

// X offset and width in points for a cell starting at gridColumnStart and covering colSpan grid
// columns. Replaces the previous per-renderer fallbacks that silently defaulted a mis-indexed
// cell to the entire table width, which drew cells on top of their neighbours.
export function cellGeometryPt(columnsPt, gridColumnStart, colSpan = 1) {
  const start = Math.max(0, Math.min(gridColumnStart, columnsPt.length - 1));
  const span = Math.max(1, colSpan);
  const widthPt = columnsPt.slice(start, start + span).reduce((sum, width) => sum + width, 0);
  const xOffsetPt = columnsPt.slice(0, start).reduce((sum, width) => sum + width, 0);
  return { xOffsetPt, widthPt: widthPt || columnsPt[start] || 0 };
}

// Largest-remainder integer-twip allocator: distributes `target` twips across the given point
// widths so every column is at least 1 twip and the columns sum exactly to `target`.
export function distributeTwips(widths, target) {
  const safeTarget = Math.max(widths.length, Math.round(target));
  const positive = widths.map((width) => (Number.isFinite(width) && width > 0 ? width : 1));
  const total = positive.reduce((sum, width) => sum + width, 0);
  const exact = positive.map((width) => (width / total) * safeTarget);
  const result = exact.map((width) => Math.max(1, Math.floor(width)));
  let remainder = safeTarget - result.reduce((sum, width) => sum + width, 0);
  const order = exact.map((width, index) => ({ index, fraction: width - Math.floor(width) }))
    .sort((left, right) => right.fraction - left.fraction || left.index - right.index);
  for (let index = 0; remainder > 0; index += 1, remainder -= 1) result[order[index % order.length].index] += 1;
  while (remainder < 0) {
    const candidate = result.findIndex((width) => width > 1);
    if (candidate < 0) break;
    result[candidate] -= 1;
    remainder += 1;
  }
  return result;
}

export function cellGridWidth(gridTwips, columnIndex, colSpan = 1) {
  const start = Math.max(0, Math.min(columnIndex, gridTwips.length - 1));
  return gridTwips.slice(start, start + Math.max(1, colSpan)).reduce((sum, width) => sum + width, 0);
}

// DOCX twip geometry derived from the SAME point widths as the PDF path. Word cannot overflow the
// page the way the fixed PDF can, so the table (and its left indent) is clamped to the usable body
// width; the shared columns are then rounded to integer twips.
export function toDocxTwips(model, item, columnsPt, totalPt) {
  const pageUsablePt = Math.max(1, model.page.width - model.page.marginLeft - model.page.marginRight);
  const bodyUsablePt = Math.max(1, Math.min(model.body.width || pageUsablePt, pageUsablePt));
  const requestedPt = Math.max(1, totalPt);
  const tablePt = Math.min(requestedPt, bodyUsablePt);
  const desiredIndentPt = Math.max(0, item.left || 0);
  const indentPt = Math.min(desiredIndentPt, Math.max(0, bodyUsablePt - tablePt));
  const tableTwips = pointsToTwips(tablePt);
  return {
    availableTwips: pointsToTwips(bodyUsablePt),
    tableTwips,
    indentTwips: pointsToTwips(indentPt),
    gridTwips: distributeTwips(columnsPt, tableTwips),
    scaledToFit: requestedPt > bodyUsablePt || desiredIndentPt > indentPt,
  };
}
