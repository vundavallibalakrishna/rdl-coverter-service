import { cellGridWidth, distributeTwips, resolveGridColumns, toDocxTwips } from './tableLayout.js';

// Backwards-compatible DOCX geometry facade. The column proportions now come from the shared
// resolveGridColumns so the PDF and DOCX grids can never diverge; this wrapper only applies the
// DOCX-specific page clamp and twip rounding.
export function computeDocxTableGeometry(model, item) {
  const { columnsPt, totalPt } = resolveGridColumns(item);
  return toDocxTwips(model, item, columnsPt, totalPt);
}

export { cellGridWidth, distributeTwips };
