import assert from 'node:assert/strict';
import test from 'node:test';
import { cellGeometryPt, resolveGridColumns, toDocxTwips } from '../src/render/tableLayout.js';

// A 7 row-header + 10 body column matrix like the Combined Assurance "Table 2", where the last
// four body columns pair a wide provider column with a narrow frequency column.
const columns = [30, 55, 58, 62, 61, 62, 62, 51, 20, 51, 20, 51, 20, 51, 20, 71, 62];
const width = columns.reduce((sum, value) => sum + value, 0);
const item = { columns, width, left: 0 };
const model = { page: { width: width + 20, marginLeft: 0, marginRight: 0 }, body: { width: width + 20 } };

test('resolveGridColumns scales columns to the tablix width and preserves the grid count', () => {
  const { columnsPt, totalPt, gridColumnCount } = resolveGridColumns(item);
  assert.equal(gridColumnCount, 17);
  assert.equal(Math.round(totalPt), width);
  assert.equal(columnsPt.length, 17);
});

test('cellGeometryPt sums covered columns and never falls back to the whole table width', () => {
  const { columnsPt } = resolveGridColumns(item);
  // "4th Line" header spans grid columns 13 and 14 (the wide provider + narrow frequency pair).
  const fourthLine = cellGeometryPt(columnsPt, 13, 2);
  assert.equal(Math.round(fourthLine.widthPt), columns[13] + columns[14]);
  // "Combined Assurance Level" is a single column at grid index 15.
  const combined = cellGeometryPt(columnsPt, 15, 1);
  assert.equal(Math.round(combined.widthPt), columns[15]);
  assert.equal(Math.round(combined.xOffsetPt), columns.slice(0, 15).reduce((sum, value) => sum + value, 0));
});

test('PDF points and DOCX twips derive from the same shared column widths', () => {
  const { columnsPt, totalPt } = resolveGridColumns(item);
  const docx = toDocxTwips(model, item, columnsPt, totalPt);
  assert.equal(docx.gridTwips.length, columnsPt.length);
  assert.equal(docx.gridTwips.reduce((sum, value) => sum + value, 0), docx.tableTwips);
  const maxDiffPt = Math.max(...columnsPt.map((value, index) => Math.abs(value - docx.gridTwips[index] / 20)));
  assert.ok(maxDiffPt < 0.1, `columns diverge by ${maxDiffPt}pt`);
});

test('a wide table is clamped to the usable body width for DOCX only', () => {
  const narrowModel = { page: { width: 200, marginLeft: 0, marginRight: 0 }, body: { width: 200 } };
  const { columnsPt, totalPt } = resolveGridColumns(item);
  const docx = toDocxTwips(narrowModel, item, columnsPt, totalPt);
  assert.equal(docx.scaledToFit, true);
  assert.ok(docx.tableTwips <= docx.availableTwips);
});
