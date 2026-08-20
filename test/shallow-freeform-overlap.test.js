// Adjacent free-form boxes in a real RDL routinely overlap by a fraction of a point: an inch-valued box
// and a centimetre-valued neighbour rarely land on the same quarter point. A fixed-layout renderer strokes
// both and nobody notices. A grid renderer cannot — Excel has no way to place two items in overlapping
// merged ranges — so the worksheet path rejected the export with
// `RDL_INVALID: RDL produced overlapping Excel merged-cell ranges`.
//
// Excel already had the rule that resolves a shared edge onto the midpoint of the two coordinates, but it
// applied only within a border stroke and only to items ADJACENT in the sorted sibling list. A page-level
// section holds dozens of free-form boxes spread down the page, so the two that share a horizontal band
// are almost never neighbours in x across the whole group; those pairs were never even examined.
//
// The rule now matches the one the page-locked Word renderer already applies to the same construct
// (coalesceShallowEdgeOverlaps): a positive overlap resolves onto the shared midpoint whenever neither edge
// has to move further than the certified 0.5pt geometry tolerance, i.e. for overlaps up to twice that.
// Deeper overlaps are intentional and still fail closed. Gaps are unchanged — a gap is real whitespace and
// only a border-stroke sliver may be closed.
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { promisify } from 'node:util';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

const execFileAsync = promisify(execFile);
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const request = { outputFileName: 'shallow-overlap', parameters: {}, datasets: {} };

function textbox(name, value, { left, top, width, height }) {
  return `<Textbox Name="${name}">
    <Paragraphs><Paragraph><TextRuns><TextRun><Value>${value}</Value></TextRun></TextRuns></Paragraph></Paragraphs>
    <Top>${top}pt</Top><Left>${left}pt</Left><Height>${height}pt</Height><Width>${width}pt</Width>
    <Style><FontFamily>Arial</FontFamily><FontSize>8pt</FontSize><Border><Style>None</Style></Border></Style>
  </Textbox>`;
}

/**
 * LEFT_BOX and RIGHT_BOX share a horizontal band and overlap by `overlap` points. SPACER sits between them
 * in the sorted-by-left order but in a different vertical band, which is exactly the arrangement that hid
 * the overlapping pair from the old consecutive-neighbours rule.
 *
 * When `axis` is 'vertical' the same shallow overlap is between stacked boxes instead.
 */
function report({ overlap = 0.6, axis = 'horizontal' } = {}) {
  const items = axis === 'horizontal'
    ? [
      textbox('LeftBox', 'LEFT_BOX', { left: 0, top: 40, width: 100, height: 20 }),
      textbox('Spacer', 'SPACER', { left: 50, top: 0, width: 100, height: 20 }),
      textbox('RightBox', 'RIGHT_BOX', { left: 100 - overlap, top: 40, width: 100, height: 20 }),
    ]
    : [
      textbox('TopBox', 'LEFT_BOX', { left: 0, top: 0, width: 100, height: 40 }),
      textbox('Spacer', 'SPACER', { left: 200, top: 20, width: 100, height: 20 }),
      textbox('BottomBox', 'RIGHT_BOX', { left: 0, top: 40 - overlap, width: 100, height: 40 }),
    ];
  return Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <ReportSections><ReportSection>
    <Body>
      <ReportItems>${items.join('\n')}</ReportItems>
      <Height>2in</Height><Style/>
    </Body>
    <Width>7in</Width>
    <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`, 'utf8');
}

async function worksheet(options, context) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-shallow-overlap-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderExcel(parseRdl(report(options)), { ...request, excelLayoutMode: 'REPORT' }, config, tempDir);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(rendered.buffer);
  return workbook.worksheets[0];
}

function textAt(sheet, needle) {
  let found = null;
  sheet.eachRow((row, rowNumber) => row.eachCell((cell, columnNumber) => {
    const value = cell.value?.richText ? cell.value.richText.map((run) => run.text).join('') : String(cell.value ?? '');
    if (!found && value.includes(needle)) found = { rowNumber, columnNumber };
  }));
  return found;
}

test('XLSX: a sub-point horizontal overlap resolves to a shared edge instead of failing closed', async (context) => {
  // Previously: RDL_INVALID "RDL produced overlapping Excel merged-cell ranges".
  const sheet = await worksheet({}, context);
  assert.ok(textAt(sheet, 'LEFT_BOX'), 'the left box must reach the worksheet');
  assert.ok(textAt(sheet, 'RIGHT_BOX'), 'the right box must reach the worksheet');
  // Resolved onto one grid line, so the two occupy disjoint column ranges.
  assert.ok(textAt(sheet, 'RIGHT_BOX').columnNumber > textAt(sheet, 'LEFT_BOX').columnNumber);
  const merges = Object.values(sheet._merges || {});
  for (const [first, second] of merges.flatMap((a, i) => merges.slice(i + 1).map((b) => [a, b]))) {
    const overlaps = first.left <= second.right && second.left <= first.right
      && first.top <= second.bottom && second.top <= first.bottom;
    assert.equal(overlaps, false, 'no two merged ranges may overlap');
  }
});

test('XLSX: the same sub-point overlap between stacked boxes resolves on the row grid', async (context) => {
  const sheet = await worksheet({ axis: 'vertical' }, context);
  assert.ok(textAt(sheet, 'LEFT_BOX'), 'the top box must reach the worksheet');
  assert.ok(textAt(sheet, 'RIGHT_BOX'), 'the bottom box must reach the worksheet');
  assert.ok(textAt(sheet, 'RIGHT_BOX').rowNumber > textAt(sheet, 'LEFT_BOX').rowNumber);
});

test('XLSX: an overlap right at twice the certified tolerance still resolves', async (context) => {
  // Each edge moves 0.5pt — the limit, not beyond it.
  const sheet = await worksheet({ overlap: 1 }, context);
  assert.ok(textAt(sheet, 'LEFT_BOX') && textAt(sheet, 'RIGHT_BOX'));
});

test('counterexample: a deep overlap is intentional and still fails closed', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-deep-overlap-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  // 6pt of genuine overlap cannot be resolved without moving an edge far past the tolerance. Silently
  // snapping it would move real content; the export must still refuse rather than approximate.
  await assert.rejects(
    () => renderExcel(parseRdl(report({ overlap: 6 })), { ...request, excelLayoutMode: 'REPORT' }, config, tempDir),
    (error) => error.code === 'RDL_INVALID' && /overlapping Excel merged-cell ranges/.test(error.message),
  );
});

test('DOCX_EDITABLE: the same construct stays native and keeps both boxes', async () => {
  const rendered = await renderEditableDocx(parseRdl(report()), request, config);
  const xml = await (await JSZip.loadAsync(rendered.buffer)).file('word/document.xml').async('string');
  assert.match(xml, /LEFT_BOX/);
  assert.match(xml, /RIGHT_BOX/);
  assert.match(xml, /<w:tbl>/);
});

test('PDF is unaffected: a fixed-layout renderer just strokes both boxes', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-shallow-overlap-pdf-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, 'overlap.pdf');
  const rendered = await renderPdf(parseRdl(report()), request, config);
  await fs.writeFile(pdfPath, rendered.buffer);
  const { stdout } = await execFileAsync('pdftotext', ['-layout', pdfPath, '-']);
  assert.match(stdout, /LEFT_BOX/);
  assert.match(stdout, /RIGHT_BOX/);
  // Even the deep overlap the grid renderers refuse is drawable here, which is why they and not the PDF
  // own the restriction.
  const deepPath = path.join(tempDir, 'deep.pdf');
  await fs.writeFile(deepPath, (await renderPdf(parseRdl(report({ overlap: 6 })), request, config)).buffer);
  const deep = await execFileAsync('pdftotext', ['-layout', deepPath, '-']);
  assert.match(deep.stdout, /LEFT_BOX/);
});
