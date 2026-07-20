// Integration tests for the XLSX renderer. The point of Excel export over PDF/DOCX is live, computable
// values, so these assert that numbers are real numbers with format codes, styling and merges survive, and
// untrusted values are stored as typed strings that Excel cannot execute.
import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { parseRdl } from '../src/rdl/parser.js';
import { loadConfig } from '../src/config.js';
import { renderExcel } from '../src/render/excel.js';
import { renderDocument, OUTPUTS } from '../src/render/index.js';

const model = parseRdl(await fs.readFile(new URL('./fixtures/basic.rdl', import.meta.url)));
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const request = { outputFileName: 'Sales', parameters: { Title: 'Sales', Choice: 'A' }, datasets: { Sales: [{ Name: 'North', Amount: 1234.5 }, { Name: 'South', Amount: 99 }], Choices: [{ Value: 'A' }] } };

async function load(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb.worksheets[0];
}
const findCell = (ws, text) => {
  let found = null;
  ws.eachRow((row) => row.eachCell((cell) => { if (cell.value === text) found = cell; }));
  return found;
};

test('XLSX is a registered output and dispatches through renderDocument', async () => {
  assert.equal(OUTPUTS.has('XLSX'), true);
  const result = await renderDocument(model, { ...request, output: 'XLSX' }, config, null);
  assert.equal(result.extension, 'xlsx');
  assert.equal(result.mimeType, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  assert.equal(result.buffer.subarray(0, 2).toString(), 'PK');
});

test('renders a valid workbook that reloads with the expected header and data text', async () => {
  const result = await renderExcel(model, request, config, null);
  const ws = await load(result.buffer);
  assert.match(ws.name, /Sales/);
  assert.ok(findCell(ws, 'Name'));
  assert.ok(findCell(ws, 'Amount'));
  assert.ok(findCell(ws, 'North'));
  assert.ok(findCell(ws, 'South'));
});

test('a Format()-wrapped numeric field becomes a live number with a translated format code', async () => {
  // basic.rdl renders Amount as =Format(Fields!Amount.Value, "N2"). It must arrive as the number 1234.5,
  // not the string "1,234.50", so it stays summable — with an N2-equivalent Excel format for display.
  const ws = await load((await renderExcel(model, request, config, null)).buffer);
  let amount = null;
  ws.eachRow((row) => row.eachCell((cell) => { if (cell.value === 1234.5) amount = cell; }));
  assert.ok(amount, 'expected a live numeric 1234.5 cell');
  assert.equal(typeof amount.value, 'number');
  assert.equal(amount.numFmt, '#,##0.00');
});

test('a text field stays text', async () => {
  const ws = await load((await renderExcel(model, request, config, null)).buffer);
  assert.equal(typeof findCell(ws, 'North').value, 'string');
});

test('cell fills and borders from the RDL survive into the workbook', async () => {
  const ws = await load((await renderExcel(model, request, config, null)).buffer);
  const header = findCell(ws, 'Amount'); // basic.rdl gives the header a #dddddd background and a solid border
  assert.equal(header.fill?.pattern, 'solid');
  assert.equal(header.fill?.fgColor?.argb, 'FFDDDDDD');
  assert.equal(header.border?.bottom?.style, 'thin');
  assert.equal(header.font?.bold, true);
});

test('merged cells are emitted for spanning cells', async () => {
  const merged = structuredClone(model);
  const tablix = merged.body.items.find((item) => item.type === 'Tablix');
  tablix.rows[0].cells[0].colSpan = 2; // span the first header cell across both columns
  const result = await renderExcel(merged, request, config, null);
  const sheetXml = await (await JSZip.loadAsync(result.buffer)).file('xl/worksheets/sheet1.xml').async('string');
  assert.match(sheetXml, /<mergeCell ref="[A-Z]+\d+:[A-Z]+\d+"/);
});

test('sheetPerTablix puts each tablix on its own worksheet with its own columns', async () => {
  const single = await renderExcel(model, request, config, null);
  assert.equal(single.sheetCount, 1); // default: everything stacked on one sheet

  const perTablix = await renderExcel(model, { ...request, excel: { sheetPerTablix: true } }, config, null);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(perTablix.buffer);
  const names = wb.worksheets.map((ws) => ws.name);
  assert.ok(names.includes('Table 1'), `expected a per-tablix sheet, got ${names.join(', ')}`);
  // basic.rdl has a free-form title textbox, so its non-tablix content lands on a leading Overview sheet.
  assert.ok(names.includes('Overview'));
  const table = wb.getWorksheet('Table 1');
  assert.ok(findCell(table, 'North'), 'the tablix data belongs on its own sheet');
});

test('columns are autofit to their content width', async () => {
  const ws = await load((await renderExcel(model, request, config, null)).buffer);
  // Every column that received content should get a concrete, bounded width rather than Excel's default.
  const widths = ws.columns.map((column) => column.width).filter((width) => typeof width === 'number');
  assert.ok(widths.length >= 2, 'expected at least the two data columns to be sized');
  assert.ok(widths.every((width) => width >= 6 && width <= 60), `widths out of clamp: ${widths.join(', ')}`);
});

test('an untrusted value beginning with = is stored as a typed string, never a live formula', async () => {
  const evil = { ...request, datasets: { ...request.datasets, Sales: [{ Name: '=1+2+cmd|calc', Amount: 1 }] } };
  const result = await renderExcel(model, evil, config, null);
  const sheetXml = await (await JSZip.loadAsync(result.buffer)).file('xl/worksheets/sheet1.xml').async('string');
  // The dangerous string must appear only inside a string cell/sharedString, and there must be no formula
  // element anywhere in the sheet.
  assert.doesNotMatch(sheetXml, /<f>/);
  const ws = await load(result.buffer);
  const cell = findCell(ws, '=1+2+cmd|calc');
  assert.ok(cell, 'value should round-trip verbatim (no apostrophe corruption)');
  assert.equal(typeof cell.value, 'string');
});
