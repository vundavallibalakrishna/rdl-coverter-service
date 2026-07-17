import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { materializeTablixRows, validateRenderInput } from '../src/rdl/validation.js';
import { computeCellPlacements } from '../src/render/tableGrid.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { MISSING_SAMPLES, hasSamples, samplePath } from '../scripts/lib/samples.js';

const serviceRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const rdlPath = path.join(serviceRoot, '..', 'lumina', 'Combined Assurance Reports (2).rdl');
const hydrationPath = samplePath('combined-assurance-hydration.json');
const excelRdlPath = samplePath('Combined Assurance Reports Excel.rdl');
const excelRequestPath = samplePath('combined-assurance-excel-request.json');

test('the Combined Assurance hydration consumes its certified row-header, style, and pagination constructs', async (context) => {
  try {
    await fs.access(rdlPath);
    await fs.access(hydrationPath);
  } catch {
    context.skip(`Supplied RDL/hydration is not available (${rdlPath})`);
    return;
  }

  const [rdl, hydrationText] = await Promise.all([
    fs.readFile(rdlPath),
    fs.readFile(hydrationPath, 'utf8'),
  ]);
  const request = JSON.parse(hydrationText);
  const model = parseRdl(rdl);
  const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
  assert.deepEqual(model.unsupported, []);
  assert.equal(validateRenderInput(model, request, config).totalRows, 11);

  const assurance = model.body.items.find((item) => item.name === 'Tablix2');
  assert.equal(assurance.rowHeaderColumns.length, 7);
  assert.equal(assurance.bodyColumns.length, 10);
  assert.equal(assurance.columns.length, 17);
  const rows = materializeTablixRows(assurance, request.datasets.MainData, request.parameters, {}, {});
  const placements = computeCellPlacements(rows, assurance.columns.length);
  assert.equal(rows[0].cells[0].colSpan, 7);
  assert.equal(rows.slice(1).some((row) => row.cells.some((cell) => cell.isRowHeader && cell.rowSpan > 1)), true);
  for (const [rowIndex, row] of rows.entries()) {
    const firstBodyCell = row.cells.findIndex((cell) => !cell.isRowHeader);
    assert.equal(placements[rowIndex][firstBodyCell], 7);
  }

  const editable = await renderEditableDocx(model, request);
  const zip = await JSZip.loadAsync(editable.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  assert.match(documentXml, /Testing Header/);
  assert.match(documentXml, /<w:vMerge w:val="restart"\/>/);
  assert.match(documentXml, /<w:vMerge w:val="continue"\/>/);
  assert.match(documentXml, /<w:top w:val="single"/);

  for (const datasetName of model.renderingDatasets) {
    const definition = model.datasets.find((dataset) => dataset.name === datasetName);
    const expectedFields = definition.fields.map((field) => field.dataField).sort();
    for (const row of request.datasets[datasetName]) {
      assert.deepEqual(Object.keys(row).sort(), expectedFields);
    }
  }

  assert.equal(request.datasets.MainData.length + request.datasets.intro.length + request.datasets.User.length, 11);
  assert.equal(Math.round(model.page.width * 100) / 100, 841.89);
  assert.equal(Math.round(model.page.height * 100) / 100, 595.28);
});

test('the client Combined Assurance sample declares all three static tablix rows as repeating headers', {
  skip: hasSamples('Combined Assurance Reports Excel.rdl', 'combined-assurance-excel-request.json') ? false : MISSING_SAMPLES,
}, async () => {
  const [rdl, requestText] = await Promise.all([
    fs.readFile(excelRdlPath),
    fs.readFile(excelRequestPath, 'utf8'),
  ]);
  const request = JSON.parse(requestText);
  const model = parseRdl(rdl);
  const assurance = model.body.items.find((item) => item.name === 'Tablix2');
  const rows = materializeTablixRows(assurance, request.datasets.MainData, request.parameters, {}, {});

  assert.deepEqual(
    assurance.rowMembers.slice(0, 3).map((member) => ({ repeatOnNewPage: member.repeatOnNewPage, keepWithGroup: member.keepWithGroup })),
    Array.from({ length: 3 }, () => ({ repeatOnNewPage: true, keepWithGroup: 'After' })),
  );
  assert.deepEqual(rows.slice(0, 3).map((row) => row.isHeader), [true, true, true]);
  assert.equal(rows[3].isHeader, false);
});
