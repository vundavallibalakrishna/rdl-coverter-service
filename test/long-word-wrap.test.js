import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { promisify } from 'node:util';
import JSZip from 'jszip';
import PDFDocument from 'pdfkit';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { measureTextboxHeight, renderPdf } from '../src/render/pdf.js';

const execFileAsync = promisify(execFile);
const sourceModel = parseRdl(await fs.readFile(new URL('./fixtures/basic.rdl', import.meta.url)));
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const narrowWidth = 45.5;

function scenario(value = 'establishment', { bold = false } = {}) {
  const model = structuredClone(sourceModel);
  const tablix = model.body.items.find((item) => item.type === 'Tablix');
  tablix.columns[0] = narrowWidth;
  tablix.width = tablix.columns.reduce((sum, width) => sum + width, 0);
  const detailRow = tablix.rows[tablix.rows.length - 1];
  const textbox = detailRow.cells[0].items.find((item) => item.type === 'Textbox');
  textbox.canGrow = true;
  textbox.style.fontFamily = 'Arial';
  textbox.style.fontSize = 8;
  textbox.style.fontWeight = bold ? 'Bold' : 'Normal';
  textbox.style.paddingLeft = 2;
  textbox.style.paddingRight = 2;
  for (const run of textbox.paragraphs.flat()) {
    run.style.fontFamily = 'Arial';
    run.style.fontSize = 8;
    run.style.fontWeight = bold ? 'Bold' : 'Normal';
  }
  const request = {
    parameters: { Title: 'Long word wrapping', Choice: 'A' },
    datasets: { Sales: [{ Name: value, Amount: 1 }], Choices: [{ Value: 'A' }] },
  };
  return { model, request, textbox };
}

test('PDF rich text wraps an oversized word at grapheme boundaries instead of clipping its suffix', async (context) => {
  const { model, request, textbox } = scenario();
  const measurementDoc = new PDFDocument({ autoFirstPage: false });
  measurementDoc.addPage();
  const contextValues = { fields: { Name: 'establishment' }, parameters: request.parameters, globals: {}, datasets: request.datasets };
  const shortHeight = measureTextboxHeight(measurementDoc, config, textbox, contextValues, 'short', narrowWidth);
  const wrappedHeight = measureTextboxHeight(measurementDoc, config, textbox, contextValues, 'establishment', narrowWidth);
  measurementDoc.end();
  assert.ok(wrappedHeight > shortHeight * 1.5, `expected a multi-line word (${wrappedHeight} > ${shortHeight})`);

  const boldScenario = scenario('establishment', { bold: true });
  const boldMeasurementDoc = new PDFDocument({ autoFirstPage: false });
  boldMeasurementDoc.addPage();
  const boldContext = {
    fields: { Name: 'establishment' },
    parameters: boldScenario.request.parameters,
    globals: {},
    datasets: boldScenario.request.datasets,
  };
  const boldShortHeight = measureTextboxHeight(
    boldMeasurementDoc,
    config,
    boldScenario.textbox,
    boldContext,
    'short',
    narrowWidth,
  );
  const boldWrappedHeight = measureTextboxHeight(
    boldMeasurementDoc,
    config,
    boldScenario.textbox,
    boldContext,
    'establishment',
    narrowWidth,
  );
  boldMeasurementDoc.end();
  assert.ok(
    boldWrappedHeight > boldShortHeight * 1.5,
    `expected a multi-line bold word (${boldWrappedHeight} > ${boldShortHeight})`,
  );

  const rendered = await renderPdf(model, request, config);
  const pdfPath = path.resolve('tmp', `long-word-wrap-${process.pid}.pdf`);
  context.after(() => fs.rm(pdfPath, { force: true }));
  await fs.writeFile(pdfPath, rendered.buffer, { mode: 0o600 });
  const { stdout } = await execFileAsync('pdftotext', ['-layout', pdfPath, '-']);
  // PDF text extraction may interleave the adjacent numeric cell between wrapped line fragments, but the
  // complete prefix and suffix must both remain selectable; the height assertion proves a real line break.
  assert.match(stdout.replace(/\s+/g, ''), /establish.*ent/s);
});

test('editable DOCX enables character-level wrapping for ordinary RDL paragraphs', async () => {
  const { model, request } = scenario();
  const rendered = await renderEditableDocx(model, request, config);
  const zip = await JSZip.loadAsync(rendered.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  assert.match(documentXml, /<w:wordWrap w:val="1"\/>/);
  assert.doesNotMatch(documentXml, /<w:wordWrap w:val="0"\/>/);
  assert.match(documentXml, /establishment/);
  assert.ok(
    documentXml.indexOf('<w:wordWrap w:val="1"/>') < documentXml.indexOf('<w:spacing '),
    'wordWrap must remain in the schema-defined paragraph-property position',
  );
});
