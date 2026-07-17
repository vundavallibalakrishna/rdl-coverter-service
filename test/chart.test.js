import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import JSZip from 'jszip';
import PDFDocument from 'pdfkit';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { color, normalizeDatasets } from '../src/render/common.js';
import { materializeChart } from '../src/render/chartData.js';
import { drawChart } from '../src/render/chart.js';
import { renderPdf } from '../src/render/pdf.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { MISSING_SAMPLES, hasSamples, samplePath } from '../scripts/lib/samples.js';

const execFileAsync = promisify(execFile);
// The client RDL lives outside version control (see scripts/lib/samples.js), so a fresh clone skips the
// sample-backed tests instead of failing on a file it was never given.
const RDL_NAME = 'Incident Dashboard Report.rdl';
const REQUEST_NAME = 'incident-dashboard-request.json';
const skip = hasSamples(RDL_NAME, REQUEST_NAME) ? false : MISSING_SAMPLES;
const model = skip ? null : parseRdl(await fs.readFile(samplePath(RDL_NAME)));
const request = skip ? null : JSON.parse(await fs.readFile(samplePath(REQUEST_NAME), 'utf8'));
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '30000' });
const chart = (name) => model.body.items.find((item) => item.name === name);
const dataFor = (name) => materializeChart(chart(name), normalizeDatasets(model, request), request.parameters, {});

test('the incident dashboard with charts passes the capability gate', { skip }, () => {
  assert.deepEqual(model.unsupported, []);
  assert.equal(model.body.items.filter((item) => item.type === 'Chart').length, 3);
});

test('parses the three chart types, series grouping, and legend declarations', { skip }, () => {
  assert.equal(chart('Chart1').chartType, 'bar');
  assert.equal(chart('Chart2').chartType, 'pie');
  assert.equal(chart('Chart4').chartType, 'column');
  assert.equal(chart('Chart1').series, null); // single static series
  assert.equal(chart('Chart4').series.group.expressions[0], '=Fields!Incident_Type.Value');
  // Chart1/Chart2 omit <ChartLegends> (manual legend items); only Chart4 declares one.
  assert.equal(chart('Chart1').legend.visible, false);
  assert.equal(chart('Chart4').legend.visible, true);
});

test('resolves CSS/X11 colour names to hex so PDFKit fills render', () => {
  assert.equal(color('SlateBlue'), '#6a5acd');
  assert.equal(color('DarkOliveGreen'), '#556b2f');
  assert.equal(color('DarkTurquoise'), '#00ced1');
  assert.equal(color('LightBlue'), '#add8e6');
  assert.equal(color('#123abc'), '#123abc'); // hex passes through
});

test('materializes bar counts per category', { skip }, () => {
  const data = dataFor('Chart1');
  const byLabel = Object.fromEntries(data.categories.map((category, index) => [category.label, data.series[0].points[index].y]));
  assert.equal(byLabel['Not Assessed'], 25);
  assert.equal(byLabel['Lost time injury'], 15);
  assert.equal(byLabel['Fatality'], 4);
});

test('materializes pie percentages and per-slice colours', { skip }, () => {
  const data = dataFor('Chart2');
  const total = data.series[0].points.reduce((sum, point) => sum + (point.y || 0), 0);
  assert.equal(total, 100); // 3 + 18 + 28 + 30 + 21
  const open = data.categories.findIndex((category) => category.label === 'Open');
  assert.equal(data.series[0].points[open].color, '#00ced1'); // DarkTurquoise
});

test('materializes clustered series with correct per-series colours', { skip }, () => {
  const data = dataFor('Chart4');
  const byLabel = Object.fromEntries(data.legend.map((entry) => [entry.label, entry.color]));
  assert.equal(byLabel['Flooding'], '#800080'); // Purple
  assert.equal(byLabel['Information Technology'], '#add8e6'); // LightBlue
  assert.equal(byLabel['Not Assessed'], '#ff0000'); // Red
});

test('draws every supported chart type (incl. scatter and stacked variants) to a valid PDF', { skip }, async () => {
  const datasets = normalizeDatasets(model, request);
  const drawn = (chartType, sourceName, stacked) => new Promise((resolve, reject) => {
    const source = chart(sourceName);
    const data = materializeChart(source, datasets, request.parameters, {});
    const doc = new PDFDocument({ autoFirstPage: false });
    const chunks = [];
    doc.on('data', (piece) => chunks.push(piece));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);
    doc.addPage({ size: [400, 300], margins: { top: 0, left: 0, right: 0, bottom: 0 } });
    drawChart(doc, config, { ...source, chartType, stacked }, data, 0, 0, 400, 300, { parameters: request.parameters, globals: {}, datasets });
    doc.end();
  });
  const cases = [
    ['doughnut', 'Chart2', 'none'], ['line', 'Chart4', 'none'], ['area', 'Chart4', 'none'],
    ['scatter', 'Chart4', 'none'], ['column', 'Chart4', 'stacked'], ['column', 'Chart4', 'percent'],
    ['bar', 'Chart1', 'stacked'], ['area', 'Chart4', 'stacked'],
  ];
  for (const [chartType, sourceName, stacked] of cases) {
    const buffer = await drawn(chartType, sourceName, stacked);
    assert.equal(buffer.subarray(0, 4).toString(), '%PDF', `${chartType}/${stacked} did not render a PDF`);
    assert.ok(buffer.length > 500);
  }
});

test('renders the incident dashboard to a selectable PDF with chart labels', { skip }, async (context) => {
  const result = await renderPdf(model, request, config);
  assert.equal(result.buffer.subarray(0, 4).toString(), '%PDF');
  assert.ok(result.pageCount > 3);
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-chart-pdf-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, 'report.pdf');
  await fs.writeFile(pdfPath, result.buffer);
  const extracted = await execFileAsync('pdftotext', [pdfPath, '-']);
  assert.match(extracted.stdout, /Breakdown By Classification/); // chart title band (textbox)
  assert.match(extracted.stdout, /18%/); // pie data label
});

test('embeds one image per chart in the editable DOCX, scaled to fit the page and landscape', { skip }, async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-chart-docx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const result = await renderEditableDocx(model, request, config, tempDir);
  const zip = await JSZip.loadAsync(result.buffer);
  const media = Object.keys(zip.files).filter((name) => /^word\/media\/.*\.png$/.test(name));
  assert.equal(media.length, 3);

  const documentXml = await zip.file('word/document.xml').async('string');
  // Orientation must be declared landscape (matching the landscape page dimensions).
  assert.match(documentXml, /<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"\/>/);
  // Every embedded chart image must fit inside the usable page body so it cannot spill onto a blank page.
  const page = model.page;
  const usableWidthPt = page.width - page.marginLeft - page.marginRight;
  const usableHeightPt = page.height - page.marginTop - (page.header?.height || 0) - page.marginBottom - (page.footer?.height || 0);
  const extents = [...documentXml.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"/g)];
  assert.equal(extents.length, 3);
  for (const [, cx, cy] of extents) {
    const widthPt = Number(cx) / 914400 * 72;
    const heightPt = Number(cy) / 914400 * 72;
    assert.ok(widthPt <= usableWidthPt + 1, `chart width ${widthPt}pt exceeds usable ${usableWidthPt}pt`);
    assert.ok(heightPt <= usableHeightPt + 1, `chart height ${heightPt}pt exceeds usable ${usableHeightPt}pt`);
  }
});
