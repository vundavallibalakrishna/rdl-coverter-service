import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import PDFDocument from 'pdfkit';
import { PNG } from 'pngjs';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { normalizeDatasets } from '../src/render/common.js';
import { materializeChart } from '../src/render/chartData.js';
import { drawChart } from '../src/render/chart.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';
import { rasterizePdf } from '../src/render/raster.js';

// SSRS calls a data point whose value is Nothing/null an EMPTY POINT. On a shape chart (pie/doughnut) it
// draws no slice, takes no legend entry, and consumes no palette colour, so the points after it keep the
// colours they would have had without it. SSRS also draws an outside shape-chart label as a callout: a
// radial stub off the slice edge, then a horizontal elbow, with the label starting at the elbow.
//
// Everything here is a synthetic minimal RDL isolating those constructs.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000' });
const PALETTE = ['#01b8aa', '#374649', '#fd625e'];

// The Y expression is the SSRS idiom that produces an empty point: a value below the threshold becomes
// Nothing rather than zero.
const chartXml = ({ name, type, subtype, labelPosition, y, label }) => `
<Chart Name="${name}">
  <ChartCategoryHierarchy><ChartMembers><ChartMember>
    <Group Name="${name}Cat"><GroupExpressions><GroupExpression>=Fields!Category.Value</GroupExpression></GroupExpressions></Group>
    <SortExpressions><SortExpression><Value>=Fields!Category.Value</Value></SortExpression></SortExpressions>
    <Label>=Fields!Category.Value</Label>
  </ChartMember></ChartMembers></ChartCategoryHierarchy>
  <ChartSeriesHierarchy><ChartMembers><ChartMember><Label>Amount</Label></ChartMember></ChartMembers></ChartSeriesHierarchy>
  <ChartData><ChartSeriesCollection><ChartSeries Name="Amount"><ChartDataPoints><ChartDataPoint>
    <ChartDataPointValues><Y>${y}</Y></ChartDataPointValues>
    <ChartDataLabel>${label ? `<Label>${label}</Label>` : '<UseValueAsLabel>true</UseValueAsLabel>'}<Visible>true</Visible>
      ${labelPosition ? `<Position>${labelPosition}</Position>` : ''}
      <Style><FontFamily>Arial</FontFamily><FontSize>8pt</FontSize></Style></ChartDataLabel>
    <Style/>
  </ChartDataPoint></ChartDataPoints><Type>${type}</Type>${subtype ? `<Subtype>${subtype}</Subtype>` : ''}
  </ChartSeries></ChartSeriesCollection></ChartData>
  <ChartAreas><ChartArea Name="Default"><Style/></ChartArea></ChartAreas>
  <ChartLegends><ChartLegend Name="Default"><Position>RightTop</Position>
    <Style><FontFamily>Arial</FontFamily><FontSize>8pt</FontSize></Style></ChartLegend></ChartLegends>
  <Palette>Custom</Palette><ChartCustomPaletteColors>
    ${PALETTE.map((entry) => `<ChartCustomPaletteColor>${entry}</ChartCustomPaletteColor>`).join('')}
  </ChartCustomPaletteColors>
  <DataSetName>D</DataSetName><Top>0.1in</Top><Left>0.1in</Left><Width>5in</Width><Height>3in</Height>
  <Style><Border><Style>Solid</Style></Border></Style>
</Chart>`;

const reportRdl = (chart) => `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <Body><ReportItems>${chart}</ReportItems><Height>3.2in</Height></Body><Width>5.2in</Width>
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="Category"><DataField>Category</DataField></Field>
    <Field Name="Amount"><DataField>Amount</DataField></Field>
  </Fields></DataSet></DataSets>
  <Page><PageWidth>5.4in</PageWidth><PageHeight>3.4in</PageHeight><TopMargin>0.1in</TopMargin>
    <BottomMargin>0.1in</BottomMargin><LeftMargin>0.1in</LeftMargin><RightMargin>0.1in</RightMargin></Page>
</Report>`;

// "Alpha" is below the threshold, so its Y is Nothing — an empty point. "Beta" and "Gamma" are real.
const EMPTY_POINT_Y = '=IIF(Sum(Fields!Amount.Value) &gt; 0, Sum(Fields!Amount.Value), Nothing)';
const PLAIN_Y = '=Sum(Fields!Amount.Value)';
const rows = [{ Category: 'Alpha', Amount: 0 }, { Category: 'Beta', Amount: 1 }, { Category: 'Gamma', Amount: 1 }];
const request = { outputFileName: 'chart-empty-point', parameters: {}, datasets: { D: rows } };

function dataFor(rdlText) {
  const model = parseRdl(rdlText);
  const chart = model.body.items.find((item) => item.type === 'Chart');
  return { model, chart, data: materializeChart(chart, normalizeDatasets(model, request), {}, {}) };
}

test('a shape-chart empty point takes no legend entry and no palette colour', () => {
  for (const subtype of ['Pie', 'ExplodedDoughnut']) {
    const { chart, data } = dataFor(reportRdl(chartXml({
      name: 'Shape', type: 'Shape', subtype, labelPosition: 'Outside', y: EMPTY_POINT_Y,
    })));
    assert.equal(chart.chartType, subtype === 'Pie' ? 'pie' : 'doughnut');
    assert.equal(data.series[0].points[0].y, null, 'Alpha must materialize as an empty point');
    assert.equal(data.series[0].points[0].empty, true);
    assert.deepEqual(data.legend.map((entry) => entry.label), ['Beta', 'Gamma'], `${subtype} legend`);
    // The two rendered points take the FIRST two palette colours: the empty point consumed none.
    assert.deepEqual(data.legend.map((entry) => entry.color), [PALETTE[0], PALETTE[1]]);
    assert.equal(data.series[0].points[1].color, PALETTE[0]);
    assert.equal(data.series[0].points[2].color, PALETTE[1]);
  }
});

test('a zero is a real value, not an empty point: it keeps its legend entry and colour', () => {
  const { data } = dataFor(reportRdl(chartXml({
    name: 'Shape', type: 'Shape', subtype: 'Pie', labelPosition: 'Outside', y: PLAIN_Y,
  })));
  assert.equal(data.series[0].points[0].y, 0);
  assert.equal(data.series[0].points[0].empty, false);
  assert.deepEqual(data.legend.map((entry) => entry.label), ['Alpha', 'Beta', 'Gamma']);
  assert.deepEqual(data.legend.map((entry) => entry.color), PALETTE);
});

test('a non-shape chart keeps its category legend and colours across an empty point', () => {
  // A column chart's legend/palette is not a per-point sequence, so an empty cell only leaves a gap in
  // the series' own points — the categories after it must not shift colour.
  const { chart, data } = dataFor(reportRdl(chartXml({
    name: 'Columns', type: 'Column', subtype: '', labelPosition: null, y: EMPTY_POINT_Y,
  })));
  assert.equal(chart.chartType, 'column');
  assert.equal(data.series[0].points[0].y, null);
  assert.deepEqual(data.legend.map((entry) => entry.label), ['Alpha', 'Beta', 'Gamma']);
  assert.deepEqual(data.legend.map((entry) => entry.color), PALETTE);
});

// Records the resolved PDFKit path and text operations for one chart draw.
function captureChart(rdlText) {
  const { model, chart, data } = dataFor(rdlText);
  const doc = new PDFDocument({ autoFirstPage: false });
  doc.addPage({ size: [chart.width + 20, chart.height + 20], margins: { top: 0, left: 0, right: 0, bottom: 0 } });
  const paths = [];
  const texts = [];
  let current = null;
  const originalMoveTo = doc.moveTo.bind(doc);
  const originalLineTo = doc.lineTo.bind(doc);
  const originalStroke = doc.stroke.bind(doc);
  const originalText = doc.text.bind(doc);
  doc.moveTo = (x, y) => { current = [[x, y]]; return originalMoveTo(x, y); };
  doc.lineTo = (x, y) => { if (current) current.push([x, y]); return originalLineTo(x, y); };
  doc.stroke = (...args) => { if (current) { paths.push(current); current = null; } return originalStroke(...args); };
  doc.text = (value, x, y, options) => {
    texts.push({ value: String(value), x, y, options, width: doc.widthOfString(String(value)) });
    return originalText(value, x, y, options);
  };
  drawChart(doc, config, chart, data, 10, 10, chart.width, chart.height, { parameters: {}, globals: {}, datasets: normalizeDatasets(model, request) });
  return { chart, data, paths, texts };
}

test('an outside shape-chart label is drawn as a radial stub plus a horizontal elbow', () => {
  const { paths, texts } = captureChart(reportRdl(chartXml({
    name: 'Shape', type: 'Shape', subtype: 'Pie', labelPosition: 'Outside', y: EMPTY_POINT_Y,
  })));
  // Two rendered slices, so two callouts. A callout is the only three-point polyline the pie draws.
  const callouts = paths.filter((points) => points.length === 3);
  assert.equal(callouts.length, 2, 'one callout per rendered slice');
  for (const [start, elbow, end] of callouts) {
    assert.ok(Math.abs(elbow[1] - end[1]) < 0.01, 'the second segment must be horizontal');
    assert.ok(Math.abs(end[0] - elbow[0]) > 1, 'the elbow must have a real horizontal run');
    // The stub leaves the slice edge along the bisector, so it is never zero-length.
    assert.ok(Math.hypot(elbow[0] - start[0], elbow[1] - start[1]) > 1, 'the radial stub must have length');
  }
  // Two slices of equal size give vertical bisectors. Their cosine is ±1e-16, and a floating-point sign
  // must not put one label on the left and the other on the right.
  const directions = callouts.map(([, elbow, end]) => Math.sign(end[0] - elbow[0]));
  assert.deepEqual(directions, [1, 1], 'a straight-up/straight-down callout points right on both slices');
  const labels = texts.filter((entry) => /^\d/.test(entry.value));
  assert.ok(labels.length >= 2);
  for (const label of labels) {
    assert.ok(
      label.options.width >= label.width,
      `label "${label.value}" must be given at least its measured width (${label.options.width} < ${label.width})`,
    );
  }
});

test('an inside shape-chart label is given its measured width so it never wraps', () => {
  // A value-plus-percentage label is the common SSRS idiom and is far wider than a bare number; a fixed
  // label box broke it across two lines, which SSRS never does.
  const { texts } = captureChart(reportRdl(chartXml({
    name: 'Shape',
    type: 'Shape',
    subtype: 'ExplodedDoughnut',
    labelPosition: null,
    y: EMPTY_POINT_Y,
    label: '=Sum(Fields!Amount.Value) &amp; " (" &amp; FormatPercent(0.5, 1) &amp; ")"',
  })));
  const labels = texts.filter((entry) => /\(\d/.test(entry.value));
  assert.equal(labels.length, 2, 'the doughnut draws one label per rendered slice');
  for (const label of labels) {
    assert.ok(label.width > 30, 'the fixture label must be wide enough to expose a fixed-width label box');
    assert.ok(label.options.width >= label.width, `inside label "${label.value}" was given a narrower box than its text`);
  }
});

test('the whole outside callout stays inside the chart rectangle', () => {
  const { chart, paths } = captureChart(reportRdl(chartXml({
    name: 'Shape', type: 'Shape', subtype: 'Pie', labelPosition: 'Outside', y: EMPTY_POINT_Y,
  })));
  const callouts = paths.filter((points) => points.length === 3);
  for (const points of callouts) {
    for (const [x, y] of points) {
      assert.ok(x >= 10 && x <= 10 + chart.width, `callout x ${x} left the chart`);
      assert.ok(y >= 10 && y <= 10 + chart.height, `callout y ${y} left the chart`);
    }
  }
});

async function paletteUseInPng(buffer) {
  const png = PNG.sync.read(buffer);
  const seen = new Set();
  for (let index = 0; index < png.data.length; index += 4) {
    if (png.data[index + 3] === 0) continue;
    const hex = `#${[0, 1, 2].map((offset) => png.data[index + offset].toString(16).padStart(2, '0')).join('')}`;
    if (PALETTE.includes(hex)) seen.add(hex);
  }
  return seen;
}

test('PDF, editable DOCX, and XLSX all paint the empty point out of the palette sequence', async (context) => {
  const rdlText = reportRdl(chartXml({
    name: 'Shape', type: 'Shape', subtype: 'Pie', labelPosition: 'Outside', y: EMPTY_POINT_Y,
  }));
  const model = parseRdl(rdlText);
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-chart-empty-'));
  context.after(() => fs.rm(tempRoot, { recursive: true, force: true }));
  const ownedConfig = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000', RDL_TEMP_ROOT: tempRoot });

  const pdf = await renderPdf(model, request, ownedConfig);
  const rasterDir = await fs.mkdtemp(path.join(tempRoot, 'raster-'));
  const [page] = await rasterizePdf(pdf.buffer, ownedConfig, rasterDir, 'chart', { dpi: 150, singleFile: true });
  const pdfColours = await paletteUseInPng(page.data);
  assert.ok(pdfColours.has(PALETTE[0]) && pdfColours.has(PALETTE[1]), 'PDF must paint the first two palette colours');
  assert.ok(!pdfColours.has(PALETTE[2]), 'PDF must not reach the third palette colour for two rendered points');

  const docx = await renderEditableDocx(model, request, ownedConfig);
  const zip = await JSZip.loadAsync(docx.buffer);
  const media = Object.keys(zip.files).filter((name) => /^word\/media\/.*\.png$/.test(name));
  assert.equal(media.length, 1, 'the chart is one embedded picture');
  const docxColours = await paletteUseInPng(await zip.file(media[0]).async('nodebuffer'));
  assert.ok(docxColours.has(PALETTE[0]) && docxColours.has(PALETTE[1]));
  assert.ok(!docxColours.has(PALETTE[2]), 'editable DOCX inherits the same palette sequence');

  const xlsx = await renderExcel(model, { ...request, excel: { layoutMode: 'REPORT' } }, ownedConfig, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(xlsx.buffer);
  assert.equal(workbook.model.media.length, 1, 'XLSX embeds the same chart image');
  const xlsxColours = await paletteUseInPng(workbook.model.media[0].buffer);
  assert.ok(xlsxColours.has(PALETTE[0]) && xlsxColours.has(PALETTE[1]));
  assert.ok(!xlsxColours.has(PALETTE[2]), 'XLSX inherits the same palette sequence');
});
