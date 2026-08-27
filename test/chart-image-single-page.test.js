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
import { renderChartPng } from '../src/render/chartImage.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';

// A chart's strings are absolutely positioned, never flowed. PDFKit, however, treats a string that reaches
// the bottom of the page as overflowing body copy and starts a NEW PAGE — carrying that label and
// everything drawn after it with it. Word and Excel embed each chart as a picture rendered onto a one-page
// document exactly the size of the chart, and only page 1 is rasterized, so the spill silently deleted
// whole slices from the picture while the same chart on a tall PDF report page looked correct. That is the
// class of "Word/Excel differ from the PDF" defect this file guards.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000' });
const PALETTE = ['#01b8aa', '#374649'];

const chartXml = ({ subtype, width, height }) => `
<Chart Name="Shape">
  <ChartCategoryHierarchy><ChartMembers><ChartMember>
    <Group Name="Cat"><GroupExpressions><GroupExpression>=Fields!Category.Value</GroupExpression></GroupExpressions></Group>
    <SortExpressions><SortExpression><Value>=Fields!Category.Value</Value></SortExpression></SortExpressions>
    <Label>=Fields!Category.Value</Label>
  </ChartMember></ChartMembers></ChartCategoryHierarchy>
  <ChartSeriesHierarchy><ChartMembers><ChartMember><Label>Amount</Label></ChartMember></ChartMembers></ChartSeriesHierarchy>
  <ChartData><ChartSeriesCollection><ChartSeries Name="Amount"><ChartDataPoints><ChartDataPoint>
    <ChartDataPointValues><Y>=Sum(Fields!Amount.Value)</Y></ChartDataPointValues>
    <ChartDataLabel><Label>=Sum(Fields!Amount.Value) &amp; " (" &amp; FormatPercent(0.5, 1) &amp; ")"</Label>
      <Visible>true</Visible><Position>Outside</Position>
      <Style><FontFamily>Arial</FontFamily><FontSize>8pt</FontSize></Style></ChartDataLabel>
    <Style/>
  </ChartDataPoint></ChartDataPoints><Type>Shape</Type><Subtype>${subtype}</Subtype>
  </ChartSeries></ChartSeriesCollection></ChartData>
  <ChartAreas><ChartArea Name="Default"><Style/></ChartArea></ChartAreas>
  <ChartLegends><ChartLegend Name="Default"><Position>RightTop</Position>
    <Style><FontFamily>Arial</FontFamily><FontSize>8pt</FontSize></Style></ChartLegend></ChartLegends>
  <Palette>Custom</Palette><ChartCustomPaletteColors>
    ${PALETTE.map((entry) => `<ChartCustomPaletteColor>${entry}</ChartCustomPaletteColor>`).join('')}
  </ChartCustomPaletteColors>
  <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Width>${width}pt</Width><Height>${height}pt</Height><Style/>
</Chart>`;

const reportRdl = (chart) => `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <Body><ReportItems>${chart}</ReportItems><Height>4in</Height></Body><Width>7in</Width>
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="Category"><DataField>Category</DataField></Field>
    <Field Name="Amount"><DataField>Amount</DataField></Field>
  </Fields></DataSet></DataSets>
  <Page><PageWidth>7.2in</PageWidth><PageHeight>4.2in</PageHeight><TopMargin>0.1in</TopMargin>
    <BottomMargin>0.1in</BottomMargin><LeftMargin>0.1in</LeftMargin><RightMargin>0.1in</RightMargin></Page>
</Report>`;

const request = {
  outputFileName: 'chart-image-single-page',
  parameters: {},
  datasets: { D: [{ Category: 'Alpha', Amount: 1 }, { Category: 'Beta', Amount: 1 }] },
};

function chartOf(rdlText) {
  const model = parseRdl(rdlText);
  const chart = model.body.items.find((item) => item.type === 'Chart');
  const datasets = normalizeDatasets(model, request);
  return { model, chart, datasets, data: materializeChart(chart, datasets, {}, {}) };
}

// Draws the chart exactly as `renderChartPng` does — onto a page that IS the chart — and counts pages.
function pagesUsed({ chart, data, datasets }) {
  const doc = new PDFDocument({ autoFirstPage: false });
  let pages = 0;
  doc.on('pageAdded', () => { pages += 1; });
  doc.addPage({ size: [chart.width, chart.height], margins: { top: 0, right: 0, bottom: 0, left: 0 } });
  drawChart(doc, config, chart, data, 0, 0, chart.width, chart.height, { parameters: {}, globals: {}, datasets });
  doc.end();
  return pages;
}

test('a chart never paginates the one-page document it is drawn onto', () => {
  // A short shape chart puts its bottom outside label hard against the page edge — the exact condition
  // that used to push the remaining slices onto a discarded second page.
  for (const subtype of ['Pie', 'Doughnut', 'ExplodedPie', 'ExplodedDoughnut']) {
    for (const height of [30, 32, 34, 40, 60, 120, 230]) {
      for (const width of [120, 335, 480]) {
        const context = chartOf(reportRdl(chartXml({ subtype, width, height })));
        assert.equal(pagesUsed(context), 1, `${subtype} ${width}x${height} left its own page`);
      }
    }
  }
});

function paletteUse(buffer) {
  const png = PNG.sync.read(buffer);
  const counts = new Map();
  for (let index = 0; index < png.data.length; index += 4) {
    if (png.data[index + 3] === 0) continue;
    const hex = `#${[0, 1, 2].map((offset) => png.data[index + offset].toString(16).padStart(2, '0')).join('')}`;
    if (PALETTE.includes(hex)) counts.set(hex, (counts.get(hex) || 0) + 1);
  }
  return counts;
}

test('the chart picture keeps every slice, so Word and Excel match the PDF', async (context) => {
  const rdlText = reportRdl(chartXml({ subtype: 'Pie', width: 335, height: 32 }));
  const { model, chart, data, datasets } = chartOf(rdlText);
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-chart-page-'));
  context.after(() => fs.rm(tempRoot, { recursive: true, force: true }));
  const ownedConfig = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000', RDL_TEMP_ROOT: tempRoot });

  const png = await renderChartPng(chart, data, ownedConfig, tempRoot, { parameters: {}, globals: {}, datasets }, 0);
  const direct = paletteUse(png.data);
  // Two equal slices: both palette colours must be present, in comparable amounts. A dropped slice showed
  // up as one colour holding the whole shape and the other reduced to its legend swatch.
  assert.equal(direct.size, 2, 'the chart picture must contain both slices');
  const [first, second] = [...direct.values()].sort((left, right) => right - left);
  assert.ok(second > first * 0.5, `both slices must be drawn (got ${first} vs ${second} pixels)`);

  const docx = await renderEditableDocx(model, request, ownedConfig);
  const zip = await JSZip.loadAsync(docx.buffer);
  const media = Object.keys(zip.files).filter((name) => /^word\/media\/.*\.png$/.test(name));
  assert.equal(media.length, 1);
  const docxCounts = paletteUse(await zip.file(media[0]).async('nodebuffer'));
  assert.equal(docxCounts.size, 2, 'editable DOCX embedded an incomplete chart');

  const xlsx = await renderExcel(model, { ...request, excel: { layoutMode: 'REPORT' } }, ownedConfig, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(xlsx.buffer);
  assert.equal(workbook.model.media.length, 1);
  assert.equal(paletteUse(workbook.model.media[0].buffer).size, 2, 'XLSX embedded an incomplete chart');
});

test('editable Word keeps a Rectangle border visible above a chart that fills the container', async (context) => {
  const chart = chartXml({ subtype: 'Pie', width: 240, height: 160 });
  const rdlText = reportRdl(`
    <Rectangle Name="Frame"><ReportItems>${chart}</ReportItems>
      <Top>0pt</Top><Left>0pt</Left><Width>240pt</Width><Height>160pt</Height>
      <Style><Border><Color>#112233</Color><Style>Solid</Style><Width>2pt</Width></Border></Style>
    </Rectangle>`);
  const model = parseRdl(rdlText);
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-chart-frame-docx-'));
  context.after(() => fs.rm(tempRoot, { recursive: true, force: true }));
  const ownedConfig = loadConfig({
    ...process.env,
    RDL_STRICT_FONTS: 'false',
    RDL_RENDER_TIMEOUT_MS: '60000',
    RDL_TEMP_ROOT: tempRoot,
  });

  const docx = await renderEditableDocx(model, request, ownedConfig);
  const zip = await JSZip.loadAsync(docx.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  // Word floats drawings above native cell borders. The converter must fit the picture inside the four
  // 2pt container edges, leaving the parent frame visible just as SSRS paints it above its children.
  const extent = documentXml.match(/<wp:extent cx="(\d+)" cy="(\d+)"/);
  assert.ok(extent, 'the chart must remain a native Word drawing');
  assert.ok(Number(extent[1]) / 12_700 <= 236, 'drawing must stay inside the 2pt left/right frame edges');
  assert.ok(Number(extent[2]) / 12_700 <= 156, 'drawing must stay inside the 2pt top/bottom frame edges');
  assert.match(documentXml, /<w:bottom w:val="single" w:color="112233" w:sz="16"\/>/);
  const offsets = [...documentXml.matchAll(/<wp:posOffset>(\d+)<\/wp:posOffset>/g)].map((match) => Number(match[1]));
  assert.ok(offsets.some((offset) => offset >= 25_400),
    'the chart must start inside the 2pt left/top frame edge');
});
