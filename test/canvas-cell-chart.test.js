import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { PNG } from 'pngjs';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

// A chart on an SSRS List / canvas cell is scoped to its group instance: it sees that instance's rows, not
// the whole dataset. Two things used to break that outside the PDF:
//
//  * Word rasterizes charts from the canonical trace, but the trace carried only the chart's geometry, so
//    the Word renderer re-materialized it from the report-level datasets — one chart, showing every
//    category in the report, in place of the per-instance ones.
//  * Excel embedded only BODY-level charts, so a report that puts every chart inside a List cell produced
//    no chart at all.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000' });

// Each group owns categories nobody else has, so a report-wide chart is unmistakable: it would carry every
// category instead of the two (or one) the instance actually has.
const rows = [
  { G: 'Alpha', Cat: 'ALPHA_ONE', Val: 3 },
  { G: 'Alpha', Cat: 'ALPHA_TWO', Val: 5 },
  { G: 'Beta', Cat: 'BETA_ONLY', Val: 7 },
];
const GROUPS = 2;

const rdl = `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="G"><DataField>G</DataField></Field>
    <Field Name="Cat"><DataField>Cat</DataField></Field>
    <Field Name="Val"><DataField>Val</DataField></Field>
  </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="List1"><TablixBody>
      <TablixColumns><TablixColumn><Width>400pt</Width></TablixColumn></TablixColumns>
      <TablixRows><TablixRow><Height>220pt</Height><TablixCells><TablixCell><CellContents>
        <Rectangle Name="Canvas"><ReportItems>
          <Textbox Name="Title"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!G.Value</Value>
            <Style><FontFamily>Arial</FontFamily><FontSize>10pt</FontSize></Style></TextRun></TextRuns></Paragraph></Paragraphs>
            <Top>0pt</Top><Left>0pt</Left><Height>16pt</Height><Width>200pt</Width><Style/></Textbox>
          <Textbox Name="Subtitle"><Paragraphs><Paragraph><TextRuns><TextRun><Value>SUBTITLE_LINE</Value>
            <Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style></TextRun></TextRuns></Paragraph></Paragraphs>
            <Top>0pt</Top><Left>210pt</Left><Height>16pt</Height><Width>180pt</Width><Style/></Textbox>
          <Line Name="Rule"><Top>20pt</Top><Left>0pt</Left><Height>0pt</Height><Width>400pt</Width>
            <Style><Border><Style>Solid</Style></Border></Style></Line>
          <Chart Name="GroupChart">
            <ChartCategoryHierarchy><ChartMembers><ChartMember>
              <Group Name="CatG"><GroupExpressions><GroupExpression>=Fields!Cat.Value</GroupExpression></GroupExpressions></Group>
              <Label>=Fields!Cat.Value</Label></ChartMember></ChartMembers></ChartCategoryHierarchy>
            <ChartSeriesHierarchy><ChartMembers><ChartMember><Label>Val</Label></ChartMember></ChartMembers></ChartSeriesHierarchy>
            <ChartData><ChartSeriesCollection><ChartSeries Name="Val"><ChartDataPoints><ChartDataPoint>
              <ChartDataPointValues><Y>=Sum(Fields!Val.Value)</Y></ChartDataPointValues>
              <ChartDataLabel><UseValueAsLabel>true</UseValueAsLabel><Visible>true</Visible>
                <Style><FontFamily>Arial</FontFamily><FontSize>8pt</FontSize></Style></ChartDataLabel>
              <Style/></ChartDataPoint></ChartDataPoints><Type>Shape</Type><Subtype>Pie</Subtype>
            </ChartSeries></ChartSeriesCollection></ChartData>
            <ChartAreas><ChartArea Name="Default"><Style/></ChartArea></ChartAreas>
            <ChartLegends><ChartLegend Name="Default"><Position>RightTop</Position>
              <Style><FontFamily>Arial</FontFamily><FontSize>8pt</FontSize></Style></ChartLegend></ChartLegends>
            <DataSetName>D</DataSetName><Top>30pt</Top><Left>10pt</Left><Width>360pt</Width><Height>170pt</Height><Style/>
          </Chart>
        </ReportItems><Top>0pt</Top><Left>0pt</Left><Width>400pt</Width><Height>220pt</Height><Style/></Rectangle>
      </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
    </TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers><TablixMember>
      <Group Name="G"><GroupExpressions><GroupExpression>=Fields!G.Value</GroupExpression></GroupExpressions></Group>
    </TablixMember></TablixMembers></TablixRowHierarchy>
    <DataSetName>D</DataSetName><Top>0pt</Top><Left>0pt</Left><Height>220pt</Height><Width>400pt</Width><Style/>
    </Tablix>
  </ReportItems><Height>440pt</Height><Style/></Body><Width>410pt</Width>
  <Page><PageWidth>460pt</PageWidth><PageHeight>620pt</PageHeight><TopMargin>10pt</TopMargin>
    <BottomMargin>10pt</BottomMargin><LeftMargin>10pt</LeftMargin><RightMargin>10pt</RightMargin></Page>
  </ReportSection></ReportSections>
</Report>`;

const request = { outputFileName: 'canvas-cell-chart', parameters: {}, datasets: { D: rows }, excel: { layoutMode: 'REPORT' } };

// A signature of which palette colours a chart picture uses and how much of each: two charts drawn from
// different category sets cannot share one.
function pictureSignature(buffer) {
  const png = PNG.sync.read(buffer);
  const counts = new Map();
  for (let index = 0; index < png.data.length; index += 4) {
    if (png.data[index + 3] === 0) continue;
    const hex = `#${[0, 1, 2].map((offset) => png.data[index + offset].toString(16).padStart(2, '0')).join('')}`;
    if (hex === '#ffffff') continue;
    counts.set(hex, (counts.get(hex) || 0) + 1);
  }
  return [...counts.entries()].filter(([, count]) => count > 200).sort((left, right) => right[1] - left[1])
    .map(([hex]) => hex).join(',');
}

test('the canonical trace carries each chart in its own group scope', async () => {
  const rendered = await renderPdf(parseRdl(rdl), request, config, { captureLayoutTrace: true });
  const charts = rendered.layoutTrace.pages
    .flatMap((page) => (page.items || []).filter((item) => item.kind === 'chart'));
  assert.equal(charts.length, GROUPS, 'one chart per group instance');
  for (const chart of charts) {
    assert.ok(chart.chartData, 'the trace must carry the resolved series, not just the geometry');
  }
  const legends = charts.map((chart) => chart.chartData.legend.map((entry) => entry.label).sort().join(','));
  assert.deepEqual(legends.sort(), ['ALPHA_ONE,ALPHA_TWO', 'BETA_ONLY']);
});

test('editable DOCX embeds the per-instance chart, not one report-wide chart', async (context) => {
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-canvas-chart-'));
  context.after(() => fs.rm(tempRoot, { recursive: true, force: true }));
  const owned = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000', RDL_TEMP_ROOT: tempRoot });
  const docx = await renderEditableDocx(parseRdl(rdl), request, owned);
  const zip = await JSZip.loadAsync(docx.buffer);
  const media = Object.keys(zip.files).filter((name) => /^word\/media\/.*\.png$/i.test(name));
  assert.equal(media.length, GROUPS, `expected one chart picture per group instance, got ${media.length}`);
  const signatures = new Set();
  for (const name of media) signatures.add(pictureSignature(await zip.file(name).async('nodebuffer')));
  assert.equal(signatures.size, GROUPS, 'each instance must draw its own categories, not the whole report');
});

test('XLSX embeds a chart that lives on a tablix cell canvas', async (context) => {
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-canvas-chart-xl-'));
  context.after(() => fs.rm(tempRoot, { recursive: true, force: true }));
  const owned = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000', RDL_TEMP_ROOT: tempRoot });
  const xlsx = await renderExcel(parseRdl(rdl), request, owned, tempRoot);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(xlsx.buffer);
  const pictures = workbook.model.media.filter((entry) => entry.extension === 'png');
  assert.equal(pictures.length, GROUPS, `expected one chart picture per group instance, got ${pictures.length}`);
  assert.equal(
    new Set(pictures.map((entry) => pictureSignature(entry.buffer))).size,
    GROUPS,
    'each instance must draw its own categories in Excel too',
  );
  // Anchored over the chart's own region rather than the whole cell.
  const anchors = workbook.worksheets[0].getImages();
  assert.equal(anchors.length, GROUPS);
  for (const anchor of anchors) {
    assert.ok(anchor.range.br.nativeRow > anchor.range.tl.nativeRow, 'the picture must span rows');
    assert.ok(anchor.range.br.nativeCol > anchor.range.tl.nativeCol, 'the picture must span columns');
  }
});

test('XLSX places each canvas textbox at its own coordinates instead of joining them', async (context) => {
  // Joining a canvas cell's children into one cell value put a whole page of headings and prose into a
  // single merged cell — present in the workbook, invisible in it.
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-canvas-text-'));
  context.after(() => fs.rm(tempRoot, { recursive: true, force: true }));
  const owned = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000', RDL_TEMP_ROOT: tempRoot });
  const xlsx = await renderExcel(parseRdl(rdl), request, owned, tempRoot);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(xlsx.buffer);
  const sheet = workbook.worksheets[0];

  const found = new Map();
  sheet.eachRow((row, rowNumber) => row.eachCell((cell) => {
    const value = typeof cell.value === 'object' && cell.value?.richText
      ? cell.value.richText.map((run) => run.text).join('')
      : String(cell.value ?? '');
    const text = value.trim();
    if (!text) return;
    const master = cell.master ? cell.master.address : cell.address;
    if (!found.has(master)) found.set(master, { text, rowNumber });
  }));
  const titles = [...found.values()].filter((entry) => entry.text === 'Alpha' || entry.text === 'Beta');
  const subtitles = [...found.values()].filter((entry) => entry.text === 'SUBTITLE_LINE');
  assert.equal(subtitles.length, GROUPS, 'the sibling textbox is placed separately, not appended to the title');
  assert.equal(titles.length, GROUPS, 'each instance writes its own canvas title');
  for (const entry of titles) {
    assert.equal(entry.text.includes('\n'), false, 'a canvas child must not be joined with its siblings');
  }
  // The two instances' titles sit on different rows, in report order.
  assert.equal(new Set(titles.map((entry) => entry.rowNumber)).size, GROUPS);
});
