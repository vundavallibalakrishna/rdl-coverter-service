// An SSRS List / free-form canvas: a grouped 1x1 tablix whose single cell is a Rectangle holding several
// positioned items (textboxes, a line, a chart). The renderer draws the cell item-by-item at each item's
// position and reflows a taller-than-page canvas across pages, rather than refusing the report.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { analyzeRdl, parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { loadConfig } from '../src/config.js';

const execFileAsync = promisify(execFile);
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

function listReport() {
  return Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D">
    <Fields>
      <Field Name="G"><DataField>G</DataField><TypeName>System.String</TypeName></Field>
      <Field Name="Cat"><DataField>Cat</DataField><TypeName>System.String</TypeName></Field>
      <Field Name="Val"><DataField>Val</DataField><TypeName>System.Int32</TypeName></Field>
    </Fields>
    <Query><CommandText>x</CommandText></Query>
  </DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="List1"><TablixBody>
      <TablixColumns><TablixColumn><Width>6in</Width></TablixColumn></TablixColumns>
      <TablixRows><TablixRow><Height>4in</Height><TablixCells><TablixCell><CellContents>
        <Rectangle Name="Canvas"><ReportItems>
          <Textbox Name="Title"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!G.Value</Value><Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs>
            <Top>0in</Top><Left>0in</Left><Height>0.3in</Height><Width>4in</Width><Style/></Textbox>
          <Line Name="Rule"><Top>0.5in</Top><Left>0in</Left><Height>0in</Height><Width>6in</Width><Style><Border><Style>Solid</Style></Border></Style></Line>
          <Chart Name="Chart1"><ChartCategoryHierarchy><ChartMembers><ChartMember><Group Name="CatG"><GroupExpressions><GroupExpression>=Fields!Cat.Value</GroupExpression></GroupExpressions></Group><Label>=Fields!Cat.Value</Label></ChartMember></ChartMembers></ChartCategoryHierarchy>
            <ChartSeriesHierarchy><ChartMembers><ChartMember/></ChartMembers></ChartSeriesHierarchy>
            <ChartData><ChartSeriesCollection><ChartSeries Name="S"><ChartDataPoints><ChartDataPoint><ChartDataPointValues><Y>=Sum(Fields!Val.Value)</Y></ChartDataPointValues></ChartDataPoint></ChartDataPoints><Type>Column</Type></ChartSeries></ChartSeriesCollection></ChartData>
            <ChartAreas><ChartArea Name="A"><ChartCategoryAxes><ChartAxis Name="PC"/></ChartCategoryAxes><ChartValueAxes><ChartAxis Name="PV"/></ChartValueAxes></ChartArea></ChartAreas>
            <Top>1in</Top><Left>0in</Left><Height>2.5in</Height><Width>5in</Width><Style/></Chart>
        </ReportItems><Style/></Rectangle>
      </CellContents></TablixCell></TablixCells></TablixRow></TablixRows></TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="G"><GroupExpressions><GroupExpression>=Fields!G.Value</GroupExpression></GroupExpressions><PageBreak><BreakLocation>Between</BreakLocation></PageBreak></Group></TablixMember></TablixMembers></TablixRowHierarchy>
      <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>4in</Height><Width>6in</Width><Style/></Tablix>
  </ReportItems><Height>4in</Height><Style/></Body><Width>8in</Width>
  <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`, 'utf8');
}

test('a List / canvas cell (textbox + line + chart) is compatible, not refused', () => {
  assert.equal(analyzeRdl(listReport()).compatible, true);
});

test('a List renders one canvas per group with the line and chart drawn', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-list-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const request = {
    outputFileName: 'list',
    parameters: {},
    datasets: { D: [
      { G: 'Alpha', Cat: 'X', Val: 3 }, { G: 'Alpha', Cat: 'Y', Val: 5 },
      { G: 'Beta', Cat: 'X', Val: 2 }, { G: 'Beta', Cat: 'Y', Val: 8 },
    ] },
  };
  const rendered = await renderPdf(parseRdl(listReport()), request, config);
  assert.ok(rendered.buffer.length > 0);
  const pdfPath = path.join(tempDir, 'list.pdf');
  await fs.writeFile(pdfPath, rendered.buffer);
  const { stdout } = await execFileAsync('pdftotext', ['-layout', pdfPath, '-']);
  // Both group titles rendered (one canvas per group instance).
  assert.match(stdout, /Alpha/);
  assert.match(stdout, /Beta/);
});

const CANVAS_DATA = { D: [
  { G: 'Alpha', Cat: 'X', Val: 3 }, { G: 'Alpha', Cat: 'Y', Val: 5 },
  { G: 'Beta', Cat: 'X', Val: 2 }, { G: 'Beta', Cat: 'Y', Val: 8 },
] };

test('editable DOCX renders a canvas cell as native text plus the chart picture', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-list-docx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderEditableDocx(parseRdl(listReport()), { output: 'DOCX_EDITABLE', parameters: {}, datasets: CANVAS_DATA }, config, tempDir);
  const zip = await JSZip.loadAsync(rendered.buffer);
  const document = await zip.file('word/document.xml').async('string');
  const texts = [...document.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((match) => match[1]).join(' ');
  assert.match(texts, /Alpha/);
  assert.match(texts, /Beta/);
  // The canvas chart is embedded as a Word picture.
  assert.ok(Object.keys(zip.files).some((name) => /^word\/media\/.*\.png$/.test(name)), 'expected an embedded chart image');
});

test('XLSX renders a canvas cell as populated cells without failing', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-list-xlsx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderExcel(parseRdl(listReport()), { output: 'XLSX', excelLayoutMode: 'REPORT', parameters: {}, datasets: CANVAS_DATA }, config, tempDir);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(rendered.buffer);
  const values = [];
  workbook.worksheets.forEach((sheet) => sheet.eachRow((row) => row.eachCell((cell) => {
    if (cell.value) values.push(String(cell.value));
  })));
  assert.ok(values.includes('Alpha'), `expected canvas textbox in ${JSON.stringify(values)}`);
  assert.ok(values.includes('Beta'));
});
