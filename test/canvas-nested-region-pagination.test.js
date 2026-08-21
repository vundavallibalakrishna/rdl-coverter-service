// A free-form canvas cell (a Rectangle inside a tablix cell) holds positioned peers, including child data
// regions. Two SSRS semantics are certified here:
//  1. A child data region that renders taller than one printable page continues its rows on the next page.
//     PDF used to draw it as one atomic block, straight through the footer band and off the sheet, which
//     the page-locked Word renderer then had to refuse as "outside the Word page canvas".
//  2. A child region that renders taller than its declared height displaces its later canvas peers by the
//     growth. XLSX used to schedule every child at its DESIGN top, so a grown region and the peer beneath
//     it were mapped onto the same worksheet rows and produced overlapping merged ranges.
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { execFile } from 'node:child_process';
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

// Canvas layout, in design coordinates inside the cell. The Line makes the cell an SSRS List / free-form
// canvas (content a single cell textbox cannot represent), which is the construct under test:
//   Rule    Top 0.25in Height 0in
//   Grower  Top 0.5in  Height 0.25in  (one detail row per source row, so it grows with the data)
//   Below   Top 1.0in  Height 0.25in  (a static child region well clear of Grower's declared bottom)
// Grower's declared bottom is 0.75in, so nothing overlaps until Grower actually grows.
const REPORT = `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="G"><DataField>G</DataField><TypeName>System.String</TypeName></Field>
    <Field Name="Cat"><DataField>Cat</DataField><TypeName>System.String</TypeName></Field>
    <Field Name="Val"><DataField>Val</DataField><TypeName>System.Int32</TypeName></Field>
  </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="List1"><TablixBody>
      <TablixColumns><TablixColumn><Width>7in</Width></TablixColumn></TablixColumns>
      <TablixRows><TablixRow><Height>2in</Height><TablixCells><TablixCell><CellContents>
        <Rectangle Name="Canvas"><ReportItems>
          <Line Name="Rule"><Top>0.25in</Top><Left>0in</Left><Height>0in</Height><Width>5in</Width><Style><Border><Style>Solid</Style></Border></Style></Line>
          <Tablix Name="Grower"><TablixBody>
            <TablixColumns><TablixColumn><Width>2in</Width></TablixColumn><TablixColumn><Width>3in</Width></TablixColumn></TablixColumns>
            <TablixRows><TablixRow><Height>0.25in</Height><TablixCells>
              <TablixCell><CellContents><Textbox Name="gc"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Cat.Value</Value><Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell>
              <TablixCell><CellContents><Textbox Name="gv"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Val.Value</Value><Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell>
            </TablixCells></TablixRow></TablixRows></TablixBody>
            <TablixColumnHierarchy><TablixMembers><TablixMember/><TablixMember/></TablixMembers></TablixColumnHierarchy>
            <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="GrowerDetails"/></TablixMember></TablixMembers></TablixRowHierarchy>
            <Top>0.5in</Top><Left>0in</Left><Height>0.25in</Height><Width>5in</Width><Style/></Tablix>
          <Tablix Name="Below"><TablixBody>
            <TablixColumns><TablixColumn><Width>5in</Width></TablixColumn></TablixColumns>
            <TablixRows><TablixRow><Height>0.25in</Height><TablixCells>
              <TablixCell><CellContents><Textbox Name="bt"><Paragraphs><Paragraph><TextRuns><TextRun><Value>BELOW_REGION</Value><Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell>
            </TablixCells></TablixRow></TablixRows></TablixBody>
            <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
            <TablixRowHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixRowHierarchy>
            <Top>1in</Top><Left>0in</Left><Height>0.25in</Height><Width>5in</Width><Style/></Tablix>
        </ReportItems><Style/></Rectangle>
      </CellContents></TablixCell></TablixCells></TablixRow></TablixRows></TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="G"><GroupExpressions><GroupExpression>=Fields!G.Value</GroupExpression></GroupExpressions></Group></TablixMember></TablixMembers></TablixRowHierarchy>
      <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>2in</Height><Width>7in</Width><Style/></Tablix>
  </ReportItems><Height>2in</Height><Style/></Body><Width>7.5in</Width>
  <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections></Report>`;

const model = () => parseRdl(Buffer.from(REPORT, 'utf8'));
// 60 detail rows at 0.25in exceed the 10in printable body, so Grower must continue on a second page.
const rowsFor = (count) => Array.from({ length: count }, (_, index) => ({ G: 'Alpha', Cat: `ROW_${index}`, Val: index }));
const requestFor = (count, output) => ({ output, parameters: {}, datasets: { D: rowsFor(count) } });

test('a canvas child region taller than a page continues on the next page instead of leaving the body', async () => {
  const rendered = await renderPdf(model(), requestFor(60, 'PDF'), config, { captureLayoutTrace: true });
  assert.ok(rendered.pageCount > 1, 'the grown child region must span more than one page');
  for (const page of rendered.layoutTrace.pages) {
    // The definition declares no page header or footer, so every traced primitive is body content and
    // must sit inside the printable body band.
    for (const item of page.items) {
      assert.ok(
        item.y + item.height <= page.bodyBottom + 0.5,
        `${item.itemName || item.kind} on page ${page.number} ends at ${item.y + item.height}, past the body bottom ${page.bodyBottom}`,
      );
      assert.ok(item.y >= page.bodyTop - 0.5, `${item.itemName || item.kind} on page ${page.number} starts above the body top`);
    }
  }
});

test('PDF draws every child row once and keeps the displaced peer after the grown region', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-canvas-nested-pdf-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, 'canvas-nested.pdf');
  const rendered = await renderPdf(model(), requestFor(60, 'PDF'), config);
  await fs.writeFile(pdfPath, rendered.buffer);
  const { stdout } = await execFileAsync('pdftotext', ['-layout', pdfPath, '-']);
  for (let index = 0; index < 60; index += 1) assert.match(stdout, new RegExp(`ROW_${index}\\b`));
  assert.equal((stdout.match(/BELOW_REGION/g) || []).length, 1);
  assert.ok(
    stdout.indexOf('BELOW_REGION') > stdout.indexOf('ROW_59'),
    'the peer below the grown region must render after its last row',
  );
});

test('DOCX_EDITABLE accepts the paginated canvas region and keeps the canonical page count', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-canvas-nested-docx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const [pdf, docx] = await Promise.all([
    renderPdf(model(), requestFor(60, 'PDF'), config),
    renderEditableDocx(model(), requestFor(60, 'DOCX_EDITABLE'), config, tempDir),
  ]);
  assert.equal(docx.pageCount, pdf.pageCount);
  const document = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  const texts = [...document.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((match) => match[1]);
  assert.equal(texts.filter((value) => value === 'BELOW_REGION').length, 1);
  assert.equal(texts.filter((value) => value === 'ROW_59').length, 1);
});

test('XLSX places a grown canvas region and its displaced peer without overlapping merges', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-canvas-nested-xlsx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderExcel(model(), { ...requestFor(60, 'XLSX'), excelLayoutMode: 'REPORT' }, config, tempDir);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(rendered.buffer);
  const positions = new Map();
  workbook.worksheets.forEach((sheet) => sheet.eachRow((row) => row.eachCell((cell) => {
    if (typeof cell.value === 'string' && cell.value) positions.set(cell.value, cell.row);
  })));
  for (let index = 0; index < 60; index += 1) {
    assert.ok(positions.has(`ROW_${index}`), `expected ROW_${index} in the workbook`);
  }
  assert.ok(positions.has('BELOW_REGION'), 'expected the displaced peer region in the workbook');
  assert.ok(
    positions.get('BELOW_REGION') > positions.get('ROW_59'),
    'the displaced peer must sit below the grown region rather than inside its rows',
  );
});

test('a canvas child region that does not grow leaves its peer at the declared coordinates', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-canvas-nested-flat-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  // One detail row exactly fills Grower's declared height, so displacement must contribute nothing and
  // both children stay on the single canvas page.
  const pdf = await renderPdf(model(), requestFor(1, 'PDF'), config, { captureLayoutTrace: true });
  assert.equal(pdf.pageCount, 1);
  const below = pdf.layoutTrace.pages[0].items.find((item) => item.text === 'BELOW_REGION');
  const grower = pdf.layoutTrace.pages[0].items.find((item) => item.text === 'ROW_0');
  // Declared gap: Below.Top (1in) - Grower.Top (0.5in) = 0.5in = 36pt.
  assert.ok(Math.abs((below.y - grower.y) - 36) <= 0.5, `expected the declared 36pt gap, saw ${below.y - grower.y}`);
  const xlsx = await renderExcel(model(), { ...requestFor(1, 'XLSX'), excelLayoutMode: 'REPORT' }, config, tempDir);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(xlsx.buffer);
  const positions = new Map();
  workbook.worksheets.forEach((sheet) => sheet.eachRow((row) => row.eachCell((cell) => {
    if (typeof cell.value === 'string' && cell.value) positions.set(cell.value, cell.row);
  })));
  assert.ok(positions.get('BELOW_REGION') > positions.get('ROW_0'));
});
