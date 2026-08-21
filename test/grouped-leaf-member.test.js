// A tablix row member that IS a group (it declares GroupExpressions) and has no child member beneath it
// is a GROUPED LEAF. SSRS renders that template once per distinct group instance, not once per source row.
// The flat materializer used to walk the row stream for such a member and repeat the template for every
// row in the instance, multiplying the region — a grouped 1x1 List then produced one canvas per row.
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
import { materializeTablixRows } from '../src/rdl/validation.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

const execFileAsync = promisify(execFile);
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

// `rowHierarchy` is the only thing that varies between the cases below, so every other part of the
// definition — columns, the detail template row, the dataset — stays byte-identical.
function reportRdl(rowHierarchy) {
  return `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="Cat"><DataField>Cat</DataField><TypeName>System.String</TypeName></Field>
    <Field Name="Region"><DataField>Region</DataField><TypeName>System.String</TypeName></Field>
    <Field Name="Amount"><DataField>Amount</DataField><TypeName>System.Int32</TypeName></Field>
  </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="T"><TablixBody>
      <TablixColumns><TablixColumn><Width>3in</Width></TablixColumn><TablixColumn><Width>2in</Width></TablixColumn></TablixColumns>
      <TablixRows><TablixRow><Height>0.25in</Height><TablixCells>
        <TablixCell><CellContents><Textbox Name="c"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Cat.Value</Value><Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell>
        <TablixCell><CellContents><Textbox Name="a"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Sum(Fields!Amount.Value)</Value><Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell>
      </TablixCells></TablixRow></TablixRows></TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers>${rowHierarchy}</TablixMembers></TablixRowHierarchy>
      <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.25in</Height><Width>5in</Width><Style/></Tablix>
  </ReportItems><Height>9in</Height><Style/></Body><Width>5in</Width>
  <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections></Report>`;
}

// Report Builder names the member "Details" when a group is set ON the details row; the name carries no
// meaning, only the presence of GroupExpressions does.
const GROUPED_LEAF = '<TablixMember><Group Name="Details"><GroupExpressions><GroupExpression>=Fields!Cat.Value</GroupExpression></GroupExpressions></Group></TablixMember>';
// Counterexample: the same group keeping a real detail member underneath. SSRS emits one row per SOURCE
// ROW there, so this must keep the unchanged per-row behaviour.
const GROUP_WITH_DETAILS = '<TablixMember><Group Name="Cat"><GroupExpressions><GroupExpression>=Fields!Cat.Value</GroupExpression></GroupExpressions></Group>'
  + '<TablixMembers><TablixMember><Group Name="Details"/></TablixMember></TablixMembers></TablixMember>';
// Variant: a grouped leaf nested inside an outer group. One row per distinct (Region, Cat) pair.
const NESTED_GROUPED_LEAF = '<TablixMember><Group Name="Region"><GroupExpressions><GroupExpression>=Fields!Region.Value</GroupExpression></GroupExpressions></Group>'
  + `<TablixMembers>${GROUPED_LEAF}</TablixMembers></TablixMember>`;
// Variant: the group expression is an arbitrary expression rather than a bare field reference.
const EXPRESSION_GROUPED_LEAF = '<TablixMember><Group Name="Details"><GroupExpressions><GroupExpression>=UCase(Fields!Cat.Value)</GroupExpression></GroupExpressions></Group></TablixMember>';

// Five source rows, three distinct Cat values, two distinct Regions.
const ROWS = [
  { Cat: 'Alpha', Region: 'North', Amount: 1 },
  { Cat: 'Alpha', Region: 'North', Amount: 2 },
  { Cat: 'Beta', Region: 'North', Amount: 4 },
  { Cat: 'Beta', Region: 'South', Amount: 8 },
  { Cat: 'Gamma', Region: 'South', Amount: 16 },
];
const renderRequest = { output: 'PDF', parameters: {}, datasets: { D: ROWS } };

const tablixOf = (rdl) => parseRdl(Buffer.from(rdl, 'utf8')).body.items.find((entry) => entry.type === 'Tablix');
const cellValues = (rows) => rows.map((row) => row.cells.map((cell) => (cell.values || []).join('')));

test('a grouped leaf member emits one row per group instance, not per source row', () => {
  assert.deepEqual(
    cellValues(materializeTablixRows(tablixOf(reportRdl(GROUPED_LEAF)), ROWS, {}, {}, { D: ROWS })),
    [['Alpha', '3'], ['Beta', '12'], ['Gamma', '16']],
  );
});

test('a group that still owns a detail member keeps emitting one row per source row', () => {
  assert.deepEqual(
    cellValues(materializeTablixRows(tablixOf(reportRdl(GROUP_WITH_DETAILS)), ROWS, {}, {}, { D: ROWS })),
    [['Alpha', '3'], ['Alpha', '3'], ['Beta', '12'], ['Beta', '12'], ['Gamma', '16']],
  );
});

test('a grouped leaf inside an outer group emits one row per distinct combination', () => {
  assert.deepEqual(
    cellValues(materializeTablixRows(tablixOf(reportRdl(NESTED_GROUPED_LEAF)), ROWS, {}, {}, { D: ROWS })),
    [['Alpha', '3'], ['Beta', '4'], ['Beta', '8'], ['Gamma', '16']],
  );
});

test('a grouped leaf keyed by an expression groups on the evaluated value', () => {
  const rows = [{ Cat: 'alpha', Amount: 1 }, { Cat: 'ALPHA', Amount: 2 }, { Cat: 'beta', Amount: 4 }];
  assert.deepEqual(
    cellValues(materializeTablixRows(tablixOf(reportRdl(EXPRESSION_GROUPED_LEAF)), rows, {}, {}, { D: rows })),
    [['alpha', '3'], ['beta', '4']],
  );
});

test('PDF draws a grouped leaf once per instance', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-grouped-leaf-pdf-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, 'grouped-leaf.pdf');
  const rendered = await renderPdf(parseRdl(Buffer.from(reportRdl(GROUPED_LEAF), 'utf8')), renderRequest, config);
  await fs.writeFile(pdfPath, rendered.buffer);
  const { stdout } = await execFileAsync('pdftotext', ['-layout', pdfPath, '-']);
  assert.equal((stdout.match(/Alpha/g) || []).length, 1);
  assert.equal((stdout.match(/Beta/g) || []).length, 1);
  assert.match(stdout, /Gamma/);
});

test('DOCX_EDITABLE draws a grouped leaf once per instance', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-grouped-leaf-docx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderEditableDocx(
    parseRdl(Buffer.from(reportRdl(GROUPED_LEAF), 'utf8')),
    { ...renderRequest, output: 'DOCX_EDITABLE' },
    config,
    tempDir,
  );
  const document = await (await JSZip.loadAsync(rendered.buffer)).file('word/document.xml').async('string');
  const texts = [...document.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((match) => match[1]);
  assert.equal(texts.filter((value) => value === 'Alpha').length, 1);
  assert.equal(texts.filter((value) => value === 'Beta').length, 1);
  assert.equal(texts.filter((value) => value === 'Gamma').length, 1);
});

test('XLSX writes a grouped leaf once per instance with the instance aggregate', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-grouped-leaf-xlsx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderExcel(
    parseRdl(Buffer.from(reportRdl(GROUPED_LEAF), 'utf8')),
    { ...renderRequest, output: 'XLSX', excelLayoutMode: 'REPORT' },
    config,
    tempDir,
  );
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(rendered.buffer);
  const values = [];
  workbook.worksheets.forEach((sheet) => sheet.eachRow((row) => row.eachCell((cell) => {
    if (cell.value !== null && cell.value !== undefined && cell.value !== '') values.push(cell.value);
  })));
  assert.equal(values.filter((value) => value === 'Alpha').length, 1);
  assert.equal(values.filter((value) => value === 'Beta').length, 1);
  assert.equal(values.filter((value) => value === 'Gamma').length, 1);
  assert.deepEqual(values.filter((value) => typeof value === 'number'), [3, 12, 16]);
});
