// A running aggregate accumulates in the order the DATA REGION processes its rows. Grouping reorders rows
// relative to the dataset even with no SortExpression, so the standard SSRS serial-number idiom
// `RunningValue(Fields!Id.Value, CountDistinct, Nothing)` must count down the visible column. Accumulating
// in dataset arrival order instead gave every row the rank it held in the raw data, so the column rendered
// scrambled in PDF, editable Word, and Excel alike.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { parseRdl } from '../src/rdl/parser.js';
import { materializeTablixRows, needsAdvancedMaterialization } from '../src/rdl/validation.js';
import { renderPdf } from '../src/render/pdf.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { loadConfig } from '../src/config.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

// Arrival order alternates the group key, so grouping must reorder the rows. Ids are already sequential in
// arrival order: the pre-fix ranks are therefore exactly the Ids, and the expected ranks are not.
const DATA = {
  D: [
    { Id: 1, Area: 'Alpha' },
    { Id: 2, Area: 'Beta' },
    { Id: 3, Area: 'Alpha' },
    { Id: 4, Area: 'Beta' },
    { Id: 5, Area: 'Alpha' },
    { Id: 6, Area: 'Beta' },
  ],
};
// Grouped processing order: the three Alpha rows, then the three Beta rows.
const EXPECTED_SERIALS = ['S1', 'S2', 'S3', 'S4', 'S5', 'S6'];
const ARRIVAL_ORDER_SERIALS = ['S1', 'S3', 'S5', 'S2', 'S4', 'S6'];

function cell(value) {
  return `<TablixCell><CellContents><Textbox Name="T${Math.abs(hash(value))}">
    <Paragraphs><Paragraph><TextRuns><TextRun><Value>${value}</Value><Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs>
    <Style/></Textbox></CellContents></TablixCell>`;
}

function hash(value) {
  let result = 0;
  for (const character of String(value)) result = (result * 31 + character.charCodeAt(0)) | 0;
  return result;
}

// A static leaf nested inside a group (the "Area" header row) is what puts this tablix on the grouped
// materialization path — the same shape as a report that prints a business-area heading above its rows.
function report() {
  return Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="Id"><DataField>Id</DataField><TypeName>System.Int32</TypeName></Field>
    <Field Name="Area"><DataField>Area</DataField><TypeName>System.String</TypeName></Field>
  </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="Plan"><TablixBody>
      <TablixColumns>
        <TablixColumn><Width>1.5in</Width></TablixColumn>
        <TablixColumn><Width>1.5in</Width></TablixColumn>
        <TablixColumn><Width>2in</Width></TablixColumn>
      </TablixColumns>
      <TablixRows>
        <TablixRow><Height>0.25in</Height><TablixCells>${cell('#')}${cell('Row')}${cell('Area')}</TablixCells></TablixRow>
        <TablixRow><Height>0.25in</Height><TablixCells>${cell('')}${cell('')}${cell('=Fields!Area.Value')}</TablixCells></TablixRow>
        <TablixRow><Height>0.25in</Height><TablixCells>${cell('="S" &amp; RunningValue(Fields!Id.Value, CountDistinct, Nothing)')}${cell('="R" &amp; RowNumber(Nothing)')}${cell('=Fields!Area.Value')}</TablixCells></TablixRow>
      </TablixRows>
    </TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/><TablixMember/><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers>
      <TablixMember/>
      <TablixMember>
        <Group Name="AreaGroup"><GroupExpressions><GroupExpression>=Fields!Area.Value</GroupExpression></GroupExpressions></Group>
        <TablixMembers>
          <TablixMember/>
          <TablixMember><Group Name="Detail"><GroupExpressions><GroupExpression>=Fields!Id.Value</GroupExpression></GroupExpressions></Group></TablixMember>
        </TablixMembers>
      </TablixMember>
    </TablixMembers></TablixRowHierarchy>
    <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>2in</Height><Width>5in</Width><Style/></Tablix>
  </ReportItems><Height>3in</Height><Style/></Body><Width>7in</Width>
  <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`, 'utf8');
}

const tablixOf = (rdl) => parseRdl(rdl).body.items.find((item) => item.type === 'Tablix');

test('the grouped tablix takes the recursive materialization path (the synthetic trigger is real)', () => {
  assert.equal(needsAdvancedMaterialization(tablixOf(report())), true);
});

test('RunningValue(..., Nothing) counts in grouped processing order, not dataset arrival order', () => {
  const rows = materializeTablixRows(tablixOf(report()), DATA.D, {}, {}, DATA);
  const serials = rows.flatMap((row) => (row.cells || []).flatMap((entry) => (entry.values || [])))
    .filter((value) => /^S\d+$/.test(String(value)));
  assert.deepEqual(serials, EXPECTED_SERIALS);
  assert.notDeepEqual(serials, ARRIVAL_ORDER_SERIALS, 'arrival order must no longer leak through');
});

test('RowNumber(Nothing) follows the same processing order', () => {
  const rows = materializeTablixRows(tablixOf(report()), DATA.D, {}, {}, DATA);
  const numbers = rows.flatMap((row) => (row.cells || []).flatMap((entry) => (entry.values || [])))
    .filter((value) => /^R\d+$/.test(String(value)));
  assert.deepEqual(numbers, ['R1', 'R2', 'R3', 'R4', 'R5', 'R6']);
});

test('the group header rows still show each group once, in order', () => {
  const rows = materializeTablixRows(tablixOf(report()), DATA.D, {}, {}, DATA);
  // Every source row is still present exactly once: reordering the running scope must not drop or duplicate.
  const areas = rows.flatMap((row) => (row.cells || []).flatMap((entry) => (entry.values || [])))
    .filter((value) => value === 'Alpha' || value === 'Beta');
  assert.equal(areas.filter((value) => value === 'Alpha').length, 4, 'one Alpha header plus three Alpha rows');
  assert.equal(areas.filter((value) => value === 'Beta').length, 4);
});

test('PDF renders the serial column top to bottom in order', async () => {
  const rendered = await renderPdf(
    parseRdl(report()),
    { outputFileName: 'serials', parameters: {}, datasets: DATA },
    config,
    { captureLayoutTrace: true },
  );
  const serials = rendered.layoutTrace.pages
    .flatMap((page) => page.items)
    .filter((item) => /^S\d+$/.test(String(item.text || '').trim()))
    .sort((left, right) => left.y - right.y)
    .map((item) => String(item.text).trim());
  assert.deepEqual(serials, EXPECTED_SERIALS);
});

test('DOCX_EDITABLE carries the same serial order', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-serials-docx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderEditableDocx(
    parseRdl(report()),
    { output: 'DOCX_EDITABLE', outputFileName: 'serials', parameters: {}, datasets: DATA },
    config,
    tempDir,
  );
  const zip = await JSZip.loadAsync(rendered.buffer);
  const document = await zip.file('word/document.xml').async('string');
  const serials = [...document.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)]
    .map((match) => match[1].trim())
    .filter((value) => /^S\d+$/.test(value));
  assert.deepEqual(serials, EXPECTED_SERIALS);
});

test('XLSX carries the same serial order', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-serials-xlsx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderExcel(
    parseRdl(report()),
    { output: 'XLSX', excelLayoutMode: 'REPORT', outputFileName: 'serials', parameters: {}, datasets: DATA },
    config,
    tempDir,
  );
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(rendered.buffer);
  const serials = [];
  workbook.worksheets.forEach((sheet) => sheet.eachRow((row) => {
    const seen = new Set();
    row.eachCell((entry) => {
      const value = String(entry.value ?? '').trim();
      if (/^S\d+$/.test(value) && !seen.has(value)) { seen.add(value); serials.push(value); }
    });
  }));
  assert.deepEqual(serials, EXPECTED_SERIALS);
});
