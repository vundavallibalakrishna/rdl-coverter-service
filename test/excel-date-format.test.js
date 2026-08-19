// An explicitly-formatted date must stay a live typed Excel cell whose number format matches what the PDF/
// DOCX text renderer shows for the same .NET format; a format we cannot translate to a live number format
// falls back to the exact formatted string. Before this, Excel ignored the RDL date format and showed a
// date-only default.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { excelDateFormat } from '../src/render/excelFormat.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderExcel } from '../src/render/excel.js';
import { loadConfig } from '../src/config.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

test('excelDateFormat translates .NET date formats to Excel number formats', () => {
  assert.equal(excelDateFormat('dd/MM/yyyy'), 'dd/mm/yyyy');
  assert.equal(excelDateFormat('MM/dd/yyyy'), 'mm/dd/yyyy');
  assert.equal(excelDateFormat('MMMM yyyy'), 'mmmm yyyy');
  assert.equal(excelDateFormat('dd/MM/yyyy HH:mm:ss'), 'dd/mm/yyyy hh:mm:ss');
  assert.equal(excelDateFormat('hh:mm tt'), 'hh:mm AM/PM');
  // Standard single-letter specifiers the text renderer special-cases.
  assert.equal(excelDateFormat('y'), 'mmmm yyyy');
  assert.equal(excelDateFormat('d'), 'dd/mm/yyyy');
  assert.equal(excelDateFormat('G'), 'dd/mm/yyyy hh:mm:ss');
  // Untranslatable / empty → null so the caller writes the formatted string instead.
  assert.equal(excelDateFormat(''), null);
  assert.equal(excelDateFormat("dd 'of' MMMM"), null);
});

// A single-column tablix whose one data cell renders a DateTime field with an explicit Format — the typed
// Excel cell path (excelCellValue) that the fix changes.
function tablixReport(format) {
  const cell = (name, value, fmt) => `<Textbox Name="${name}"><Paragraphs><Paragraph><TextRuns><TextRun>`
    + `<Value>${value}</Value><Style><FontFamily>Arial</FontFamily>${fmt ? `<Format>${fmt}</Format>` : ''}</Style>`
    + `</TextRun></TextRuns></Paragraph></Paragraphs><Style>${fmt ? `<Format>${fmt}</Format>` : ''}</Style></Textbox>`;
  return Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="Rows">
    <Fields><Field Name="D"><DataField>D</DataField><TypeName>System.DateTime</TypeName></Field></Fields>
    <Query><CommandText>x</CommandText></Query>
  </DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="Grid">
      <TablixBody>
        <TablixColumns><TablixColumn><Width>2in</Width></TablixColumn></TablixColumns>
        <TablixRows>
          <TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents>${cell('Head', 'Date', null)}</CellContents></TablixCell></TablixCells></TablixRow>
          <TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents>${cell('Val', '=Fields!D.Value', format)}</CellContents></TablixCell></TablixCells></TablixRow>
        </TablixRows>
      </TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="g"/></TablixMember><TablixMember/></TablixMembers></TablixRowHierarchy>
      <DataSetName>Rows</DataSetName>
      <Top>0in</Top><Left>0in</Left><Height>0.5in</Height><Width>2in</Width>
    </Tablix>
  </ReportItems><Height>1in</Height><Style/></Body><Width>7in</Width>
  <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`, 'utf8');
}

test('an explicitly-formatted date is a typed Excel cell with the translated number format', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-excel-datefmt-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const request = { outputFileName: 'df', parameters: {}, datasets: { Rows: [{ D: '2026-03-04T08:04:53Z' }] }, excelLayoutMode: 'REPORT' };
  const rendered = await renderExcel(parseRdl(tablixReport('dd/MM/yyyy')), request, config, tempDir);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(rendered.buffer);
  let dateCell = null;
  workbook.worksheets.forEach((sheet) => sheet.eachRow((row) => row.eachCell((cell) => {
    if (cell.value instanceof Date) dateCell = cell;
  })));
  assert.ok(dateCell, 'expected a typed Date cell');
  assert.equal(dateCell.numFmt, 'dd/mm/yyyy');
});
