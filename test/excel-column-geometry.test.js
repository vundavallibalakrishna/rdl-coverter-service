import assert from 'node:assert/strict';
import test from 'node:test';
import ExcelJS from 'exceljs';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderExcel } from '../src/render/excel.js';

// Excel renders a stored column width `w` at exactly `w * 7` device pixels — it does not add the per-cell
// inset back. Subtracting that inset when writing the width therefore made every grid column ~5 px narrower
// than the RDL declared, and an RDL column that the shared grid splits into N slices lost 5N px. Text just
// clipped, so the loss went unnoticed; a date or number is never clipped by Excel — it becomes `#####` and
// the value disappears entirely. These tests pin the column geometry to the RDL's own points.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const EXCEL_MAX_DIGIT_WIDTH_PX = 7;
const PT_PER_PX = 72 / 96;
const pointsOf = (width) => (width || 0) * EXCEL_MAX_DIGIT_WIDTH_PX * PT_PER_PX;

const DATE_COLUMN_PT = 53.171; // an ordinary "Due Date" column: wide enough for the value, not generous
const OTHER_COLUMN_PT = 96;

const textbox = (name, value, format) => `<Textbox Name="${name}"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns>`
  + `<TextRun><Value>${value}</Value><Style><FontFamily>Arial</FontFamily><FontSize>8pt</FontSize>`
  + `${format ? `<Format>${format}</Format>` : ''}</Style></TextRun>`
  + '</TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border></Style></Textbox>';

// A second item whose edges fall INSIDE the date column forces the shared grid to slice that column into
// three, which is what multiplied the per-column loss in the real report.
const rdl = (withSlicer) => `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <Body><ReportItems>
    <Tablix Name="Actions"><TablixBody>
      <TablixColumns>
        <TablixColumn><Width>${OTHER_COLUMN_PT}pt</Width></TablixColumn>
        <TablixColumn><Width>${DATE_COLUMN_PT}pt</Width></TablixColumn>
      </TablixColumns>
      <TablixRows><TablixRow><Height>18pt</Height><TablixCells>
        <TablixCell><CellContents>${textbox('NameCell', '=Fields!Name.Value')}</CellContents></TablixCell>
        <TablixCell><CellContents>${textbox('DueCell', '=Fields!Due.Value', 'dd/MM/yyyy')}</CellContents></TablixCell>
      </TablixCells></TablixRow></TablixRows>
    </TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="Details"/></TablixMember></TablixMembers></TablixRowHierarchy>
    <DataSetName>D</DataSetName><Top>0pt</Top><Left>0pt</Left><Height>18pt</Height>
    <Width>${OTHER_COLUMN_PT + DATE_COLUMN_PT}pt</Width><Style/>
    </Tablix>
    ${withSlicer ? `<Textbox Name="Slicer"><Paragraphs><Paragraph><TextRuns><TextRun><Value>slice</Value>
      <Style><FontFamily>Arial</FontFamily><FontSize>8pt</FontSize></Style></TextRun></TextRuns></Paragraph></Paragraphs>
      <Top>40pt</Top><Left>${OTHER_COLUMN_PT + 14}pt</Left><Width>20pt</Width><Height>12pt</Height>
      <Style><Border><Style>Solid</Style></Border></Style></Textbox>` : ''}
  </ReportItems><Height>80pt</Height></Body><Width>${OTHER_COLUMN_PT + DATE_COLUMN_PT}pt</Width>
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="Name"><DataField>Name</DataField></Field>
    <Field Name="Due"><DataField>Due</DataField><rd:TypeName xmlns:rd="http://schemas.microsoft.com/SQLServer/reporting/reportdesigner">System.DateTime</rd:TypeName></Field>
  </Fields></DataSet></DataSets>
  <Page><PageWidth>8in</PageWidth><PageHeight>4in</PageHeight><TopMargin>0.2in</TopMargin>
    <BottomMargin>0.2in</BottomMargin><LeftMargin>0.2in</LeftMargin><RightMargin>0.2in</RightMargin></Page>
</Report>`;

const request = {
  outputFileName: 'excel-column-geometry',
  parameters: {},
  datasets: { D: [{ Name: 'Courier Services Action', Due: '2026-05-15T00:00:00' }] },
  excel: { layoutMode: 'REPORT' },
};

async function workbookFor(withSlicer) {
  const rendered = await renderExcel(parseRdl(rdl(withSlicer)), request, config, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(rendered.buffer);
  return workbook.worksheets[0];
}

// The columns a cell actually occupies: its own, plus every column of the merge it masters.
function spanOf(sheet, row, column) {
  const merges = (sheet.model.merges || []).map((range) => {
    const [start, end] = range.split(':');
    return { start: sheet.getCell(start), end: sheet.getCell(end || start) };
  });
  const hit = merges.find((merge) => merge.start.row <= row && merge.end.row >= row
    && merge.start.col <= column && merge.end.col >= column);
  return hit ? { from: hit.start.col, to: hit.end.col } : { from: column, to: column };
}

function dateCell(sheet) {
  let found = null;
  sheet.eachRow((row, rowNumber) => row.eachCell((cell, columnNumber) => {
    if (!found && cell.value instanceof Date) found = { rowNumber, columnNumber, cell };
  }));
  return found;
}

test('a REPORT column reproduces the RDL point width Excel will actually render', async () => {
  for (const withSlicer of [false, true]) {
    const sheet = await workbookFor(withSlicer);
    const date = dateCell(sheet);
    assert.ok(date, 'the date cell must be a live typed date');
    const span = spanOf(sheet, date.rowNumber, date.columnNumber);
    let points = 0;
    for (let column = span.from; column <= span.to; column += 1) points += pointsOf(sheet.getColumn(column).width);
    if (withSlicer) assert.ok(span.to > span.from, 'the fixture must slice the date column into several grid columns');
    assert.ok(
      Math.abs(points - DATE_COLUMN_PT) <= 1,
      `date column rendered at ${points.toFixed(2)}pt, RDL declares ${DATE_COLUMN_PT}pt (sliced: ${withSlicer})`,
    );
  }
});

test('a date column is wide enough that Excel shows the value instead of #####', async () => {
  // Excel refuses to clip or overflow a number/date: if the formatted value does not fit the (merged)
  // column it renders `#####`. Arial 8pt digits are 0.556em, the separators narrower, so the formatted
  // date needs this much room plus Excel's per-cell inset.
  const sheet = await workbookFor(true);
  const date = dateCell(sheet);
  const span = spanOf(sheet, date.rowNumber, date.columnNumber);
  let points = 0;
  for (let column = span.from; column <= span.to; column += 1) points += pointsOf(sheet.getColumn(column).width);
  const fontSize = date.cell.font?.size || 11;
  const displayed = '15/05/2026';
  const digits = displayed.replace(/\D/g, '').length;
  const separators = displayed.length - digits;
  const textPoints = (digits * 0.556 + separators * 0.278) * fontSize;
  const insetPoints = 5 * PT_PER_PX;
  assert.ok(
    points >= textPoints + insetPoints,
    `the date needs ${(textPoints + insetPoints).toFixed(2)}pt but the column gives ${points.toFixed(2)}pt`,
  );
});

test('every REPORT grid column keeps a positive rendered width', async () => {
  const sheet = await workbookFor(true);
  sheet.columns.forEach((column, index) => {
    if (column.width === undefined) return;
    assert.ok(column.width > 0, `column ${index + 1} has no width`);
    assert.ok(pointsOf(column.width) < 2000, `column ${index + 1} is implausibly wide`);
  });
});
