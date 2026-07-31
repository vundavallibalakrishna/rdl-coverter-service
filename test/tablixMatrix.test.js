import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { analyzeRdl, parseRdl } from '../src/rdl/parser.js';
import { materializeTablixColumns, materializeTablixRows, needsAdvancedMaterialization } from '../src/rdl/validation.js';
import { computeCellPlacements } from '../src/render/tableGrid.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderPdf } from '../src/render/pdf.js';
import { renderExcel } from '../src/render/excel.js';
import { renderVisualDocx } from '../src/render/visualDocx.js';
import { loadConfig } from '../src/config.js';

// A cross-tab / matrix: row group Region (row header), column group Product (dynamic columns), a
// single body column of Sum(Amount), plus a TablixCorner label in the top-left.
function matrixRdl() {
  return `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields><Field Name="Region"><DataField>Region</DataField></Field><Field Name="Product"><DataField>Product</DataField></Field><Field Name="Amount"><DataField>Amount</DataField></Field></Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="M"><TablixBody>
    <TablixColumns><TablixColumn><Width>1.5in</Width></TablixColumn></TablixColumns>
    <TablixRows><TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents><Textbox Name="v"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Sum(Fields!Amount.Value)</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell></TablixCells></TablixRow></TablixRows></TablixBody>
    <TablixCorner><TablixCornerRows><TablixCornerRow><TablixCornerCell><CellContents><Textbox Name="cor"><Paragraphs><Paragraph><TextRuns><TextRun><Value>Region</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCornerCell></TablixCornerRow></TablixCornerRows></TablixCorner>
    <TablixColumnHierarchy><TablixMembers><TablixMember>
      <Group Name="ProdGroup"><GroupExpressions><GroupExpression>=Fields!Product.Value</GroupExpression></GroupExpressions></Group>
      <TablixHeader><Size>0.25in</Size><CellContents><Textbox Name="ch"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Product.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixHeader>
    </TablixMember></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers><TablixMember>
      <Group Name="RegGroup"><GroupExpressions><GroupExpression>=Fields!Region.Value</GroupExpression></GroupExpressions></Group>
      <TablixHeader><Size>1.5in</Size><CellContents><Textbox Name="rh"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Region.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixHeader>
    </TablixMember></TablixMembers></TablixRowHierarchy>
    <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.25in</Height><Width>1.5in</Width><Style/></Tablix>
 </ReportItems><Height>3in</Height><Style/></Body><Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth></Page></ReportSection></ReportSections></Report>`;
}

// A heatmap-style matrix whose rotated/axis parent is static, whose first child is a dynamic row group,
// and whose second child is a static footer band. The parent header owns all expanded descendant rows;
// it must not be repeated once per row-group instance or once again for the footer.
function staticAncestorMatrixRdl({ groupedOuter = false } = {}) {
  const rowHierarchy = groupedOuter
    ? `<TablixMember>
        <Group Name="Categories"><GroupExpressions><GroupExpression>=Fields!Category.Value</GroupExpression></GroupExpressions></Group>
        <TablixHeader><Size>0.6in</Size><CellContents>${textbox('category', '=Fields!Category.Value')}</CellContents></TablixHeader>
        <TablixMembers><TablixMember>
          <TablixHeader><Size>0.3in</Size><CellContents>${textbox('axis', 'Axis')}</CellContents></TablixHeader>
          <TablixMembers>${axisChildren()}</TablixMembers>
        </TablixMember></TablixMembers>
      </TablixMember>`
    : `<TablixMember>
        <TablixHeader><Size>0.3in</Size><CellContents>${textbox('axis', 'Axis')}</CellContents></TablixHeader>
        <TablixMembers>${axisChildren()}</TablixMembers>
      </TablixMember>`;
  return `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields>
  <Field Name="Category"><DataField>Category</DataField></Field>
  <Field Name="RowKey"><DataField>RowKey</DataField></Field>
  <Field Name="ColumnKey"><DataField>ColumnKey</DataField></Field>
  <Field Name="Value"><DataField>Value</DataField></Field>
 </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="AxisMatrix"><TablixBody>
   <TablixColumns><TablixColumn><Width>0.8in</Width></TablixColumn></TablixColumns>
   <TablixRows>
    <TablixRow><Height>0.35in</Height><TablixCells><TablixCell><CellContents>${textbox('value', '=Sum(Fields!Value.Value)', true)}</CellContents></TablixCell></TablixCells></TablixRow>
    <TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents>${textbox('columnLabel', '=Fields!ColumnKey.Value')}</CellContents></TablixCell></TablixCells></TablixRow>
   </TablixRows>
  </TablixBody>
  <TablixColumnHierarchy><TablixMembers><TablixMember>
   <Group Name="Columns"><GroupExpressions><GroupExpression>=Fields!ColumnKey.Value</GroupExpression></GroupExpressions></Group>
  </TablixMember></TablixMembers></TablixColumnHierarchy>
  <TablixRowHierarchy><TablixMembers>${rowHierarchy}</TablixMembers></TablixRowHierarchy>
  <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.95in</Height><Width>3in</Width><Style/>
  </Tablix>
 </ReportItems><Height>3in</Height><Style/></Body><Width>6in</Width>
 <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
 </ReportSection></ReportSections></Report>`;
}

function textbox(name, value, bordered = false) {
  return `<Textbox Name="${name}"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun><Value>${value}</Value><Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns><Style><TextAlign>Center</TextAlign></Style></Paragraph></Paragraphs><Style>${bordered ? '<Border><Style>Solid</Style></Border>' : ''}</Style></Textbox>`;
}

function axisChildren() {
  return `<TablixMember>
    <Group Name="Rows"><GroupExpressions><GroupExpression>=Fields!RowKey.Value</GroupExpression></GroupExpressions></Group>
    <TablixHeader><Size>0.8in</Size><CellContents>${textbox('rowLabel', '=Fields!RowKey.Value')}</CellContents></TablixHeader>
   </TablixMember>
   <TablixMember><TablixHeader><Size>0.8in</Size><CellContents>${textbox('footerStub', '')}</CellContents></TablixHeader></TablixMember>`;
}

const tablixOf = (rdl) => parseRdl(rdl).body.items.find((item) => item.type === 'Tablix');
const rows = [
  { Region: 'East', Product: 'A', Amount: 10 },
  { Region: 'East', Product: 'B', Amount: 20 },
  { Region: 'West', Product: 'A', Amount: 5 },
];
const axisRows = [
  { Category: 'A', RowKey: 'R1', ColumnKey: 'C1', Value: 11 },
  { Category: 'A', RowKey: 'R1', ColumnKey: 'C2', Value: 12 },
  { Category: 'A', RowKey: 'R2', ColumnKey: 'C1', Value: 21 },
  { Category: 'A', RowKey: 'R2', ColumnKey: 'C2', Value: 22 },
];
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

test('a matrix and its TablixCorner are compatible (no longer fail closed)', () => {
  const analysis = analyzeRdl(matrixRdl());
  assert.equal(analysis.compatible, true);
  assert.equal(analysis.capabilities.rejected.some(({ path }) => path.includes('TablixCorner') || path.includes('TablixHeader')), false);
});

test('a static ancestor header spans its dynamic descendants and trailing static footer exactly once', () => {
  const tablix = tablixOf(staticAncestorMatrixRdl());
  const materialized = materializeTablixRows(tablix, axisRows, {}, {}, { D: axisRows });

  assert.equal(materialized.length, 3, 'two dynamic row instances plus the static footer');
  assert.equal(materialized[0].cells[0].items[0].name, 'axis');
  assert.equal(materialized[0].cells[0].rowSpan, 3);
  assert.equal(materialized.slice(1).flatMap((row) => row.cells).filter((cell) => cell.items?.[0]?.name === 'axis').length, 0);

  const columns = materializeTablixColumns(tablix, axisRows, {}, {}, { D: axisRows });
  assert.deepEqual(computeCellPlacements(materialized, columns.length), [
    [0, 1, 2, 3],
    [1, 2, 3],
    [1, 2, 3],
  ]);
});

test('a static ancestor nested inside an outer dynamic group restarts once per outer group instance', () => {
  const source = [
    ...axisRows,
    ...axisRows.map((row) => ({ ...row, Category: 'B', Value: row.Value + 100 })),
  ];
  const tablix = tablixOf(staticAncestorMatrixRdl({ groupedOuter: true }));
  const materialized = materializeTablixRows(tablix, source, {}, {}, { D: source });
  const axisOwners = materialized
    .flatMap((row, rowIndex) => row.cells.map((cell) => ({ cell, rowIndex })))
    .filter(({ cell }) => cell.items?.[0]?.name === 'axis');

  assert.deepEqual(axisOwners.map(({ rowIndex, cell }) => [rowIndex, cell.rowSpan]), [[0, 3], [3, 3]]);
  assert.deepEqual(
    materialized.flatMap((row) => row.cells).filter((cell) => cell.items?.[0]?.name === 'category').map((cell) => cell.rowSpan),
    [3, 3],
  );
});

test('PDF, editable DOCX, XLSX, and visual DOCX preserve one static matrix-axis owner', async () => {
  const model = parseRdl(staticAncestorMatrixRdl());
  const request = { outputFileName: 'axis-matrix', parameters: {}, datasets: { D: axisRows } };
  const pdf = await renderPdf(model, request, config, { captureLayoutTrace: true });
  const tracedAxis = pdf.layoutTrace.pages.flatMap((page) => page.items).filter((item) => item.text === 'Axis');
  assert.equal(tracedAxis.length, 1);
  assert.equal(tracedAxis[0].rowSpan, 3);

  const docx = await renderEditableDocx(model, request, config);
  const docxZip = await JSZip.loadAsync(docx.buffer);
  const documentXml = await docxZip.file('word/document.xml').async('string');
  const nativeText = [...documentXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((match) => match[1]).join('');
  assert.equal((nativeText.match(/Axis/g) || []).length, 1);
  assert.match(documentXml, /<w:vMerge w:val="restart"\/>/);

  const excel = await renderExcel(model, request, config, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(excel.buffer);
  const worksheet = workbook.worksheets[0];
  const axisMerge = worksheet.model.merges.find((range) => worksheet.getCell(range.split(':')[0]).value === 'Axis');
  assert.ok(axisMerge, 'the static axis header remains one native vertical merge in Excel REPORT mode');
  const [axisStart, axisEnd] = axisMerge.split(':');
  assert.equal(worksheet.getCell(axisEnd).row - worksheet.getCell(axisStart).row + 1, 3);

  await fs.mkdir(config.tempRoot, { recursive: true, mode: 0o700 });
  const tempDir = await fs.mkdtemp(path.join(config.tempRoot, 'axis-matrix-'));
  try {
    const visual = await renderVisualDocx(model, request, config, tempDir);
    const visualZip = await JSZip.loadAsync(visual.buffer);
    assert.equal(visual.pageCount, pdf.pageCount);
    assert.equal(Object.keys(visualZip.files).filter((name) => /^word\/media\/.+\.png$/.test(name)).length, pdf.pageCount);
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
});

test('a matrix expands columns to rowHeaderColumns + keys x bodyColumns with intersection cell values', () => {
  const tablix = tablixOf(matrixRdl());
  assert.equal(tablix.hasColumnGroups, true);
  assert.equal(needsAdvancedMaterialization(tablix), true);

  // 1 row-header column + 2 distinct column keys (A, B) x 1 body column = 3 grid columns.
  const columns = materializeTablixColumns(tablix, rows, {}, {}, {});
  assert.equal(columns.length, tablix.rowHeaderColumns.length + 2 * tablix.bodyColumns.length);
  assert.equal(columns.length, 3);

  const materialized = materializeTablixRows(tablix, rows, {}, {}, {});
  const text = materialized.map((row) => row.cells.map((cell) => (cell.values || []).join('')).join('|'));
  assert.deepEqual(text, [
    'Region|A|B', // TablixCorner label + the two dynamic column headers
    'East|10|20', // East x A = 10, East x B = 20 (intersections)
    'West|5|0', // West x A = 5, West x B = empty intersection -> 0
  ]);

  // The first output row is the repeatable column-header row; the corner spans the row-header column.
  assert.equal(materialized[0].isHeader, true);
  assert.equal(materialized[0].cells[0].colSpan, tablix.rowHeaderColumns.length);

  // Body cells land on their intended grid columns: row header at 0, the two product columns at 1 and 2.
  const placements = computeCellPlacements(materialized, columns.length);
  assert.deepEqual(placements[1], [0, 1, 2]);
});

test('a mixed static and dynamic column hierarchy repeats only the dynamic leaf column', () => {
  const rdl = `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields>
  <Field Name="RowKey"><DataField>RowKey</DataField></Field>
  <Field Name="ColumnKey"><DataField>ColumnKey</DataField></Field>
  <Field Name="Value"><DataField>Value</DataField></Field>
 </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="Mixed"><TablixBody>
   <TablixColumns><TablixColumn><Width>1in</Width></TablixColumn><TablixColumn><Width>0.5in</Width></TablixColumn></TablixColumns>
   <TablixRows><TablixRow><Height>0.25in</Height><TablixCells>
    <TablixCell><CellContents><Textbox Name="row"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!RowKey.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell>
    <TablixCell><CellContents><Textbox Name="value"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Value.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell>
   </TablixCells></TablixRow></TablixRows>
  </TablixBody>
  <TablixColumnHierarchy><TablixMembers>
   <TablixMember/>
   <TablixMember><Group Name="Columns"><GroupExpressions><GroupExpression>=Fields!ColumnKey.Value</GroupExpression></GroupExpressions></Group></TablixMember>
  </TablixMembers></TablixColumnHierarchy>
  <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="Rows"><GroupExpressions><GroupExpression>=Fields!RowKey.Value</GroupExpression></GroupExpressions></Group></TablixMember></TablixMembers></TablixRowHierarchy>
  <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.25in</Height><Width>1.5in</Width><Style/>
  </Tablix>
 </ReportItems><Height>3in</Height><Style/></Body><Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth></Page></ReportSection></ReportSections></Report>`;
  const tablix = tablixOf(rdl);
  const source = [
    { RowKey: 'R1', ColumnKey: 'C1', Value: 11 },
    { RowKey: 'R1', ColumnKey: 'C2', Value: 12 },
    { RowKey: 'R2', ColumnKey: 'C1', Value: 21 },
    { RowKey: 'R2', ColumnKey: 'C2', Value: 22 },
  ];

  assert.deepEqual(materializeTablixColumns(tablix, source, {}, {}, {}), [72, 36, 36]);
  assert.deepEqual(
    materializeTablixRows(tablix, source, {}, {}, {}).map((row) => row.cells.map((cell) => (cell.values || []).join(''))),
    [
      ['R1', '11', '12'],
      ['R2', '21', '22'],
    ],
  );
});

test('a dynamic column group honors its declared sort instead of incoming row order', () => {
  const rdl = matrixRdl().replace(
    '</Group>\n      <TablixHeader>',
    '</Group><SortExpressions><SortExpression><Value>=Fields!Product.Value</Value></SortExpression></SortExpressions>\n      <TablixHeader>',
  );
  const tablix = tablixOf(rdl);
  const source = [
    { Region: 'East', Product: 'C', Amount: 30 },
    { Region: 'East', Product: 'A', Amount: 10 },
    { Region: 'East', Product: 'B', Amount: 20 },
  ];

  assert.deepEqual(
    materializeTablixRows(tablix, source, {}, {}, {})[0].cells.map((cell) => (cell.values || []).join('')),
    ['Region', 'A', 'B', 'C'],
  );
});

test('matrix cells preserve their row-column intersection for expression-backed styles', () => {
  const rdl = matrixRdl().replace(
    '<Style/></Textbox></CellContents></TablixCell></TablixCells></TablixRow>',
    '<Style><BackgroundColor>=IIF(Fields!Product.Value = "A", "Red", "Blue")</BackgroundColor></Style></Textbox></CellContents></TablixCell></TablixCells></TablixRow>',
  );
  const tablix = tablixOf(rdl);
  const materialized = materializeTablixRows(tablix, rows, {}, {}, {});
  assert.equal(materialized[1].cells[1].fields.Product, 'A');
  assert.equal(materialized[1].cells[2].fields.Product, 'B');
  assert.equal(materialized[1].cells[1].scopeDataset.length, 1);
  assert.equal(materialized[1].cells[2].scopeDataset.length, 1);
});

test('a matrix renders native editable DOCX with the expanded column grid', async () => {
  const model = parseRdl(matrixRdl());
  const result = await renderEditableDocx(model, { outputFileName: 'matrix', parameters: {}, datasets: { D: rows } });
  const zip = await JSZip.loadAsync(result.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  // The page canvas adds blank-space boundaries around the three matrix columns, while preserving all
  // expanded matrix labels as native editable text.
  assert.equal((documentXml.match(/<w:gridCol\b/g) || []).length >= 3, true);
  assert.match(documentXml, /<w:tblLayout w:type="fixed"\/>/);
  assert.match(documentXml, />A</);
  assert.match(documentXml, />B</);
  assert.match(documentXml, />Region</);
});
