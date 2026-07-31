import assert from 'node:assert/strict';
import test from 'node:test';
import JSZip from 'jszip';
import { analyzeRdl, parseRdl } from '../src/rdl/parser.js';
import { materializeTablixColumns, materializeTablixRows, needsAdvancedMaterialization } from '../src/rdl/validation.js';
import { computeCellPlacements } from '../src/render/tableGrid.js';
import { renderEditableDocx } from '../src/render/docx.js';

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

const tablixOf = (rdl) => parseRdl(rdl).body.items.find((item) => item.type === 'Tablix');
const rows = [
  { Region: 'East', Product: 'A', Amount: 10 },
  { Region: 'East', Product: 'B', Amount: 20 },
  { Region: 'West', Product: 'A', Amount: 5 },
];

test('a matrix and its TablixCorner are compatible (no longer fail closed)', () => {
  const analysis = analyzeRdl(matrixRdl());
  assert.equal(analysis.compatible, true);
  assert.equal(analysis.capabilities.rejected.some(({ path }) => path.includes('TablixCorner') || path.includes('TablixHeader')), false);
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
