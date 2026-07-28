import assert from 'node:assert/strict';
import test from 'node:test';
import JSZip from 'jszip';
import { PDFDocument } from 'pdf-lib';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { materializeTablixRows } from '../src/rdl/validation.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderPdf } from '../src/render/pdf.js';
import { computeCellPlacements } from '../src/render/tableGrid.js';

const textbox = (name, value) => `<Textbox Name="${name}"><Paragraphs><Paragraph><TextRuns><TextRun><Value>${value}</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border></Style></Textbox>`;
const bodyRow = (name, value, hidden = null) => `<TablixRow><Height>0.2in</Height>${hidden ? `<Visibility><Hidden>${hidden}</Hidden></Visibility>` : ''}<TablixCells><TablixCell><CellContents>${textbox(name, value)}</CellContents></TablixCell></TablixCells></TablixRow>`;
const rowHeader = (name, value) => `<TablixHeader><Size>0.5in</Size><CellContents>${textbox(name, value)}</CellContents></TablixHeader>`;

function nestedGridRdl() {
  return `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields>
  <Field Name="Category"><DataField>Category</DataField></Field>
  <Field Name="Subcategory"><DataField>Subcategory</DataField></Field>
  <Field Name="Amount"><DataField>Amount</DataField></Field>
  <Field Name="Hide"><DataField>Hide</DataField></Field>
 </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="NestedGrid"><TablixBody>
   <TablixColumns><TablixColumn><Width>2in</Width></TablixColumn></TablixColumns>
   <TablixRows>
    ${bodyRow('reportBandBody', '')}
    ${bodyRow('categoryBand', 'CATEGORY BAND')}
    ${bodyRow('subcategoryBand', 'SUBCATEGORY BAND')}
    ${bodyRow('detail', '=Fields!Amount.Value', '=Fields!Hide.Value = &quot;yes&quot;')}
    ${bodyRow('subcategoryTotal', '=Sum(Fields!Amount.Value, &quot;SubcategoryGroup&quot;)')}
    ${bodyRow('categoryTotal', '=Sum(Fields!Amount.Value, &quot;CategoryGroup&quot;)')}
   </TablixRows>
  </TablixBody>
  <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
  <TablixRowHierarchy><TablixMembers>
  <TablixMember>${rowHeader('reportBand', 'REPORT BAND').replace('<Size>0.5in</Size>', '<Size>1in</Size>')}</TablixMember>
  <TablixMember>
   <Group Name="CategoryGroup"><GroupExpressions><GroupExpression>=Fields!Category.Value</GroupExpression></GroupExpressions></Group>
   ${rowHeader('categoryHeader', '=Fields!Category.Value')}
   <TablixMembers>
    <TablixMember/>
    <TablixMember>
     <Group Name="SubcategoryGroup"><GroupExpressions><GroupExpression>=Fields!Subcategory.Value</GroupExpression></GroupExpressions></Group>
     ${rowHeader('subcategoryHeader', '=Fields!Subcategory.Value')}
     <TablixMembers>
      <TablixMember/>
      <TablixMember><Group Name="Details"/></TablixMember>
      <TablixMember/>
     </TablixMembers>
    </TablixMember>
    <TablixMember/>
   </TablixMembers>
  </TablixMember></TablixMembers></TablixRowHierarchy>
  <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>1in</Height><Width>3in</Width><Style><Border><Style>Solid</Style></Border></Style></Tablix>
 </ReportItems><Height>4in</Height><Style/></Body>
 <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
 </ReportSection></ReportSections>
</Report>`;
}

const data = [
  { Category: 'A', Subcategory: 'A1', Amount: 10, Hide: 'no' },
  { Category: 'A', Subcategory: 'A2', Amount: 20, Hide: 'yes' },
  { Category: 'B', Subcategory: 'B1', Amount: 30, Hide: 'no' },
];

test('deep group headers retain stable hierarchy columns after hidden rows are removed', () => {
  const model = parseRdl(nestedGridRdl());
  const tablix = model.body.items[0];
  const rows = materializeTablixRows(tablix, data, {}, {}, { D: data });
  const placements = computeCellPlacements(rows, tablix.columns.length);

  assert.deepEqual(tablix.rowHeaderColumns, [36, 36]);
  assert.equal(tablix.columns.length, 3);
  assert.equal(rows.filter((row) => row.role === 'detail').some((row) => row.cells.some((cell) => cell.values?.includes('20'))), false);
  assert.equal(rows[0].cells[0].values?.[0], 'REPORT BAND');
  assert.equal(rows[0].cells[0].colSpan, 2);

  for (const [rowIndex, row] of rows.entries()) {
    for (const [cellIndex, cell] of row.cells.entries()) {
      assert.ok(placements[rowIndex][cellIndex] + (cell.colSpan || 1) <= tablix.columns.length);
    }
    const bodyCellIndex = row.cells.findIndex((cell) => cell.values?.some((value) => /^(CATEGORY|SUBCATEGORY) BAND$|^10$|^30$/.test(String(value))));
    if (bodyCellIndex >= 0) assert.equal(placements[rowIndex][bodyCellIndex], 2);
  }

  const categoryOwners = rows.flatMap((row, rowIndex) => row.cells.map((cell, cellIndex) => ({ rowIndex, cell, column: placements[rowIndex][cellIndex] })))
    .filter(({ cell }) => cell.values?.some((value) => value === 'A' || value === 'B'));
  assert.deepEqual(categoryOwners.map(({ column, cell }) => [column, cell.colSpan || 1]), [[0, 1], [0, 1]]);
});

test('the same stable nested grid renders through PDF and editable DOCX', async () => {
  const model = parseRdl(nestedGridRdl());
  const request = { outputFileName: 'nested-grid', parameters: {}, datasets: { D: data } };
  const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

  const pdfResult = await renderPdf(model, request, config);
  const pdf = await PDFDocument.load(pdfResult.buffer);
  assert.equal(pdf.getPageCount(), 1);

  const docxResult = await renderEditableDocx(model, request, config);
  const zip = await JSZip.loadAsync(docxResult.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  assert.equal((documentXml.match(/<w:gridCol\b/g) || []).length, 3);
  assert.match(documentXml, />CATEGORY BAND</);
  assert.match(documentXml, />SUBCATEGORY BAND</);
});
