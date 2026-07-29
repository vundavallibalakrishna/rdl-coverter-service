import assert from 'node:assert/strict';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { analyzeRdl, parseRdl } from '../src/rdl/parser.js';
import { materializeTablixRows } from '../src/rdl/validation.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const request = {
  outputFileName: 'nested-tablix',
  parameters: {},
  datasets: {
    D: [
      { Key: 'A', Detail: 'Alpha' },
      { Key: 'B', Detail: 'Beta' },
    ],
  },
  excel: { layoutMode: 'REPORT' },
};

const textbox = (name, value) => `<Textbox Name="${name}"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun><Value>${value}</Value><Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style></TextRun></TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border><PaddingLeft>2pt</PaddingLeft><PaddingRight>2pt</PaddingRight></Style></Textbox>`;

const rdl = `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields>
  <Field Name="Key"><DataField>Key</DataField></Field><Field Name="Detail"><DataField>Detail</DataField></Field>
 </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems><Rectangle Name="BodyContainer"><ReportItems>
  <Tablix Name="Parent"><TablixBody>
   <TablixColumns><TablixColumn><Width>3in</Width></TablixColumn></TablixColumns>
   <TablixRows><TablixRow><Height>0.35in</Height><TablixCells><TablixCell><CellContents>
    <Rectangle Name="CellContainer"><ReportItems>
     <Tablix Name="Child"><TablixBody>
      <TablixColumns><TablixColumn><Width>1in</Width></TablixColumn><TablixColumn><Width>2in</Width></TablixColumn></TablixColumns>
      <TablixRows><TablixRow><Height>0.3in</Height><TablixCells>
       <TablixCell><CellContents>${textbox('ChildKey', '=Fields!Key.Value')}</CellContents></TablixCell>
       <TablixCell><CellContents>${textbox('ChildDetail', '=Fields!Detail.Value')}</CellContents></TablixCell>
      </TablixCells></TablixRow></TablixRows>
     </TablixBody><TablixColumnHierarchy><TablixMembers><TablixMember/><TablixMember/></TablixMembers></TablixColumnHierarchy>
     <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="ChildDetails"/></TablixMember></TablixMembers></TablixRowHierarchy>
     <DataSetName>D</DataSetName><Top>0.02in</Top><Left>0in</Left><Height>0.3in</Height><Width>3in</Width><Style/>
     </Tablix>
    </ReportItems><Style/></Rectangle>
   </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
  </TablixBody><TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
  <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="ParentDetails"/></TablixMember></TablixMembers></TablixRowHierarchy>
  <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.7in</Height><Width>3in</Width><Style><Border><Style>Solid</Style></Border></Style>
  </Tablix>
 </ReportItems><Top>0in</Top><Left>0in</Left><Height>0.7in</Height><Width>3in</Width><Style/></Rectangle>
 </ReportItems><Height>1in</Height><Style/></Body><Width>3in</Width>
 <Page><PageWidth>4in</PageWidth><PageHeight>4in</PageHeight><TopMargin>0.25in</TopMargin><BottomMargin>0.25in</BottomMargin><LeftMargin>0.25in</LeftMargin><RightMargin>0.25in</RightMargin></Page>
 </ReportSection></ReportSections></Report>`;

test('nested tablix uses the current parent detail scope and stays compatible', () => {
  assert.equal(analyzeRdl(rdl).compatible, true);
  const model = parseRdl(rdl);
  const parent = model.body.items[0].items[0];
  const rows = materializeTablixRows(parent, request.datasets.D, {}, {}, { D: request.datasets.D });
  assert.equal(rows.length, 2);
  assert.deepEqual(rows.map((row) => (
    row.cells[0].nestedTablixes[0].rows.map((childRow) => childRow.cells.map((cell) => cell.values[0]))
  )), [[['A', 'Alpha']], [['B', 'Beta']]]);
});

test('nested tablix renders as native grids in PDF, editable DOCX, and XLSX REPORT', async () => {
  const model = parseRdl(rdl);
  const pdf = await renderPdf(model, request, config);
  assert.ok(pdf.buffer.length > 1000);
  assert.equal(pdf.pageCount, 1);

  const docx = await renderEditableDocx(model, request, config);
  const zip = await JSZip.loadAsync(docx.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  assert.equal((documentXml.match(/<w:tbl[ >]/g) || []).length, 3);
  for (const value of ['Alpha', 'Beta']) assert.match(documentXml, new RegExp(value));

  const xlsx = await renderExcel(model, request, config, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(xlsx.buffer);
  const values = [];
  workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => {
    const value = typeof cell.value === 'object' && cell.value?.richText
      ? cell.value.richText.map((run) => run.text).join('')
      : String(cell.value ?? '');
    if (value) values.push(value);
  }));
  for (const value of ['A', 'Alpha', 'B', 'Beta']) assert.ok(values.includes(value), `missing ${value} in XLSX`);
});

test('XLSX REPORT propagates a nested tablix CanGrow height into its editable parent grid', async () => {
  const model = parseRdl(rdl);
  const baseline = await renderExcel(model, {
    ...request,
    datasets: { D: [{ Key: 'A', Detail: 'Short' }] },
  }, config, null);
  const longText = Array.from(
    { length: 28 },
    (_, index) => `Nested line ${String(index + 1).padStart(2, '0')} must remain visible`,
  ).join('\n');
  const grown = await renderExcel(model, {
    ...request,
    datasets: { D: [{ Key: 'A', Detail: longText }] },
  }, config, null);
  const rowHeightTotal = async (buffer) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    let total = 0;
    workbook.worksheets[0].eachRow({ includeEmpty: true }, (row) => { total += row.height || 0; });
    return { workbook, total };
  };
  const baselineResult = await rowHeightTotal(baseline.buffer);
  const grownResult = await rowHeightTotal(grown.buffer);
  assert.ok(
    grownResult.total > baselineResult.total + 100,
    `expected nested CanGrow content to increase physical row height (${grownResult.total} > ${baselineResult.total})`,
  );
  let renderedText = null;
  grownResult.workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => {
    const value = typeof cell.value === 'object' && cell.value?.richText
      ? cell.value.richText.map((run) => run.text).join('')
      : cell.value;
    if (value === longText) renderedText = cell;
  }));
  assert.ok(renderedText, 'expected all nested text to remain in one editable Excel cell');
  assert.equal(renderedText.alignment?.wrapText, true);
});
