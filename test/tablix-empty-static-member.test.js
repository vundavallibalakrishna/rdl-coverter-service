import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import { promisify } from 'node:util';
import test from 'node:test';
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

function emptyDataRdl({ hideIfNoRows = false } = {}) {
  return `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields>
  <Field Name="Description"><DataField>Description</DataField></Field>
 </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="T"><TablixBody>
   <TablixColumns><TablixColumn><Width>3in</Width></TablixColumn></TablixColumns>
   <TablixRows>
    <TablixRow><Height>0.3in</Height><TablixCells><TablixCell><CellContents>
     <Textbox Name="fallback"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun>
      <Value>=IIF(Fields!Description.Value &lt;&gt; Nothing, Fields!Description.Value, "DEFAULT EMPTY TEXT")</Value>
     </TextRun></TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border></Style></Textbox>
    </CellContents></TablixCell></TablixCells></TablixRow>
    <TablixRow><Height>0.3in</Height><TablixCells><TablixCell><CellContents>
     <Textbox Name="detail"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Description.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox>
    </CellContents></TablixCell></TablixCells></TablixRow>
   </TablixRows>
  </TablixBody>
  <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
  <TablixRowHierarchy><TablixMembers>
   <TablixMember>${hideIfNoRows ? '<HideIfNoRows>true</HideIfNoRows>' : ''}</TablixMember>
   <TablixMember><Group Name="Details"/></TablixMember>
  </TablixMembers></TablixRowHierarchy>
  <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.6in</Height><Width>3in</Width><Style/>
  </Tablix>
 </ReportItems><Height>2in</Height><Style/></Body>
 <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
 </ReportSection></ReportSections>
</Report>`;
}

const tablixOf = (model) => model.body.items.find((item) => item.type === 'Tablix');
const rowText = (row) => row.cells.flatMap((cell) => cell.values || []).join('');

test('a static member evaluates its Fields fallback once when the dataset is empty', () => {
  const tablix = tablixOf(parseRdl(emptyDataRdl()));
  const rows = materializeTablixRows(tablix, [], {}, {}, { D: [] });
  assert.equal(rows.length, 1);
  assert.equal(rows[0].isStatic, true);
  assert.equal(rowText(rows[0]), 'DEFAULT EMPTY TEXT');
});

test('HideIfNoRows hides an otherwise-visible static member while detail members remain data-driven', () => {
  const hiddenTablix = tablixOf(parseRdl(emptyDataRdl({ hideIfNoRows: true })));
  assert.equal(materializeTablixRows(hiddenTablix, [], {}, {}, { D: [] }).length, 0);

  const visibleTablix = tablixOf(parseRdl(emptyDataRdl()));
  const source = [{ Description: 'First' }, { Description: 'Second' }];
  const rows = materializeTablixRows(visibleTablix, source, {}, {}, { D: source });
  assert.deepEqual(rows.map(rowText), ['First', 'First', 'Second']);
  assert.deepEqual(rows.map((row) => row.role), ['static', 'detail', 'detail']);
});

test('the static fallback handles null and empty field values without replacing populated values', () => {
  const tablix = tablixOf(parseRdl(emptyDataRdl()));
  const materialize = (Description) => materializeTablixRows(
    tablix,
    [{ Description }],
    {},
    {},
    { D: [{ Description }] },
  );
  assert.equal(rowText(materialize(null)[0]), 'DEFAULT EMPTY TEXT');
  assert.equal(rowText(materialize('')[0]), 'DEFAULT EMPTY TEXT');
  assert.equal(rowText(materialize('Supplied text')[0]), 'Supplied text');
});

test('the shared empty-static-member semantics reach PDF, editable DOCX, and XLSX', async () => {
  const model = parseRdl(emptyDataRdl());
  const request = { outputFileName: 'empty-static-member', parameters: {}, datasets: { D: [] } };

  const pdf = await renderPdf(model, request, config);
  const pdfPath = new URL(`../tmp/empty-static-member-${process.pid}.pdf`, import.meta.url);
  await fs.mkdir(new URL('../tmp/', import.meta.url), { recursive: true });
  await fs.writeFile(pdfPath, pdf.buffer);
  try {
    const extracted = await execFileAsync('pdftotext', [pdfPath.pathname, '-']);
    assert.match(extracted.stdout, /DEFAULT EMPTY TEXT/);
  } finally {
    await fs.unlink(pdfPath).catch(() => {});
  }

  const docx = await renderEditableDocx(model, request, config);
  const documentXml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  const nativeText = [...documentXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
    .map((match) => match[1])
    .join('');
  assert.match(nativeText, /DEFAULT EMPTY TEXT/);

  const xlsx = await renderExcel(model, request, config, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(xlsx.buffer);
  const values = [];
  workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => values.push(cell.value)));
  assert.ok(values.includes('DEFAULT EMPTY TEXT'));
});
