import assert from 'node:assert/strict';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { resolveBundledSubreports } from '../src/rdl/subreports.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

// SSRS breaks a page at the deepest boundary that can still fill it. A parent tablix row whose cell holds a
// child data region (nested tablix or bundled subreport) is therefore not atomic: the break falls between
// the CHILD region's own rows, so the current page is filled and the remainder continues on the next page.
// Moving the whole parent row to a fresh page is correct only when KeepTogether is declared on the child
// region (or on the owning tablix member) and the row still fits there.
//
// These fixtures are synthetic and construct-driven: one parent group whose child fits, followed by one
// whose child is taller than what remains but shorter than a whole page. Nothing here is tuned to a report.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

const FIT_ROWS = 8;
const OVERFLOW_ROWS = 16;
const rows = [
  ...Array.from({ length: FIT_ROWS }, (_, index) => ({ Key: 'A', Label: `A-${String(index + 1).padStart(2, '0')}` })),
  ...Array.from({ length: OVERFLOW_ROWS }, (_, index) => ({ Key: 'B', Label: `B-${String(index + 1).padStart(2, '0')}` })),
];

const textbox = (name, value) => `<Textbox Name="${name}"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns>`
  + `<TextRun><Value>${value}</Value><Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style></TextRun>`
  + '</TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border>'
  + '<PaddingLeft>2pt</PaddingLeft><PaddingRight>2pt</PaddingRight></Style></Textbox>';

const page = '<Page><PageWidth>4in</PageWidth><PageHeight>6in</PageHeight><TopMargin>0.25in</TopMargin>'
  + '<BottomMargin>0.25in</BottomMargin><LeftMargin>0.25in</LeftMargin><RightMargin>0.25in</RightMargin></Page>';

const parentDataset = `<DataSets><DataSet Name="D"><Fields>
  <Field Name="Key"><DataField>Key</DataField></Field>
  <Field Name="Label"><DataField>Label</DataField></Field>
 </Fields><Query><CommandText>never executed</CommandText></Query></DataSet></DataSets>`;

// Parent tablix: one row per Key group; the cell holds `cellContents`, whatever child region that is.
const parentRdl = (cellContents) => `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 ${parentDataset}
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="Parent"><TablixBody>
   <TablixColumns><TablixColumn><Width>3.5in</Width></TablixColumn></TablixColumns>
   <TablixRows><TablixRow><Height>0.2in</Height><TablixCells><TablixCell><CellContents>
    ${cellContents}
   </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
  </TablixBody>
  <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
  <TablixRowHierarchy><TablixMembers><TablixMember>
   <Group Name="ParentKey"><GroupExpressions><GroupExpression>=Fields!Key.Value</GroupExpression></GroupExpressions></Group>
  </TablixMember></TablixMembers></TablixRowHierarchy>
  <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.2in</Height><Width>3.5in</Width>
  <Style><Border><Style>Solid</Style></Border></Style>
  </Tablix>
 </ReportItems><Height>0.2in</Height><Style/></Body><Width>3.5in</Width>
 ${page}
 </ReportSection></ReportSections></Report>`;

const nestedTablix = (keepTogether) => `<Tablix Name="Child">${keepTogether ? '<KeepTogether>true</KeepTogether>' : ''}<TablixBody>
      <TablixColumns><TablixColumn><Width>3.5in</Width></TablixColumn></TablixColumns>
      <TablixRows><TablixRow><Height>0.25in</Height><TablixCells>
       <TablixCell><CellContents>${textbox('ChildLabel', '=Fields!Label.Value')}</CellContents></TablixCell>
      </TablixCells></TablixRow></TablixRows>
     </TablixBody>
     <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
     <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="ChildDetails"/></TablixMember></TablixMembers></TablixRowHierarchy>
     <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.25in</Height><Width>3.5in</Width><Style/>
     </Tablix>`;

const subreportCall = (keepTogether) => `<Subreport Name="ChildCall">
  <ReportName>/Children/Child</ReportName>
  ${keepTogether ? '<KeepTogether>true</KeepTogether>' : ''}
  <Parameters><Parameter Name="GroupKey"><Value>=Fields!Key.Value</Value></Parameter></Parameters>
  <Top>0in</Top><Left>0in</Left><Height>0.25in</Height><Width>3.5in</Width>
</Subreport>`;

const childReportRdl = `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="C"><Fields>
  <Field Name="Label"><DataField>Label</DataField></Field>
 </Fields><Query><QueryParameters><QueryParameter Name="@GroupKey"><Value>=Parameters!GroupKey.Value</Value></QueryParameter></QueryParameters><CommandText>never executed</CommandText></Query></DataSet></DataSets>
 <ReportParameters><ReportParameter Name="GroupKey"><DataType>String</DataType><Prompt>GroupKey</Prompt></ReportParameter></ReportParameters>
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="ChildTable"><TablixBody>
   <TablixColumns><TablixColumn><Width>3.5in</Width></TablixColumn></TablixColumns>
   <TablixRows><TablixRow><Height>0.25in</Height><TablixCells>
    <TablixCell><CellContents>${textbox('ChildLabel', '=Fields!Label.Value')}</CellContents></TablixCell>
   </TablixCells></TablixRow></TablixRows>
  </TablixBody>
  <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
  <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="ChildDetails"/></TablixMember></TablixMembers></TablixRowHierarchy>
  <DataSetName>C</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.25in</Height><Width>3.5in</Width><Style/>
  </Tablix>
 </ReportItems><Height>0.25in</Height><Style/></Body><Width>3.5in</Width>
 ${page}
 </ReportSection></ReportSections></Report>`;

const baseRequest = {
  outputFileName: 'nested-region-page-fill',
  parameters: {},
  datasets: { D: rows },
  excel: { layoutMode: 'REPORT' },
};

const subreportRequest = {
  ...baseRequest,
  output: 'PDF',
  subreports: {
    '/Children/Child': {
      rdlBase64: Buffer.from(childReportRdl).toString('base64'),
      instances: ['A', 'B'].map((key) => ({
        parameters: { GroupKey: key },
        datasets: { C: rows.filter((row) => row.Key === key).map(({ Label }) => ({ Label })) },
      })),
    },
  },
};

async function traceFor(rdlText, request) {
  const model = parseRdl(rdlText);
  if (request.subreports) resolveBundledSubreports(model, request, config);
  const rendered = await renderPdf(model, request, config, { captureLayoutTrace: true });
  return rendered.layoutTrace;
}

const labelsOnPage = (trace, pageNumber) => new Set((trace.pages[pageNumber - 1].tablixFragments || [])
  .flatMap((fragment) => (fragment.cells || []).map((cell) => (cell.text || '').trim()))
  .filter((text) => /^[AB]-\d\d$/.test(text)));

const allLabels = (trace) => trace.pages.flatMap((tracePage, index) => [...labelsOnPage(trace, index + 1)]);

const usedHeight = (trace, pageNumber) => {
  const tracePage = trace.pages[pageNumber - 1];
  const bottom = Math.max(tracePage.bodyTop, ...(tracePage.tablixFragments || []).map((f) => f.y + f.height));
  return (bottom - tracePage.bodyTop) / (tracePage.bodyBottom - tracePage.bodyTop);
};

test('a nested tablix taller than the page remainder splits at a child row boundary and fills the page', async () => {
  const trace = await traceFor(parentRdl(nestedTablix(false)), baseRequest);
  const firstPage = labelsOnPage(trace, 1);
  // The whole "A" group fits, so the break must fall inside "B" rather than after "A".
  for (let index = 1; index <= FIT_ROWS; index += 1) {
    assert.ok(firstPage.has(`A-${String(index).padStart(2, '0')}`), `A-${index} must stay on page 1`);
  }
  assert.ok(firstPage.has('B-01'), 'the child region must start on the page it no longer fits on');
  assert.ok(
    usedHeight(trace, 1) > 0.9,
    `expected the split to fill page 1 (used ${(usedHeight(trace, 1) * 100).toFixed(1)}%)`,
  );
  // Every child row is rendered exactly once across the fragments.
  const rendered = allLabels(trace);
  assert.equal(rendered.length, FIT_ROWS + OVERFLOW_ROWS);
  assert.equal(new Set(rendered).size, FIT_ROWS + OVERFLOW_ROWS);
  assert.ok(labelsOnPage(trace, 2).has(`B-${String(OVERFLOW_ROWS).padStart(2, '0')}`), 'the tail continues on page 2');
});

test('KeepTogether on the nested tablix moves the whole parent row to a fresh page instead', async () => {
  const trace = await traceFor(parentRdl(nestedTablix(true)), baseRequest);
  const firstPage = labelsOnPage(trace, 1);
  assert.ok(firstPage.has(`A-${String(FIT_ROWS).padStart(2, '0')}`), 'the fitting group still fills page 1');
  for (let index = 1; index <= OVERFLOW_ROWS; index += 1) {
    const label = `B-${String(index).padStart(2, '0')}`;
    assert.ok(!firstPage.has(label), `${label} must move with its KeepTogether region`);
    assert.ok(labelsOnPage(trace, 2).has(label), `${label} must land on page 2`);
  }
});

test('a bundled subreport taller than the page remainder splits at a child row boundary', async () => {
  const trace = await traceFor(parentRdl(subreportCall(false)), subreportRequest);
  const firstPage = labelsOnPage(trace, 1);
  assert.ok(firstPage.has('B-01'), 'the subreport must start on the page it no longer fits on');
  assert.ok(
    usedHeight(trace, 1) > 0.9,
    `expected the split to fill page 1 (used ${(usedHeight(trace, 1) * 100).toFixed(1)}%)`,
  );
  const rendered = allLabels(trace);
  assert.equal(new Set(rendered).size, FIT_ROWS + OVERFLOW_ROWS);
});

test('KeepTogether on the invoking Subreport keeps its content on one page', async () => {
  const trace = await traceFor(parentRdl(subreportCall(true)), subreportRequest);
  const firstPage = labelsOnPage(trace, 1);
  for (let index = 1; index <= OVERFLOW_ROWS; index += 1) {
    assert.ok(!firstPage.has(`B-${String(index).padStart(2, '0')}`), 'a KeepTogether subreport must not split');
  }
  assert.ok(labelsOnPage(trace, 2).has(`B-${String(OVERFLOW_ROWS).padStart(2, '0')}`), 'it renders whole on page 2');
});

test('editable DOCX inherits the same child-row split from the canonical PDF trace', async () => {
  const model = parseRdl(parentRdl(nestedTablix(false)));
  const pdf = await renderPdf(model, baseRequest, config);
  const docx = await renderEditableDocx(model, baseRequest, config);
  assert.equal(docx.pageCount, pdf.pageCount);
  const zip = await JSZip.loadAsync(docx.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  assert.ok(!/<w:drawing[ >]/.test(documentXml), 'the split must stay native Word content, never a page image');
  const nativeText = [...documentXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((match) => match[1]).join('\n');
  for (let index = 1; index <= OVERFLOW_ROWS; index += 1) {
    assert.match(nativeText, new RegExp(`B-${String(index).padStart(2, '0')}`), 'every child row survives the split');
  }
  // Word receives one next-page section per canonical PDF page, so the split is visible as a section break.
  assert.equal((documentXml.match(/<w:sectPr[ >]/g) || []).length, pdf.pageCount);
});

test('XLSX has no pagination, so every child row of the split region stays in one continuous grid', async () => {
  const model = parseRdl(parentRdl(nestedTablix(false)));
  const xlsx = await renderExcel(model, baseRequest, config, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(xlsx.buffer);
  const values = new Set();
  workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => {
    const value = typeof cell.value === 'object' && cell.value?.richText
      ? cell.value.richText.map((run) => run.text).join('')
      : String(cell.value ?? '');
    if (value) values.add(value.trim());
  }));
  for (let index = 1; index <= OVERFLOW_ROWS; index += 1) {
    assert.ok(values.has(`B-${String(index).padStart(2, '0')}`), `XLSX must keep B-${index}`);
  }
});
