import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { promisify } from 'node:util';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { createConverter } from '../src/index.js';

const execFileAsync = promisify(execFile);
const parentFixture = await fs.readFile(new URL('./fixtures/basic.rdl', import.meta.url), 'utf8');
const amountCell = '<Textbox Name="AmountCell"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Format(Fields!Amount.Value, "N2")</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style><FontFamily>Arial</FontFamily><TextAlign>Right</TextAlign><Border><Style>Solid</Style></Border></Style></Textbox>';
const subreportCell = `<Subreport Name="ChildCall">
  <ReportName>/Children/Child</ReportName>
  <Parameters><Parameter Name="EntityID"><Value>=Fields!Amount.Value</Value></Parameter></Parameters>
  <Height>0.25in</Height><Width>2in</Width>
</Subreport>`;
const parentRdl = parentFixture.replace(amountCell, subreportCell);
const growingParentRdl = parentRdl.replace(
  '<Textbox Name="NameCell">',
  '<Textbox Name="NameCell"><CanGrow>true</CanGrow>',
);
const rowSpanParentRdl = parentFixture.replace(
  '<TablixRowHierarchy><TablixMembers><TablixMember/><TablixMember><Group Name="Details"/></TablixMember></TablixMembers></TablixRowHierarchy>',
  `<TablixRowHierarchy><TablixMembers>
    <TablixMember/>
    <TablixMember>
      <Group Name="ParentName"><GroupExpressions><GroupExpression>=Fields!Name.Value</GroupExpression></GroupExpressions></Group>
      <TablixHeader><Size>2in</Size><CellContents>${subreportCell}</CellContents></TablixHeader>
      <TablixMembers><TablixMember><Group Name="Details"/></TablixMember></TablixMembers>
    </TablixMember>
  </TablixMembers></TablixRowHierarchy>`,
);

const childRdl = `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <Description>Bundled Child</Description>
  <DataSets>
    <DataSet Name="ChildData">
      <Fields>
        <Field Name="EntityID"><DataField>EntityID</DataField><TypeName>System.Int32</TypeName></Field>
        <Field Name="Label"><DataField>Label</DataField><TypeName>System.String</TypeName></Field>
      </Fields>
      <Query><QueryParameters><QueryParameter Name="@EntityID"><Value>=Parameters!EntityID.Value</Value></QueryParameter></QueryParameters><CommandText>never executed</CommandText></Query>
    </DataSet>
  </DataSets>
  <ReportParameters>
    <ReportParameter Name="EntityID"><DataType>Integer</DataType><Prompt>EntityID</Prompt></ReportParameter>
  </ReportParameters>
  <ReportSections><ReportSection>
    <Body><ReportItems>
      <Tablix Name="ChildTable">
        <TablixBody>
          <TablixColumns><TablixColumn><Width>2in</Width></TablixColumn></TablixColumns>
          <TablixRows><TablixRow><Height>0.25in</Height><TablixCells>
            <TablixCell><CellContents><Textbox Name="ChildLabel"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Label.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style><FontFamily>Arial</FontFamily><Border><Style>Solid</Style></Border></Style></Textbox></CellContents></TablixCell>
          </TablixCells></TablixRow></TablixRows>
        </TablixBody>
        <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
        <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="ChildDetails"/></TablixMember></TablixMembers></TablixRowHierarchy>
        <DataSetName>ChildData</DataSetName><Height>0.25in</Height><Width>2in</Width><Style><Border><Style>None</Style></Border></Style>
      </Tablix>
    </ReportItems><Height>1in</Height><Style/></Body>
    <Width>2in</Width>
    <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`;

function bundledRequest(instances) {
  return {
    output: 'PDF',
    outputFileName: 'subreport-test',
    parameters: { Title: 'Parent', Choice: 'A' },
    datasets: {
      Sales: [
        { Name: 'Parent One', Amount: 1 },
        { Name: 'Parent Two', Amount: 2 },
      ],
    },
    subreports: {
      '/Children/Child': {
        rdlBase64: Buffer.from(childRdl).toString('base64'),
        instances,
      },
    },
  };
}

async function extractPdfText(context, buffer, prefix) {
  await fs.mkdir(path.resolve('tmp'), { recursive: true });
  const pdfPath = path.resolve('tmp', `${prefix}-${randomUUID()}.pdf`);
  context.after(() => fs.rm(pdfPath, { force: true }));
  await fs.writeFile(pdfPath, buffer);
  return (await execFileAsync('pdftotext', [pdfPath, '-'])).stdout;
}

test('renders invocation-scoped bundled subreports inside parent tablix cells', async (context) => {
  const converter = await createConverter({
    env: { ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '30000' },
  });
  context.after(() => converter.close());
  const result = await converter.render({
    rdl: parentRdl,
    ...bundledRequest([
      { parameters: { EntityID: 1 }, datasets: { ChildData: [{ EntityID: 1, Label: 'CHILD_ALPHA' }] } },
      { parameters: { EntityID: 2 }, datasets: { ChildData: [{ EntityID: 2, Label: 'CHILD_BETA' }] } },
    ]),
  });
  const extracted = await extractPdfText(context, result.buffer, 'subreport-cell-test');
  assert.match(extracted, /Parent One/);
  assert.match(extracted, /CHILD_ALPHA/);
  assert.match(extracted, /Parent Two/);
  assert.match(extracted, /CHILD_BETA/);
});

test('renders a bundled subreport inside a row-header cell spanning parent detail rows', async (context) => {
  const converter = await createConverter({
    env: { ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '30000' },
  });
  context.after(() => converter.close());
  const request = bundledRequest([
    { parameters: { EntityID: 1 }, datasets: { ChildData: [{ EntityID: 1, Label: 'ROW_SPAN_CHILD' }] } },
  ]);
  request.datasets.Sales = [
    { Name: 'Grouped Parent', Amount: 1 },
    { Name: 'Grouped Parent', Amount: 1 },
  ];
  const result = await converter.render({ rdl: rowSpanParentRdl, ...request });
  const extracted = await extractPdfText(context, result.buffer, 'subreport-row-span-test');
  assert.equal((extracted.match(/ROW_SPAN_CHILD/g) || []).length, 1);
});

test('paginates a bundled child tablix taller than the parent printable page', async (context) => {
  const converter = await createConverter({
    env: { ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '30000' },
  });
  context.after(() => converter.close());
  const request = bundledRequest([
    {
      parameters: { EntityID: 1 },
      datasets: {
        ChildData: Array.from({ length: 80 }, (_, index) => ({
          EntityID: 1,
          Label: `TALL_CHILD_${String(index + 1).padStart(3, '0')}`,
        })),
      },
    },
  ]);
  request.datasets.Sales = [{ Name: 'Tall Parent', Amount: 1 }];
  const result = await converter.render({ rdl: parentRdl, ...request });
  assert.equal(result.pageCount > 1, true);
  const extracted = await extractPdfText(context, result.buffer, 'subreport-tall-test');
  for (let index = 1; index <= 80; index += 1) {
    const marker = `TALL_CHILD_${String(index).padStart(3, '0')}`;
    assert.equal((extracted.match(new RegExp(marker, 'g')) || []).length, 1);
  }
});

test('fails closed when an invoked subreport instance is missing', async (context) => {
  const converter = await createConverter({
    env: { ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '30000' },
  });
  context.after(() => converter.close());
  await assert.rejects(
    converter.render({
      rdl: parentRdl,
      ...bundledRequest([
        { parameters: { EntityID: 1 }, datasets: { ChildData: [{ EntityID: 1, Label: 'ONLY_ONE' }] } },
      ]),
    }),
    (error) => error.code === 'DATASET_MISSING',
  );
});

test('fails closed when the parent references an unbundled subreport', async (context) => {
  const converter = await createConverter({
    env: { ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '30000' },
  });
  context.after(() => converter.close());
  const request = bundledRequest([]);
  delete request.subreports;
  await assert.rejects(
    converter.render({ rdl: parentRdl, ...request }),
    (error) => error.code === 'UNSUPPORTED_FEATURE',
  );
});

test('renders bundled subreports through the PDF trace and as native cells in Excel REPORT mode', async (context) => {
  const converter = await createConverter({
    env: { ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '30000' },
  });
  context.after(() => converter.close());
  const request = bundledRequest([
    { parameters: { EntityID: 1 }, datasets: { ChildData: [{ EntityID: 1, Label: 'CHILD_ALPHA' }] } },
    { parameters: { EntityID: 2 }, datasets: { ChildData: [{ EntityID: 2, Label: 'CHILD_BETA' }] } },
  ]);
  const docx = await converter.render({ rdl: parentRdl, ...request, output: 'DOCX_EDITABLE' });
  const zip = await JSZip.loadAsync(docx.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  const nativeText = [...documentXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
    .map((match) => match[1])
    .join('');
  assert.match(nativeText, /CHILD_ALPHA/);
  assert.match(nativeText, /CHILD_BETA/);

  // REPORT is the default Excel mode. The child tablix must remain native/editable and retain each
  // invocation's data instead of being flattened, omitted, or converted to a drawing.
  const xlsx = await converter.render({ rdl: parentRdl, ...request, output: 'XLSX' });
  assert.equal(xlsx.layoutMode, 'report-sections');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(xlsx.buffer);
  const values = [];
  for (const worksheet of workbook.worksheets) {
    worksheet.eachRow((row) => row.eachCell((cell) => values.push(String(cell.text))));
  }
  assert.ok(values.includes('CHILD_ALPHA'));
  assert.ok(values.includes('CHILD_BETA'));
  assert.equal(workbook.model.media.length, 0);

  // DATA mode cannot preserve a child grid within its invoking parent cell, so it must remain fail-closed.
  await assert.rejects(
    converter.render({
      rdl: parentRdl,
      ...request,
      output: 'XLSX',
      excel: { layoutMode: 'DATA' },
    }),
    (error) => error.code === 'UNSUPPORTED_FEATURE',
  );
  await assert.rejects(
    converter.render({
      rdl: parentRdl,
      ...request,
      output: 'XLSX',
      excel: { sheetPerTablix: true },
    }),
    (error) => error.code === 'UNSUPPORTED_FEATURE',
  );
});

test('preserves a child-grid bottom edge when another parent cell grows past it in Excel REPORT mode', async (context) => {
  const converter = await createConverter({
    env: { ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '30000' },
  });
  context.after(() => converter.close());
  const request = bundledRequest([
    { parameters: { EntityID: 1 }, datasets: { ChildData: [{ EntityID: 1, Label: 'SHORT_CHILD_ALPHA' }] } },
    { parameters: { EntityID: 2 }, datasets: { ChildData: [{ EntityID: 2, Label: 'SHORT_CHILD_BETA' }] } },
  ]);
  request.datasets.Sales = [
    { Name: 'A growing parent cell '.repeat(20), Amount: 1 },
    { Name: 'Another growing parent cell '.repeat(20), Amount: 2 },
  ];
  const rendered = await converter.render({ rdl: growingParentRdl, ...request, output: 'XLSX' });
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(rendered.buffer);
  const values = [];
  workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => values.push(cell.text)));
  assert.ok(values.includes('SHORT_CHILD_ALPHA'));
  assert.ok(values.includes('SHORT_CHILD_BETA'));
  assert.ok(workbook.worksheets[0].rowCount > 4, 'grown parent rows should retain child-grid split boundaries');
});
