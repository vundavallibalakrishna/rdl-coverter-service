// A Subreport declared OUTSIDE a data region — directly on the body canvas or inside a Rectangle — is
// invoked exactly once, in the scope of the item that declares it, and SSRS lays the child report's body
// out inline at that position. These tests pin that construct down with synthetic definitions: the child
// content must appear in every renderer, keep its own coordinates and column widths, grow past the
// invoking placeholder, and carry the child body's own page breaks into the parent flow.
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { promisify } from 'node:util';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { resolveBundledSubreports } from '../src/rdl/subreports.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';
import { renderVisualDocx } from '../src/render/visualDocx.js';

const execFileAsync = promisify(execFile);
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

const childTablix = ({ name, left, width, dataset = 'ChildData', pageBreak = '' }) => `
<Tablix Name="${name}">
  <TablixBody>
    <TablixColumns><TablixColumn><Width>${width}in</Width></TablixColumn></TablixColumns>
    <TablixRows><TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents>
      <Textbox Name="${name}Cell"><CanGrow>true</CanGrow>
        <Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Label.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs>
        <Style><FontFamily>Arial</FontFamily><Border><Style>Solid</Style></Border></Style></Textbox>
    </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
  </TablixBody>
  <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
  <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="${name}Details"/></TablixMember></TablixMembers></TablixRowHierarchy>
  <DataSetName>${dataset}</DataSetName>${pageBreak}
  <Top>0in</Top><Left>${left}in</Left><Height>0.25in</Height><Width>${width}in</Width>
  <Style><Border><Style>None</Style></Border></Style>
</Tablix>`;

const childChart = `
<Chart Name="ChildChart">
  <ChartCategoryHierarchy><ChartMembers><ChartMember>
    <Group Name="ChildChartCategory"><GroupExpressions><GroupExpression>=Fields!Label.Value</GroupExpression></GroupExpressions></Group>
    <Label>=Fields!Label.Value</Label>
  </ChartMember></ChartMembers></ChartCategoryHierarchy>
  <ChartSeriesHierarchy><ChartMembers><ChartMember><Label>Amount</Label></ChartMember></ChartMembers></ChartSeriesHierarchy>
  <ChartData><ChartSeriesCollection><ChartSeries Name="AmountSeries"><ChartDataPoints><ChartDataPoint>
    <ChartDataPointValues><Y>=Fields!Amount.Value</Y></ChartDataPointValues>
    <ChartDataLabel><Label>=Fields!Amount.Value &amp; " units"</Label><Visible>true</Visible><Style/></ChartDataLabel>
    <Style/>
  </ChartDataPoint></ChartDataPoints><Type>Shape</Type><Subtype>Pie</Subtype></ChartSeries></ChartSeriesCollection></ChartData>
  <ChartAreas><ChartArea Name="Default"><Style/></ChartArea></ChartAreas>
  <ChartTitles><ChartTitle Name="Default"><Caption>CHILD_CHART_TITLE</Caption>
    <Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style></ChartTitle></ChartTitles>
  <DataSetName>ChildData</DataSetName>
  <Top>0.4in</Top><Left>0in</Left><Width>2.4in</Width><Height>1.8in</Height><Style/>
</Chart>`;

const childReport = (items, bodyHeight, bodyWidth) => `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="ChildData"><Fields>
    <Field Name="Label"><DataField>Label</DataField><TypeName>System.String</TypeName></Field>
    <Field Name="Amount"><DataField>Amount</DataField><TypeName>System.Int32</TypeName></Field>
  </Fields><Query><CommandText>never executed</CommandText></Query></DataSet></DataSets>
  <ReportParameters><ReportParameter Name="Key"><DataType>String</DataType><Prompt>Key</Prompt></ReportParameter></ReportParameters>
  <ReportSections><ReportSection>
    <Body><ReportItems>${items}</ReportItems><Height>${bodyHeight}in</Height><Style/></Body>
    <Width>${bodyWidth}in</Width>
    <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin>
      <RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`;

// One narrow child tablix. Body 1in tall, 2in wide.
const simpleChild = childReport(childTablix({ name: 'ChildTable', left: 0, width: 2 }), 1, 2);
// The child body is WIDER than the placeholder every parent below declares for it (2in).
const wideChild = childReport(childTablix({ name: 'ChildTable', left: 2.5, width: 2 }), 1, 4.5);
const chartChild = childReport(childChart, 2.4, 2.4);
const brokenChild = childReport(
  `${childTablix({ name: 'FirstTable', left: 0, width: 2, pageBreak: '<PageBreak><BreakLocation>End</BreakLocation></PageBreak>' })}
   ${childTablix({ name: 'SecondTable', left: 0, width: 2 }).replace('<Top>0in</Top>', '<Top>0.5in</Top>')}`,
  1,
  2,
);

const subreportItem = (extra = '') => `
<Subreport Name="StandaloneCall">
  <ReportName>/Children/Child</ReportName>
  <Parameters><Parameter Name="Key"><Value>ALPHA</Value></Parameter></Parameters>
  <Top>0.5in</Top><Left>0in</Left><Height>0.25in</Height><Width>2in</Width>${extra}
  <Style><Border><Style>None</Style></Border></Style>
</Subreport>`;

const trailingMarker = `
<Textbox Name="Trailing"><Paragraphs><Paragraph><TextRuns><TextRun><Value>TRAILING_MARKER</Value></TextRun></TextRuns></Paragraph></Paragraphs>
  <Top>1in</Top><Left>0in</Left><Height>0.25in</Height><Width>3in</Width>
  <Style><FontFamily>Arial</FontFamily></Style></Textbox>`;

const parentReport = (items) => `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <ReportSections><ReportSection>
    <Body><ReportItems>
      <Textbox Name="ParentTitle"><Paragraphs><Paragraph><TextRuns><TextRun><Value>PARENT_TITLE</Value></TextRun></TextRuns></Paragraph></Paragraphs>
        <Top>0in</Top><Left>0in</Left><Height>0.3in</Height><Width>4in</Width>
        <Style><FontFamily>Arial</FontFamily></Style></Textbox>
      ${items}
    </ReportItems><Height>4in</Height><Style/></Body>
    <Width>7.5in</Width>
    <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin>
      <RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`;

// The invoking Subreport sits inside a Rectangle, which is how report designers place one on a body band.
const rectangleParent = parentReport(`
  <Rectangle Name="Holder">
    <ReportItems>${subreportItem().replace('<Top>0.5in</Top>', '<Top>0in</Top>')}</ReportItems>
    <Top>0.5in</Top><Left>0in</Left><Height>0.3in</Height><Width>4in</Width>
    <Style><Border><Style>None</Style></Border></Style>
  </Rectangle>${trailingMarker}`);
// The same call placed directly on the body canvas: the construct is the Subreport, not its container.
const bodyParent = parentReport(`${subreportItem()}${trailingMarker}`);

// The same child invoked from inside a tablix cell, which is the other place SSRS evaluates a Subreport.
const cellParent = `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="Parent"><Fields>
    <Field Name="Key"><DataField>Key</DataField><TypeName>System.String</TypeName></Field>
  </Fields><Query><CommandText>never executed</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection>
    <Body><ReportItems>
      <Tablix Name="ParentTable">
        <TablixBody>
          <TablixColumns><TablixColumn><Width>2in</Width></TablixColumn></TablixColumns>
          <TablixRows><TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents>
            <Subreport Name="CellCall"><ReportName>/Children/Child</ReportName>
              <Parameters><Parameter Name="Key"><Value>=Fields!Key.Value</Value></Parameter></Parameters>
              <Top>0in</Top><Left>0in</Left><Height>0.25in</Height><Width>2in</Width>
              <Style><Border><Style>None</Style></Border></Style></Subreport>
          </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
        </TablixBody>
        <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
        <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="ParentRows"><GroupExpressions>
          <GroupExpression>=Fields!Key.Value</GroupExpression></GroupExpressions></Group></TablixMember></TablixMembers></TablixRowHierarchy>
        <DataSetName>Parent</DataSetName>
        <Top>0.5in</Top><Left>0in</Left><Height>0.25in</Height><Width>2in</Width>
        <Style><Border><Style>None</Style></Border></Style>
      </Tablix>
    </ReportItems><Height>4in</Height><Style/></Body>
    <Width>7.5in</Width>
    <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin>
      <RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`;

const childRows = (count) => Array.from(
  { length: count },
  (_, index) => ({ Label: `CHILD_ROW_${String(index).padStart(2, '0')}`, Amount: index }),
);

function prepared(parentRdl, childRdl, output, { instances, datasets = {} } = {}) {
  const model = parseRdl(Buffer.from(parentRdl, 'utf8'));
  const request = {
    output,
    parameters: {},
    datasets,
    subreports: {
      '/Children/Child': { rdlBase64: Buffer.from(childRdl).toString('base64'), instances },
    },
  };
  resolveBundledSubreports(model, request, config);
  return { model, request };
}

const oneInstance = (rows) => [{ parameters: { Key: 'ALPHA' }, datasets: { ChildData: rows } }];

async function pdfText(context, buffer, prefix) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), `rdl-${prefix}-`));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, `${prefix}.pdf`);
  await fs.writeFile(pdfPath, buffer);
  const { stdout } = await execFileAsync('pdftotext', ['-layout', pdfPath, '-']);
  return stdout;
}

test('renders a Subreport declared outside a data region in PDF, both Word modes, and Excel REPORT mode', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-standalone-subreport-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rows = childRows(3);

  const pdfInput = prepared(rectangleParent, simpleChild, 'PDF', { instances: oneInstance(rows) });
  const pdf = await renderPdf(pdfInput.model, pdfInput.request, config);
  const text = await pdfText(context, pdf.buffer, 'standalone-subreport');
  for (const row of rows) assert.equal((text.match(new RegExp(row.Label, 'g')) || []).length, 1, `${row.Label} must render once`);
  assert.match(text, /PARENT_TITLE/);

  const docxInput = prepared(rectangleParent, simpleChild, 'DOCX_EDITABLE', { instances: oneInstance(rows) });
  const docx = await renderEditableDocx(docxInput.model, docxInput.request, config, tempDir);
  const zip = await JSZip.loadAsync(docx.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  const nativeText = [...documentXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((match) => match[1]).join('|');
  // The child content must be native Word text, not a page picture of it.
  for (const row of rows) assert.match(nativeText, new RegExp(row.Label));
  assert.equal(docx.pageCount, pdf.pageCount);

  const visualInput = prepared(rectangleParent, simpleChild, 'DOCX_VISUAL', { instances: oneInstance(rows) });
  const visual = await renderVisualDocx(visualInput.model, visualInput.request, config, tempDir);
  const visualZip = await JSZip.loadAsync(visual.buffer);
  assert.equal(
    Object.keys(visualZip.files).filter((name) => /^word\/media\/.+/.test(name)).length,
    pdf.pageCount,
    'visual DOCX carries exactly one page image per canonical PDF page',
  );

  const excelInput = prepared(rectangleParent, simpleChild, 'XLSX', { instances: oneInstance(rows) });
  const excel = await renderExcel(excelInput.model, excelInput.request, config, tempDir);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(excel.buffer);
  const values = [];
  for (const worksheet of workbook.worksheets) {
    worksheet.eachRow((row) => row.eachCell((cell) => {
      try { values.push(String(cell.text ?? '')); } catch { /* a merged mirror of an empty master */ }
    }));
  }
  // REPORT mode writes the child grid as native typed cells, never as a picture of it.
  for (const row of rows) assert.ok(values.includes(row.Label), `${row.Label} must be a native Excel cell`);
  assert.equal(workbook.model.media.length, 0);
});

test('renders a Subreport placed directly on the body canvas, not only inside a container', async (context) => {
  const rows = childRows(3);
  const { model, request } = prepared(bodyParent, simpleChild, 'PDF', { instances: oneInstance(rows) });
  const pdf = await renderPdf(model, request, config);
  const text = await pdfText(context, pdf.buffer, 'body-subreport');
  for (const row of rows) assert.match(text, new RegExp(row.Label));
});

test('a standalone subreport grows past its placeholder and displaces the body items below it', async () => {
  const rows = childRows(8);
  const { model, request } = prepared(bodyParent, simpleChild, 'PDF', { instances: oneInstance(rows) });
  const pdf = await renderPdf(model, request, config, { captureLayoutTrace: true });
  const traced = pdf.layoutTrace.pages.flatMap((page) => page.items);
  const trailing = traced.find((item) => item.text === 'TRAILING_MARKER');
  const lastChildRow = traced.find((item) => item.text === 'CHILD_ROW_07');
  assert.ok(trailing, 'the item following the subreport must still be drawn');
  assert.ok(lastChildRow, 'every child row must be drawn');
  // The placeholder is 0.25in tall and the child grew to 2in. SSRS pushes the following item below the
  // RENDERED extent; keeping the declared box would overprint the child's last rows.
  assert.ok(
    trailing.y >= lastChildRow.y + lastChildRow.height - 0.5,
    `TRAILING_MARKER at ${trailing.y} must follow the grown child content ending at ${lastChildRow.y + lastChildRow.height}`,
  );
});

test('carries a page break declared in the child body into the parent flow', async (context) => {
  const rows = childRows(2);
  const { model, request } = prepared(bodyParent, brokenChild, 'PDF', { instances: oneInstance(rows) });
  const pdf = await renderPdf(model, request, config);
  assert.equal(pdf.pageCount, 2, 'the child body break must end the parent page');
  const text = await pdfText(context, pdf.buffer, 'broken-subreport');
  const [firstPage, secondPage] = text.split('\f');
  assert.match(firstPage, /CHILD_ROW_00/);
  assert.match(secondPage, /CHILD_ROW_00/);
});

test('keeps a bundled subreport at its own natural width instead of squeezing it into the placeholder', async () => {
  const rows = childRows(2);
  const { model, request } = prepared(cellParent, wideChild, 'PDF', {
    instances: [{ parameters: { Key: 'ALPHA' }, datasets: { ChildData: rows } }],
    datasets: { Parent: [{ Key: 'ALPHA' }] },
  });
  const pdf = await renderPdf(model, request, config, { captureLayoutTrace: true });
  const childCell = pdf.layoutTrace.pages
    .flatMap((page) => page.items)
    .find((item) => item.text === 'CHILD_ROW_00');
  assert.ok(childCell, 'the child row must be drawn');
  // The child column is 2in (144pt) wide and starts 2.5in into a child body invoked through a 2in
  // placeholder. SSRS never scales report content: the child keeps its declared column and overflows the
  // placeholder. Scaling it to what is left of the placeholder collapsed the column to a few points.
  assert.ok(
    Math.abs(childCell.width - 144) <= 1,
    `child column rendered ${childCell.width}pt; its declared width is 144pt`,
  );
});

test('renders a chart declared in a bundled subreport body through every renderer', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-subreport-chart-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rows = [{ Label: 'Alpha', Amount: 3 }, { Label: 'Beta', Amount: 1 }];

  for (const parent of [bodyParent, cellParent]) {
    const datasets = parent === cellParent ? { Parent: [{ Key: 'ALPHA' }] } : {};
    const { model, request } = prepared(parent, chartChild, 'PDF', { instances: oneInstance(rows), datasets });
    const pdf = await renderPdf(model, request, config);
    const text = await pdfText(context, pdf.buffer, 'subreport-chart');
    // The chart is a canvas item of the child body. Both invocation paths — inline on the parent canvas
    // and materialized inside a parent tablix cell — must draw it, not silently drop it.
    assert.match(text, /CHILD_CHART_TITLE/);
    assert.match(text, /3 units/);
  }

  const docxInput = prepared(bodyParent, chartChild, 'DOCX_EDITABLE', { instances: oneInstance(rows) });
  const docx = await renderEditableDocx(docxInput.model, docxInput.request, config, tempDir);
  const zip = await JSZip.loadAsync(docx.buffer);
  assert.ok(
    Object.keys(zip.files).some((name) => /^word\/media\/.+/.test(name)),
    'the child chart must reach Word as a drawing',
  );

  const excelInput = prepared(bodyParent, chartChild, 'XLSX', { instances: oneInstance(rows) });
  const excel = await renderExcel(excelInput.model, excelInput.request, config, tempDir);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(excel.buffer);
  assert.equal(workbook.model.media.length, 1, 'the child chart must reach Excel as exactly one picture');
});

test('does not require invocation data for a standalone subreport hidden by SSRS visibility', async () => {
  const hiddenParent = parentReport(subreportItem('<Visibility><Hidden>true</Hidden></Visibility>'));
  const { model, request } = prepared(hiddenParent, simpleChild, 'PDF', {
    instances: [{ parameters: {}, datasets: {} }],
  });
  const pdf = await renderPdf(model, request, config);
  assert.equal(pdf.pageCount, 1);
  assert.equal(pdf.buffer.subarray(0, 4).toString(), '%PDF');
});

test('fails closed when a standalone subreport has no bundled definition', async () => {
  const model = parseRdl(Buffer.from(bodyParent, 'utf8'));
  await assert.rejects(
    async () => resolveBundledSubreports(model, { output: 'PDF', parameters: {}, datasets: {} }, config),
    (error) => error.code === 'UNSUPPORTED_FEATURE',
  );
});

test('anchors a repeated child region picture at its own invoking row in Excel', async () => {
  const rows = [{ Label: 'Alpha', Amount: 3 }, { Label: 'Beta', Amount: 1 }];
  const keys = ['ALPHA', 'BETA', 'GAMMA'];
  const model = parseRdl(Buffer.from(cellParent, 'utf8'));
  const request = {
    output: 'XLSX',
    parameters: {},
    datasets: { Parent: keys.map((key) => ({ Key: key })) },
    subreports: {
      '/Children/Child': {
        rdlBase64: Buffer.from(chartChild).toString('base64'),
        instances: keys.map((key) => ({ parameters: { Key: key }, datasets: { ChildData: rows } })),
      },
    },
  };
  resolveBundledSubreports(model, request, config);
  // A chart that only exists inside a subreport invoked per row is not visible to inspection before that
  // row materializes. Passing no workspace proves the renderer still provisions one for it.
  const excel = await renderExcel(model, request, config, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(excel.buffer);
  const anchors = workbook.worksheets
    .flatMap((worksheet) => worksheet.getImages())
    .map((image) => ({ top: image.range.tl.nativeRow, bottom: image.range.br.nativeRow }));
  assert.equal(anchors.length, keys.length, 'one picture per invocation');
  // Every invocation belongs to a different parent row, so each picture must occupy its own rows. Sharing
  // one anchor stacks them and leaves every picture but the last invisible.
  const tops = anchors.map((anchor) => anchor.top);
  assert.equal(new Set(tops).size, keys.length, `pictures share an anchor: ${JSON.stringify(anchors)}`);
  const ordered = [...anchors].sort((left, right) => left.top - right.top);
  for (let index = 1; index < ordered.length; index += 1) {
    assert.ok(
      ordered[index].top >= ordered[index - 1].bottom - 1,
      `picture ${index} at row ${ordered[index].top} overlaps the one ending at ${ordered[index - 1].bottom}`,
    );
  }
});
