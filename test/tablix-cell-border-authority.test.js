import assert from 'node:assert/strict';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { resolveBundledSubreports } from '../src/rdl/subreports.js';
import { cellBorderStyle, tablixRows } from '../src/render/common.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

// RDL puts a cell's edges on the report item inside its CellContents. Tablix/Style/Border is the DATA
// REGION's outer rectangle — SSRS draws it once around the whole table and never replicates it into
// interior cells. Treating it as a per-cell default painted a full grid through every cell holding
// something other than a textbox (a subreport, image, or chart), including cells whose only content is
// hidden and which SSRS leaves completely blank.
const SUBREPORT_COLOR = '#3366cc';
const REGION_COLOR = '#000000';
const WRAPPED_TEXTBOX_COLOR = '#663399';

const childRdl = `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="ChildData">
    <Fields><Field Name="Label"><DataField>Label</DataField><TypeName>System.String</TypeName></Field></Fields>
    <Query><QueryParameters><QueryParameter Name="@Key"><Value>=Parameters!Key.Value</Value></QueryParameter></QueryParameters><CommandText>never executed</CommandText></Query>
  </DataSet></DataSets>
  <ReportParameters><ReportParameter Name="Key"><DataType>Integer</DataType><Prompt>Key</Prompt></ReportParameter></ReportParameters>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="ChildTable">
      <TablixBody>
        <TablixColumns><TablixColumn><Width>1.5in</Width></TablixColumn></TablixColumns>
        <TablixRows><TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents>
          <Textbox Name="ChildLabel"><CanGrow>true</CanGrow>
            <Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Label.Value</Value>
              <Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs>
            <Style><Border><Style>None</Style></Border></Style>
          </Textbox>
        </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
      </TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="ChildDetails"/></TablixMember></TablixMembers></TablixRowHierarchy>
      <DataSetName>ChildData</DataSetName><Height>0.25in</Height><Width>1.5in</Width>
      <Style><Border><Style>None</Style></Border></Style>
    </Tablix>
  </ReportItems><Height>0.25in</Height><Style/></Body><Width>1.5in</Width>
  <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`;

// The subreport mirrors the real construct: its own visible edges on three sides, no bottom edge, and a
// visibility expression that leaves it out of most rows.
const subreportCell = (hidden) => `<Subreport Name="ChildCall">
  <ReportName>/Children/Child</ReportName>
  <Parameters><Parameter Name="Key"><Value>=Fields!Key.Value</Value></Parameter></Parameters>
  ${hidden ? '<Visibility><Hidden>=IIF(Fields!Flag.Value = 1, False, True)</Hidden></Visibility>' : ''}
  <Style>
    <Border><Style>Solid</Style></Border>
    <TopBorder><Color>${SUBREPORT_COLOR}</Color><Style>Solid</Style><Width>1pt</Width></TopBorder>
    <BottomBorder><Color>${SUBREPORT_COLOR}</Color><Style>None</Style><Width>1pt</Width></BottomBorder>
    <LeftBorder><Color>${SUBREPORT_COLOR}</Color><Style>Solid</Style><Width>1pt</Width></LeftBorder>
    <RightBorder><Color>${SUBREPORT_COLOR}</Color><Style>Solid</Style><Width>1pt</Width></RightBorder>
  </Style>
</Subreport>`;

// The tablix declares a Solid region border on all four sides — the trigger. Every body cell but the
// subreport cell declares Border=None, so any black interior edge can only have come from the region style.
const parentRdl = (hidden) => `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="Key"><DataField>Key</DataField><TypeName>System.Int32</TypeName></Field>
    <Field Name="Flag"><DataField>Flag</DataField><TypeName>System.Int32</TypeName></Field>
    <Field Name="Label"><DataField>Label</DataField><TypeName>System.String</TypeName></Field>
  </Fields><Query><CommandText>never executed</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="Grid">
      <TablixBody>
        <TablixColumns><TablixColumn><Width>1.5in</Width></TablixColumn><TablixColumn><Width>1.5in</Width></TablixColumn></TablixColumns>
        <TablixRows><TablixRow><Height>0.4in</Height><TablixCells>
          <TablixCell><CellContents>
            <Textbox Name="LabelCell"><CanGrow>true</CanGrow>
              <Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Label.Value</Value>
                <Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs>
              <Style><Border><Style>None</Style></Border></Style>
            </Textbox>
          </CellContents></TablixCell>
          <TablixCell><CellContents>${subreportCell(hidden)}</CellContents></TablixCell>
        </TablixCells></TablixRow></TablixRows>
      </TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="Details"/></TablixMember></TablixMembers></TablixRowHierarchy>
      <DataSetName>D</DataSetName><Top>0.2in</Top><Left>0.2in</Left><Width>3in</Width><Height>0.8in</Height>
      <Style>
        <Border><Style>Solid</Style></Border>
        <TopBorder><Color>${REGION_COLOR}</Color><Style>Solid</Style><Width>1pt</Width></TopBorder>
        <BottomBorder><Color>${REGION_COLOR}</Color><Style>Solid</Style><Width>1pt</Width></BottomBorder>
        <LeftBorder><Color>${REGION_COLOR}</Color><Style>Solid</Style><Width>1pt</Width></LeftBorder>
        <RightBorder><Color>${REGION_COLOR}</Color><Style>Solid</Style><Width>1pt</Width></RightBorder>
      </Style>
    </Tablix>
  </ReportItems><Height>3in</Height><Style/></Body><Width>7in</Width>
  <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`;

// Row 1 shows the subreport; row 2 hides it. Both rows exist in the parent data region either way.
const rows = [
  { Key: 1, Flag: 1, Label: 'First' },
  { Key: 2, Flag: 0, Label: 'Second' },
];

function buildRequest(output = 'PDF') {
  return {
    output,
    outputFileName: 'tablix-cell-border-authority',
    parameters: {},
    datasets: { D: rows },
    subreports: {
      '/Children/Child': {
        rdlBase64: Buffer.from(childRdl, 'utf8').toString('base64'),
        instances: rows.map((row) => ({ parameters: { Key: row.Key }, datasets: { ChildData: [] } })),
      },
    },
  };
}

function prepared(hidden, output = 'PDF') {
  const model = parseRdl(Buffer.from(parentRdl(hidden), 'utf8'));
  const request = buildRequest(output);
  resolveBundledSubreports(model, request, config);
  return { model, request };
}

function preparedWrappedTextbox(output = 'PDF') {
  const { model, request } = prepared(false, output);
  const tablix = model.body.items.find((item) => item.name === 'Grid');
  const cell = tablix.rows[0].cells[0];
  const textbox = cell.items[0];
  const visible = { style: 'Solid', color: WRAPPED_TEXTBOX_COLOR, width: 1 };
  textbox.style.border = { ...visible };
  textbox.style.borders = { top: { ...visible }, right: { ...visible }, bottom: { ...visible }, left: { ...visible } };
  cell.items = [{
    type: 'Rectangle',
    name: 'StructuralWrapper',
    left: 0,
    top: 0,
    width: 108,
    height: 28.8,
    style: { borders: { top: { style: 'None' }, right: { style: 'None' }, bottom: { style: 'None' }, left: { style: 'None' } } },
    items: [textbox],
  }];
  return { model, request };
}

const subreportCellBorders = (trace) => trace.pages
  .flatMap((page) => page.items)
  .filter((item) => item.kind === 'tablixCell' && item.columnIndex === 1)
  .sort((left, right) => left.y - right.y)
  .map((item) => ({
    top: item.borders?.top?.color ?? null,
    left: item.borders?.left?.color ?? null,
    right: item.borders?.right?.color ?? null,
  }));

test('a cell holding a subreport takes that item as its border authority, not the tablix region border', () => {
  const { model, request } = prepared(false);
  const tablix = model.body.items.find((item) => item.name === 'Grid');
  const { rows: materialized } = tablixRows(tablix, request, {}, model);
  const cell = materialized[0].cells[1];
  assert.equal(cell.items[0].type, 'Subreport');
  // The tablix declares Solid black on every side; the resolved authority must be the subreport instead.
  assert.equal(cellBorderStyle(cell, tablix)?.borders?.top?.color, SUBREPORT_COLOR);
  assert.equal(cellBorderStyle(cell, tablix)?.borders?.bottom?.style, 'None');
  // A textbox cell is unchanged: its own style stays the authority.
  assert.equal(cellBorderStyle(materialized[0].cells[0], tablix)?.borders?.top?.style, 'None');
});

test('a cell whose only content is hidden contributes no border of its own', () => {
  const { model, request } = prepared(true);
  const tablix = model.body.items.find((item) => item.name === 'Grid');
  const { rows: materialized } = tablixRows(tablix, request, {}, model);
  assert.equal(cellBorderStyle(materialized[0].cells[1], tablix)?.borders?.top?.color, SUBREPORT_COLOR);
  assert.equal(cellBorderStyle(materialized[1].cells[1], tablix), null);
});

test('a Rectangle-wrapped textbox with explicit borders remains the cell border authority in PDF, DOCX, and XLSX', async () => {
  const { model, request } = preparedWrappedTextbox();
  const tablix = model.body.items.find((item) => item.name === 'Grid');
  const { rows: materialized } = tablixRows(tablix, request, {}, model);
  assert.equal(materialized[0].cells[0].containerWrapped, true);
  assert.equal(cellBorderStyle(materialized[0].cells[0], tablix)?.borders?.top?.color, WRAPPED_TEXTBOX_COLOR);

  const pdf = await renderPdf(model, request, config, { captureLayoutTrace: true });
  const pdfBorders = pdf.layoutTrace.pages.flatMap((page) => page.items)
    .filter((item) => item.kind === 'tablixCell' && item.columnIndex === 0 && item.text === 'First')
    .map((item) => item.borders?.top?.color);
  assert.deepEqual(pdfBorders, [WRAPPED_TEXTBOX_COLOR]);

  const docx = await renderEditableDocx(model, request, config);
  const documentXml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  assert.match(documentXml, /w:color="663399"/i);

  const xlsx = await renderExcel(...Object.values(preparedWrappedTextbox('XLSX')), config, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(xlsx.buffer);
  const sheet = workbook.worksheets[0];
  let border = null;
  sheet.eachRow((row) => row.eachCell((cell) => {
    if (cell.value === 'First') border = cell.border;
  }));
  assert.equal(border?.top?.color?.argb, 'FF663399');
  assert.equal(border?.bottom?.color?.argb, 'FF663399');
});

test('PDF draws the subreport edges only where the subreport renders, and keeps the region border', async () => {
  const visible = await renderPdf(...Object.values(prepared(false)), config, { captureLayoutTrace: true });
  assert.deepEqual(subreportCellBorders(visible.layoutTrace), [
    { top: SUBREPORT_COLOR, left: SUBREPORT_COLOR, right: SUBREPORT_COLOR },
    { top: SUBREPORT_COLOR, left: SUBREPORT_COLOR, right: SUBREPORT_COLOR },
  ], 'both rows render the subreport, so both carry its declared edges');

  const conditional = await renderPdf(...Object.values(prepared(true)), config, { captureLayoutTrace: true });
  assert.deepEqual(subreportCellBorders(conditional.layoutTrace), [
    { top: SUBREPORT_COLOR, left: SUBREPORT_COLOR, right: SUBREPORT_COLOR },
    // Row 2 hides the subreport. Its own edges vanish; the top does NOT fall back to the tablix border,
    // and the row-1 bottom it shares is declared None, so no rule is drawn between the two rows.
    { top: null, left: null, right: null },
  ], 'the hidden row must not inherit the data-region border');

  // Regression guard: the data region's own rectangle is still stroked around the whole tablix.
  const regionEdges = conditional.layoutTrace.pages
    .flatMap((page) => page.items)
    .filter((item) => item.traceRole === 'resolvedTablixFragmentBorder');
  assert.ok(regionEdges.length > 0, 'the tablix outer border must still be drawn');
  assert.ok(regionEdges.every((edge) => edge.line.color === REGION_COLOR));
});

test('editable DOCX and XLSX apply the same cell border authority', async () => {
  const docx = await renderEditableDocx(...Object.values(prepared(true)), config);
  const documentXml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  // Exactly one row's subreport cell renders, contributing its three coloured edges.
  assert.equal((documentXml.match(/w:color="3366CC"/gi) || []).length, 3);

  const { model, request } = prepared(true, 'XLSX');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load((await renderExcel(model, request, config, null)).buffer);
  // Locate the subreport column relative to the label cell rather than assuming the sheet origin.
  const sheet = workbook.worksheets[0];
  const borders = [];
  sheet.eachRow((row) => {
    row.eachCell((cell, column) => {
      if (cell.value !== 'First' && cell.value !== 'Second') return;
      const subreportCellBorder = row.getCell(column + 1).border || {};
      borders.push({
        top: subreportCellBorder.top?.color?.argb ?? null,
        left: subreportCellBorder.left?.color?.argb ?? null,
        right: subreportCellBorder.right?.color?.argb ?? null,
      });
    });
  });
  assert.deepEqual(borders, [
    { top: 'FF3366CC', left: 'FF3366CC', right: 'FF3366CC' },
    { top: null, left: null, right: null },
  ]);
});
