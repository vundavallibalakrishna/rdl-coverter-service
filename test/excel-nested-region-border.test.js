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

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

// Excel cannot merge a cell that hosts a nested data region: the child grid needs the individual cells to
// place its own rows. That physical region still represents ONE logical RDL cell, so its resolved edges
// belong on the region's perimeter. Writing all four sides onto the anchor cell instead drew the cell's
// bottom rule across its first physical row — a horizontal rule through the middle of a tall cell that
// neither SSRS nor the PDF renderer draws.
const CELL_COLOR = '#c04030';

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
        <TablixColumns><TablixColumn><Width>1.2in</Width></TablixColumn></TablixColumns>
        <TablixRows><TablixRow><Height>0.2in</Height><TablixCells><TablixCell><CellContents>
          <Textbox Name="ChildLabel"><CanGrow>true</CanGrow>
            <Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Label.Value</Value>
              <Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs>
            <Style><Border><Style>None</Style></Border></Style>
          </Textbox>
        </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
      </TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="ChildDetails"/></TablixMember></TablixMembers></TablixRowHierarchy>
      <DataSetName>ChildData</DataSetName><Height>0.2in</Height><Width>1.2in</Width>
      <Style><Border><Style>None</Style></Border></Style>
    </Tablix>
  </ReportItems><Height>0.2in</Height><Style/></Body><Width>1.2in</Width>
  <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`;

// The subreport sits in the group header, so one logical cell spans every detail row of its group — the
// same shape as a row-spanning report item that also hosts a child report.
const parentRdl = `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="Key"><DataField>Key</DataField><TypeName>System.Int32</TypeName></Field>
    <Field Name="Detail"><DataField>Detail</DataField><TypeName>System.String</TypeName></Field>
  </Fields><Query><CommandText>never executed</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="Grid">
      <TablixBody>
        <TablixColumns><TablixColumn><Width>1.5in</Width></TablixColumn></TablixColumns>
        <TablixRows><TablixRow><Height>0.3in</Height><TablixCells><TablixCell><CellContents>
          <Textbox Name="DetailCell"><CanGrow>true</CanGrow>
            <Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Detail.Value</Value>
              <Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs>
            <Style><Border><Style>Solid</Style><Color>Black</Color><Width>1pt</Width></Border></Style>
          </Textbox>
        </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
      </TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers><TablixMember>
        <Group Name="KeyGroup"><GroupExpressions><GroupExpression>=Fields!Key.Value</GroupExpression></GroupExpressions></Group>
        <TablixHeader><Size>1.2in</Size><CellContents>
          <Subreport Name="ChildCall">
            <ReportName>/Children/Child</ReportName>
            <Parameters><Parameter Name="Key"><Value>=Fields!Key.Value</Value></Parameter></Parameters>
            <Style>
              <Border><Style>Solid</Style><Color>${CELL_COLOR}</Color><Width>1pt</Width></Border>
            </Style>
          </Subreport>
        </CellContents></TablixHeader>
        <TablixMembers><TablixMember><Group Name="Details"/></TablixMember></TablixMembers>
      </TablixMember></TablixMembers></TablixRowHierarchy>
      <DataSetName>D</DataSetName><Top>0.2in</Top><Left>0.2in</Left><Width>2.7in</Width><Height>1.2in</Height>
      <Style><Border><Style>None</Style></Border></Style>
    </Tablix>
  </ReportItems><Height>3in</Height><Style/></Body><Width>7in</Width>
  <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`;

// One group key with three detail rows, so the subreport cell spans several physical Excel rows.
const rows = [
  { Key: 1, Detail: 'First' },
  { Key: 1, Detail: 'Second' },
  { Key: 1, Detail: 'Third' },
];

function prepared(output) {
  const model = parseRdl(Buffer.from(parentRdl, 'utf8'));
  const request = {
    output,
    outputFileName: 'excel-nested-region-border',
    excelLayoutMode: 'REPORT',
    parameters: {},
    datasets: { D: rows },
    subreports: {
      '/Children/Child': {
        rdlBase64: Buffer.from(childRdl, 'utf8').toString('base64'),
        instances: [{ parameters: { Key: 1 }, datasets: { ChildData: [{ Label: 'Child A' }] } }],
      },
    },
  };
  resolveBundledSubreports(model, request, config);
  return { model, request };
}

// The physical Excel rows the subreport cell occupies, in order, with the sides actually set on each.
async function subreportColumnEdges() {
  const { model, request } = prepared('XLSX');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load((await renderExcel(model, request, config, null)).buffer);
  const sheet = workbook.worksheets[0];
  // Locate the subreport column: the one immediately left of the column holding the detail text.
  let column = null;
  let firstRow = null;
  sheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, cellColumn) => {
      // 'First' anchors a merged range, so keep the FIRST occurrence: that is where the region starts.
      if (cell.value !== 'First' || column !== null) return;
      column = cellColumn - 1;
      firstRow = rowNumber;
    });
  });
  assert.ok(column && firstRow, 'the detail column must be present');
  const edges = [];
  for (let rowNumber = firstRow; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const border = sheet.getRow(rowNumber).getCell(column).border;
    if (!border) break;
    edges.push(['top', 'bottom', 'left', 'right'].filter((side) => border[side]).join('') || 'none');
  }
  return edges;
}

test('a nested-region cell closes at the bottom of its Excel region, not on its first row', async () => {
  const edges = await subreportColumnEdges();
  assert.ok(edges.length > 1, 'the subreport cell must span several physical Excel rows');
  // Top edge only on the first row, bottom edge only on the last, verticals on every row: one box.
  assert.equal(edges[0], 'topleftright');
  assert.equal(edges.at(-1), 'bottomleftright');
  for (const interior of edges.slice(1, -1)) assert.equal(interior, 'leftright');
});

test('the PDF draws that same cell as a single unbroken box', async () => {
  const { model, request } = prepared('PDF');
  const pdf = await renderPdf(model, request, config, { captureLayoutTrace: true });
  const cells = pdf.layoutTrace.pages
    .flatMap((page) => page.items)
    .filter((item) => item.kind === 'tablixCell' && item.borders?.top?.color === CELL_COLOR);
  // A single traced cell carrying all four edges over its full height — the reference the Excel region
  // must reproduce, and the reason this defect is Excel-only rather than a shared layout error.
  assert.equal(cells.length, 1);
  assert.equal(cells[0].borders.bottom?.color, CELL_COLOR);
  assert.ok(cells[0].height > 0);
});

test('editable DOCX already places that cell box on its region perimeter', async () => {
  const { model, request } = prepared('DOCX_EDITABLE');
  const docx = await renderEditableDocx(model, request, config);
  const documentXml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  const sides = (side) => (documentXml.match(new RegExp(`<w:${side} w:val="single" w:color="C04030"`, 'gi')) || []).length;
  // Word derives its grid from the PDF geometry, so the cell's top lands once — on the first sub-row —
  // while the verticals repeat down every sub-row. This is the arrangement Excel now reproduces; Excel
  // needed the explicit fix only because it cannot merge a cell that hosts a nested data region.
  assert.equal(sides('top'), 1);
  assert.ok(sides('left') > 1 && sides('left') === sides('right'));
});
