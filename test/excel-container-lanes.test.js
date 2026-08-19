// An RDL Rectangle is a coordinate container, not a flow unit: its children keep the positions the report
// declared, so two children in disjoint horizontal lanes stay side by side. The XLSX renderer honours that
// through its coordinate scheduler, but the scheduler needs a flat item list. A container holding a tablix
// used to be excluded from it and fell back to the flow path, which appends each child below the previous
// one regardless of where it sits horizontally - turning a two-column infographic into a single column with
// the whole right-hand column stranded under the left one.
//
// Expanding those containers to absolute body coordinates restores the declared layout. A container that
// paints keeps a childless copy of itself, so expansion never loses a fill or border.
//
// This is XLSX-specific: it is a property of that renderer's section scheduler. PDF lays every item out at
// its resolved coordinates and DOCX_EDITABLE inherits the canonical PDF trace; the PDF assertion below is
// the counterexample proving both already place these two columns side by side.
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import ExcelJS from 'exceljs';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const request = {
  outputFileName: 'excel-container-lanes',
  parameters: {},
  datasets: { Rows: [{ Label: 'r1' }, { Label: 'r2' }, { Label: 'r3' }, { Label: 'r4' }] },
};

function cell(name, value) {
  return `<Textbox Name="${name}">
    <CanGrow>true</CanGrow>
    <Paragraphs><Paragraph><TextRuns><TextRun><Value>${value}</Value></TextRun></TextRuns></Paragraph></Paragraphs>
    <Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style>
  </Textbox>`;
}

function tablix(name, width, cellName, cellValue) {
  return `<Tablix Name="${name}">
    <TablixBody>
      <TablixColumns><TablixColumn><Width>${width}in</Width></TablixColumn></TablixColumns>
      <TablixRows>
        <TablixRow><Height>0.25in</Height><TablixCells>
          <TablixCell><CellContents>${cell(`${cellName}Head`, cellValue)}</CellContents></TablixCell>
        </TablixCells></TablixRow>
        <TablixRow><Height>0.25in</Height><TablixCells>
          <TablixCell><CellContents>${cell(cellName, '=Fields!Label.Value')}</CellContents></TablixCell>
        </TablixCells></TablixRow>
      </TablixRows>
    </TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers>
      <TablixMember/>
      <TablixMember><Group Name="${name}Detail"/></TablixMember>
    </TablixMembers></TablixRowHierarchy>
    <DataSetName>Rows</DataSetName>
    <Top>0in</Top><Left>0in</Left><Height>0.5in</Height><Width>${width}in</Width>
    <Style><FontFamily>Arial</FontFamily><Border><Style>Solid</Style><Color>Black</Color><Width>1pt</Width></Border></Style>
  </Tablix>`;
}

// `Frame` holds two columns that never overlap horizontally. `RightBox` also paints a border, so the
// expansion has to preserve it.
const report = Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSources><DataSource Name="S"><DataSourceReference>/x</DataSourceReference></DataSource></DataSources>
  <DataSets><DataSet Name="Rows">
    <Fields><Field Name="Label"><DataField>Label</DataField><TypeName>System.String</TypeName></Field></Fields>
    <Query><DataSourceName>S</DataSourceName><CommandText>ignored</CommandText></Query>
  </DataSet></DataSets>
  <ReportSections><ReportSection>
    <Body>
      <ReportItems>
        <Rectangle Name="Frame">
          <ReportItems>
            <Rectangle Name="LeftCol">
              <ReportItems>${tablix('GridL', 2.5, 'LeftCell', 'LEFT_HEAD')}</ReportItems>
              <Top>0in</Top><Left>0in</Left><Height>2.5in</Height><Width>2.5in</Width>
              <Style><Border><Style>None</Style></Border></Style>
            </Rectangle>
            <Rectangle Name="RightBox">
              <ReportItems>${tablix('GridR', 2.5, 'RightCell', 'RIGHT_HEAD')}</ReportItems>
              <Top>0in</Top><Left>3in</Left><Height>0.6in</Height><Width>2.5in</Width>
              <Style><Border><Style>Solid</Style><Color>Red</Color><Width>1pt</Width></Border></Style>
            </Rectangle>
          </ReportItems>
          <Top>0in</Top><Left>0in</Left><Height>2.6in</Height><Width>5.5in</Width>
          <Style><Border><Style>None</Style></Border></Style>
        </Rectangle>
      </ReportItems>
      <Height>3in</Height><Style/>
    </Body>
    <Width>7in</Width>
    <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`, 'utf8');

async function worksheet() {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-lane-'));
  try {
    const rendered = await renderExcel(parseRdl(report), { ...request, excelLayoutMode: 'REPORT' }, config, tempDir);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(rendered.buffer);
    return workbook.worksheets[0];
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

function find(sheet, text) {
  let found = null;
  sheet.eachRow((row, rowNumber) => row.eachCell((target, columnNumber) => {
    const value = target.value?.richText
      ? target.value.richText.map((run) => run.text).join('')
      : String(target.value ?? '');
    if (!found && value.includes(text)) found = { rowNumber, columnNumber, cell: target };
  }));
  assert.ok(found, `expected a cell containing ${text}`);
  return found;
}

test('side-by-side containers holding tablixes keep their lanes instead of stacking', async () => {
  const sheet = await worksheet();
  const leftHead = find(sheet, 'LEFT_HEAD');
  const rightHead = find(sheet, 'RIGHT_HEAD');
  assert.ok(
    rightHead.columnNumber > leftHead.columnNumber,
    'the right column must stay to the right of the left one',
  );
  // Both columns are declared at Top=0, so their headers belong on the same worksheet row.
  assert.equal(
    rightHead.rowNumber,
    leftHead.rowNumber,
    'a container in its own horizontal lane must not be pushed below the other column',
  );
});

test('expanding a painted container preserves the border it declared', async () => {
  const sheet = await worksheet();
  const rightHead = find(sheet, 'RIGHT_HEAD');
  // RightBox declares a red 1pt outline around its tablix; expansion must not drop it.
  let sawRed = false;
  sheet.eachRow((row) => row.eachCell((target) => {
    for (const side of ['top', 'right', 'bottom', 'left']) {
      const argb = target.border?.[side]?.color?.argb;
      if (argb && /FF0000$/i.test(argb)) sawRed = true;
    }
  }));
  assert.ok(sawRed, 'the painted container must keep its declared border after expansion');
  assert.ok(rightHead.rowNumber >= 1);
});

test('PDF is the oracle: it already places the two columns side by side', async () => {
  const captured = await renderPdf(parseRdl(report), request, config, { captureLayoutTrace: true });
  const cells = captured.layoutTrace.pages[0].items.filter((item) => item.kind === 'tablixCell');
  const left = cells.find((item) => String(item.text).includes('LEFT_HEAD'));
  const right = cells.find((item) => String(item.text).includes('RIGHT_HEAD'));
  assert.ok(left && right, 'both column headers must be traced');
  assert.ok(right.x > left.x + left.width, 'they occupy disjoint horizontal lanes');
  assert.ok(Math.abs(right.y - left.y) < 0.5, 'and the same vertical band');
});
