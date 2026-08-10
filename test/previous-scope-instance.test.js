import assert from 'node:assert/strict';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { evaluateExpression } from '../src/rdl/expression.js';
import { parseRdl } from '../src/rdl/parser.js';
import { cellText, cellTextbox, materializedCellContext, styleValue, tablixRows } from '../src/render/common.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

// SSRS evaluates Previous(expression) against the previous INSTANCE of the current scope. In a group
// scope that is the previous group instance; the current instance's own representative row is always its
// first row, so walking back one row inside the instance's aggregate rows resolves to Nothing on every
// instance. Reports use exactly this to draw a rule only where an outer key changes, so getting it wrong
// turns a group-boundary separator into a rule on every single row.
const CONDITIONAL_TOP = '=IIF(Previous(Fields!Owner.Value) &lt;&gt; Fields!Owner.Value, "Solid", "None")';
const BOUNDARY_COLOR = '#204060';

const detailTextbox = (name, topStyle) => `
  <Textbox Name="${name}"><CanGrow>true</CanGrow>
    <Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Detail.Value</Value>
      <Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs>
    <Style>
      <Border><Style>None</Style></Border>
      <TopBorder><Style>${topStyle}</Style><Color>${BOUNDARY_COLOR}</Color><Width>1pt</Width></TopBorder>
      <BottomBorder><Style>None</Style></BottomBorder>
      <LeftBorder><Style>None</Style></LeftBorder>
      <RightBorder><Style>None</Style></RightBorder>
    </Style>
  </Textbox>`;

const bannerTextbox = `
  <Textbox Name="OwnerBanner"><CanGrow>true</CanGrow>
    <Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Owner.Value</Value>
      <Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs>
    <Style><Border><Style>None</Style></Border></Style>
  </Textbox>`;

const PAGE = `<Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth>
  <LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin>
  <TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>`;

const DATASET = `<DataSets><DataSet Name="D"><Fields>
    <Field Name="Owner"><DataField>Owner</DataField></Field>
    <Field Name="Item"><DataField>Item</DataField></Field>
    <Field Name="Detail"><DataField>Detail</DataField></Field>
  </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>`;

// A leaf group member whose rows carry the conditional top border, optionally preceded by a static group
// header member. The static leaf inside a group is what routes the tablix through the recursive
// (advanced) materializer, so the same construct exercises both materialization paths.
function boundaryReport({ withGroupHeader, topStyle = CONDITIONAL_TOP }) {
  const rows = [
    withGroupHeader
      ? `<TablixRow><Height>0.3in</Height><TablixCells><TablixCell><CellContents>${bannerTextbox}</CellContents></TablixCell></TablixCells></TablixRow>`
      : '',
    `<TablixRow><Height>0.3in</Height><TablixCells><TablixCell><CellContents>${detailTextbox('ItemCell', topStyle)}</CellContents></TablixCell></TablixCells></TablixRow>`,
  ].join('');
  const itemMember = `<TablixMember><Group Name="ItemGroup"><GroupExpressions><GroupExpression>=Fields!Item.Value</GroupExpression></GroupExpressions></Group></TablixMember>`;
  const children = withGroupHeader ? `<TablixMember/>${itemMember}` : itemMember;
  return parseRdl(`<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  ${DATASET}
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="Boundary">
      <TablixBody>
        <TablixColumns><TablixColumn><Width>2in</Width></TablixColumn></TablixColumns>
        <TablixRows>${rows}</TablixRows>
      </TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers><TablixMember>
        <Group Name="OwnerGroup"><GroupExpressions><GroupExpression>=Fields!Owner.Value</GroupExpression></GroupExpressions></Group>
        <TablixMembers>${children}</TablixMembers>
      </TablixMember></TablixMembers></TablixRowHierarchy>
      <DataSetName>D</DataSetName><Top>0.2in</Top><Left>0.2in</Left><Width>2in</Width><Height>2in</Height><Style/>
    </Tablix>
  </ReportItems><Height>4in</Height><Style/></Body><Width>7in</Width>
  ${PAGE}
  </ReportSection></ReportSections>
</Report>`);
}

const request = {
  outputFileName: 'previous-scope-instance',
  parameters: {},
  datasets: { D: [
    { Owner: 'O1', Item: 'I1', Detail: 'D1' },
    { Owner: 'O1', Item: 'I2', Detail: 'D2' },
    { Owner: 'O2', Item: 'I3', Detail: 'D3' },
    { Owner: 'O2', Item: 'I4', Detail: 'D4' },
  ] },
};

// Resolves the conditional top border exactly as a renderer does: from the materialized cell, through the
// style helpers, never from the raw expression string.
function resolvedTopBorders(model) {
  const tablix = model.body.items.find((item) => item.name === 'Boundary');
  const { rows } = tablixRows(tablix, request, {}, model);
  return rows
    .map((row) => {
      const cell = row.cells[0];
      const border = cellTextbox(cell)?.style?.borders?.top;
      if (!border || cellTextbox(cell)?.name !== 'ItemCell') return null;
      const context = materializedCellContext(cell, row, {
        parameters: request.parameters, globals: {}, dataset: request.datasets.D, datasets: { D: request.datasets.D },
      });
      return { text: cellText(cell), style: String(styleValue(border.style, context, 'None')) };
    })
    .filter(Boolean);
}

test('Previous() resolves the previous group instance, not the first row of the current one', () => {
  for (const withGroupHeader of [false, true]) {
    const resolved = resolvedTopBorders(boundaryReport({ withGroupHeader }));
    assert.deepEqual(resolved, [
      { text: 'D1', style: 'Solid' }, // first instance of the scope: Previous is Nothing, so Nothing <> "O1"
      { text: 'D2', style: 'None' }, // same owner as the previous instance
      { text: 'D3', style: 'Solid' }, // owner changed
      { text: 'D4', style: 'None' }, // same owner again
    ], `withGroupHeader=${withGroupHeader}`);
  }
});

test('an explicit scope argument steps through instances of that named group', () => {
  const model = boundaryReport({ withGroupHeader: true });
  const tablix = model.body.items.find((item) => item.name === 'Boundary');
  const { rows } = tablixRows(tablix, request, {}, model);
  const seen = rows
    .filter((row) => cellTextbox(row.cells[0])?.name === 'ItemCell')
    .map((row) => {
      const context = materializedCellContext(row.cells[0], row, {
        parameters: {}, globals: {}, dataset: request.datasets.D, datasets: { D: request.datasets.D },
      });
      return evaluateExpression('=Previous(Fields!Owner.Value, "OwnerGroup")', context);
    });
  // A named scope steps by THAT scope: both item rows inside the first owner have no preceding owner
  // instance, and both rows inside the second owner see the first owner. This is deliberately different
  // from the no-scope form above, which steps once per item row.
  assert.deepEqual(seen, [null, null, 'O1', 'O1']);
});

test('a detail scope steps one rendered row at a time and crosses group boundaries', () => {
  const model = parseRdl(`<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  ${DATASET}
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="Boundary">
      <TablixBody>
        <TablixColumns><TablixColumn><Width>2in</Width></TablixColumn></TablixColumns>
        <TablixRows><TablixRow><Height>0.3in</Height><TablixCells><TablixCell><CellContents>
          ${detailTextbox('ItemCell', CONDITIONAL_TOP)}
        </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
      </TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers><TablixMember>
        <Group Name="OwnerGroup"><GroupExpressions><GroupExpression>=Fields!Owner.Value</GroupExpression></GroupExpressions></Group>
        <TablixMembers><TablixMember><Group Name="Details"/></TablixMember></TablixMembers>
      </TablixMember></TablixMembers></TablixRowHierarchy>
      <DataSetName>D</DataSetName><Top>0.2in</Top><Left>0.2in</Left><Width>2in</Width><Height>2in</Height><Style/>
    </Tablix>
  </ReportItems><Height>4in</Height><Style/></Body><Width>7in</Width>
  ${PAGE}
  </ReportSection></ReportSections>
</Report>`);
  // The detail scope spans the whole data region: only the very first detail row has no predecessor, and
  // the first row of the second owner still sees the last row of the first owner.
  assert.deepEqual(resolvedTopBorders(model).map((entry) => entry.style), ['Solid', 'None', 'Solid', 'None']);
});

test('row-scope functions keep the data-region scope when a renderer re-evaluates a style', () => {
  const model = boundaryReport({ withGroupHeader: true });
  const tablix = model.body.items.find((item) => item.name === 'Boundary');
  const { rows } = tablixRows(tablix, request, {}, model);
  const numbers = rows
    .filter((row) => cellTextbox(row.cells[0])?.name === 'ItemCell')
    .map((row) => {
      const context = materializedCellContext(row.cells[0], row, {
        parameters: {}, globals: {}, dataset: request.datasets.D, datasets: { D: request.datasets.D },
      });
      return evaluateExpression('=RowNumber(Nothing)', context);
    });
  // RowNumber(Nothing) counts across the data region. Losing the region scope in the render-time context
  // collapsed it to the innermost group and returned 1 for every row, which silently forced the first
  // branch of every `IIF(RowNumber(Nothing) = 1, ...)` border rule.
  assert.deepEqual(numbers, [1, 2, 3, 4]);
});

test('PDF, editable DOCX, and XLSX all draw the group-boundary rule only where the owner changes', async () => {
  const model = boundaryReport({ withGroupHeader: true });

  const pdf = await renderPdf(model, request, config, { captureLayoutTrace: true });
  const traced = pdf.layoutTrace.pages.flatMap((page) => page.items)
    .filter((item) => item.kind === 'tablixCell' && /^D\d$/.test(String(item.text || '')))
    .sort((left, right) => left.y - right.y)
    .map((item) => [item.text, item.borders?.top?.color ?? null]);
  assert.deepEqual(traced, [['D1', BOUNDARY_COLOR], ['D2', null], ['D3', BOUNDARY_COLOR], ['D4', null]]);

  const docx = await renderEditableDocx(model, request, config);
  const documentXml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  const wordRules = documentXml.match(/<w:top w:val="single" w:color="204060"[^/]*\/>/g) || [];
  assert.equal(wordRules.length, 2, 'Word must draw the boundary rule exactly twice');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load((await renderExcel(model, request, config, null)).buffer);
  const excelTops = [];
  workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => {
    if (/^D\d$/.test(String(cell.value ?? ''))) excelTops.push([cell.value, cell.border?.top?.color?.argb ?? null]);
  }));
  assert.deepEqual(excelTops, [['D1', 'FF204060'], ['D2', null], ['D3', 'FF204060'], ['D4', null]]);
});
