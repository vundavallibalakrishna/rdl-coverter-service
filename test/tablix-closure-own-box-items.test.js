import assert from 'node:assert/strict';
import test from 'node:test';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { shouldEnforceTablixBottom, tablixRows } from '../src/render/common.js';
import { renderPdf } from '../src/render/pdf.js';

// What counts as "this tablix is a bordered grid" when deciding to close a data fragment.
//
// A data tablix whose cells declare grid edges is closed with a rule where it ends, so a bordered grid is
// never left open. The test for that intent reads the cells' direct content — and it used to accept ANY
// item whose Style declared a border. A `Line` report item always does: RDL gives a line no stroke
// property, so its rule IS its Style.Border. A borderless form/layout tablix that merely holds separator
// lines therefore read as a bordered grid, and, finding no border anywhere to reuse, the closure fell back
// to a synthesized solid black rule across the full width — after the last field on the page, or, when the
// block happened to end low, hard above the page footer, where it looked like a footer border the report
// never declared.
//
// A Line draws itself at its own coordinates, so its border tells you nothing about the cell holding it.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000' });

const textbox = (name, value, border = 'None') => `<Textbox Name="${name}"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns>`
  + `<TextRun><Value>${value}</Value><Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style></TextRun>`
  + '</TextRuns></Paragraph></Paragraphs>'
  + `<Style><Border><Style>${border}</Style><Color>#000000</Color><Width>1pt</Width></Border></Style></Textbox>`;

// A borderless one-column data tablix whose detail cell also holds one of the "own box" items.
const report = (extra, cellBorder = 'None') => Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="DS"><Fields><Field Name="V"><DataField>V</DataField><TypeName>System.String</TypeName></Field>
  </Fields><Query><CommandText>never executed</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="T"><TablixBody>
      <TablixColumns><TablixColumn><Width>4in</Width></TablixColumn></TablixColumns>
      <TablixRows><TablixRow><Height>0.4in</Height><TablixCells><TablixCell><CellContents>
        <Rectangle Name="Cell"><ReportItems>
          ${textbox('Value', '=Fields!V.Value', cellBorder)}
          ${extra}
        </ReportItems><Top>0in</Top><Left>0in</Left><Width>4in</Width><Height>0.4in</Height>
        <Style><Border><Style>None</Style></Border></Style></Rectangle>
      </CellContents></TablixCell></TablixCells></TablixRow></TablixRows></TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers><TablixMember>
        <Group Name="D"><GroupExpressions><GroupExpression>=Fields!V.Value</GroupExpression></GroupExpressions></Group>
      </TablixMember></TablixMembers></TablixRowHierarchy>
      <DataSetName>DS</DataSetName><Top>0in</Top><Left>0in</Left><Width>4in</Width><Height>0.4in</Height>
      <Style><Border><Style>None</Style></Border></Style></Tablix>
  </ReportItems><Height>6in</Height><Style/></Body><Width>6in</Width>
  <Page><PageHeight>8in</PageHeight><PageWidth>6in</PageWidth><LeftMargin>0.3in</LeftMargin><RightMargin>0.3in</RightMargin>
    <TopMargin>0.3in</TopMargin><BottomMargin>0.3in</BottomMargin></Page>
  </ReportSection></ReportSections></Report>`, 'utf8');

const LINE = '<Line Name="Rule"><Top>0.3in</Top><Left>0in</Left><Height>0in</Height><Width>3in</Width>'
  + '<Style><Border><Style>Solid</Style><Color>#000000</Color><Width>1pt</Width></Border></Style></Line>';

const rows = [{ V: 'ROW_1' }, { V: 'ROW_2' }, { V: 'ROW_3' }];
const request = { outputFileName: 'closure-own-box', parameters: {}, datasets: { DS: rows } };

const intent = (definition) => {
  const model = parseRdl(definition);
  const tablix = model.body.items.find((item) => item.type === 'Tablix');
  return shouldEnforceTablixBottom(tablixRows(tablix, request, { PageNumber: 1, TotalPages: 1 }, model).rows, tablix);
};

const closures = async (definition) => {
  const rendered = await renderPdf(parseRdl(definition), request, config, { captureLayoutTrace: true });
  return rendered.layoutTrace.pages.flatMap((page) => (page.items || [])
    .filter((item) => item.traceRole === 'resolvedTablixFragmentBorder' && item.fragmentSide === 'bottom'));
};

test('a separator Line in a cell does not make a borderless tablix a bordered grid', async () => {
  assert.equal(intent(report(LINE)), false, 'a line is a rule of its own, not a declared cell edge');
  assert.deepEqual(await closures(report(LINE)), [], 'so no closing rule is synthesized after the last row');

  // The line itself is still drawn — only the invented full-width rule is gone.
  const rendered = await renderPdf(parseRdl(report(LINE)), request, config, { captureLayoutTrace: true });
  const lines = rendered.layoutTrace.pages.flatMap((page) => (page.items || []).filter((item) => item.itemName === 'Rule'));
  assert.equal(lines.length, rows.length, 'every row still draws its own separator line');
});

test('a cell that really declares a grid edge is still closed', async () => {
  // The counterexample: the same tablix with the border on the textbox that fills the cell — the ordinary
  // way an RDL grid declares its edges — keeps its closing rule, and reuses the declared style rather than
  // inventing one.
  assert.equal(intent(report(LINE, 'Solid')), true);
  const closing = await closures(report(LINE, 'Solid'));
  assert.equal(closing.length, 1, 'the bordered grid closes exactly once, where it ends');
  assert.equal(closing[0].line.style, 'Solid');
  assert.equal(closing[0].line.color, '#000000');
});
