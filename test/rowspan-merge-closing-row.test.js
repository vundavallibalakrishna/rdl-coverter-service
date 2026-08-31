import assert from 'node:assert/strict';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';

// Which ROW owns the height a vertical merge adds.
//
// A merged (row-span) cell taller than the rows it spans makes its group grow. SSRS sizes a row to the
// tallest content that ENDS in it, so that growth belongs to the merge's LAST spanned row: every cell of
// that row is painted and ruled at the grown height, and the grid stays closed all the way down.
//
// This renderer appended the difference after the last row instead, as a band no row owned. The merged
// cell drew its own box through it, but the other columns' cells stopped at their natural height — so the
// band showed one tall cell flanked by columns with no borders at all, and the row's fills stopped short
// of the row it belongs to. XLSX already grows the final row of the span (`excel.js`); this is the same
// rule in the PDF layer, which DOCX_EDITABLE and DOCX_VISUAL inherit through the canonical trace.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

const textbox = (name, value) => `<Textbox Name="${name}"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns>`
  + `<TextRun><Value>${value}</Value><Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style></TextRun>`
  + '</TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style><Color>#000000</Color><Width>1pt</Width></Border>'
  + '<PaddingLeft>2pt</PaddingLeft><PaddingRight>2pt</PaddingRight></Style></Textbox>';

// One group of DETAILS physical rows, plus a merged group header whose height is the variable under test.
const report = (headerLines) => Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="DS"><Fields>
    <Field Name="G"><DataField>G</DataField><TypeName>System.String</TypeName></Field>
    <Field Name="D"><DataField>D</DataField><TypeName>System.String</TypeName></Field>
    <Field Name="E"><DataField>E</DataField><TypeName>System.String</TypeName></Field>
  </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="T"><TablixBody>
      <TablixColumns><TablixColumn><Width>1.5in</Width></TablixColumn><TablixColumn><Width>1.5in</Width></TablixColumn></TablixColumns>
      <TablixRows><TablixRow><Height>0.25in</Height><TablixCells>
        <TablixCell><CellContents>${textbox('Detail', '=Fields!D.Value')}</CellContents></TablixCell>
        <TablixCell><CellContents>${textbox('Side', '=Fields!E.Value')}</CellContents></TablixCell>
      </TablixCells></TablixRow></TablixRows></TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers>
        <TablixMember><Group Name="G"><GroupExpressions><GroupExpression>=Fields!G.Value</GroupExpression></GroupExpressions></Group>
          <TablixHeader><Size>1.5in</Size><CellContents>${textbox('GroupHeader', Array.from({ length: headerLines }, (unused, index) => `HDRLINE_${index + 1}`).join('&#xA;'))}</CellContents></TablixHeader>
          <TablixMembers>
            <TablixMember><Group Name="D"><GroupExpressions><GroupExpression>=Fields!D.Value</GroupExpression></GroupExpressions></Group></TablixMember>
          </TablixMembers>
        </TablixMember>
      </TablixMembers></TablixRowHierarchy>
      <DataSetName>DS</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.25in</Height><Width>4.5in</Width><Style/></Tablix>
  </ReportItems><Height>8in</Height><Style/></Body><Width>8in</Width>
  <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.3in</LeftMargin><RightMargin>0.3in</RightMargin><TopMargin>0.3in</TopMargin><BottomMargin>0.3in</BottomMargin></Page>
  </ReportSection></ReportSections></Report>`, 'utf8');

const DETAILS = 2;
const rows = Array.from({ length: DETAILS }, (unused, index) => ({
  G: 'G1', D: `DETAIL_${index + 1}`, E: `SIDE_${index + 1}`,
}));
const request = (output) => ({ output, outputFileName: 'merge-closing-row', parameters: {}, datasets: { DS: rows } });

const cells = async (headerLines) => {
  const rendered = await renderPdf(parseRdl(report(headerLines)), request('PDF'), config, { captureLayoutTrace: true });
  const items = rendered.layoutTrace.pages.flatMap((page) => page.items || []);
  const withText = (value) => items.find((item) => item.kind === 'tablixCell' && (item.text || '').startsWith(value));
  return {
    header: withText('HDRLINE_1'),
    detail: Array.from({ length: DETAILS }, (unused, index) => withText(`DETAIL_${index + 1}`)),
    side: Array.from({ length: DETAILS }, (unused, index) => withText(`SIDE_${index + 1}`)),
  };
};

const bottom = (item) => item.y + item.height;

test('a merge taller than its rows grows its last spanned row, not a band below it', async () => {
  const tall = await cells(14);
  assert.ok(tall.header, 'the merged group header is drawn');

  // The merge really does outgrow its rows; without that this test would prove nothing.
  const natural = tall.detail[0].height;
  assert.ok(
    tall.header.height > natural * DETAILS + 1,
    `the merged header must be taller than the rows it spans (${tall.header.height} vs ${natural * DETAILS})`,
  );

  // Every cell of the last spanned row reaches the group's bottom edge: the detail column and the
  // neighbouring column that carries no growing content of its own, not only the merged cell.
  for (const item of [tall.detail[DETAILS - 1], tall.side[DETAILS - 1]]) {
    assert.ok(
      Math.abs(bottom(item) - bottom(tall.header)) <= 0.5,
      `cell ${item.text} must close at the group bottom ${bottom(tall.header)}, got ${bottom(item)}`,
    );
  }

  // The earlier spanned rows keep their natural height and stay contiguous, so only the LAST row grew.
  assert.ok(Math.abs(bottom(tall.detail[0]) - tall.detail[1].y) <= 0.5, 'the spanned rows stay contiguous');
  assert.ok(tall.detail[DETAILS - 1].height > natural + 1, 'the last spanned row carries the growth');
});

test('a merge that fits its rows changes no row height', async () => {
  const short = await cells(1);
  assert.ok(
    Math.abs(short.detail[0].height - short.detail[DETAILS - 1].height) <= 0.5,
    'with a short merge every spanned row keeps the same height',
  );
  assert.ok(
    Math.abs(bottom(short.side[DETAILS - 1]) - bottom(short.detail[DETAILS - 1])) <= 0.5,
    'and the columns still close together',
  );
});

test('the grown closing row reaches editable Word and Excel', async () => {
  const model = parseRdl(report(14));
  const canonical = await cells(14);
  // The height the closing row must have: from its own top down to where the merge ends.
  const grownTwips = Math.round((bottom(canonical.header) - canonical.detail[DETAILS - 1].y) * 20);

  // Editable Word is built from that canonical trace: the Word row that HOLDS the last spanned detail must
  // itself carry the grown height, so its cells rule the same box the PDF does.
  const editable = await renderEditableDocx(model, request('DOCX_EDITABLE'), config);
  const documentXml = await (await JSZip.loadAsync(editable.buffer)).file('word/document.xml').async('string');
  const closingRow = documentXml.split('<w:tr>').find((chunk) => chunk.includes(`>DETAIL_${DETAILS}<`));
  assert.ok(closingRow, 'the last spanned detail is written into a Word row');
  const rowHeight = Number(closingRow.match(/<w:trHeight w:val="(\d+)"/)?.[1]);
  assert.ok(
    Math.abs(rowHeight - grownTwips) <= 20,
    `the Word row holding DETAIL_${DETAILS} must be the grown height (${grownTwips} twips), got ${rowHeight}`,
  );

  // Excel has no pages and already grows the final row of a vertical span; assert the same shape holds so
  // the two renderers cannot drift apart.
  const excel = await renderExcel(model, request('XLSX'), config, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(excel.buffer);
  const sheet = workbook.worksheets[0];
  const rowOf = (value) => {
    let found = null;
    sheet.eachRow((row) => row.eachCell((cell) => { if (cell.value === value) found = row; }));
    return found;
  };
  const first = rowOf('DETAIL_1');
  const last = rowOf(`DETAIL_${DETAILS}`);
  assert.ok(first && last, 'both detail rows are written');
  assert.ok(last.height > first.height + 1, 'the last spanned Excel row carries the merge growth');
});
