import assert from 'node:assert/strict';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';

// Where a vertical merge's extra height goes.
//
// A row-header cell that merges every physical row of its group is content of ALL those rows, not of the
// row it happens to start in. When the merged cell is taller than the rows it spans, SSRS leaves each
// spanned row at its own natural height and grows the GROUP past its last row so the merge can finish.
//
// This renderer sized the merge's first spanned row from the merge's whole content instead, so every later
// row of the group was pushed down by it: with a tall group header the second detail row started hundreds
// of points below its own data and no longer lined up with the header it belongs to.
//
// The rule is a property of the merge, not of what fills it, so it is exercised here for a merged VALUE
// and for a merged child data region, plus the counterexample where the merge is shorter than its rows.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

const textbox = (name, value, extra = '') => `<Textbox Name="${name}"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns>`
  + `<TextRun><Value>${value}</Value><Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style></TextRun>`
  + `</TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border>${extra}`
  + '<PaddingLeft>2pt</PaddingLeft><PaddingRight>2pt</PaddingRight></Style></Textbox>';

// The group-header cell is either a textbox carrying `headerValue`, or a static child tablix of
// `childRows` rows — the two ways a merge can outgrow the rows beneath it.
const report = ({ headerValue = null, childRows = 0, childRowHeight = 0.4, pageHeight = 11, keepTogether = false }) => Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="DS"><Fields>
    <Field Name="G"><DataField>G</DataField><TypeName>System.String</TypeName></Field>
    <Field Name="D"><DataField>D</DataField><TypeName>System.String</TypeName></Field>
  </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="T"><TablixBody>
      <TablixColumns><TablixColumn><Width>3in</Width></TablixColumn></TablixColumns>
      <TablixRows>
        <TablixRow><Height>0.3in</Height><TablixCells><TablixCell><CellContents>${textbox('Body1', '=Fields!D.Value')}</CellContents></TablixCell></TablixCells></TablixRow>
      </TablixRows></TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers>
        <TablixMember><Group Name="G"><GroupExpressions><GroupExpression>=Fields!G.Value</GroupExpression></GroupExpressions></Group>${keepTogether ? '<KeepTogether>true</KeepTogether>' : ''}
          <TablixHeader><Size>1.5in</Size><CellContents>${headerValue !== null ? textbox('GroupHeader', headerValue) : `
            <Tablix Name="Child"><TablixBody>
              <TablixColumns><TablixColumn><Width>1.5in</Width></TablixColumn></TablixColumns>
              <TablixRows>${Array.from({ length: childRows }, (unused, index) => (
    `<TablixRow><Height>${childRowHeight}in</Height><TablixCells><TablixCell><CellContents>`
                + `${textbox(`ChildCell${index + 1}`, `CHILD_${index + 1}`)}</CellContents></TablixCell></TablixCells></TablixRow>`
  )).join('')}</TablixRows>
            </TablixBody>
            <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
            <TablixRowHierarchy><TablixMembers>${'<TablixMember/>'.repeat(childRows)}</TablixMembers></TablixRowHierarchy>
            <Top>0in</Top><Left>0in</Left><Height>${childRowHeight * childRows}in</Height><Width>1.5in</Width><Style/></Tablix>`}
          </CellContents></TablixHeader>
          <TablixMembers>
            <TablixMember><Group Name="D"><GroupExpressions><GroupExpression>=Fields!D.Value</GroupExpression></GroupExpressions></Group></TablixMember>
          </TablixMembers>
        </TablixMember>
      </TablixMembers></TablixRowHierarchy>
      <DataSetName>DS</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.3in</Height><Width>4.5in</Width><Style/></Tablix>
  </ReportItems><Height>6in</Height><Style/></Body><Width>8in</Width>
  <Page><PageHeight>${pageHeight}in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.3in</LeftMargin><RightMargin>0.3in</RightMargin><TopMargin>0.3in</TopMargin><BottomMargin>0.3in</BottomMargin></Page>
  </ReportSection></ReportSections></Report>`, 'utf8');

const DETAILS = 3;
const rows = Array.from({ length: DETAILS }, (unused, index) => ({ G: 'G1', D: `DETAIL_${index + 1}` }));
const request = (output) => ({ output, outputFileName: 'merge-growth', parameters: {}, datasets: { DS: rows } });

const trace = async (definition) => renderPdf(
  parseRdl(report(definition)),
  request('PDF'),
  config,
  { captureLayoutTrace: true },
);
const items = (rendered) => rendered.layoutTrace.pages.flatMap((page) => page.items || []);
const byText = (rendered, value) => items(rendered).find((item) => (item.text || '').startsWith(value));
// The detail rows' own tops, in order. These are what a merge must not displace.
const detailTops = (rendered) => Array.from(
  { length: DETAILS },
  (unused, index) => byText(rendered, `DETAIL_${index + 1}`).y,
);

// A one-line group header is shorter than the three detail rows, so nothing grows: this fixes the detail
// row pitch that the oversized cases below must reproduce exactly.
const baseline = await trace({ headerValue: 'short' });
const BASE_TOPS = detailTops(baseline);
const BASE_PITCH = BASE_TOPS[1] - BASE_TOPS[0];

test('the detail rows of a group are unmoved by a merged value taller than they are', async () => {
  // Enough lines in the merged group header to exceed all three detail rows put together.
  const tall = await trace({ headerValue: Array.from({ length: 40 }, (unused, index) => `line ${index + 1}`).join('&#xA;') });
  const tops = detailTops(tall);

  assert.deepEqual(
    tops.map((top) => Number((top - tops[0]).toFixed(2))),
    BASE_TOPS.map((top) => Number((top - BASE_TOPS[0]).toFixed(2))),
    'each spanned row keeps its natural height; the merge must not push the rows below it down',
  );

  // The group did grow — the merged cell is taller than its three rows, and it is the merge that carries
  // the extra height, not the first row.
  const header = byText(tall, 'line 1');
  assert.equal(header.height > tops[DETAILS - 1] - tops[0] + BASE_PITCH + 1, true, 'the group grew for the merge');
  assert.equal(Math.abs(header.y - tops[0]) < 0.5, true, 'the merge still starts at its first spanned row');
});

test('the detail rows of a group are unmoved by a merged child region taller than they are', async () => {
  const tall = await trace({ childRows: 12 });
  const tops = detailTops(tall);

  assert.deepEqual(
    tops.map((top) => Number((top - tops[0]).toFixed(2))),
    BASE_TOPS.map((top) => Number((top - BASE_TOPS[0]).toFixed(2))),
    'a child data region in a merged cell must not inflate the row it starts in',
  );

  // Every child row is drawn, below its predecessor, starting at the merge's own top.
  const childTops = Array.from({ length: 12 }, (unused, index) => byText(tall, `CHILD_${index + 1}`).y);
  assert.equal(Math.abs(childTops[0] - tops[0]) < 0.5, true);
  for (let index = 1; index < childTops.length; index += 1) {
    assert.equal(childTops[index] > childTops[index - 1], true, `CHILD_${index + 1} must follow CHILD_${index}`);
  }
  // ...and it reaches past the last detail row, which is the whole point of growing the group.
  assert.equal(childTops[childTops.length - 1] > tops[DETAILS - 1], true);
});

test('a merge shorter than its spanned rows leaves the group exactly as tall as those rows', async () => {
  const short = await trace({ childRows: 2 });
  assert.deepEqual(detailTops(short), BASE_TOPS, 'a merge that fits must change no geometry at all');
});

test('the merged-cell growth rule reaches editable Word and Excel', async () => {
  const model = parseRdl(report({ childRows: 12 }));
  const canonical = await renderPdf(model, request('PDF'), config, { captureLayoutTrace: true });
  const canonicalTops = detailTops(canonical);

  // Editable Word is built from the canonical PDF trace. Word rows are exact-height twips, so the row
  // holding the first detail must not have absorbed the merge: its height stays the natural row pitch.
  const editable = await renderEditableDocx(model, request('DOCX_EDITABLE'), config);
  assert.equal(editable.pageCount, canonical.pageCount);
  const documentXml = await (await JSZip.loadAsync(editable.buffer)).file('word/document.xml').async('string');
  for (let index = 1; index <= DETAILS; index += 1) {
    assert.match(documentXml, new RegExp(`<w:t[^>]*>DETAIL_${index}</w:t>`));
  }
  for (let index = 1; index <= 12; index += 1) {
    assert.match(documentXml, new RegExp(`<w:t[^>]*>CHILD_${index}</w:t>`));
  }
  const naturalPitch = canonicalTops[1] - canonicalTops[0];
  const firstDetailRowTwips = Math.round(naturalPitch * 20);
  const rowHeights = [...documentXml.matchAll(/<w:trHeight w:val="(\d+)"/g)].map((match) => Number(match[1]));
  assert.equal(
    rowHeights.some((height) => Math.abs(height - firstDetailRowTwips) <= 20),
    true,
    'a Word row of the natural detail height must exist — the merge did not swallow the first row',
  );

  // Excel has no pages and lays the merge out as a spanned block; every part must still be present and the
  // detail rows must stay in their own cells rather than being displaced by the merge.
  const excel = await renderExcel(model, request('XLSX'), config, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(excel.buffer);
  const values = [];
  workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => {
    if (cell.isMerged && cell.master !== cell) return;
    values.push(cell.value);
  }));
  for (let index = 1; index <= DETAILS; index += 1) assert.equal(values.filter((value) => value === `DETAIL_${index}`).length, 1);
  for (let index = 1; index <= 12; index += 1) assert.equal(values.filter((value) => value === `CHILD_${index}`).length, 1);
});

// KeepTogether reserves the WHOLE merge, not just the rows it spans. Row heights no longer carry a merge's
// content, so the reservation has to ask the merge itself — otherwise a member that declares KeepTogether
// would still be split at the page boundary by the very merge that made it too tall.
test('a KeepTogether member reserves the merge that makes its group too tall', async () => {
  const twoGroups = [
    { G: 'G0', D: 'FILLER' },
    ...Array.from({ length: DETAILS }, (unused, index) => ({ G: 'G1', D: `DETAIL_${index + 1}` })),
  ];
  const render = async (keepTogether) => renderPdf(
    parseRdl(report({ childRows: 6, pageHeight: 4.1, keepTogether })),
    { output: 'PDF', outputFileName: 'merge-keep', parameters: {}, datasets: { DS: twoGroups } },
    config,
    { captureLayoutTrace: true },
  );
  // Where G1 starts, and whether its merge finishes on that same page. CHILD_* text also appears in G0's
  // own copy of the merge, so the check is anchored to items below G1's first detail row.
  const groupPage = (rendered) => rendered.layoutTrace.pages
    .find((page) => (page.items || []).some((item) => (item.text || '').startsWith('DETAIL_1')));
  const mergeFinishesWithGroup = (rendered) => {
    const page = groupPage(rendered);
    const start = (page.items || []).find((item) => (item.text || '').startsWith('DETAIL_1')).y;
    return (page.items || []).some((item) => (item.text || '').startsWith('CHILD_6') && item.y > start);
  };

  // G1's merge is taller than its rows and than what is left of page 1, but fits a fresh page.
  const split = await render(false);
  assert.equal(groupPage(split).number, 1, 'without KeepTogether the group fills the current page');
  assert.equal(mergeFinishesWithGroup(split), false, 'and its merge continues on the next page');

  const kept = await render(true);
  assert.equal(groupPage(kept).number, 2, 'KeepTogether moves the whole group to a page that holds it');
  assert.equal(mergeFinishesWithGroup(kept), true, 'the reserved page holds the complete merge');
});

// A page break forced by a merge that cannot finish is a cut inside the group, not the end of the table.
// Closing the fragment there stroked the tablix's synthesized bottom rule hard against the printable body
// boundary — which is exactly where a page footer draws its own rule — so the two fused into one bar. The
// rule belongs at the row boundary where the table really ends, on the last page only.
const footerRuleReport = ({ childRows }) => Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="DS"><Fields>
    <Field Name="G"><DataField>G</DataField><TypeName>System.String</TypeName></Field>
    <Field Name="D"><DataField>D</DataField><TypeName>System.String</TypeName></Field>
  </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="T"><TablixBody>
      <TablixColumns><TablixColumn><Width>3in</Width></TablixColumn></TablixColumns>
      <TablixRows>
        <TablixRow><Height>0.3in</Height><TablixCells><TablixCell><CellContents>
          <Textbox Name="Body1"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns>
            <TextRun><Value>=Fields!D.Value</Value><Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style></TextRun>
          </TextRuns></Paragraph></Paragraphs>
          <Style><Border><Style>None</Style></Border><BottomBorder><Style>Dotted</Style><Color>Black</Color><Width>1pt</Width></BottomBorder></Style></Textbox>
        </CellContents></TablixCell></TablixCells></TablixRow>
      </TablixRows></TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers>
        <TablixMember><Group Name="G"><GroupExpressions><GroupExpression>=Fields!G.Value</GroupExpression></GroupExpressions></Group>
          <TablixHeader><Size>1.5in</Size><CellContents>
            <Tablix Name="Child"><TablixBody>
              <TablixColumns><TablixColumn><Width>1.5in</Width></TablixColumn></TablixColumns>
              <TablixRows>${Array.from({ length: childRows }, (unused, index) => (
    `<TablixRow><Height>0.4in</Height><TablixCells><TablixCell><CellContents>`
                + `${textbox(`ChildCell${index + 1}`, `CHILD_${index + 1}`)}</CellContents></TablixCell></TablixCells></TablixRow>`
  )).join('')}</TablixRows>
            </TablixBody>
            <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
            <TablixRowHierarchy><TablixMembers>${'<TablixMember/>'.repeat(childRows)}</TablixMembers></TablixRowHierarchy>
            <Top>0in</Top><Left>0in</Left><Height>${0.4 * childRows}in</Height><Width>1.5in</Width><Style/></Tablix>
          </CellContents></TablixHeader>
          <TablixMembers>
            <TablixMember><Group Name="D"><GroupExpressions><GroupExpression>=Fields!D.Value</GroupExpression></GroupExpressions></Group></TablixMember>
          </TablixMembers>
        </TablixMember>
      </TablixMembers></TablixRowHierarchy>
      <DataSetName>DS</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.3in</Height><Width>4.5in</Width>
      <Style><Border><Style>None</Style></Border></Style></Tablix>
  </ReportItems><Height>6in</Height><Style/></Body><Width>8in</Width>
  <Page><PageHeight>4in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.3in</LeftMargin><RightMargin>0.3in</RightMargin><TopMargin>0.3in</TopMargin><BottomMargin>0.3in</BottomMargin>
    <PageFooter><Height>0.5in</Height><PrintOnFirstPage>true</PrintOnFirstPage><PrintOnLastPage>true</PrintOnLastPage><ReportItems>
      <Line Name="FooterRule"><Top>0.05in</Top><Left>0in</Left><Height>0in</Height><Width>7.9in</Width>
        <Style><Border><Style>Solid</Style><Color>Black</Color><Width>1pt</Width></Border></Style></Line>
    </ReportItems><Style/></PageFooter>
  </Page></ReportSection></ReportSections></Report>`, 'utf8');

test('a page break forced by an unfinished merge does not rule the body boundary', async () => {
  const rendered = await renderPdf(
    parseRdl(footerRuleReport({ childRows: 12 })),
    { output: 'PDF', outputFileName: 'merge-fragment-border', parameters: {}, datasets: { DS: rows } },
    config,
    { captureLayoutTrace: true },
  );
  const pages = rendered.layoutTrace.pages;
  assert.ok(pages.length >= 2, `the merge must cross a page to isolate the cut, got ${pages.length}`);

  const closuresAtBodyBottom = (page) => (page.items || []).filter((item) => (
    item.traceRole === 'resolvedTablixFragmentBorder'
    && item.fragmentSide === 'bottom'
    && Math.abs(item.y - page.bodyBottom) <= 0.5
  ));

  // Every page but the last is cut inside the group: nothing may close there.
  for (const page of pages.slice(0, -1)) {
    assert.deepEqual(
      closuresAtBodyBottom(page).map((item) => item.itemName || item.kind),
      [],
      `page ${page.number} continues the merge, so the table must not close at the body boundary`,
    );
  }

  // The table does close, on the last page, at the row boundary where it really ends.
  const last = pages.at(-1);
  const closing = (last.items || []).filter((item) => (
    item.traceRole === 'resolvedTablixFragmentBorder' && item.fragmentSide === 'bottom'
  ));
  assert.equal(closing.length, 1, 'the table closes exactly once, where it ends');
  assert.ok(closing[0].y < last.bodyBottom - 1, 'and clear of the body boundary');

  // The footer keeps its own rule on every page; the two are never the same line.
  for (const page of pages) {
    const footer = (page.items || []).filter((item) => item.itemName === 'FooterRule');
    assert.equal(footer.length, 1, `page ${page.number} draws the report footer rule`);
    assert.ok(footer[0].y > page.bodyBottom, 'the footer rule belongs to the footer band, not the body');
  }

  // Both Word modes are built from this trace, so the closure they can draw is the one asserted above.
  // Rendering them here proves the traced geometry stays representable, page for page.
  const model = parseRdl(footerRuleReport({ childRows: 12 }));
  const wordRequest = (output) => ({ output, outputFileName: 'merge-fragment-border', parameters: {}, datasets: { DS: rows } });
  const editable = await renderEditableDocx(model, wordRequest('DOCX_EDITABLE'), config);
  assert.equal(editable.pageCount, pages.length);
  const documentXml = await (await JSZip.loadAsync(editable.buffer)).file('word/document.xml').async('string');
  assert.match(documentXml, /<w:t[^>]*>CHILD_12<\/w:t>/);

  // Excel has no page fragments at all, so the construct cannot arise there; the table still closes once,
  // on its real last row.
  const excel = await renderExcel(model, wordRequest('XLSX'), config, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(excel.buffer);
  const sheet = workbook.worksheets[0];
  const ruled = [];
  sheet.eachRow((row, number) => row.eachCell((cell) => {
    if (cell.border?.bottom && !ruled.includes(number)) ruled.push(number);
  }));
  assert.ok(ruled.length > 0, 'the Excel table keeps its row rules');
  assert.ok(ruled.at(-1) <= sheet.rowCount, 'and none of them fall past the last written row');
});
