import assert from 'node:assert/strict';
import test from 'node:test';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';

// A tablix ROW is SSRS's indivisible pagination unit.
//
// When a row does not fit what is left of the page, SSRS moves the WHOLE row to the next page; it splits a
// row only when the row cannot fit on a page at all. This engine split any row with continuation-able text
// as soon as it overflowed the remainder, which had two visible consequences: a value SSRS keeps together
// was broken across pages, and — because a cut inside a row is deliberately left open (no closing rule, see
// `tablix-split-row-open-edge.test.js`) — the table stopped flush on the printable body boundary, right
// against the page footer's own rule, instead of closing at its last complete row with the gap that leaves.
//
// The counterexample below keeps the split path alive for the case that genuinely needs it: a row taller
// than a whole page has nowhere to move to and must still be split.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000' });

// A small page with a footer that draws its own rule, so "did the table close against the footer?" is a
// question about real geometry rather than about this one report.
const rdl = `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields><Field Name="V"><DataField>V</DataField></Field></Fields>
   <Query><CommandText>never executed</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="T"><TablixBody><TablixColumns><TablixColumn><Width>3in</Width></TablixColumn></TablixColumns>
   <TablixRows>
    <TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents>
      <Textbox Name="h"><Paragraphs><Paragraph><TextRuns><TextRun><Value>COLHDR</Value>
        <Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style></TextRun></TextRuns></Paragraph></Paragraphs>
        <Style><Border><Style>Solid</Style><Color>#000000</Color><Width>1pt</Width></Border></Style></Textbox>
    </CellContents></TablixCell></TablixCells></TablixRow>
    <TablixRow><Height>0.15in</Height><TablixCells><TablixCell><CellContents>
      <Textbox Name="d"><CanGrow>true</CanGrow>
        <Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!V.Value</Value>
        <Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style></TextRun></TextRuns></Paragraph></Paragraphs>
        <Style><Border><Style>Solid</Style><Color>#000000</Color><Width>1pt</Width></Border></Style></Textbox>
    </CellContents></TablixCell></TablixCells></TablixRow>
   </TablixRows></TablixBody>
   <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
   <TablixRowHierarchy><TablixMembers><TablixMember/><TablixMember>
     <Group Name="g"><GroupExpressions><GroupExpression>=Fields!V.Value</GroupExpression></GroupExpressions></Group>
   </TablixMember></TablixMembers></TablixRowHierarchy>
   <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Width>3in</Width>
   <Style><Border><Style>None</Style></Border></Style></Tablix>
 </ReportItems><Height>2in</Height><Style/></Body>
 <Page><PageHeight>3in</PageHeight><PageWidth>4in</PageWidth>
   <TopMargin>0.25in</TopMargin><BottomMargin>0.25in</BottomMargin>
   <LeftMargin>0.25in</LeftMargin><RightMargin>0.25in</RightMargin>
   <PageFooter><Height>0.3in</Height><PrintOnFirstPage>true</PrintOnFirstPage><PrintOnLastPage>true</PrintOnLastPage>
     <ReportItems><Line Name="FooterRule"><Top>0.0125in</Top><Left>0in</Left><Height>0in</Height><Width>3.5in</Width>
       <Style><Border><Style>Solid</Style><Color>#000000</Color><Width>1pt</Width></Border></Style></Line></ReportItems>
     <Style/></PageFooter>
 </Page></ReportSection></ReportSections></Report>`;

// Each row is two explicit lines, so its measured height is deterministic whatever font is installed, and
// "was this row split?" is answerable by asking whether its two halves landed on the same page.
const ROWS = 9;
const rows = Array.from({ length: ROWS }, (unused, index) => ({ V: `ROW${index + 1}_TOP\nROW${index + 1}_BOTTOM` }));
const render = async (datasetRows) => renderPdf(
  parseRdl(rdl),
  { outputFileName: 'tablix-row-page-atomicity', parameters: {}, datasets: { D: datasetRows } },
  config,
  { captureLayoutTrace: true },
);

const pageOf = (pages, needle) => pages
  .filter((page) => (page.items || []).some((item) => (item.text || '').includes(needle)))
  .map((page) => page.number);

// The lowest edge any tablix cell reaches on a page: where the table's last complete row closes.
const tableBottom = (page) => Math.max(
  0,
  ...(page.items || []).filter((item) => item.kind === 'tablixCell').map((item) => item.y + item.height),
);

test('a row that does not fit the remainder moves whole instead of splitting', async () => {
  const rendered = await render(rows);
  const pages = rendered.layoutTrace.pages;
  assert.ok(pages.length >= 2, `the table must cross a page boundary, got ${pages.length} page(s)`);

  for (let index = 1; index <= ROWS; index += 1) {
    const top = pageOf(pages, `ROW${index}_TOP`);
    const bottomHalf = pageOf(pages, `ROW${index}_BOTTOM`);
    assert.equal(top.length, 1, `ROW${index} is drawn once`);
    assert.deepEqual(bottomHalf, top, `both halves of ROW${index} belong to one page`);
  }
});

test('the table closes at its last complete row, clear of the page footer band', async () => {
  const rendered = await render(rows);
  const pages = rendered.layoutTrace.pages;

  for (const page of pages) {
    const closes = tableBottom(page);
    assert.ok(closes > 0, `page ${page.number} carries table rows`);
    assert.ok(
      closes < page.bodyBottom - 0.5,
      `page ${page.number} must close its rows above the printable body boundary (${closes} vs ${page.bodyBottom})`,
    );
    const footer = (page.items || []).filter((item) => item.itemName === 'FooterRule');
    assert.equal(footer.length, 1, `page ${page.number} draws the report footer rule`);
    assert.ok(footer[0].y > page.bodyBottom, 'the footer rule belongs to the footer band');
  }

  // A moved row leaves real slack behind it — a split row would have consumed the remainder exactly.
  assert.ok(
    pages.slice(0, -1).some((page) => page.bodyBottom - tableBottom(page) > 5),
    'a page the table continues past must show the gap its last complete row leaves',
  );
});

test('a row too tall for any page is still split', async () => {
  // One row of 60 lines cannot fit a 3in page, so it has nowhere to move to: splitting is the only way to
  // render it, and that path must stay alive.
  const giant = [{ V: Array.from({ length: 60 }, (unused, index) => `GIANT_${index + 1}`).join('\n') }];
  const rendered = await render(giant);
  const pages = rendered.layoutTrace.pages;
  assert.ok(pages.length >= 2, 'the oversized row spans several pages');
  assert.notDeepEqual(
    pageOf(pages, 'GIANT_1'),
    pageOf(pages, 'GIANT_60'),
    'the row that cannot fit a page is split across pages',
  );
});
