import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderPdf } from '../src/render/pdf.js';

// A page break INSIDE a tablix row is not an edge of anything: the row resumes at the top of the next page.
// Both the row's own cell borders and the synthesized fragment closure used to be stroked at that cut, so a
// table that broke mid-row drew a horizontal rule hard against the printable body boundary — which is where
// a page footer's own rule sits — while an otherwise identical table that broke between rows showed the gap
// its last row leaves. Same document, two different bottoms. A rule now appears only at a real row
// boundary, so the spacing below the table is the same on every page.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000' });

// One long sentence per row. The body is deliberately short, so a single row's text consumes several whole
// pages: the middle ones carry nothing but the repeated header and the middle of the split row, which is
// what makes "how many rules does a page that ends inside a row draw?" answerable by counting.
const LONG = Array.from({ length: 120 }, (_, index) => `sentence ${index} of the split row`).join(', ');

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
        <Style><Border><Style>None</Style></Border>
          <BottomBorder><Style>Dotted</Style><Color>Black</Color><Width>1pt</Width></BottomBorder></Style></Textbox>
    </CellContents></TablixCell></TablixCells></TablixRow>
    <TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents>
      <Textbox Name="d"><CanGrow>true</CanGrow>
        <Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!V.Value</Value>
        <Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style></TextRun></TextRuns></Paragraph></Paragraphs>
        <Style><Border><Style>None</Style></Border>
          <BottomBorder><Style>Dotted</Style><Color>Black</Color><Width>1pt</Width></BottomBorder></Style></Textbox>
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

const request = { outputFileName: 'tablix-split-row-open-edge', parameters: {}, datasets: { D: [{ V: LONG }] } };

// Every horizontal rule the table draws on a page: a cell's resolved bottom border, plus the synthesized
// outer closure. `atBodyBottom` is the subset that lands on the printable body boundary — the cut.
function rulesOnPage(page) {
  const atY = (y) => Math.abs(y - page.bodyBottom) <= 0.5;
  const cellBottoms = (page.items || []).filter((item) => (
    item.kind === 'tablixCell' && item.borders?.bottom && !/^none$/i.test(String(item.borders.bottom.style))
  ));
  const closures = (page.items || []).filter((item) => (
    item.traceRole === 'resolvedTablixFragmentBorder' && item.fragmentSide === 'bottom'
  ));
  return {
    cellBottoms,
    closures,
    atBodyBottom: [
      ...cellBottoms.filter((item) => atY(item.y + item.height)),
      ...closures.filter((item) => atY(item.y)),
    ],
  };
}

test('a page break inside a row leaves the row open instead of ruling the cut', async () => {
  const rendered = await renderPdf(parseRdl(rdl), request, config, { captureLayoutTrace: true });
  const pages = rendered.layoutTrace.pages;
  assert.ok(pages.length >= 3, `the row must span several pages to isolate the cut, got ${pages.length}`);

  // Every page but the last ends inside the row: none of them may draw a rule at the cut.
  for (const page of pages.slice(0, -1)) {
    const { atBodyBottom } = rulesOnPage(page);
    assert.deepEqual(
      atBodyBottom.map((item) => item.itemName || item.kind),
      [],
      `page ${page.number} ends inside the row, so nothing closes at the body boundary`,
    );
  }

  // The row ends on the last page, and there its own rule is drawn — the table is never left open.
  const last = pages.at(-1);
  const closing = rulesOnPage(last);
  assert.ok(closing.cellBottoms.length > 0, 'the row that ends draws its own bottom rule');
  const lowest = Math.max(...closing.cellBottoms.map((item) => item.y + item.height));
  assert.ok(
    lowest < last.bodyBottom - 1,
    `the closing rule sits at the row's end (${lowest}), clear of the body boundary (${last.bodyBottom})`,
  );

  // The footer keeps its own rule on every page; the two are never the same line.
  for (const page of pages) {
    const footer = (page.items || []).filter((item) => item.itemName === 'FooterRule');
    assert.equal(footer.length, 1, `page ${page.number} draws the report footer rule`);
    assert.ok(footer[0].y > page.bodyBottom, 'the footer rule belongs to the footer band, not the body');
  }
});

test('a repeated header still rules its own boundary on every page', async () => {
  const rendered = await renderPdf(parseRdl(rdl), request, config, { captureLayoutTrace: true });
  for (const page of rendered.layoutTrace.pages) {
    const { cellBottoms } = rulesOnPage(page);
    assert.ok(
      cellBottoms.some((item) => item.repeatedHeader),
      `page ${page.number} repeats the column header and rules under it`,
    );
  }
});

test('editable DOCX leaves the same cuts open', async (context) => {
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-split-row-edge-'));
  context.after(() => fs.rm(tempRoot, { recursive: true, force: true }));
  const owned = loadConfig({
    ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000', RDL_TEMP_ROOT: tempRoot,
  });
  const model = parseRdl(rdl);
  const trace = (await renderPdf(model, request, owned, { captureLayoutTrace: true })).layoutTrace;
  const docx = await renderEditableDocx(model, request, owned);
  const xml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');

  // One Word section per canonical page. A page that ends inside the row carries exactly the rules that
  // page really has — the repeated header's — and none for the cut; before the fix each carried one more.
  const sections = xml.split(/<w:sectPr\b/).slice(0, trace.pages.length);
  assert.equal(sections.length, trace.pages.length, 'one Word section per canonical PDF page');
  const dotted = (chunk) => (chunk.match(/<w:bottom [^>]*w:val="dotted"/g) || []).length;
  for (const [index, chunk] of sections.entries()) {
    const page = trace.pages[index];
    const expected = rulesOnPage(page).cellBottoms.length;
    assert.equal(
      dotted(chunk),
      expected,
      `Word page ${page.number} must draw the ${expected} rule(s) the canonical page draws`,
    );
  }
  // The construct is only meaningful if some page genuinely ends inside the row.
  assert.ok(
    sections.slice(0, -1).every((chunk) => dotted(chunk) === 1),
    'a page consumed by the middle of a split row rules only under the repeated header',
  );
});
