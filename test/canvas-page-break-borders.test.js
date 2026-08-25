import assert from 'node:assert/strict';
import test from 'node:test';
import zlib from 'node:zlib';
import { PDFDocument, PDFName } from 'pdf-lib';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';

// Tablix borders are collected as edges and stroked as merged runs when a page closes. Only the tablix's
// own continuation path used to flush them, so a page break taken by any OTHER pagination path in the
// renderer — the SSRS List / canvas reflow, and the nested-region continuation — carried a whole page of
// borders forward. The table then had NO borders on the pages it actually occupies, and the accumulated
// runs were painted, at the earlier pages' coordinates, over whatever sat on the page where the flush
// finally landed. Borders belong to the page their cells are on; a page must never close with edges
// pending.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

const cell = (name, value) => `<Textbox Name="${name}"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns>`
  + `<TextRun><Value>${value}</Value><Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style></TextRun>`
  + '</TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border>'
  + '<TopBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></TopBorder>'
  + '<BottomBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></BottomBorder>'
  + '<LeftBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></LeftBorder>'
  + '<RightBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></RightBorder>'
  + '<PaddingLeft>2pt</PaddingLeft><PaddingRight>2pt</PaddingRight></Style></Textbox>';

// A grouped 1x1 tablix whose cell is a free-form canvas (the Line makes it one), holding a bordered
// nested tablix that is far taller than a page. Drawing it walks the canvas reflow and the nested-region
// continuation — both of which add pages without going through the tablix's own continuation path.
const rdl = `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="Label"><DataField>Label</DataField></Field>
  </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="Canvas"><TablixBody>
      <TablixColumns><TablixColumn><Width>5in</Width></TablixColumn></TablixColumns>
      <TablixRows><TablixRow><Height>2in</Height><TablixCells><TablixCell><CellContents>
        <Rectangle Name="Surface"><ReportItems>
          <Line Name="Rule"><Top>0in</Top><Left>0in</Left><Height>0in</Height><Width>5in</Width>
            <Style><Border><Style>Solid</Style></Border></Style></Line>
          <Tablix Name="Inner"><TablixBody>
            <TablixColumns><TablixColumn><Width>1in</Width></TablixColumn><TablixColumn><Width>4in</Width></TablixColumn></TablixColumns>
            <TablixRows><TablixRow><Height>0.25in</Height><TablixCells>
              <TablixCell><CellContents>${cell('InnerKey', '=RowNumber(Nothing)')}</CellContents></TablixCell>
              <TablixCell><CellContents>${cell('InnerLabel', '=Fields!Label.Value')}</CellContents></TablixCell>
            </TablixCells></TablixRow></TablixRows>
          </TablixBody>
          <TablixColumnHierarchy><TablixMembers><TablixMember/><TablixMember/></TablixMembers></TablixColumnHierarchy>
          <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="InnerDetails"/></TablixMember></TablixMembers></TablixRowHierarchy>
          <DataSetName>D</DataSetName><Top>0.2in</Top><Left>0in</Left><Height>0.25in</Height><Width>5in</Width><Style/>
          </Tablix>
        </ReportItems><Top>0in</Top><Left>0in</Left><Width>5in</Width><Height>2in</Height><Style/></Rectangle>
      </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
    </TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="All"/></TablixMember></TablixMembers></TablixRowHierarchy>
    <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>2in</Height><Width>5in</Width><Style/>
    </Tablix>
  </ReportItems><Height>2in</Height><Style/></Body><Width>5.2in</Width>
  <Page><PageHeight>4in</PageHeight><PageWidth>6in</PageWidth><TopMargin>0.3in</TopMargin>
    <BottomMargin>0.3in</BottomMargin><LeftMargin>0.3in</LeftMargin><RightMargin>0.3in</RightMargin></Page>
  </ReportSection></ReportSections>
</Report>`;

const request = {
  outputFileName: 'canvas-page-break-borders',
  parameters: {},
  datasets: { D: Array.from({ length: 60 }, (_, index) => ({ Label: `ROW_${String(index + 1).padStart(3, '0')}` })) },
};

// Stroke operations per page, in page order. PDFKit writes each border run as `x y m` / `x y l` / `S`.
async function strokesPerPage(buffer) {
  const doc = await PDFDocument.load(buffer);
  const counts = [];
  for (const page of doc.getPages()) {
    const contents = page.node.context.lookup(page.node.get(PDFName.of('Contents')));
    let data = contents.constructor.name === 'PDFArray'
      ? Buffer.concat(contents.asArray().map((ref) => Buffer.from(page.node.context.lookup(ref).getContents())))
      : Buffer.from(contents.getContents());
    try { data = zlib.inflateSync(data); } catch { /* already flat */ }
    counts.push((data.toString('latin1').match(/^S$/gm) || []).length);
  }
  return counts;
}

test('every page of a canvas-hosted nested tablix strokes its own borders', async () => {
  const rendered = await renderPdf(parseRdl(rdl), request, config, { captureLayoutTrace: true });
  assert.ok(rendered.pageCount > 3, `the fixture must cross several pages (got ${rendered.pageCount})`);
  const strokes = await strokesPerPage(rendered.buffer);
  assert.equal(strokes.length, rendered.pageCount);

  // Every page carrying bordered cells must carry strokes of its own.
  const borderedPages = rendered.layoutTrace.pages
    .filter((page) => (page.tablixFragments || []).some((fragment) => (fragment.cells || [])
      .some((traced) => traced.borders && Object.values(traced.borders).some(Boolean))))
    .map((page) => page.number);
  assert.ok(borderedPages.length > 3, 'the fixture must place bordered cells on several pages');
  const cellsPerPage = rendered.layoutTrace.pages
    .map((page) => (page.tablixFragments || []).flatMap((fragment) => fragment.cells || []).length);
  for (const number of borderedPages) {
    // Merging collapses shared edges, so a page of C bordered cells strokes fewer than 4C runs — but it
    // can never be down at one or two. Carrying the edges forward left exactly that: a visibly borderless
    // table with one stray rule on it.
    const expected = Math.min(4, cellsPerPage[number - 1]);
    assert.ok(
      strokes[number - 1] >= expected,
      `page ${number} holds ${cellsPerPage[number - 1]} bordered cells but strokes only ${strokes[number - 1]} runs`,
    );
  }

  // And no page may carry a hoarded pile from the pages before it. A page can only stroke the edges of
  // the cells it actually holds: four sides each, plus a little slack for the fragment outline and any
  // body-level rules. Pre-fix, the flush page carried hundreds of runs against two cells of its own.
  for (const [index, count] of strokes.entries()) {
    assert.ok(
      count <= cellsPerPage[index] * 4 + 8,
      `page ${index + 1} strokes ${count} runs for ${cellsPerPage[index]} cells — borders from other pages landed here`,
    );
  }
});

test('the canvas fixture keeps every row exactly once across the page breaks', async () => {
  const rendered = await renderPdf(parseRdl(rdl), request, config, { captureLayoutTrace: true });
  const labels = rendered.layoutTrace.pages
    .flatMap((page) => (page.tablixFragments || []).flatMap((fragment) => (fragment.cells || []).map((traced) => traced.text)))
    .filter((text) => /^ROW_\d\d\d$/.test(String(text || '')));
  assert.equal(new Set(labels).size, 60, 'every row must render');
  assert.equal(labels.length, 60, 'no row may render twice');
});
