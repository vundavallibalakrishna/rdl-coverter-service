// A row-span (merged) group-header cell whose group crosses a page boundary must not emit two copies of the
// header at the same position. The trigger: the trailing row of a row-span group is text-less (its columns
// are covered by the merged header, its body cell empty). When such a row does not fit at a page boundary,
// the text-split loop cannot advance it — before the fix it was dropped, its open row-span stayed open, and
// the header's residual then closed against the NEXT group's cursor, painting a duplicate header at one
// origin. In PDF the duplicate paints over itself (invisible); native Word cells cannot overlap, so
// DOCX_EDITABLE fail-closed with "Overlapping editable PDF regions". The fix moves the unsplittable row to a
// fresh page so it draws in order and its span closes against its own row.
import test from 'node:test';
import assert from 'node:assert/strict';
import os from 'node:os';
import path from 'node:path';
import fs from 'node:fs/promises';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';
import { loadConfig } from '../src/config.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

// Outer group G with a row-header cell (Ghdr) that row-spans two STATIC inner member rows: a body row with
// text and a trailing row whose body cell is EMPTY (text-less). A short page forces a group's text-less
// trailing row past the page bottom, exercising the drop-and-overlap path.
function report() {
  return Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="DS"><Fields>
    <Field Name="G"><DataField>G</DataField><TypeName>System.String</TypeName></Field>
    <Field Name="D"><DataField>D</DataField><TypeName>System.String</TypeName></Field>
  </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="T"><TablixBody>
      <TablixColumns><TablixColumn><Width>4in</Width></TablixColumn></TablixColumns>
      <TablixRows>
        <TablixRow><Height>0.5in</Height><TablixCells><TablixCell><CellContents><Textbox Name="Body1"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!D.Value</Value><Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border></Style></Textbox></CellContents></TablixCell></TablixCells></TablixRow>
        <TablixRow><Height>0.35in</Height><TablixCells><TablixCell><CellContents><Textbox Name="Body2"><Paragraphs><Paragraph><TextRuns><TextRun><Value></Value><Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border></Style></Textbox></CellContents></TablixCell></TablixCells></TablixRow>
      </TablixRows></TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers>
        <TablixMember><Group Name="G"><GroupExpressions><GroupExpression>=Fields!G.Value</GroupExpression></GroupExpressions></Group>
          <TablixHeader><Size>1in</Size><CellContents><Textbox Name="Ghdr"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!G.Value</Value><Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border></Style></Textbox></CellContents></TablixHeader>
          <TablixMembers><TablixMember/><TablixMember/></TablixMembers>
        </TablixMember>
      </TablixMembers></TablixRowHierarchy>
      <DataSetName>DS</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.85in</Height><Width>5in</Width><Style/></Tablix>
  </ReportItems><Height>2in</Height><Style/></Body><Width>8in</Width>
  <Page><PageHeight>3in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.3in</LeftMargin><RightMargin>0.3in</RightMargin><TopMargin>0.3in</TopMargin><BottomMargin>0.3in</BottomMargin></Page>
  </ReportSection></ReportSections></Report>`, 'utf8');
}

const EPS = 0.13;
const overlaps = (a, b) => (
  Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x) > EPS
  && Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y) > EPS
);

test('a row-span group header crossing a page boundary is not emitted twice at one origin', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-rowspan-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rows = Array.from({ length: 20 }, (_, g) => ({ G: `G${g}`, D: `body ${g}` }));
  const rendered = await renderPdf(
    parseRdl(report()),
    { output: 'PDF', outputFileName: 'rowspan', parameters: {}, datasets: { DS: rows } },
    config,
    { captureLayoutTrace: true },
  );

  // No two owner cells sharing an item name may occupy overlapping area on any page — the exact condition
  // the native-Word cell builder enforces. Before the fix this produced same-named "Ghdr" cells at one
  // origin on the continuation pages.
  const kinds = ['textbox', 'tablixCell', 'image', 'chart'];
  const sameNameOverlaps = [];
  for (const page of rendered.layoutTrace.pages) {
    const owners = (page.items || []).filter((item) => kinds.includes(item.kind));
    for (let i = 0; i < owners.length; i += 1) {
      for (let j = i + 1; j < owners.length; j += 1) {
        if (owners[i].itemName && owners[i].itemName === owners[j].itemName && overlaps(owners[i], owners[j])) {
          sameNameOverlaps.push(`page ${page.number}: ${owners[i].itemName}`);
        }
      }
    }
  }
  assert.deepEqual(sameNameOverlaps, [], `unexpected same-name owner overlaps: ${JSON.stringify(sameNameOverlaps)}`);

  // The report still paginates and every group header is present (content is not dropped).
  const rendered2 = await renderPdf(parseRdl(report()), { output: 'PDF', outputFileName: 'rowspan', parameters: {}, datasets: { DS: rows } }, config, {});
  assert.ok(rendered2.pageCount > 1, 'expected the report to span multiple pages');
});
