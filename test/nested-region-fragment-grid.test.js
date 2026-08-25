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

// A child data region taller than a page is drawn one fragment at a time, each fragment a SLICE of its
// rows. Two things must survive that slicing:
//
//  * Cell placement. It walks the rows in order, carrying each row-span's occupancy forward, so computing
//    it from a slice alone loses every span that began earlier — the cells to the right of a spanned row
//    header shift left into the header's own columns and the child's grid re-flows mid-table.
//  * The open span itself. SSRS repeats a row header that spans into a continuation fragment, clipped to
//    the rows that fragment shows; dropping it leaves those columns empty and unruled for the rest of the
//    region.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const GROUP_COLUMN_PT = 80;
const ITEM_COLUMN_PT = 220;
const DETAIL_COUNT = 40;

const textbox = (name, value) => `<Textbox Name="${name}"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns>`
  + `<TextRun><Value>${value}</Value><Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style></TextRun>`
  + '</TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border>'
  + '<TopBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></TopBorder>'
  + '<BottomBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></BottomBorder>'
  + '<LeftBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></LeftBorder>'
  + '<RightBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></RightBorder>'
  + '<PaddingLeft>2pt</PaddingLeft><PaddingRight>2pt</PaddingRight></Style></Textbox>';

// The inner region: a row-header column carrying the group name (which spans every detail row of that
// group) and a body column carrying the detail. The Line makes the outer cell a free-form canvas, which is
// the path that reflows the child across pages.
const rdl = `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="G"><DataField>G</DataField></Field>
    <Field Name="Item"><DataField>Item</DataField></Field>
  </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="Canvas"><TablixBody>
      <TablixColumns><TablixColumn><Width>300pt</Width></TablixColumn></TablixColumns>
      <TablixRows><TablixRow><Height>60pt</Height><TablixCells><TablixCell><CellContents>
        <Rectangle Name="Surface"><ReportItems>
          <Line Name="Rule"><Top>0pt</Top><Left>0pt</Left><Height>0pt</Height><Width>300pt</Width>
            <Style><Border><Style>Solid</Style></Border></Style></Line>
          <Tablix Name="Inner"><TablixBody>
            <TablixColumns><TablixColumn><Width>${ITEM_COLUMN_PT}pt</Width></TablixColumn></TablixColumns>
            <TablixRows>
              <TablixRow><Height>16pt</Height><TablixCells><TablixCell><CellContents>${textbox('HeadItem', 'Item')}</CellContents></TablixCell></TablixCells></TablixRow>
              <TablixRow><Height>16pt</Height><TablixCells><TablixCell><CellContents>${textbox('ItemCell', '=Fields!Item.Value')}</CellContents></TablixCell></TablixCells></TablixRow>
            </TablixRows>
          </TablixBody>
          <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
          <TablixRowHierarchy><TablixMembers>
            <TablixMember>
              <TablixHeader><Size>${GROUP_COLUMN_PT}pt</Size><CellContents>${textbox('HeadGroup', 'Group')}</CellContents></TablixHeader>
              <TablixMembers><TablixMember/></TablixMembers>
            </TablixMember>
            <TablixMember>
              <Group Name="G"><GroupExpressions><GroupExpression>=Fields!G.Value</GroupExpression></GroupExpressions></Group>
              <TablixHeader><Size>${GROUP_COLUMN_PT}pt</Size><CellContents>${textbox('GroupCell', '=Fields!G.Value')}</CellContents></TablixHeader>
              <TablixMembers><TablixMember><Group Name="Detail"/></TablixMember></TablixMembers>
            </TablixMember>
          </TablixMembers></TablixRowHierarchy>
          <DataSetName>D</DataSetName><Top>4pt</Top><Left>0pt</Left><Height>32pt</Height>
          <Width>${GROUP_COLUMN_PT + ITEM_COLUMN_PT}pt</Width><Style/>
          </Tablix>
        </ReportItems><Top>0pt</Top><Left>0pt</Left><Width>300pt</Width><Height>60pt</Height><Style/></Rectangle>
      </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
    </TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers><TablixMember>
      <Group Name="All"><GroupExpressions><GroupExpression>=1</GroupExpression></GroupExpressions></Group>
    </TablixMember></TablixMembers></TablixRowHierarchy>
    <DataSetName>D</DataSetName><Top>0pt</Top><Left>0pt</Left><Height>60pt</Height><Width>300pt</Width><Style/>
    </Tablix>
  </ReportItems><Height>60pt</Height><Style/></Body><Width>310pt</Width>
  <Page><PageWidth>360pt</PageWidth><PageHeight>230pt</PageHeight><TopMargin>10pt</TopMargin>
    <BottomMargin>10pt</BottomMargin><LeftMargin>10pt</LeftMargin><RightMargin>10pt</RightMargin></Page>
  </ReportSection></ReportSections>
</Report>`;

const rows = Array.from({ length: DETAIL_COUNT }, (_, index) => ({ G: 'ONLY_GROUP', Item: `ITEM_${String(index + 1).padStart(3, '0')}` }));
const request = { outputFileName: 'nested-fragment-grid', parameters: {}, datasets: { D: rows }, excel: { layoutMode: 'REPORT' } };

const tracedCells = (trace) => trace.pages.flatMap((page) => (page.tablixFragments || [])
  .flatMap((fragment) => (fragment.cells || []).map((cell) => ({ ...cell, page: page.number }))));

test('a child region keeps its column grid across every page fragment', async () => {
  const rendered = await renderPdf(parseRdl(rdl), request, config, { captureLayoutTrace: true });
  assert.ok(rendered.pageCount > 2, `the fixture must split the child across pages (got ${rendered.pageCount})`);
  const cells = tracedCells(rendered.layoutTrace);
  const details = cells.filter((cell) => /^ITEM_\d\d\d$/.test((cell.text || '').trim()));
  assert.equal(details.length, DETAIL_COUNT, 'every detail renders exactly once');
  assert.ok(new Set(details.map((cell) => cell.page)).size > 2, 'the details must span several pages');

  // Every detail belongs in the body column, on every page. Recomputing placement from a slice put the
  // ones after the first fragment into the row-header column instead.
  const columns = new Set(details.map((cell) => cell.columnIndex));
  assert.deepEqual([...columns], [1], `details drifted out of the body column: ${[...columns].join(',')}`);
  const widths = new Set(details.map((cell) => Math.round(cell.width)));
  assert.equal(widths.size, 1, 'every detail cell keeps the body column width');
  assert.ok(
    Math.abs([...widths][0] - ITEM_COLUMN_PT) < 2,
    `detail cells are ${[...widths][0]}pt wide, expected the ${ITEM_COLUMN_PT}pt body column`,
  );
});

test('a row header that spans into a continuation fragment is repeated there', async () => {
  const rendered = await renderPdf(parseRdl(rdl), request, config, { captureLayoutTrace: true });
  const cells = tracedCells(rendered.layoutTrace);
  const detailPages = new Set(cells
    .filter((cell) => /^ITEM_\d\d\d$/.test((cell.text || '').trim()))
    .map((cell) => cell.page));
  const groupPages = new Set(cells
    .filter((cell) => (cell.text || '').trim() === 'ONLY_GROUP')
    .map((cell) => cell.page));
  for (const page of detailPages) {
    assert.ok(groupPages.has(page), `page ${page} shows details but not the group header that spans them`);
  }
  // The repeated header stays in the row-header column at its declared width.
  const groupCells = cells.filter((cell) => (cell.text || '').trim() === 'ONLY_GROUP');
  for (const cell of groupCells) {
    assert.equal(cell.columnIndex, 0);
    assert.ok(Math.abs(cell.width - GROUP_COLUMN_PT) < 2, `the group header is ${cell.width}pt wide`);
  }
});

test('a continued child region is labelled on the page it continues onto', async () => {
  // The "continued from previous page" annotation belongs to whichever region crossed the boundary. Only
  // the top-level tablix's continuation path emitted it, so a report whose tables are all CHILD regions
  // carried no marker at all even with the option on.
  const marked = await renderPdf(parseRdl(rdl), { ...request, pagination: { continuationMarkers: true } }, config, { captureLayoutTrace: true });
  const markerPages = new Set(marked.layoutTrace.pages
    .filter((page) => (page.items || []).some((item) => item.traceRole === 'continuationMarker'))
    .map((page) => page.number));
  const detailPages = [...new Set(tracedCells(marked.layoutTrace)
    .filter((cell) => /^ITEM_\d\d\d$/.test((cell.text || '').trim()))
    .map((cell) => cell.page))].sort((left, right) => left - right);
  assert.ok(detailPages.length > 2, 'the fixture must continue across several pages');
  assert.ok(!markerPages.has(detailPages[0]), 'the page the table starts on is not a continuation');
  for (const page of detailPages.slice(1)) {
    assert.ok(markerPages.has(page), `page ${page} continues the child region but carries no marker`);
  }

  // The option is opt-in: without it the annotation must not appear.
  const plain = await renderPdf(parseRdl(rdl), request, config, { captureLayoutTrace: true });
  assert.equal(
    plain.layoutTrace.pages.filter((page) => (page.items || []).some((item) => item.traceRole === 'continuationMarker')).length,
    0,
    'markers must stay off unless the request asks for them',
  );
});

test('XLSX has no pagination, so the child region keeps one continuous grid', async (context) => {
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-fragment-grid-'));
  context.after(() => fs.rm(tempRoot, { recursive: true, force: true }));
  const owned = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_TEMP_ROOT: tempRoot });
  const xlsx = await renderExcel(parseRdl(rdl), request, owned, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(xlsx.buffer);
  const sheet = workbook.worksheets[0];
  const columnsUsed = new Set();
  let found = 0;
  sheet.eachRow((row) => row.eachCell((cell, columnNumber) => {
    const value = typeof cell.value === 'object' && cell.value?.richText
      ? cell.value.richText.map((run) => run.text).join('')
      : String(cell.value ?? '');
    if (/^ITEM_\d\d\d$/.test(value.trim())) {
      found += 1;
      columnsUsed.add(cell.master ? cell.master.col : columnNumber);
    }
  }));
  assert.equal(found > 0, true, 'the details reach Excel');
  assert.equal(columnsUsed.size, 1, `details landed in ${columnsUsed.size} different Excel columns`);
});
