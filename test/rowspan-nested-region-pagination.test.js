import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderVisualDocx } from '../src/render/visualDocx.js';
import { renderExcel } from '../src/render/excel.js';

// A nested data region living in a row-header cell that vertically merges several physical parent rows.
// The region is taller than one printable page, so the parent row is drawn once per page fragment while
// the child is windowed to that fragment's rows.
//
// Two defects met here. drawRowContent pushed a NEW open span for the merged cell on every fragment of its
// origin row, so one merged cell ended up with two concurrent spans drawing its value, fill, and borders
// twice at one origin. And the duplicate span was opened after the fragment loop had moved on, so when it
// finally closed it measured the WHOLE child against a single page segment and failed the render with
// UNSUPPORTED_FEATURE ("Row-spanned nested tablix ... exceeds its printable page segment") — while the
// first span, already marked drawn, silently dropped every child row after the first fragment.
//
// The merged cell must open once, and each fragment must draw exactly the child rows it was laid out with.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

const textbox = (name, value) => `<Textbox Name="${name}"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns>`
  + `<TextRun><Value>${value}</Value><Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style></TextRun>`
  + '</TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border>'
  + '<PaddingLeft>2pt</PaddingLeft><PaddingRight>2pt</PaddingRight></Style></Textbox>';

// `childRowCount` static rows of `childRowHeight` inside the group header cell. Static rows keep the child
// height independent of the data, so one fixture can be taller than a page and its counterexample shorter.
const report = ({ childRowCount, childRowHeight }) => Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
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
        <TablixMember><Group Name="G"><GroupExpressions><GroupExpression>=Fields!G.Value</GroupExpression></GroupExpressions></Group>
          <TablixHeader><Size>2in</Size><CellContents>
            <Tablix Name="Child"><TablixBody>
              <TablixColumns><TablixColumn><Width>2in</Width></TablixColumn></TablixColumns>
              <TablixRows>${Array.from({ length: childRowCount }, (unused, index) => (
    `<TablixRow><Height>${childRowHeight}in</Height><TablixCells><TablixCell><CellContents>`
                + `${textbox(`ChildCell${index + 1}`, `CHILD_${index + 1}`)}</CellContents></TablixCell></TablixCells></TablixRow>`
  )).join('')}</TablixRows>
            </TablixBody>
            <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
            <TablixRowHierarchy><TablixMembers>${'<TablixMember/>'.repeat(childRowCount)}</TablixMembers></TablixRowHierarchy>
            <Top>0in</Top><Left>0in</Left><Height>${childRowHeight * childRowCount}in</Height><Width>2in</Width><Style/></Tablix>
          </CellContents></TablixHeader>
          <TablixMembers>
            <TablixMember><Group Name="D"><GroupExpressions><GroupExpression>=Fields!D.Value</GroupExpression></GroupExpressions></Group></TablixMember>
          </TablixMembers>
        </TablixMember>
      </TablixMembers></TablixRowHierarchy>
      <DataSetName>DS</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.3in</Height><Width>5in</Width><Style/></Tablix>
  </ReportItems><Height>2in</Height><Style/></Body><Width>8in</Width>
  <Page><PageHeight>3in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.3in</LeftMargin><RightMargin>0.3in</RightMargin><TopMargin>0.3in</TopMargin><BottomMargin>0.3in</BottomMargin></Page>
  </ReportSection></ReportSections></Report>`, 'utf8');

// Two groups of four detail rows each, so the group header cell row-spans four physical rows.
const GROUPS = ['G1', 'G2'];
const DETAILS_PER_GROUP = 4;
const rows = GROUPS.flatMap((group) => Array.from(
  { length: DETAILS_PER_GROUP },
  (unused, index) => ({ G: group, D: `${group}-D${index + 1}` }),
));
const request = (output) => ({ output, outputFileName: 'rowspan-nested', parameters: {}, datasets: { DS: rows } });

// The body is 3in - 0.6in of margins = 2.4in tall, so four child rows of 0.8in must split across pages and
// two child rows of 0.5in must not.
const OVERSIZED = { childRowCount: 4, childRowHeight: 0.8 };
const FITTING = { childRowCount: 2, childRowHeight: 0.5 };

const EPS = 0.13;
const overlaps = (a, b) => (
  Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x) > EPS
  && Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y) > EPS
);
const drawnTexts = (rendered) => rendered.layoutTrace.pages
  .flatMap((page) => (page.items || []).map((item) => item.text).filter(Boolean));

test('a nested region in a row-spanned cell paginates instead of failing closed', async () => {
  const rendered = await renderPdf(parseRdl(report(OVERSIZED)), request('PDF'), config, { captureLayoutTrace: true });
  const texts = drawnTexts(rendered);

  // Every child row is drawn exactly once per group: fragments after the first were dropped before the
  // fix, and a duplicated span would have drawn some of them twice.
  for (let index = 1; index <= OVERSIZED.childRowCount; index += 1) {
    assert.equal(
      texts.filter((value) => value === `CHILD_${index}`).length,
      GROUPS.length,
      `CHILD_${index} must appear once per group`,
    );
  }
  // The child really did split: for each group, its last row lands on a later page than its first.
  const pagesCarrying = (value) => rendered.layoutTrace.pages
    .filter((page) => (page.items || []).some((item) => item.text === value))
    .map((page) => page.number);
  const firstRowPages = pagesCarrying('CHILD_1');
  const lastRowPages = pagesCarrying(`CHILD_${OVERSIZED.childRowCount}`);
  assert.equal(firstRowPages.length, GROUPS.length);
  assert.equal(lastRowPages.length, GROUPS.length);
  for (const [index, page] of lastRowPages.entries()) {
    assert.equal(page > firstRowPages[index], true, 'the child must continue onto a later page');
  }

  // No merged cell may be emitted twice at one origin — the condition the native-Word cell builder enforces.
  for (const page of rendered.layoutTrace.pages) {
    const owners = (page.items || []).filter((item) => item.itemName);
    for (let i = 0; i < owners.length; i += 1) {
      for (let j = i + 1; j < owners.length; j += 1) {
        assert.equal(
          owners[i].itemName === owners[j].itemName && overlaps(owners[i], owners[j]),
          false,
          `page ${page.number}: ${owners[i].itemName} drawn twice at one origin`,
        );
      }
    }
  }
});

test('a nested region that fits its row-spanned cell is still drawn exactly once', async () => {
  const rendered = await renderPdf(parseRdl(report(FITTING)), request('PDF'), config, { captureLayoutTrace: true });
  const texts = drawnTexts(rendered);
  for (let index = 1; index <= FITTING.childRowCount; index += 1) {
    assert.equal(texts.filter((value) => value === `CHILD_${index}`).length, GROUPS.length);
  }
});

test('the paginated row-spanned child reaches editable Word, visual Word, and Excel', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-rowspan-nested-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const model = parseRdl(report(OVERSIZED));
  const canonical = await renderPdf(model, request('PDF'), config, { captureLayoutTrace: true });

  // Editable Word is built from the canonical PDF trace: same page count, every child row as native text.
  const editable = await renderEditableDocx(model, request('DOCX_EDITABLE'), config);
  assert.equal(editable.pageCount, canonical.pageCount);
  const documentXml = await (await JSZip.loadAsync(editable.buffer)).file('word/document.xml').async('string');
  for (let index = 1; index <= OVERSIZED.childRowCount; index += 1) {
    assert.equal(
      (documentXml.match(new RegExp(`<w:t[^>]*>CHILD_${index}</w:t>`, 'g')) || []).length,
      GROUPS.length,
      `CHILD_${index} must be native Word text once per group`,
    );
  }

  // Visual Word rasterizes the same canonical PDF: exactly one page image per canonical page.
  const visual = await renderVisualDocx(model, request('DOCX_VISUAL'), config, tempDir);
  assert.equal(visual.pageCount, canonical.pageCount);
  const visualZip = await JSZip.loadAsync(visual.buffer);
  const media = Object.keys(visualZip.files).filter((name) => /^word\/media\/.+\.png$/.test(name));
  assert.equal(media.length, canonical.pageCount);

  // Excel has no pages, so the construct is never split there — but every child row must still be present,
  // which is the same "no dropped fragment" guarantee expressed in the format that can hold it.
  const excel = await renderExcel(model, request('XLSX'), config, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(excel.buffer);
  const values = [];
  workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => {
    // A merged block reports its value from every covered cell; count the block once.
    if (cell.isMerged && cell.master !== cell) return;
    values.push(cell.value);
  }));
  for (let index = 1; index <= OVERSIZED.childRowCount; index += 1) {
    assert.equal(values.filter((value) => value === `CHILD_${index}`).length, GROUPS.length);
  }
});
