// A tablix cell whose CellContents is a Rectangle is a CANVAS: the flattened children keep their own
// declared Top/Left/Width/Height. Drawing one of them stretched over the whole cell put it in the wrong
// place and made it swallow the nested data region beside it, which the page-locked Word renderer then
// refused as "Overlapping editable PDF regions". A textbox whose sibling is a nested tablix must therefore
// be drawn at its declared geometry in every renderer that can express it.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { loadConfig } from '../src/config.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const PT = 72;
const DATA = { D: [{ Code: 'AU', Label: 'Assurance Services' }] };

const INNER_TABLIX = `
  <Tablix Name="Inner">
    <TablixBody>
      <TablixColumns><TablixColumn><Width>1in</Width></TablixColumn><TablixColumn><Width>2in</Width></TablixColumn></TablixColumns>
      <TablixRows><TablixRow><Height>0.25in</Height><TablixCells>
        <TablixCell><CellContents><Textbox Name="InnerCode"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Code.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell>
        <TablixCell><CellContents><Textbox Name="InnerLabel"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Label.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell>
      </TablixCells></TablixRow></TablixRows>
    </TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixRowHierarchy>
    <DataSetName>D</DataSetName>
    <Top>0.5in</Top><Left>0.25in</Left><Height>0.25in</Height><Width>3in</Width><Style/>
  </Tablix>`;

const CAPTION = `
  <Textbox Name="Caption">
    <Paragraphs><Paragraph><TextRuns><TextRun><Value>Legend</Value></TextRun></TextRuns></Paragraph></Paragraphs>
    <Top>0.1in</Top><Left>0.25in</Left><Height>0.25in</Height><Width>1.5in</Width><Style/>
  </Textbox>`;

// `cellContents` becomes the single report item of the outer tablix's only cell.
function report(cellContents) {
  return Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="Code"><DataField>Code</DataField><TypeName>System.String</TypeName></Field>
    <Field Name="Label"><DataField>Label</DataField><TypeName>System.String</TypeName></Field>
  </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="Outer"><TablixBody>
      <TablixColumns><TablixColumn><Width>5in</Width></TablixColumn></TablixColumns>
      <TablixRows><TablixRow><Height>1.5in</Height><TablixCells>
        <TablixCell><CellContents>${cellContents}</CellContents></TablixCell>
      </TablixCells></TablixRow></TablixRows>
    </TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="G"/></TablixMember></TablixMembers></TablixRowHierarchy>
    <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>1.5in</Height><Width>5in</Width><Style/></Tablix>
  </ReportItems><Height>3in</Height><Style/></Body><Width>7in</Width>
  <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`, 'utf8');
}

// The same caption with no declared Width/Height. RDL defaults those to zero, so there is no free-form box
// to honour and the fill-the-cell path must be kept rather than drawing the text into a 0x0 rectangle.
const UNSIZED_CAPTION = `
  <Textbox Name="Caption">
    <Paragraphs><Paragraph><TextRuns><TextRun><Value>Legend</Value></TextRun></TextRuns></Paragraph></Paragraphs>
    <Top>0.1in</Top><Left>0.25in</Left><Style/>
  </Textbox>`;

const canvasReport = () => report(`<Rectangle Name="Canvas"><ReportItems>${CAPTION}${INNER_TABLIX}</ReportItems><Style/></Rectangle>`);
const wrappedSingleTextboxReport = () => report(`<Rectangle Name="Canvas"><ReportItems>${CAPTION}</ReportItems><Style/></Rectangle>`);
const unsizedCanvasReport = () => report(`<Rectangle Name="Canvas"><ReportItems>${UNSIZED_CAPTION}${INNER_TABLIX}</ReportItems><Style/></Rectangle>`);

async function trace(rdl) {
  const rendered = await renderPdf(
    parseRdl(rdl),
    { outputFileName: 'canvas', parameters: {}, datasets: DATA },
    config,
    { captureLayoutTrace: true },
  );
  return rendered.layoutTrace;
}

function positiveOverlap(left, right) {
  return Math.min(left.x + left.width, right.x + right.width) - Math.max(left.x, right.x) > 0.01
    && Math.min(left.y + left.height, right.y + right.height) - Math.max(left.y, right.y) > 0.01;
}

const round = (value) => Math.round(value * 100) / 100;

test('PDF draws a caption sharing its cell with a nested tablix at its declared position and size', async () => {
  const layout = await trace(canvasReport());
  const [page] = layout.pages;
  const caption = page.items.find((entry) => entry.itemName === 'Caption');
  assert.ok(caption, 'the caption must be traced as its own positioned item');
  // The outer tablix sits at the body origin, so the caption lands at the 0.5in margin plus its own Left/Top.
  assert.equal(round(caption.x), round((0.5 + 0.25) * PT));
  assert.equal(round(caption.y), round((0.5 + 0.1) * PT));
  // Its declared 1.5in width, NOT the 5in cell width it used to be stretched to.
  assert.equal(round(caption.width), 1.5 * PT);
  assert.equal(round(caption.height), 0.25 * PT);
});

test('PDF leaves no overlap between the caption and the nested tablix it shares a cell with', async () => {
  const layout = await trace(canvasReport());
  const [page] = layout.pages;
  const caption = page.items.find((entry) => entry.itemName === 'Caption');
  const nested = page.items.filter((entry) => entry.tablixName === 'Inner' && entry.kind === 'tablixCell');
  assert.equal(nested.length, 2, 'both nested cells must render');
  for (const cell of nested) {
    assert.equal(positiveOverlap(caption, cell), false, `Caption must not overlap ${cell.itemName}`);
  }
});

test('a Rectangle-wrapped cell whose only content is one textbox still fills the cell (fast path unchanged)', async () => {
  const layout = await trace(wrappedSingleTextboxReport());
  const [page] = layout.pages;
  const cell = page.items.find((entry) => entry.kind === 'tablixCell' && entry.tablixName === 'Outer');
  assert.ok(cell, 'the outer cell must still be traced as one textbox-bearing tablix cell');
  assert.equal(cell.itemName, 'Caption');
  assert.equal(cell.text, 'Legend');
  assert.equal(round(cell.width), 5 * PT);
  assert.equal(page.items.some((entry) => entry.kind === 'textbox' && entry.itemName === 'Caption'), false);
});

test('a caption that declares no free-form box keeps the fill-the-cell path instead of vanishing', async () => {
  const layout = await trace(unsizedCanvasReport());
  const [page] = layout.pages;
  const cell = page.items.find((entry) => entry.kind === 'tablixCell' && entry.tablixName === 'Outer');
  assert.equal(cell.itemName, 'Caption');
  assert.equal(cell.text, 'Legend');
  assert.equal(round(cell.width), 5 * PT);
  // No zero-sized positioned copy was emitted in its place.
  assert.equal(page.items.some((entry) => entry.kind === 'textbox' && entry.itemName === 'Caption'), false);
});

test('DOCX_EDITABLE renders the caption and the nested region instead of refusing overlapping regions', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-canvas-docx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderEditableDocx(
    parseRdl(canvasReport()),
    { output: 'DOCX_EDITABLE', outputFileName: 'canvas', parameters: {}, datasets: DATA },
    config,
    tempDir,
  );
  const zip = await JSZip.loadAsync(rendered.buffer);
  const document = await zip.file('word/document.xml').async('string');
  const texts = [...document.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((match) => match[1]).join(' ');
  assert.match(texts, /Legend/);
  assert.match(texts, /AU/);
  // The label keeps the canonical PDF's explicit wrap points, so its words are separate native runs.
  assert.match(texts, /Assurance/);
  assert.match(texts, /Services/);
  assert.match(document, /<w:tbl>/);
  assert.equal(Object.keys(zip.files).some((name) => /^word\/media\//.test(name)), false, 'no page screenshot');
});

test('XLSX writes both the caption and the nested region content', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-canvas-xlsx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderExcel(
    parseRdl(canvasReport()),
    { output: 'XLSX', excelLayoutMode: 'REPORT', outputFileName: 'canvas', parameters: {}, datasets: DATA },
    config,
    tempDir,
  );
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(rendered.buffer);
  const values = [];
  workbook.worksheets.forEach((sheet) => sheet.eachRow((row) => row.eachCell((cell) => {
    if (cell.value) values.push(String(cell.value));
  })));
  assert.ok(values.includes('Legend'), `expected the caption in ${JSON.stringify(values)}`);
  assert.ok(values.includes('AU'), `expected the nested code in ${JSON.stringify(values)}`);
  assert.ok(values.includes('Assurance Services'), `expected the nested label in ${JSON.stringify(values)}`);
});
