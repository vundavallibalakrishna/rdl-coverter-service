// RDL designers make adjacent free-form boxes overlap by their shared border stroke so a fixed-layout
// renderer draws one continuous rule. Excel has no overlapping merged ranges, so those two coordinates are
// resolved to one grid edge. The resolver only compared each box with its immediate predecessor in the
// left-sorted sibling list, so any unrelated item whose left fell between two genuine neighbours hid the
// pair and the export failed closed with RDL_INVALID. It now compares every vertically-overlapping pair and
// resolves the recorded equivalences as clusters, so the canonical edge does not depend on discovery order.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { loadConfig } from '../src/config.js';

const execFileAsync = promisify(execFile);
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const DATA = { D: [{ N: 'row' }] };

const BORDER = `<Style>
  <Border><Style>Solid</Style></Border>
  <TopBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></TopBorder>
  <BottomBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></BottomBorder>
  <LeftBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></LeftBorder>
  <RightBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></RightBorder>
</Style>`;

function box(name, value, { left, top, width, height }) {
  return `<Textbox Name="${name}">
    <Paragraphs><Paragraph><TextRuns><TextRun><Value>${value}</Value></TextRun></TextRuns></Paragraph></Paragraphs>
    <Top>${top}in</Top><Left>${left}in</Left><Height>${height}in</Height><Width>${width}in</Width>
    ${BORDER}
  </Textbox>`;
}

function report(bandItems) {
  return Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D"><Fields><Field Name="N"><DataField>N</DataField><TypeName>System.String</TypeName></Field></Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Rectangle Name="Band"><ReportItems>${bandItems}</ReportItems>
      <Top>0in</Top><Left>0in</Left><Height>2in</Height><Width>7in</Width><Style/></Rectangle>
  </ReportItems><Height>2in</Height><Style/></Body><Width>7.5in</Width>
  <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`, 'utf8');
}

// `Divider` is the interleaver: its Left (1.5in) sorts between `Signature` (0.5in) and `Date` (2.49in), so
// the two boxes that actually share a rule are not adjacent in the left-sorted list. It sits in its own
// vertical band, so it shares no edge with either.
const SIGNATURE = box('Signature', 'Signature', { left: 0.5, top: 1, width: 2, height: 0.25 });
const DIVIDER = box('Divider', 'Divider', { left: 1.5, top: 0.2, width: 1, height: 0.25 });
// Overlaps Signature's right edge by 0.01in (0.72pt) — inside the 1pt border both boxes declare.
const DATE = box('Date', 'Date', { left: 2.49, top: 1, width: 2, height: 0.25 });
// Overlaps by 0.1in (7.2pt) — far beyond any shared stroke, so it must stay a refused overlap.
const DEEP_DATE = box('Date', 'Date', { left: 2.4, top: 1, width: 2, height: 0.25 });

const interleavedReport = () => report(SIGNATURE + DIVIDER + DATE);
const deepOverlapReport = () => report(SIGNATURE + DIVIDER + DEEP_DATE);
// Two stacked left boxes whose right edges differ slightly both meet one tall right box on the same rule.
// The shared `Date` left edge appears in two equivalence pairs, so a last-write-wins alias map would strand
// one of them on a stale canonical edge and leave the ranges overlapping.
const sharedRuleReport = () => report(
  box('TopLeft', 'Top', { left: 0.5, top: 1, width: 2, height: 0.25 })
  + box('BottomLeft', 'Bottom', { left: 0.5, top: 1.25, width: 1.9931, height: 0.25 })
  + DIVIDER
  + box('Date', 'Date', { left: 2.4896, top: 1, width: 2, height: 0.5 }),
);

async function renderWorkbook(rdl, context) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-shared-edge-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderExcel(
    parseRdl(rdl),
    { output: 'XLSX', excelLayoutMode: 'REPORT', outputFileName: 'edges', parameters: {}, datasets: DATA },
    config,
    tempDir,
  );
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(rendered.buffer);
  return workbook;
}

// The 1-based column span each value occupies, across every sheet row.
function columnsOf(workbook, wanted) {
  const columns = [];
  workbook.worksheets.forEach((sheet) => sheet.eachRow((row) => row.eachCell((cell, column) => {
    if (String(cell.value ?? '') === wanted) columns.push(column);
  })));
  assert.ok(columns.length > 0, `expected ${wanted} to occupy at least one column`);
  return { start: Math.min(...columns), end: Math.max(...columns) };
}

test('XLSX resolves a shared edge between boxes that are not adjacent in the left-sorted order', async (context) => {
  const workbook = await renderWorkbook(interleavedReport(), context);
  const signature = columnsOf(workbook, 'Signature');
  const date = columnsOf(workbook, 'Date');
  assert.ok(signature.end < date.start, `Signature ${JSON.stringify(signature)} must not overlap Date ${JSON.stringify(date)}`);
  assert.equal(date.start, signature.end + 1, 'the shared stroke must become one grid edge, leaving no gap band');
});

test('XLSX resolves one rule shared by three boxes to a single edge regardless of pair order', async (context) => {
  const workbook = await renderWorkbook(sharedRuleReport(), context);
  const top = columnsOf(workbook, 'Top');
  const bottom = columnsOf(workbook, 'Bottom');
  const date = columnsOf(workbook, 'Date');
  assert.equal(top.end, bottom.end, 'both left boxes must end on the same resolved grid edge');
  assert.equal(date.start, top.end + 1, 'the right box must start immediately after that edge');
});

test('XLSX still fails closed on an overlap far beyond any shared border stroke', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-shared-edge-deep-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  await assert.rejects(
    () => renderExcel(
      parseRdl(deepOverlapReport()),
      { output: 'XLSX', excelLayoutMode: 'REPORT', outputFileName: 'edges', parameters: {}, datasets: DATA },
      config,
      tempDir,
    ),
    (error) => error.code === 'RDL_INVALID' && /overlapping Excel merged-cell ranges/.test(error.message),
  );
});

test('PDF draws both boxes of the shared-edge pair (fixed layout keeps the declared geometry)', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-shared-edge-pdf-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderPdf(parseRdl(interleavedReport()), { outputFileName: 'edges', parameters: {}, datasets: DATA }, config);
  const pdfPath = path.join(tempDir, 'edges.pdf');
  await fs.writeFile(pdfPath, rendered.buffer);
  const { stdout } = await execFileAsync('pdftotext', ['-layout', pdfPath, '-']);
  assert.match(stdout, /Signature/);
  assert.match(stdout, /Date/);
});

test('DOCX_EDITABLE renders the same shared-edge pair as native cells', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-shared-edge-docx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderEditableDocx(
    parseRdl(interleavedReport()),
    { output: 'DOCX_EDITABLE', outputFileName: 'edges', parameters: {}, datasets: DATA },
    config,
    tempDir,
  );
  const zip = await JSZip.loadAsync(rendered.buffer);
  const document = await zip.file('word/document.xml').async('string');
  const texts = [...document.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((match) => match[1]).join(' ');
  assert.match(texts, /Signature/);
  assert.match(texts, /Date/);
  assert.match(document, /<w:tbl>/);
});
