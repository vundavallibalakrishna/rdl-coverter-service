// A Style/Format date pattern is resolved once, in rdl/expression.js `formatValue`, and every renderer
// consumes the result. That function used to carry its own date fast path whose token table held only
// {yyyy, MMM, MM, dd, HH, mm, ss} and substituted them through one first-match-wins alternation. A longer
// token was therefore cut short and its remainder emitted literally — "MMMM yyyy" rendered "AugM 2026"
// instead of "August 2026" — while the single-character forms and tt were not tokens at all and survived
// verbatim. These tests pin the whole custom-pattern vocabulary at that shared layer and then prove the
// resolved text reaches PDF, editable DOCX, and XLSX.
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { promisify } from 'node:util';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { formatValue } from '../src/rdl/expression.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';

const execFileAsync = promisify(execFile);
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const SAMPLE = new Date(Date.UTC(2026, 7, 13, 8, 4, 53));

test('custom date patterns resolve every .NET token instead of truncating the longest form', () => {
  // Month names: the reported defect. MMMM must not be served by the MMM branch.
  assert.equal(formatValue(SAMPLE, 'MMMM yyyy'), 'August 2026');
  assert.equal(formatValue(SAMPLE, 'MMMM'), 'August');
  assert.equal(formatValue(SAMPLE, 'MMM'), 'Aug');
  assert.equal(formatValue(SAMPLE, 'MM'), '08');
  assert.equal(formatValue(SAMPLE, 'M'), '8');
  // Day names truncated the same way, and the single-character day was not a token at all.
  assert.equal(formatValue(SAMPLE, 'dddd'), 'Thursday');
  assert.equal(formatValue(SAMPLE, 'ddd'), 'Thu');
  assert.equal(formatValue(SAMPLE, 'dd'), '13');
  assert.equal(formatValue(SAMPLE, 'd MMMM yyyy'), '13 August 2026');
  assert.equal(formatValue(SAMPLE, 'dddd, MMMM d, yyyy'), 'Thursday, August 13, 2026');
  // 12-hour clock and the AM/PM designator.
  assert.equal(formatValue(SAMPLE, 'hh:mm tt'), '08:04 AM');
  assert.equal(formatValue(SAMPLE, 'h:mm tt'), '8:04 AM');
});

test('the tokens that already worked are byte-for-byte unchanged, in UTC', () => {
  assert.equal(formatValue(SAMPLE, 'yyyy'), '2026');
  assert.equal(formatValue(SAMPLE, 'HH'), '08');
  assert.equal(formatValue(SAMPLE, 'mm'), '04');
  assert.equal(formatValue(SAMPLE, 'ss'), '53');
  assert.equal(formatValue(SAMPLE, 'dd/MM/yyyy'), '13/08/2026');
  assert.equal(formatValue(SAMPLE, 'MM/dd/yyyy'), '08/13/2026');
  assert.equal(formatValue(SAMPLE, 'dd/MM/yyyy HH:mm:ss'), '13/08/2026 08:04:53');
  // No format at all resolves to .NET/SSRS general date/time (date AND time, with seconds), the value's
  // default ToString — not a bare date and not a raw JS Date string.
  assert.equal(formatValue(SAMPLE, undefined), '13/08/2026 08:04:53');
  // A pattern must resolve in UTC, never in the host time zone: 22:30Z is still the 13th, not the 14th.
  const lateUtc = new Date(Date.UTC(2026, 7, 13, 22, 30, 0));
  assert.equal(formatValue(lateUtc, 'dd MMMM yyyy'), '13 August 2026');
  assert.equal(formatValue(lateUtc, 'HH:mm'), '22:30');
});

test('the standard single-letter specifiers keep their existing shortcut output', () => {
  assert.equal(formatValue(SAMPLE, 'y'), 'August 2026');
  assert.equal(formatValue(SAMPLE, 'Y'), 'August 2026');
  assert.equal(formatValue(SAMPLE, 'd'), '13/08/2026');
  assert.equal(formatValue(SAMPLE, 'D'), '13 August 2026');
});

test('every month abbreviates the way .NET does, including the four-letter CLDR outlier', () => {
  // en-GB CLDR abbreviates September as "Sept"; .NET (and therefore SSRS) uses "Sep". The engine must
  // follow .NET, since the format string is a .NET format string.
  const expected = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  for (let month = 0; month < 12; month += 1) {
    assert.equal(formatValue(new Date(Date.UTC(2026, month, 5)), 'MMM'), expected[month]);
  }
});

// One synthetic textbox carrying the reported Style/Format, so the renderer assertions differ from the unit
// assertions above only by the transport.
function report() {
  return Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <ReportSections><ReportSection>
    <Body>
      <ReportItems>
        <Textbox Name="IssueDate">
          <CanGrow>true</CanGrow>
          <Paragraphs><Paragraph><TextRuns><TextRun>
            <Value>=Parameters!Stamp.Value</Value>
            <Style><FontFamily>Arial</FontFamily><Format>MMMM yyyy</Format></Style>
          </TextRun></TextRuns></Paragraph></Paragraphs>
          <Top>0in</Top><Left>0in</Left><Height>0.3in</Height><Width>5in</Width>
          <Style><FontFamily>Arial</FontFamily><Format>MMMM yyyy</Format></Style>
        </Textbox>
      </ReportItems>
      <Height>1in</Height><Style/>
    </Body>
    <Width>7in</Width>
    <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
  <ReportParameters>
    <ReportParameter Name="Stamp"><DataType>DateTime</DataType><Prompt>Stamp</Prompt></ReportParameter>
  </ReportParameters>
</Report>`, 'utf8');
}

const request = {
  outputFileName: 'date-format-tokens',
  parameters: { Stamp: '2026-08-13T08:04:53Z' },
  datasets: {},
};

test('PDF renders the resolved full month name, not the truncated token', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-date-format-pdf-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, 'date-format.pdf');
  const rendered = await renderPdf(parseRdl(report()), request, config);
  await fs.writeFile(pdfPath, rendered.buffer);
  const { stdout } = await execFileAsync('pdftotext', ['-layout', pdfPath, '-']);
  assert.match(stdout, /August 2026/);
  assert.doesNotMatch(stdout, /AugM/);
});

test('editable DOCX carries the resolved month name as native text', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-date-format-docx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderEditableDocx(parseRdl(report()), request, config, tempDir);
  const zip = await JSZip.loadAsync(rendered.buffer);
  const document = await zip.file('word/document.xml').async('string');
  // Word splits a paragraph into runs at style boundaries, so the date arrives as several w:t elements.
  // Join them back before asserting on the displayed string.
  const texts = (document.match(/<w:t[^>]*>[^<]*<\/w:t>/g) || []).map((run) => run.replace(/<[^>]+>/g, '')).join('');
  assert.match(texts, /August 2026/);
  assert.doesNotMatch(texts, /AugM/);
});

test('XLSX writes the resolved month name for a date format Excel cannot represent natively', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-date-format-xlsx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderExcel(parseRdl(report()), { ...request, excelLayoutMode: 'REPORT' }, config, tempDir);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(rendered.buffer);
  const values = [];
  workbook.worksheets.forEach((sheet) => sheet.eachRow((row) => row.eachCell((cell) => {
    const value = String(cell.value ?? '').trim();
    if (value) values.push(value);
  })));
  assert.ok(values.some((value) => value.includes('August 2026')), `expected "August 2026" in ${JSON.stringify(values)}`);
  assert.equal(values.some((value) => value.includes('AugM')), false);
});

// Same synthetic textbox, but with NO Format — SSRS renders a DateTime's general date/time default
// (date AND time), not a bare date and not a raw Date string. Every renderer must agree.
function reportNoFormat() {
  return Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <ReportSections><ReportSection>
    <Body>
      <ReportItems>
        <Textbox Name="Stamp">
          <CanGrow>true</CanGrow>
          <Paragraphs><Paragraph><TextRuns><TextRun>
            <Value>=Parameters!Stamp.Value</Value>
            <Style><FontFamily>Arial</FontFamily></Style>
          </TextRun></TextRuns></Paragraph></Paragraphs>
          <Top>0in</Top><Left>0in</Left><Height>0.3in</Height><Width>5in</Width>
          <Style><FontFamily>Arial</FontFamily></Style>
        </Textbox>
      </ReportItems>
      <Height>1in</Height><Style/>
    </Body>
    <Width>7in</Width>
    <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
  <ReportParameters>
    <ReportParameter Name="Stamp"><DataType>DateTime</DataType><Prompt>Stamp</Prompt></ReportParameter>
  </ReportParameters>
</Report>`, 'utf8');
}

test('a no-format DateTime renders general date/time (date and time) in PDF', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-date-noformat-pdf-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, 'date-noformat.pdf');
  const rendered = await renderPdf(parseRdl(reportNoFormat()), request, config);
  await fs.writeFile(pdfPath, rendered.buffer);
  const { stdout } = await execFileAsync('pdftotext', ['-layout', pdfPath, '-']);
  assert.match(stdout, /13\/08\/2026 08:04:53/);
});

test('a no-format DateTime is a typed Excel date with the general date/time number format', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-date-noformat-xlsx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderExcel(parseRdl(reportNoFormat()), { ...request, excelLayoutMode: 'REPORT' }, config, tempDir);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(rendered.buffer);
  let dateCell = null;
  workbook.worksheets.forEach((sheet) => sheet.eachRow((row) => row.eachCell((cell) => {
    if (cell.value instanceof Date) dateCell = cell;
  })));
  assert.ok(dateCell, 'expected a typed Date cell');
  assert.equal(dateCell.numFmt, 'yyyy-mm-dd hh:mm:ss');
});
