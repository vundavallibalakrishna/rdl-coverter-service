// The report `Language` (culture) drives every value's formatting — separators, currency symbol, month/day
// names, and standard date ordering — exactly as SSRS does. A report with no Language keeps the engine's
// legacy defaults so nothing regresses. These assertions cover the format engine, the parser, and the full
// render plumbing (globals.culture -> cultureFor -> formatValue).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { formatValue } from '../src/rdl/expression.js';
import { canonicalizeCulture, currencyForCulture } from '../src/rdl/culture.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';
import { loadConfig } from '../src/config.js';

const execFileAsync = promisify(execFile);
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const SAMPLE = new Date(Date.UTC(2026, 2, 4, 8, 4, 53)); // 4 March 2026

test('culture helpers resolve locale and its default currency', () => {
  assert.equal(canonicalizeCulture('en-ZA'), 'en-ZA');
  assert.equal(canonicalizeCulture('  de-DE '), 'de-DE');
  assert.equal(canonicalizeCulture(''), null);
  assert.equal(canonicalizeCulture('not a locale!!'), null);
  assert.equal(currencyForCulture('en-US'), 'USD');
  assert.equal(currencyForCulture('en-GB'), 'GBP');
  assert.equal(currencyForCulture('en-ZA'), 'ZAR');
  assert.equal(currencyForCulture('de-DE'), 'EUR');
  assert.equal(currencyForCulture(null), 'USD');
});

test('formatValue with no culture keeps the legacy defaults (no regression)', () => {
  assert.equal(formatValue(1234.5, 'C2', null), '$1,234.50');
  assert.equal(formatValue(1234.5, 'N2', null), '1,234.50');
  assert.equal(formatValue(SAMPLE, 'd', null), '04/03/2026');
});

test('formatValue follows a declared culture for currency, separators, and month names', () => {
  assert.equal(formatValue(1234.5, 'C2', 'en-GB'), '£1,234.50');
  assert.equal(formatValue(1234.5, 'C2', 'de-DE'), '1.234,50 €');
  assert.equal(formatValue(1234.5, 'N2', 'de-DE'), '1.234,50');
  assert.equal(formatValue(SAMPLE, 'MMMM', 'de-DE'), 'März');
  assert.equal(formatValue(SAMPLE, 'dddd', 'fr-FR'), 'mercredi');
  assert.equal(formatValue(SAMPLE, 'd', 'en-US'), '3/4/26');
  assert.equal(formatValue(SAMPLE, 'd', 'en-ZA'), '2026/03/04');
});

test('the parser records the report culture (literal Language), null when absent', () => {
  const withLang = parseRdl(reportRdl('de-DE'));
  assert.equal(withLang.culture, 'de-DE');
  assert.equal(parseRdl(reportRdl('')).culture, null);
});

function reportRdl(language) {
  return Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  ${language ? `<Language>${language}</Language>` : ''}
  <ReportSections><ReportSection><Body><ReportItems>
    <Textbox Name="Amt"><Paragraphs><Paragraph><TextRuns><TextRun>
      <Value>=1234.5</Value><Style><FontFamily>Arial</FontFamily><Format>C2</Format></Style>
    </TextRun></TextRuns></Paragraph></Paragraphs><Top>0in</Top><Left>0in</Left><Height>0.3in</Height><Width>4in</Width><Style><Format>C2</Format></Style></Textbox>
  </ReportItems><Height>1in</Height><Style/></Body><Width>7in</Width>
  <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections></Report>`, 'utf8');
}

test('PDF render applies the declared culture end to end', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-culture-pdf-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const request = { outputFileName: 'culture', parameters: {}, datasets: {} };

  const de = await renderPdf(parseRdl(reportRdl('de-DE')), request, config);
  const dePath = path.join(tempDir, 'de.pdf');
  await fs.writeFile(dePath, de.buffer);
  const { stdout: deText } = await execFileAsync('pdftotext', ['-layout', dePath, '-']);
  assert.match(deText, /1\.234,50/);

  const none = await renderPdf(parseRdl(reportRdl('')), request, config);
  const nonePath = path.join(tempDir, 'none.pdf');
  await fs.writeFile(nonePath, none.buffer);
  const { stdout: noneText } = await execFileAsync('pdftotext', ['-layout', nonePath, '-']);
  assert.match(noneText, /\$1,234\.50/);
});
