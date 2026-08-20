// An RDL Style property can be PRESENT BUT EMPTY (`<FontFamily />`, `<TextAlign></TextAlign>`) — report
// writers emit those elements for properties the author never set. SSRS ignores an empty element and uses
// the inherited/default value. Carrying the empty string into the normalized model instead made the property
// look declared-but-blank, and the blank family reached font selection as a nameless font: under strict
// fonts every output of such a report failed closed with `FONT_MISSING: Required font is unavailable: `
// (PDF, editable DOCX, and XLSX alike — Excel measures rows through the PDF text measurer), and under
// non-strict fonts the report silently dropped to base-14 Helvetica. The same must hold for an =expression
// that evaluates to "".
import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { parseRdl } from '../src/rdl/parser.js';
import { styleText, styleValue } from '../src/render/common.js';
import { renderPdf } from '../src/render/pdf.js';
import { renderExcel } from '../src/render/excel.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { resolveFontFile } from '../src/render/fonts.js';
import { loadConfig } from '../src/config.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const request = { parameters: {}, datasets: {}, outputFileName: 'empty-style' };
const context = { fields: {}, parameters: {}, globals: {}, dataset: [], datasets: {} };

// Minimal free-form report: one textbox whose Style is whatever the case under test declares.
function reportRdl(styleXml, { defaultFontFamily = null } = {}) {
  const defaultFontNs = defaultFontFamily === null
    ? ''
    : ` xmlns:df="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition/defaultfontfamily" MustUnderstand="df"`;
  const defaultFontElement = defaultFontFamily === null ? '' : `<df:DefaultFontFamily>${defaultFontFamily}</df:DefaultFontFamily>`;
  return `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition"${defaultFontNs}>
 ${defaultFontElement}
 <ReportSections><ReportSection><Body><ReportItems>
  <Textbox Name="t"><Paragraphs><Paragraph><TextRuns><TextRun><Value>hello</Value></TextRun></TextRuns></Paragraph></Paragraphs>
    <Top>0in</Top><Left>0in</Left><Width>2in</Width><Height>0.25in</Height>
    <Style>${styleXml}</Style></Textbox>
 </ReportItems><Height>1in</Height><Style/></Body>
 <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth></Page></ReportSection></ReportSections></Report>`;
}

const textboxOf = (model) => model.body.items.find((item) => item.type === 'Textbox');

test('an empty <FontFamily /> element normalizes to the report default, not a blank family', () => {
  const blank = textboxOf(parseRdl(reportRdl('<FontFamily />')));
  assert.equal(blank.style.fontFamily, 'Arial');
  assert.equal(blank.paragraphs[0][0].style.fontFamily, 'Arial'); // the TextRun style inherits it too

  // A whitespace-only element is equally undeclared, and a declared family is still honoured verbatim.
  assert.equal(textboxOf(parseRdl(reportRdl('<FontFamily>   </FontFamily>'))).style.fontFamily, 'Arial');
  assert.equal(textboxOf(parseRdl(reportRdl('<FontFamily>Times New Roman</FontFamily>'))).style.fontFamily, 'Times New Roman');
});

test('an empty <FontFamily /> falls back to df:DefaultFontFamily when the report declares one', () => {
  const declared = parseRdl(reportRdl('<FontFamily />', { defaultFontFamily: 'Segoe UI' }));
  assert.equal(textboxOf(declared).style.fontFamily, 'Segoe UI');

  // An empty default-font element is itself undeclared and must not become the blank family either.
  const blankDefault = parseRdl(reportRdl('<FontFamily />', { defaultFontFamily: '' }));
  assert.equal(blankDefault.defaultFontFamily, 'Arial');
  assert.equal(textboxOf(blankDefault).style.fontFamily, 'Arial');
});

test('other empty style elements normalize to their defaults instead of a blank string', () => {
  const style = textboxOf(parseRdl(reportRdl(
    '<Color /><FontWeight /><FontStyle /><TextDecoration /><TextAlign /><VerticalAlign /><WritingMode /><Format /><BackgroundColor />',
  ))).style;
  assert.equal(style.color, '#000000');
  assert.equal(style.fontWeight, 'Normal');
  assert.equal(style.fontStyle, 'Normal');
  assert.equal(style.textDecoration, 'None');
  assert.equal(style.textAlign, 'Left');
  assert.equal(style.verticalAlign, 'Top');
  assert.equal(style.writingMode, 'Default');
  assert.equal(style.format, null);
  assert.equal(style.backgroundColor, null);
});

test('an empty side-border element inherits the general Border instead of shadowing it', () => {
  const style = textboxOf(parseRdl(reportRdl(
    '<Border><Style>Solid</Style><Color>#ff0000</Color><Width>2pt</Width></Border>'
    + '<TopBorder><Style /><Color /><Width /></TopBorder>'
    + '<BottomBorder><Style>None</Style></BottomBorder>',
  ))).style;
  assert.equal(style.borders.top.style, 'Solid'); // empty -> inherited, not "None"
  assert.equal(style.borders.top.color, '#ff0000');
  assert.equal(style.borders.top.width, 2);
  assert.equal(style.borders.bottom.style, 'None'); // an explicit None is still an explicit value
});

test('styleText resolves a blank expression result to the fallback, unlike styleValue', () => {
  const empty = '=IIf(1 = 1, "", "Segoe UI")';
  assert.equal(styleValue(empty, context, 'Arial'), ''); // raw resolution keeps the empty string
  assert.equal(styleText(empty, context, 'Arial'), 'Arial');
  assert.equal(styleText('=IIf(1 = 1, "Times New Roman", "")', context, 'Arial'), 'Times New Roman');
  assert.equal(styleText('   ', context, 'Arial'), 'Arial');
  assert.equal(styleText(undefined, context, 'Arial'), 'Arial');
  assert.equal(styleText('Segoe UI', context, 'Arial'), 'Segoe UI');
});

test('every output renders an empty <FontFamily /> under strict fonts instead of failing FONT_MISSING', async () => {
  const strict = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'true' });
  // Strict mode fails closed on any family this host lacks, so the assertion is only meaningful where the
  // default family itself resolves; elsewhere the non-strict renderer assertions below still run.
  if (!resolveFontFile(strict.fontDir, 'Arial', false, false)) return;
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-empty-style-strict-'));
  try {
    // PDF is the canonical layout; XLSX measures rows through the same PDF text measurer and editable DOCX
    // is built from the PDF trace, so a blank family failed all three closed.
    const pdf = await renderPdf(parseRdl(reportRdl('<FontFamily /><FontSize>10pt</FontSize>')), request, strict);
    assert.ok(pdf.buffer.length > 0);
    assert.equal(pdf.pageCount, 1);
    assert.ok((await renderExcel(parseRdl(reportRdl('<FontFamily />')), request, strict, tempDir)).buffer.length > 0);
    assert.ok((await renderEditableDocx(parseRdl(reportRdl('<FontFamily />')), request, strict)).buffer.length > 0);

    // An expression-valued family that evaluates to "" must not fail the export either.
    const expressionRdl = reportRdl('<FontFamily>=IIf(1 = 1, "", "Arial")</FontFamily>');
    assert.ok((await renderPdf(parseRdl(expressionRdl), request, strict)).buffer.length > 0);
    assert.ok((await renderExcel(parseRdl(expressionRdl), request, strict, tempDir)).buffer.length > 0);
    assert.ok((await renderEditableDocx(parseRdl(expressionRdl), request, strict)).buffer.length > 0);
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
});

test('XLSX writes the default family for an empty <FontFamily />, never a blank font name', async () => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-empty-style-xlsx-'));
  try {
    const rendered = await renderExcel(parseRdl(reportRdl('<FontFamily />')), request, config, tempDir);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(rendered.buffer);
    const names = new Set();
    workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => names.add(cell.font?.name)));
    assert.ok(names.size > 0);
    for (const name of names) assert.equal(name, 'Arial');
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
});

test('editable DOCX writes the default family for an empty <FontFamily />, never a blank one', async () => {
  const rendered = await renderEditableDocx(parseRdl(reportRdl('<FontFamily />')), request, config);
  const xml = await (await JSZip.loadAsync(rendered.buffer)).file('word/document.xml').async('string');
  assert.ok(xml.includes('hello'));
  assert.match(xml, /w:ascii="Arial"/);
  assert.doesNotMatch(xml, /w:ascii=""/); // a blank family must never reach the OOXML run properties
});
