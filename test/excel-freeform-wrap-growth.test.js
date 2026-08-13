// A free-form CanGrow textbox is written to a merged, wrap-enabled Excel cell whose row band is sized from
// declared RDL geometry alone. Excel then re-wraps the text with its own engine inside a narrower usable
// width, and a merged cell cannot spill sideways — so every line past the declared height is simply not
// displayed. The band has to be raised to what the text needs.
//
// Raising it is only half the fix. RDL VerticalAlign describes the DECLARED box, where a grown box is
// exactly its own text and Middle and Top coincide. Keeping Middle against the enlarged band instead
// splits the surplus symmetrically, which reads as a large blank gap between a heading and its paragraph.
// A grown cell therefore anchors its text to the top, so any surplus trails below it.
//
// XLSX-specific: PDF grows a CanGrow textbox during layout and DOCX inherits the canonical PDF trace. The
// PDF assertion below is the counterexample proving that.
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { promisify } from 'node:util';
import ExcelJS from 'exceljs';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

const execFileAsync = promisify(execFile);
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const request = { outputFileName: 'excel-freeform-wrap-growth', parameters: {}, datasets: {} };

const SHORT = 'SHORT_LABEL';
// Sits at or past the wrap point of the 1.6in box below, which is the case Excel reflows into a line the
// declared height cannot show.
const LONG = 'A DELIBERATELY LONG SECTION HEADING THAT MUST WRAP ONTO FURTHER LINES IN ANY ENGINE';

// `verticalAlign` is what separates the two alignment scenarios; `body` separates the growth ones.
function report({ body = SHORT, verticalAlign = 'Middle' } = {}) {
  return Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <ReportSections><ReportSection>
    <Body>
      <ReportItems>
        <Textbox Name="Heading">
          <CanGrow>true</CanGrow>
          <Paragraphs><Paragraph><TextRuns><TextRun><Value>HEADING_MARKER</Value></TextRun></TextRuns></Paragraph></Paragraphs>
          <Top>0in</Top><Left>0in</Left><Height>0.2in</Height><Width>4in</Width>
          <Style><FontFamily>Arial</FontFamily><FontSize>11pt</FontSize><FontWeight>Bold</FontWeight></Style>
        </Textbox>
        <Textbox Name="Body">
          <CanGrow>true</CanGrow>
          <Paragraphs><Paragraph><TextRuns><TextRun><Value>${body}</Value></TextRun></TextRuns></Paragraph></Paragraphs>
          <Top>0.4in</Top><Left>0in</Left><Height>0.2in</Height><Width>1.6in</Width>
          <Style><FontFamily>Arial</FontFamily><FontSize>11pt</FontSize><VerticalAlign>${verticalAlign}</VerticalAlign></Style>
        </Textbox>
      </ReportItems>
      <Height>1.2in</Height><Style/>
    </Body>
    <Width>7in</Width>
    <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`, 'utf8');
}

async function bodyCell(options) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-xlsx-grow-'));
  try {
    const rendered = await renderExcel(parseRdl(report(options)), { ...request, excelLayoutMode: 'REPORT' }, config, tempDir);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(rendered.buffer);
    const sheet = workbook.worksheets[0];
    const marker = String(options.body ?? SHORT).slice(0, 20);
    let found = null;
    sheet.eachRow((row, rowNumber) => row.eachCell((cell) => {
      const value = cell.value?.richText
        ? cell.value.richText.map((run) => run.text).join('')
        : String(cell.value ?? '');
      if (!found && value.includes(marker)) found = { cell, rowNumber };
    }));
    assert.ok(found, `expected a cell containing ${marker}`);
    // The item's band can span several worksheet rows, and only the anchor row carries the value. Height
    // is therefore the sum over the merged range, not the anchor row alone.
    const merge = (sheet.model.merges || [])
      .map((entry) => entry.match(/^[A-Z]+(\d+):[A-Z]+(\d+)$/))
      .find((match) => match && Number(match[1]) <= found.rowNumber && found.rowNumber <= Number(match[2]));
    const startRow = merge ? Number(merge[1]) : found.rowNumber;
    const endRow = merge ? Number(merge[2]) : found.rowNumber;
    let height = 0;
    for (let row = startRow; row <= endRow; row += 1) height += sheet.getRow(row).height || 0;
    return { cell: found.cell, height };
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

test('a free-form CanGrow textbox gets a band tall enough for Excel to show every line', async () => {
  const short = await bodyCell({ body: SHORT });
  const long = await bodyCell({ body: LONG });
  // Counterexample first: content that fits keeps its declared compact geometry, so this is not a blanket
  // "make every row taller".
  assert.ok(short.height < 20, `a fitting single line must keep its one-line height, got ${short.height}`);
  assert.ok(
    long.height > short.height * 1.5,
    `wrapping content must raise the band (short ${short.height} vs long ${long.height})`,
  );
});

test('a grown cell anchors its text to the top so the surplus cannot become a gap above it', async () => {
  const long = await bodyCell({ body: LONG, verticalAlign: 'Middle' });
  assert.ok(long.height > 20, 'precondition: this cell must have grown');
  assert.equal(long.cell.alignment?.vertical, 'top');
});

test('a cell that did not grow keeps the RDL VerticalAlign it declared', async () => {
  const middle = await bodyCell({ body: SHORT, verticalAlign: 'Middle' });
  assert.ok(middle.height < 20, 'precondition: this cell must not have grown');
  assert.equal(middle.cell.alignment?.vertical, 'middle');
});

test('a block past one row ceiling spreads across several rows instead of being capped', async () => {
  const paragraph = Array.from({ length: 400 }, (unused, index) => `SENTENCE_${index}`).join(' ');
  const long = await bodyCell({ body: paragraph });
  assert.ok(long.height > 100, 'a very long block must still grow');
  assert.ok(long.height > 409, "a block past one row ceiling must spread across several rows");
});

test('PDF is unaffected: it grows the textbox during layout either way', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-xlsx-grow-pdf-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, 'grow.pdf');
  const rendered = await renderPdf(parseRdl(report({ body: LONG })), request, config);
  await fs.writeFile(pdfPath, rendered.buffer);
  const { stdout } = await execFileAsync('pdftotext', ['-layout', pdfPath, '-']);
  assert.match(stdout, /HEADING_MARKER/);
  assert.match(stdout, /DELIBERATELY/);
  assert.match(stdout, /ANY ENGINE/);
});
