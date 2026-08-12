// RDL declares PageBreak on ReportItemType — every report item carries it, not only a direct <Body> child.
// The renderers used to consult it exclusively on top-level body items, so a break declared on a rectangle
// (or tablix) nested inside another rectangle was parsed, ignored, and silently dropped from the output.
// These tests pin the construct generically: the break must move the page wherever it is declared, must
// still be spent on whatever follows when the declaring item is the last child of its container, and must
// stay inert when RDL says it is disabled.
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
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderVisualDocx } from '../src/render/visualDocx.js';

const execFileAsync = promisify(execFile);
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const request = { outputFileName: 'nested-page-break', parameters: {}, datasets: {} };
const tmpRoot = path.resolve(path.dirname(new URL(import.meta.url).pathname).replace(/^\/(?=[A-Za-z]:)/, ''), '..', 'tmp');

function textbox(name, value, top) {
  return `
    <Textbox Name="${name}">
      <CanGrow>true</CanGrow>
      <Paragraphs><Paragraph><TextRuns><TextRun><Value>${value}</Value></TextRun></TextRuns></Paragraph></Paragraphs>
      <Top>${top}in</Top><Left>0in</Left><Height>0.3in</Height><Width>5in</Width>
      <Style><FontFamily>Arial</FontFamily></Style>
    </Textbox>`;
}

// One synthetic shape, parameterised over where the break sits and what it says, so every assertion below
// differs from its counterexample by exactly the RDL property under test.
function report({ breakXml = '', outerStyle = '<Style/>', depth = 1 } = {}) {
  const inner = `
    <Rectangle Name="InnerBreak">
      <ReportItems>${textbox('Second', 'BLOCK_TWO', 0)}</ReportItems>
      ${breakXml}
      <KeepTogether>true</KeepTogether>
      <Top>0.4in</Top><Left>0in</Left><Height>0.4in</Height><Width>5in</Width><Style/>
    </Rectangle>`;
  const nested = depth === 2
    ? `<Rectangle Name="Middle"><ReportItems>${inner}</ReportItems>
         <Top>0in</Top><Left>0in</Left><Height>0.9in</Height><Width>5in</Width><Style/></Rectangle>`
    : inner;
  return Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <ReportSections><ReportSection>
    <Body>
      <ReportItems>
        <Rectangle Name="Outer">
          <ReportItems>
            ${textbox('First', 'BLOCK_ONE', 0)}
            ${depth === 2 ? `<Rectangle Name="Wrapper"><ReportItems>${nested}</ReportItems><Top>0.4in</Top><Left>0in</Left><Height>0.4in</Height><Width>5in</Width><Style/></Rectangle>` : nested}
            ${textbox('Third', 'BLOCK_THREE', 0.9)}
          </ReportItems>
          <Top>0in</Top><Left>0in</Left><Height>1.3in</Height><Width>5in</Width>${outerStyle}
        </Rectangle>
      </ReportItems>
      <Height>2in</Height><Style/>
    </Body>
    <Width>7in</Width>
    <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`, 'utf8');
}

const END_BREAK = '<PageBreak><BreakLocation>End</BreakLocation></PageBreak>';
const START_BREAK = '<PageBreak><BreakLocation>Start</BreakLocation></PageBreak>';

async function pdfPages(context, model, name) {
  await fs.mkdir(tmpRoot, { recursive: true }).catch(() => {});
  const pdfPath = path.join(tmpRoot, `${name}-${process.pid}.pdf`);
  context.after(() => fs.rm(pdfPath, { force: true }));
  const rendered = await renderPdf(model, request, config);
  await fs.writeFile(pdfPath, rendered.buffer);
  const pages = [];
  for (let page = 1; page <= rendered.pageCount; page += 1) {
    const { stdout } = await execFileAsync('pdftotext', ['-f', String(page), '-l', String(page), '-layout', pdfPath, '-']);
    pages.push(stdout.replace(/\s+/g, ' ').trim());
  }
  return { rendered, pages, pdfPath };
}

test('PDF spends a PageBreak declared on a nested rectangle on the following sibling', async (context) => {
  const { rendered, pages } = await pdfPages(context, parseRdl(report({ breakXml: END_BREAK })), 'nested-end');
  assert.equal(rendered.pageCount, 2);
  assert.match(pages[0], /BLOCK_ONE/);
  assert.match(pages[0], /BLOCK_TWO/);
  assert.doesNotMatch(pages[0], /BLOCK_THREE/);
  assert.match(pages[1], /BLOCK_THREE/);
});

test('PDF applies a nested BreakLocation Start before the declaring item, not after it', async (context) => {
  const { rendered, pages } = await pdfPages(context, parseRdl(report({ breakXml: START_BREAK })), 'nested-start');
  assert.equal(rendered.pageCount, 2);
  assert.match(pages[0], /BLOCK_ONE/);
  assert.doesNotMatch(pages[0], /BLOCK_TWO/);
  assert.match(pages[1], /BLOCK_TWO/);
  // Start breaks before, so the following sibling stays with it rather than moving again.
  assert.match(pages[1], /BLOCK_THREE/);
});

test('PDF carries a nested break out through every enclosing container level', async (context) => {
  const { rendered, pages } = await pdfPages(context, parseRdl(report({ breakXml: END_BREAK, depth: 2 })), 'nested-deep');
  assert.equal(rendered.pageCount, 2);
  assert.match(pages[0], /BLOCK_TWO/);
  assert.doesNotMatch(pages[0], /BLOCK_THREE/);
  assert.match(pages[1], /BLOCK_THREE/);
});

// Counterexamples. Without these the assertions above would also pass for a renderer that simply broke the
// page at every rectangle boundary.
test('PDF keeps identical content on one page when no nested break is declared', async (context) => {
  const { rendered, pages } = await pdfPages(context, parseRdl(report()), 'nested-none');
  assert.equal(rendered.pageCount, 1);
  assert.match(pages[0], /BLOCK_ONE.*BLOCK_TWO.*BLOCK_THREE/);
});

test('PDF ignores a nested break whose Disabled expression resolves true', async (context) => {
  const disabled = '<PageBreak><BreakLocation>End</BreakLocation><Disabled>=1=1</Disabled></PageBreak>';
  const { rendered } = await pdfPages(context, parseRdl(report({ breakXml: disabled })), 'nested-disabled');
  assert.equal(rendered.pageCount, 1);
});

test('PDF honours a nested break whose Disabled expression resolves false', async (context) => {
  const enabled = '<PageBreak><BreakLocation>End</BreakLocation><Disabled>=1=2</Disabled></PageBreak>';
  const { rendered } = await pdfPages(context, parseRdl(report({ breakXml: enabled })), 'nested-enabled');
  assert.equal(rendered.pageCount, 2);
});

test('PDF fails closed when a nested break would fragment a rectangle that paints its own extent', async () => {
  const outerStyle = '<Style><BackgroundColor>#eeeeee</BackgroundColor></Style>';
  await assert.rejects(
    () => renderPdf(parseRdl(report({ breakXml: END_BREAK, outerStyle })), request, config),
    (error) => error.code === 'UNSUPPORTED_FEATURE',
  );
});

test('editable DOCX inherits the nested break as a separate Word page section with native text', async () => {
  const rendered = await renderEditableDocx(parseRdl(report({ breakXml: END_BREAK })), request, config);
  assert.equal(rendered.pageCount, 2);
  const zip = await JSZip.loadAsync(rendered.buffer);
  const document = await zip.file('word/document.xml').async('string');
  // One next-page section per canonical PDF page.
  assert.equal((document.match(/<w:sectPr/g) || []).length, 2);
  const texts = (document.match(/<w:t[^>]*>[^<]*<\/w:t>/g) || []).map((run) => run.replace(/<[^>]+>/g, ''));
  assert.deepEqual(texts.filter((value) => /^BLOCK_/.test(value)), ['BLOCK_ONE', 'BLOCK_TWO', 'BLOCK_THREE']);
  assert.equal(document.includes('<w:pict'), false);
});

test('visual DOCX rasterizes one page image per canonical page created by the nested break', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-nested-visual-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const model = parseRdl(report({ breakXml: END_BREAK }));
  const rendered = await renderVisualDocx(model, request, config, tempDir);
  assert.equal(rendered.pageCount, 2);
  const zip = await JSZip.loadAsync(rendered.buffer);
  assert.equal(Object.keys(zip.files).filter((name) => /^word\/media\/.+\.png$/.test(name)).length, 2);
});

async function excelSheets(model) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-nested-xlsx-'));
  try {
    const rendered = await renderExcel(model, { ...request, excelLayoutMode: 'REPORT' }, config, tempDir);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(rendered.buffer);
    return workbook.worksheets.map((sheet) => {
      const values = [];
      sheet.eachRow((row) => row.eachCell((cell) => {
        const value = String(cell.value ?? '').trim();
        if (value) values.push(value);
      }));
      return values;
    });
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

test('XLSX partitions worksheets at a nested break and not at an ordinary container boundary', async () => {
  assert.deepEqual(
    await excelSheets(parseRdl(report({ breakXml: END_BREAK }))),
    [['BLOCK_ONE', 'BLOCK_TWO'], ['BLOCK_THREE']],
  );
  assert.deepEqual(
    await excelSheets(parseRdl(report())),
    [['BLOCK_ONE', 'BLOCK_TWO', 'BLOCK_THREE']],
  );
});

test('XLSX fails closed on the same unfragmentable painted rectangle the PDF refuses', async () => {
  const outerStyle = '<Style><BackgroundColor>#eeeeee</BackgroundColor></Style>';
  await assert.rejects(
    () => excelSheets(parseRdl(report({ breakXml: END_BREAK, outerStyle }))),
    (error) => error.code === 'UNSUPPORTED_FEATURE',
  );
});
