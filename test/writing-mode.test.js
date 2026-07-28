import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { promisify } from 'node:util';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { analyzeRdl, parseRdl } from '../src/rdl/parser.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

const execFileAsync = promisify(execFile);
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const request = { outputFileName: 'writing-mode', parameters: {}, datasets: {}, excel: { layoutMode: 'REPORT' } };

const rdl = (secondMode = 'Vertical') => `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <Body><ReportItems>
    <Textbox Name="BottomToTop"><CanGrow>false</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun>
      <Value>ROTATE270</Value><Style><FontFamily>Arial</FontFamily><FontSize>12pt</FontSize><FontWeight>Bold</FontWeight></Style>
    </TextRun></TextRuns></Paragraph></Paragraphs><Top>0.1in</Top><Left>0.1in</Left><Width>0.5in</Width><Height>1.5in</Height>
      <Style><WritingMode>Rotate270</WritingMode><TextAlign>Center</TextAlign><VerticalAlign>Middle</VerticalAlign><Border><Style>Solid</Style></Border></Style>
    </Textbox>
    <Textbox Name="TopToBottom"><CanGrow>false</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun>
      <Value>VERTICAL</Value><Style><FontFamily>Arial</FontFamily><FontSize>12pt</FontSize></Style>
    </TextRun></TextRuns></Paragraph></Paragraphs><Top>0.1in</Top><Left>0.8in</Left><Width>0.5in</Width><Height>1.5in</Height>
      <Style><WritingMode>${secondMode}</WritingMode><TextAlign>Center</TextAlign><VerticalAlign>Middle</VerticalAlign><Border><Style>Solid</Style></Border></Style>
    </Textbox>
  </ReportItems><Height>1.8in</Height></Body>
  <Width>1.5in</Width><Page><PageWidth>2in</PageWidth><PageHeight>2in</PageHeight>
    <TopMargin>0in</TopMargin><BottomMargin>0in</BottomMargin><LeftMargin>0in</LeftMargin><RightMargin>0in</RightMargin>
  </Page>
</Report>`;

test('WritingMode is normalized and unknown modes fail closed', () => {
  const model = parseRdl(rdl());
  assert.equal(model.body.items[0].style.writingMode, 'Rotate270');
  assert.equal(model.body.items[1].style.writingMode, 'Vertical');
  const analysis = analyzeRdl(rdl('Diagonal'));
  assert.equal(analysis.compatible, false);
  assert.deepEqual(analysis.blockingErrors, [{ code: 'UNSUPPORTED_FEATURE', feature: 'WritingMode:Diagonal' }]);
});

test('PDF writes Rotate270 and Vertical text on a vertical physical axis', async (context) => {
  const result = await renderPdf(parseRdl(rdl()), request, config);
  const pdfPath = path.resolve('tmp', `writing-mode-${process.pid}.pdf`);
  context.after(() => fs.rm(pdfPath, { force: true }));
  await fs.writeFile(pdfPath, result.buffer, { mode: 0o600 });
  const { stdout } = await execFileAsync('pdftotext', ['-bbox-layout', pdfPath, '-']);
  for (const word of ['ROTATE270', 'VERTICAL']) {
    const match = stdout.match(new RegExp(`<word xMin="([^"]+)" yMin="([^"]+)" xMax="([^"]+)" yMax="([^"]+)">${word}</word>`));
    assert.ok(match, `missing ${word} in PDF text layer`);
    const [, xMin, yMin, xMax, yMax] = match.map(Number);
    assert.ok(yMax - yMin > xMax - xMin, `${word} did not render vertically`);
  }
});

test('editable DOCX uses native table-cell text directions', async () => {
  const result = await renderEditableDocx(parseRdl(rdl()), request, config);
  const zip = await JSZip.loadAsync(result.buffer);
  const xml = await zip.file('word/document.xml').async('string');
  assert.match(xml, /<w:textDirection w:val="btLr"\s*\/>/);
  assert.match(xml, /<w:textDirection w:val="tbRl"\s*\/>/);
  assert.match(xml, /ROTATE270/);
  assert.match(xml, /VERTICAL/);
});

test('XLSX REPORT uses native editable cell rotations', async () => {
  const result = await renderExcel(parseRdl(rdl()), request, config, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(result.buffer);
  const rotations = new Map();
  workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => {
    const value = typeof cell.value === 'object' && cell.value?.richText
      ? cell.value.richText.map((run) => run.text).join('')
      : String(cell.value ?? '');
    if (value) rotations.set(value, cell.alignment?.textRotation);
  }));
  assert.equal(rotations.get('ROTATE270'), 90);
  assert.equal(rotations.get('VERTICAL'), -90);
});
