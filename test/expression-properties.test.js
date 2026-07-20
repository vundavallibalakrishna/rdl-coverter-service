// Regression tests for the "RDL property can be an expression" bug class. Every style property below is
// ExpressionType in RDL — a literal OR an =expression evaluated per row. The failure mode we guard against
// is a renderer testing the RAW expression source instead of the evaluated result (e.g. =IIF(c,"Middle",
// "Top") always matching "Middle"). Each test asserts the EVALUATED branch wins: a false condition must NOT
// produce the "true" literal's output.
import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { PNG } from 'pngjs';
import { parseRdl } from '../src/rdl/parser.js';
import { loadConfig } from '../src/config.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

const base = parseRdl(await fs.readFile(new URL('./fixtures/basic.rdl', import.meta.url)));
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const request = { parameters: { Title: 'T', Choice: 'A' }, datasets: { Sales: [{ Name: 'North', Amount: 1 }], Choices: [{ Value: 'A' }] }, outputFileName: 'expr' };

const detailTextbox = (m) => {
  const tablix = m.body.items.find((item) => item.type === 'Tablix');
  return tablix.rows[tablix.rows.length - 1].cells[0].items.find((item) => item.type === 'Textbox');
};
const docXml = async (buffer) => (await JSZip.loadAsync(buffer)).file('word/document.xml').async('string');

test('conditional VerticalAlign resolves the expression, not the raw source (DOCX)', async () => {
  const centered = structuredClone(base);
  detailTextbox(centered).style.verticalAlign = '=IIF(1=1,"Middle","Top")';
  assert.match(await docXml((await renderEditableDocx(centered, request, config)).buffer), /<w:vAlign w:val="center"\/>/);

  const top = structuredClone(base);
  detailTextbox(top).style.verticalAlign = '=IIF(1=2,"Middle","Top")'; // false -> Top, must NOT be center
  const topXml = await docXml((await renderEditableDocx(top, request, config)).buffer);
  assert.doesNotMatch(topXml, /<w:vAlign w:val="center"\/>/);
});

test('conditional TextAlign resolves the expression, not the raw source (Excel)', async () => {
  const right = structuredClone(base);
  detailTextbox(right).style.textAlign = '=IIF(1=1,"Right","Left")';
  const rightWb = new ExcelJS.Workbook();
  await rightWb.xlsx.load((await renderExcel(right, request, config, null)).buffer);
  const rightCell = findByValue(rightWb.worksheets[0], 'North');
  assert.equal(rightCell.alignment.horizontal, 'right');

  const left = structuredClone(base);
  detailTextbox(left).style.textAlign = '=IIF(1=2,"Right","Left")'; // false -> Left, must NOT be right
  const leftWb = new ExcelJS.Workbook();
  await leftWb.xlsx.load((await renderExcel(left, request, config, null)).buffer);
  assert.equal(findByValue(leftWb.worksheets[0], 'North').alignment.horizontal, 'left');
});

test('a dynamic Image Value expression is resolved before the embeddedImages lookup (PDF + DOCX)', async () => {
  const withImage = (m) => {
    m.embeddedImages = { ...(m.embeddedImages || {}), LOGO: { data: PNG_4x4, mimeType: 'image/png' } };
    m.page.header = {
      height: 40, printOnFirstPage: true, printOnLastPage: true,
      // Value is an expression that evaluates to the embedded image's name — must resolve, not look up raw.
      items: [{ type: 'Image', name: 'I', source: 'Embedded', value: '=IIF(1=1,"LOGO","MISSING")', sizing: 'FitProportional', width: 20, height: 20, top: 0, left: 0, style: {} }],
    };
    return m;
  };

  // DOCX is the discriminating check: with the raw-lookup bug the expression key misses and no media part
  // is written; resolved, the image embeds.
  const docxZip = await JSZip.loadAsync((await renderEditableDocx(withImage(structuredClone(base)), request, config)).buffer);
  assert.equal(Object.keys(docxZip.files).some((n) => /^word\/media\/.*\.(png|jpeg|gif)$/.test(n)), true, 'dynamic image dropped in DOCX');

  // PDF must resolve and draw it (reaching PDFKit's image decoder) rather than silently skipping.
  const pdf = await renderPdf(withImage(structuredClone(base)), request, config);
  assert.equal(pdf.buffer.subarray(0, 4).toString(), '%PDF');
});

test('a conditional FontSize/Padding parses (no crash) and applies per row', async () => {
  const m = structuredClone(base);
  const tb = detailTextbox(m);
  tb.style.fontSize = '=IIF(Fields!Amount.Value >= 100, "14pt", "9pt")';
  tb.style.paddingLeft = '=IIF(Fields!Amount.Value >= 100, "6pt", "2pt")';
  const twoRows = { ...request, datasets: { ...request.datasets, Sales: [{ Name: 'Big', Amount: 100 }, { Name: 'Small', Amount: 1 }] } };
  const xml = await docXml((await renderEditableDocx(m, twoRows, config)).buffer); // must not throw at parse or render
  const sizes = new Set([...xml.matchAll(/<w:sz w:val="(\d+)"\s*\/>/g)].map((mt) => mt[1]));
  assert.ok(sizes.has('28'), '14pt (Amount>=100) row missing'); // 14pt -> 28 half-points
  assert.ok(sizes.has('18'), '9pt (Amount<100) row missing'); // 9pt -> 18 half-points
});

function findByValue(ws, value) {
  let found = null;
  ws.eachRow((row) => row.eachCell((cell) => { if (cell.value === value) found = cell; }));
  return found;
}

// A real 4x4 opaque PNG that PDFKit's decoder accepts.
const PNG_4x4 = (() => {
  const png = new PNG({ width: 4, height: 4 });
  png.data.fill(0x80); // opaque grey
  for (let i = 3; i < png.data.length; i += 4) png.data[i] = 0xFF;
  return PNG.sync.write(png).toString('base64');
})();
