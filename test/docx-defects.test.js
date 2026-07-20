// Regression tests for editable/fixed DOCX defects found in the end-to-end Word-engine review.
// Each of these was a latent crash or silent-drop that no prior test covered: a construct just outside
// the happy path of the sampled client reports. They assert on the generated OpenXML (and that the render
// does not throw), which is what each fix changed.
import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { parseRdl } from '../src/rdl/parser.js';
import { loadConfig } from '../src/config.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderFixedEditableDocx } from '../src/render/fixedDocx.js';

const model = parseRdl(await fs.readFile(new URL('./fixtures/basic.rdl', import.meta.url)));
const request = { parameters: { Title: 'T', Choice: 'A' }, datasets: { Sales: [{ Name: 'North', Amount: 1 }], Choices: [{ Value: 'A' }] } };
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

const documentXml = async (buffer) => (await JSZip.loadAsync(buffer)).file('word/document.xml').async('string');
const partXml = async (buffer, part) => (await JSZip.loadAsync(buffer)).file(part).async('string');
// The textbox inside the last (detail) row's first cell — a real table cell that carries cell borders.
function detailCellTextbox(m) {
  const tablix = m.body.items.find((item) => item.type === 'Tablix');
  return tablix.rows[tablix.rows.length - 1].cells[0].items.find((item) => item.type === 'Textbox');
}
const NONE = { style: 'None', color: '#000000', width: 1 };

test('a conditional FontSize expression renders (does not throw NaN) at the evaluated size', async () => {
  const m = structuredClone(model);
  detailCellTextbox(m).style.fontSize = '=IIF(1=1,14,10)';
  const result = await renderEditableDocx(m, request, config); // previously threw "Invalid value 'NaN'"
  const xml = await documentXml(result.buffer);
  assert.doesNotMatch(xml, /w:sz w:val="NaN"/);
  // 14pt -> 28 half-points on the run.
  assert.match(xml, /<w:sz w:val="28"\s*\/>/);
});

test('a Solid border with no explicit color renders as a black single line (SSRS default), not dropped', async () => {
  const m = structuredClone(model);
  detailCellTextbox(m).style.borders = { top: { style: 'Solid', color: undefined, width: 1 }, right: { ...NONE }, bottom: { ...NONE }, left: { ...NONE } };
  const xml = await documentXml((await renderEditableDocx(m, request, config)).buffer);
  assert.match(xml, /<w:top w:val="single" w:color="000000"/);
});

test('a conditional BorderWidth expression yields a numeric size, never NaN', async () => {
  const m = structuredClone(model);
  detailCellTextbox(m).style.borders = { top: { style: 'Solid', color: '#000000', width: '=IIF(1=1,3,1)' }, right: { ...NONE }, bottom: { ...NONE }, left: { ...NONE } };
  const xml = await documentXml((await renderEditableDocx(m, request, config)).buffer);
  assert.doesNotMatch(xml, /w:sz="NaN"/);
  // 3pt -> sz 24 eighths of a point.
  assert.match(xml, /<w:top w:val="single"[^>]*w:sz="24"/);
});

test('a Double border style is preserved as a double rule, not flattened to single', async () => {
  const m = structuredClone(model);
  detailCellTextbox(m).style.borders = { top: { style: 'Double', color: '#123456', width: 1 }, right: { ...NONE }, bottom: { ...NONE }, left: { ...NONE } };
  const xml = await documentXml((await renderEditableDocx(m, request, config)).buffer);
  assert.match(xml, /<w:top w:val="double" w:color="123456"/);
});

test('a Line item with no border style renders without throwing', async () => {
  const m = structuredClone(model);
  m.body.items = [{ type: 'Line', name: 'L', top: 0, left: 0, width: 100, height: 0, style: {} }];
  await assert.doesNotReject(() => renderEditableDocx(m, request, config)); // previously TypeError on .border.color
});

test('PageNumber/TotalPages inside a tablix cell become live Word fields, not a frozen "1"', async () => {
  const m = structuredClone(model);
  const tb = detailCellTextbox(m);
  tb.value = '="Page " & Globals!PageNumber & " of " & Globals!TotalPages';
  tb.paragraphs = [[{ value: tb.value, markupType: 'None' }]];
  const xml = await documentXml((await renderEditableDocx(m, request, config)).buffer);
  assert.match(xml, /w:instr="PAGE"/);
  assert.match(xml, /w:instr="NUMPAGES"/);
  assert.doesNotMatch(xml, /Page 1 of 1/);
});

test('an embedded image is typed from its real bytes, not a wrong declared MIMEType', async () => {
  // A GIF mislabelled as image/png must still be emitted as a .gif part, or Word shows a broken image.
  const gif = Buffer.concat([Buffer.from('GIF89a', 'ascii'), Buffer.from([0x0A, 0x00, 0x05, 0x00]), Buffer.alloc(40)]);
  const m = structuredClone(model);
  m.embeddedImages = { ...(m.embeddedImages || {}), LOGO: { data: gif.toString('base64'), mimeType: 'image/png' } };
  m.page.header = { height: 40, printOnFirstPage: true, printOnLastPage: true, items: [{ type: 'Image', name: 'I', value: 'LOGO', source: 'Embedded', sizing: 'FitProportional', width: 20, height: 20, top: 0, left: 0, style: {} }] };
  const zip = await JSZip.loadAsync((await renderEditableDocx(m, request, config)).buffer);
  assert.equal(Object.keys(zip.files).some((name) => /^word\/media\/.*\.gif$/.test(name)), true);
});

test('a visible chart that cannot be rendered fails closed instead of silently vanishing', async () => {
  const m = structuredClone(model);
  // A minimal visible Chart item is enough to reach the chart branch; the guard fires before materialization.
  m.body.items = [{ type: 'Chart', name: 'DroppedChart', top: 0, left: 0, width: 100, height: 100, hidden: 'false', chartType: 'column', style: {} }];
  // No config/tempDir => the chart image cannot be produced. It must throw, not produce a chartless document.
  await assert.rejects(
    () => renderEditableDocx(m, request),
    (error) => error.code === 'RENDER_FAILED' && /chart/i.test(error.message),
  );
});

test('fixed-editable renders a genuinely empty page instead of failing closed', async () => {
  const m = structuredClone(model);
  const textbox = m.body.items.find((item) => item.type === 'Textbox');
  m.body.items = [{ ...structuredClone(textbox), value: '', paragraphs: [['']], top: 0, left: 0 }];
  const result = await renderFixedEditableDocx(m, request, config); // previously threw "contains no positioned objects"
  assert.equal(result.buffer.subarray(0, 2).toString(), 'PK');
  assert.equal(result.pageCount, 1);
});
