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
  assert.match(xml, /<w:instrText[^>]*>PAGE<\/w:instrText>/);
  assert.match(xml, /<w:instrText[^>]*>NUMPAGES<\/w:instrText>/);
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

test('a free-form body Rectangle preserves its children vertical spacing (not crammed to the top)', async () => {
  // A coordinate-designed cover: a body Rectangle whose children sit at absolute Top offsets. The editable
  // DOCX must translate those gaps into vertical whitespace (exact-height spacer paragraphs), not collapse
  // everything to the top of the page.
  const m = structuredClone(model);
  const tb = (name, top, text) => ({ type: 'Textbox', name, top, left: 0, width: 400, height: 20, hidden: 'false',
    style: { paddingLeft: 2, paddingRight: 2, paddingTop: 2, paddingBottom: 2 }, paragraphs: [[text]] });
  m.body.items = [{
    type: 'Rectangle', name: 'Cover', top: 0, left: 0, width: 400, height: 300, hidden: 'false', style: {},
    items: [tb('T1', 0, 'TOP'), tb('T2', 120, 'MIDDLE')], // T1 bottom = 0+20; T2 top = 120 => 100pt whitespace
  }];
  const zip = await JSZip.loadAsync((await renderEditableDocx(m, request, config)).buffer);
  const xml = await zip.file('word/document.xml').async('string');
  // 100pt whitespace (T2.top - T1.bottom) => a 2000-twip exact-height spacer paragraph between the lines.
  assert.match(xml, /w:line="2000" w:lineRule="exact"/);
  assert.match(xml, /TOP/);
  assert.match(xml, /MIDDLE/);
});

test('a centered free-form image is horizontally centered from its box position (generic, not per-report)', async () => {
  const { PNG } = await import('pngjs');
  const png = new PNG({ width: 4, height: 4 });
  png.data.fill(0x80);
  for (let i = 3; i < png.data.length; i += 4) png.data[i] = 0xFF;
  const data = PNG.sync.write(png).toString('base64');

  const m = structuredClone(model);
  m.embeddedImages = { ...(m.embeddedImages || {}), LOGO: { data, mimeType: 'image/png' } };
  // Container width 400; image at left=150 width=100 => right margin also 150 => centered by position.
  m.body.items = [{
    type: 'Rectangle', name: 'Cover', top: 0, left: 0, width: 400, height: 200, hidden: 'false', style: {},
    items: [{ type: 'Image', name: 'L', source: 'Embedded', value: 'LOGO', sizing: 'FitProportional', top: 0, left: 150, width: 100, height: 60, style: {} }],
  }];
  const xml = await documentXml((await renderEditableDocx(m, request, config)).buffer);
  // The image's paragraph must carry centre justification (derived from Left/Width), and the image embeds.
  assert.match(xml, /<w:jc w:val="center"\/>[\s\S]*?<w:drawing>/);
});

test('a narrow shaded free-form box confines its fill to the RDL Width (single-cell table), not full column', async () => {
  // A small shaded date chip on a cover: box Width 120 inside a 400-wide container. Paragraph shading would
  // fill the whole text column; it must instead become a fixed-width single-cell table of the box Width.
  const m = structuredClone(model);
  m.body.items = [{
    type: 'Rectangle', name: 'Cover', top: 0, left: 0, width: 400, height: 200, hidden: 'false', style: {},
    items: [{ type: 'Textbox', name: 'Chip', top: 20, left: 140, width: 120, height: 24, hidden: 'false',
      style: { backgroundColor: '#D9D9D9', paddingLeft: 2, paddingRight: 2, paddingTop: 2, paddingBottom: 2 },
      paragraphs: [['July 2026']] }],
  }];
  const xml = await documentXml((await renderEditableDocx(m, request, config)).buffer);
  // A table of the box width (120pt -> 2400 twips) carrying the fill on its cell, not a full-width w:p shading.
  assert.match(xml, /<w:tblW w:type="dxa" w:w="2400"\/>/);
  assert.match(xml, /<w:shd[^>]*w:fill="D9D9D9"/);
  // No report-declared border => the box must have an explicit no-border table, not the docx default grid.
  assert.doesNotMatch(xml, /<w:tblBorders>[\s\S]*?w:val="single"[\s\S]*?<\/w:tblBorders>/);
  assert.match(xml, /July 2026/);
});

test('a full-width shaded free-form bar stays a paragraph (no fixed-width table regression)', async () => {
  // A shaded bar as wide as the container must NOT be wrapped — it keeps paragraph shading.
  const m = structuredClone(model);
  m.body.items = [{
    type: 'Rectangle', name: 'Cover', top: 0, left: 0, width: 400, height: 200, hidden: 'false', style: {},
    items: [{ type: 'Textbox', name: 'Bar', top: 0, left: 0, width: 400, height: 24, hidden: 'false',
      style: { backgroundColor: '#003366', paddingLeft: 2, paddingRight: 2, paddingTop: 2, paddingBottom: 2 },
      paragraphs: [['TITLE']] }],
  }];
  const xml = await documentXml((await renderEditableDocx(m, request, config)).buffer);
  assert.doesNotMatch(xml, /<w:tblW w:w="8000" w:type="dxa"/); // 400pt = 8000 twips: no fixed-width table
  assert.match(xml, /<w:shd[^>]*w:fill="003366"/); // shading survives on the paragraph
  assert.match(xml, /TITLE/);
});

test('a run that overrides the textbox style (White/Bold over black/Normal) renders per-run, not the textbox style', async () => {
  // The SSRS pattern behind the Combined Assurance header's last two cells: the textbox style is black /
  // Normal, but the visible run explicitly sets White / Bold. The DOCX previously styled every run from the
  // textbox (black / Normal) and lost it; it must now honour each run's own style, matching the PDF.
  const m = structuredClone(model);
  const tb = detailCellTextbox(m);
  tb.style = { ...tb.style, color: '#000000', fontWeight: 'Normal' };
  const runStyle = { ...tb.style, color: '#FFFFFF', fontWeight: 'Bold' };
  tb.value = 'HEADER';
  tb.paragraphs = [[{ value: 'HEADER', markupType: 'None', style: runStyle }]];
  const xml = await documentXml((await renderEditableDocx(m, request, config)).buffer);
  const i = xml.indexOf('HEADER');
  const rpr = xml.slice(xml.lastIndexOf('<w:r>', i), i); // run properties of the HEADER run
  assert.match(rpr, /<w:b\/>/); // bold from the run, not the textbox's Normal
  assert.match(rpr, /<w:color w:val="FFFFFF"\/>/i); // white from the run, not the textbox's black
});
