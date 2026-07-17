import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { promisify } from 'node:util';
import { PDFDocument } from 'pdf-lib';
import JSZip from 'jszip';
import { PNG } from 'pngjs';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderVisualDocx } from '../src/render/visualDocx.js';
import { cellGridWidth, computeDocxTableGeometry } from '../src/render/docxTableLayout.js';

const fixture = await fs.readFile(new URL('./fixtures/basic.rdl', import.meta.url));
const model = parseRdl(fixture);
const request = {
  outputFileName: 'sales', parameters: { Title: 'Sales', Choice: 'A' },
  datasets: { Sales: [{ Name: 'North', Amount: 1234.5 }, { Name: 'South', Amount: 99 }] },
};
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '30000' });
const execFileAsync = promisify(execFile);

test('renders a valid selectable PDF with expected text, dimensions, checksum, and raster geometry', async (context) => {
  const result = await renderPdf(model, request, config);
  assert.equal(result.buffer.subarray(0, 4).toString(), '%PDF');
  const pdf = await PDFDocument.load(result.buffer);
  assert.equal(pdf.getPageCount(), result.pageCount);
  assert.equal(Math.round(pdf.getPage(0).getWidth()), 612);
  assert.equal(Math.round(pdf.getPage(0).getHeight()), 792);
  assert.match(createHash('sha256').update(result.buffer).digest('hex'), /^[a-f0-9]{64}$/);
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-pdf-verify-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, 'report.pdf');
  const pngPrefix = path.join(tempDir, 'page');
  await fs.writeFile(pdfPath, result.buffer);
  const extracted = await execFileAsync('pdftotext', [pdfPath, '-']);
  assert.match(extracted.stdout, /North/);
  assert.match(extracted.stdout, /1,234\.50/);
  await execFileAsync(config.pdftoppmPath, ['-f', '1', '-singlefile', '-png', '-r', '144', pdfPath, pngPrefix]);
  const raster = PNG.sync.read(await fs.readFile(`${pngPrefix}.png`));
  assert.equal(raster.width, 1224);
  assert.equal(raster.height, 1584);
});

test('honors explicit start page breaks without adding blank trailing pages', async () => {
  const textbox = model.body.items.find((item) => item.type === 'Textbox');
  const second = structuredClone(textbox);
  second.name = 'SecondPage';
  second.value = 'Second page';
  second.paragraphs = [['Second page']];
  second.pageBreak = { location: 'Start', disabled: 'false' };
  const pagedModel = structuredClone(model);
  pagedModel.body.items = [textbox, second];
  const result = await renderPdf(pagedModel, request, config);
  assert.equal(result.pageCount, 2);
});

test('does not render children or borders of hidden rectangles', async (context) => {
  const textbox = structuredClone(model.body.items.find((item) => item.type === 'Textbox'));
  textbox.name = 'HiddenRectangleChild';
  textbox.value = 'HIDDEN_RECTANGLE_CONTENT';
  textbox.paragraphs = [['HIDDEN_RECTANGLE_CONTENT']];
  textbox.top = 0;
  textbox.left = 0;
  const hiddenModel = structuredClone(model);
  hiddenModel.body.items = [{
    type: 'Rectangle',
    name: 'HiddenRectangle',
    top: 0,
    left: 0,
    width: 200,
    height: 100,
    hidden: 'true',
    style: textbox.style,
    items: [textbox],
  }];
  const result = await renderPdf(hiddenModel, request, config);
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-hidden-rectangle-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, 'hidden.pdf');
  await fs.writeFile(pdfPath, result.buffer);
  const extracted = await execFileAsync('pdftotext', [pdfPath, '-']);
  assert.doesNotMatch(extracted.stdout, /HIDDEN_RECTANGLE_CONTENT/);
});

test('continues a PDF table cell taller than a page without clipping or duplicating text', async (context) => {
  const longCell = `PDF_GIANT_START\n${Array.from({ length: 180 }, (_, index) => `PDF_GIANT_LINE_${String(index + 1).padStart(3, '0')} deterministic overflow content`).join('\n')}\nPDF_GIANT_END`;
  const overflowRequest = {
    ...request,
    datasets: {
      Sales: [
        { Name: longCell, Amount: 1 },
        ...Array.from({ length: 30 }, (_, index) => ({ Name: `PDF_FOLLOWING_ROW_${String(index + 1).padStart(3, '0')}`, Amount: index + 2 })),
      ],
    },
  };
  const result = await renderPdf(model, overflowRequest, config);
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-pdf-overflow-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, 'overflow.pdf');
  await fs.writeFile(pdfPath, result.buffer);
  const extracted = await execFileAsync('pdftotext', ['-layout', pdfPath, '-']);
  assert.equal((extracted.stdout.match(/PDF_GIANT_START/g) || []).length, 1);
  assert.equal((extracted.stdout.match(/PDF_GIANT_END/g) || []).length, 1);
  for (let index = 1; index <= 180; index += 1) {
    const marker = `PDF_GIANT_LINE_${String(index).padStart(3, '0')}`;
    assert.equal((extracted.stdout.match(new RegExp(marker, 'g')) || []).length, 1);
  }
  for (let index = 1; index <= 30; index += 1) {
    const marker = `PDF_FOLLOWING_ROW_${String(index).padStart(3, '0')}`;
    assert.equal((extracted.stdout.match(new RegExp(marker, 'g')) || []).length, 1);
  }
});

test('clips overflowing row-span cells at the body boundary so they never enter the footer', async (context) => {
  const protectedModel = structuredClone(model);
  protectedModel.page.footer = { height: 60, printOnFirstPage: true, printOnLastPage: true, items: [] };
  const tablix = protectedModel.body.items.find((item) => item.type === 'Tablix');
  const staticHeader = structuredClone(tablix.rows[0].cells[0]);
  staticHeader.items[0].value = 'Group';
  staticHeader.items[0].paragraphs = [['Group']];
  const dynamicHeader = structuredClone(tablix.rows[1].cells[0]);
  dynamicHeader.items[0].value = '=Fields!Group.Value';
  dynamicHeader.items[0].paragraphs = [['=Fields!Group.Value']];
  const staticMember = { group: null, repeatOnNewPage: true, keepTogether: true, keepWithGroup: 'After', fixedData: false, hidden: 'false', sortExpressions: [], header: { size: 40, cell: staticHeader }, children: [] };
  const dynamicMember = { group: { name: 'FooterGroup', expressions: ['=Fields!Group.Value'], pageBreak: 'None' }, repeatOnNewPage: false, keepTogether: false, keepWithGroup: 'None', fixedData: false, hidden: 'false', sortExpressions: [], header: { size: 40, cell: dynamicHeader }, children: [] };
  tablix.rowHeaderColumns = [40];
  tablix.columns = [40, ...tablix.bodyColumns];
  tablix.rowMembers = [staticMember, dynamicMember];
  tablix.rowMemberPaths = [[staticMember], [dynamicMember]];

  const tallGroup = Array.from({ length: 190 }, (_, index) => `FOOTER_PROTECTED_LINE_${String(index + 1).padStart(3, '0')}`).join('\n');
  const protectedRequest = {
    ...request,
    datasets: { Sales: Array.from({ length: 3 }, (_, index) => ({ Name: `Row ${index + 1}`, Amount: index + 1, Group: tallGroup })) },
  };
  const result = await renderPdf(protectedModel, protectedRequest, config);
  assert.equal(result.pageCount > 1, true);
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-footer-protection-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, 'footer-protection.pdf');
  const pngPrefix = path.join(tempDir, 'page');
  await fs.writeFile(pdfPath, result.buffer);
  await execFileAsync(config.pdftoppmPath, ['-f', '1', '-singlefile', '-png', '-r', '144', pdfPath, pngPrefix]);
  const raster = PNG.sync.read(await fs.readFile(`${pngPrefix}.png`));
  const scale = 2;
  const tableLeft = Math.round((protectedModel.page.marginLeft + tablix.left) * scale);
  const protectedTop = Math.ceil((protectedModel.page.height - protectedModel.page.marginBottom - protectedModel.page.footer.height + 2) * scale);
  const protectedBottom = Math.floor((protectedModel.page.height - protectedModel.page.marginBottom - 2) * scale);
  let darkPixels = 0;
  for (let y = protectedTop; y < protectedBottom; y += 1) {
    for (let x = tableLeft - 1; x <= tableLeft + 1; x += 1) {
      const offset = (y * raster.width + x) * 4;
      if (raster.data[offset] < 96 && raster.data[offset + 1] < 96 && raster.data[offset + 2] < 96) darkPixels += 1;
    }
  }
  assert.equal(darkPixels, 0);
});

test('draws the tablix-level bottom border after the final row when detail cells omit it', async (context) => {
  const borderedModel = structuredClone(model);
  borderedModel.page.header = null;
  borderedModel.page.footer = null;
  const tablix = borderedModel.body.items.find((item) => item.type === 'Tablix');
  borderedModel.body.items = [tablix];
  tablix.top = 0;
  tablix.left = 0;
  tablix.style.borders.bottom = { style: 'Solid', color: '#000000', width: 1 };
  for (const row of tablix.rows) {
    for (const cell of row.cells) {
      const textbox = cell.items.find((item) => item.type === 'Textbox');
      if (textbox) textbox.style.borders.bottom = { style: 'None', color: '#000000', width: 1 };
    }
  }

  const borderedRequest = { ...request, datasets: { Sales: [{ Name: 'Only row', Amount: 1 }] } };
  const result = await renderPdf(borderedModel, borderedRequest, config);
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-tablix-outer-border-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, 'outer-border.pdf');
  const pngPrefix = path.join(tempDir, 'page');
  await fs.writeFile(pdfPath, result.buffer);
  await execFileAsync(config.pdftoppmPath, ['-f', '1', '-singlefile', '-png', '-r', '144', pdfPath, pngPrefix]);
  const raster = PNG.sync.read(await fs.readFile(`${pngPrefix}.png`));
  const scale = 2;
  const expectedY = Math.round((borderedModel.page.marginTop + tablix.rows[0].height + tablix.rows[1].height) * scale);
  const startX = Math.round(borderedModel.page.marginLeft * scale);
  const endX = Math.round((borderedModel.page.marginLeft + tablix.width) * scale);
  let strongestHorizontalLine = 0;
  for (let y = expectedY - 2; y <= expectedY + 2; y += 1) {
    let darkPixels = 0;
    for (let x = startX; x <= endX; x += 1) {
      const offset = (y * raster.width + x) * 4;
      if (raster.data[offset] < 96 && raster.data[offset + 1] < 96 && raster.data[offset + 2] < 96) darkPixels += 1;
    }
    strongestHorizontalLine = Math.max(strongestHorizontalLine, darkPixels);
  }
  assert.equal(strongestHorizontalLine > (endX - startX) * 0.9, true);
});

test('closes the tablix-level bottom border on every physical page fragment', async (context) => {
  const pagedModel = structuredClone(model);
  pagedModel.page.height = 200;
  pagedModel.page.marginTop = 20;
  pagedModel.page.marginBottom = 20;
  pagedModel.page.header = null;
  pagedModel.page.footer = null;
  const tablix = pagedModel.body.items.find((item) => item.type === 'Tablix');
  pagedModel.body.items = [tablix];
  tablix.top = 0;
  tablix.left = 0;
  tablix.rows[1].height = 70;
  tablix.style.borders.bottom = { style: 'Solid', color: '#000000', width: 1 };
  for (const row of tablix.rows) {
    for (const cell of row.cells) {
      const textbox = cell.items.find((item) => item.type === 'Textbox');
      if (textbox) textbox.style.borders.bottom = { style: 'None', color: '#000000', width: 1 };
    }
  }

  const pagedRequest = {
    ...request,
    datasets: { Sales: [{ Name: 'First row', Amount: 1 }, { Name: 'Second row', Amount: 2 }] },
  };
  const result = await renderPdf(pagedModel, pagedRequest, config);
  assert.equal(result.pageCount, 2);
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-tablix-fragment-border-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, 'fragment-border.pdf');
  const pngPrefix = path.join(tempDir, 'page');
  await fs.writeFile(pdfPath, result.buffer);
  await execFileAsync(config.pdftoppmPath, ['-f', '1', '-singlefile', '-png', '-r', '144', pdfPath, pngPrefix]);
  const raster = PNG.sync.read(await fs.readFile(`${pngPrefix}.png`));
  const scale = 2;
  const expectedY = Math.round((pagedModel.page.marginTop + tablix.rows[0].height + tablix.rows[1].height) * scale);
  const startX = Math.round(pagedModel.page.marginLeft * scale);
  const endX = Math.round((pagedModel.page.marginLeft + tablix.width) * scale);
  let strongestHorizontalLine = 0;
  for (let y = expectedY - 2; y <= expectedY + 2; y += 1) {
    let darkPixels = 0;
    for (let x = startX; x <= endX; x += 1) {
      const offset = (y * raster.width + x) * 4;
      if (raster.data[offset] < 96 && raster.data[offset + 1] < 96 && raster.data[offset + 2] < 96) darkPixels += 1;
    }
    strongestHorizontalLine = Math.max(strongestHorizontalLine, darkPixels);
  }
  assert.equal(strongestHorizontalLine > (endX - startX) * 0.9, true);
});

test('closes a fragment boundary using cell shared-edge borders when the tablix has no outer border', async (context) => {
  // SSRS shared-edge model: cells declare only a top border (a group/row separator) and rely on the
  // NEXT cell to supply the shared line. With no tablix outer border, the row that ends a page must
  // still get a bottom line resolved from the continuing cell's top border.
  const pagedModel = structuredClone(model);
  pagedModel.page.height = 200;
  pagedModel.page.marginTop = 20;
  pagedModel.page.marginBottom = 20;
  pagedModel.page.header = null;
  pagedModel.page.footer = null;
  const tablix = pagedModel.body.items.find((item) => item.type === 'Tablix');
  pagedModel.body.items = [tablix];
  tablix.top = 0;
  tablix.left = 0;
  tablix.rows[1].height = 70;
  const none = { style: 'None', color: '#000000', width: 1 };
  tablix.style.borders = { top: { ...none }, right: { ...none }, bottom: { ...none }, left: { ...none } };
  for (const row of tablix.rows) {
    for (const cell of row.cells) {
      const textbox = cell.items.find((item) => item.type === 'Textbox');
      if (textbox) textbox.style.borders = { top: { style: 'Solid', color: '#000000', width: 1 }, right: { ...none }, bottom: { ...none }, left: { ...none } };
    }
  }

  const pagedRequest = { ...request, datasets: { Sales: [{ Name: 'First row', Amount: 1 }, { Name: 'Second row', Amount: 2 }] } };
  const result = await renderPdf(pagedModel, pagedRequest, config);
  assert.equal(result.pageCount, 2);
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-shared-edge-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, 'shared-edge.pdf');
  const pngPrefix = path.join(tempDir, 'page');
  await fs.writeFile(pdfPath, result.buffer);
  await execFileAsync(config.pdftoppmPath, ['-f', '1', '-singlefile', '-png', '-r', '144', pdfPath, pngPrefix]);
  const raster = PNG.sync.read(await fs.readFile(`${pngPrefix}.png`));
  const scale = 2;
  const expectedY = Math.round((pagedModel.page.marginTop + tablix.rows[0].height + tablix.rows[1].height) * scale);
  const startX = Math.round(pagedModel.page.marginLeft * scale);
  const endX = Math.round((pagedModel.page.marginLeft + tablix.width) * scale);
  let strongestHorizontalLine = 0;
  for (let y = expectedY - 2; y <= expectedY + 2; y += 1) {
    let darkPixels = 0;
    for (let x = startX; x <= endX; x += 1) {
      const offset = (y * raster.width + x) * 4;
      if (raster.data[offset] < 96 && raster.data[offset + 1] < 96 && raster.data[offset + 2] < 96) darkPixels += 1;
    }
    strongestHorizontalLine = Math.max(strongestHorizontalLine, darkPixels);
  }
  assert.equal(strongestHorizontalLine > (endX - startX) * 0.9, true);
});

test('renders editable DOCX with native OpenXML tables and text', async () => {
  const result = await renderEditableDocx(model, request);
  assert.equal(result.buffer.subarray(0, 2).toString(), 'PK');
  const zip = await JSZip.loadAsync(result.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  assert.match(documentXml, /<w:tbl>/);
  assert.match(documentXml, /North/);
});

test('preserves textbox background and foreground colors in editable DOCX page headers', async () => {
  const headerModel = structuredClone(model);
  const title = structuredClone(model.body.items.find((item) => item.type === 'Textbox'));
  title.name = 'ColoredHeaderTitle';
  title.value = 'VISIBLE_HEADER_TITLE';
  title.paragraphs = [['VISIBLE_HEADER_TITLE']];
  title.style.backgroundColor = 'Navy';
  title.style.color = 'White';
  headerModel.page.header = {
    height: 40,
    printOnFirstPage: true,
    printOnLastPage: true,
    items: [title],
  };

  const result = await renderEditableDocx(headerModel, request);
  const zip = await JSZip.loadAsync(result.buffer);
  const headerXml = await zip.file('word/header1.xml').async('string');
  assert.match(headerXml, /VISIBLE_HEADER_TITLE/);
  assert.match(headerXml, /<w:shd[^>]*w:fill="000080"/);
  assert.match(headerXml, /<w:color w:val="ffffff"/i);
});

test('preserves tablix-level outer borders in editable DOCX', async () => {
  const borderedModel = structuredClone(model);
  const tablix = borderedModel.body.items.find((item) => item.type === 'Tablix');
  tablix.style.borders.bottom = { style: 'Solid', color: '#112233', width: 1 };
  const result = await renderEditableDocx(borderedModel, request);
  const zip = await JSZip.loadAsync(result.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  assert.match(documentXml, /<w:tblBorders>[\s\S]*?<w:bottom w:val="single" w:color="112233" w:sz="8"[\s\S]*?<\/w:tblBorders>/);
});

test('keeps wide multi-page DOCX tables inside the page, repeats only declared headers, and never duplicates overflowing cells', async () => {
  const stressModel = structuredClone(model);
  const tablix = stressModel.body.items.find((item) => item.type === 'Tablix');
  tablix.width = 900;
  tablix.columns = [260, 640];
  tablix.rowMembers[0].repeatOnNewPage = true;
  const longCell = `LONG_CELL_START\n${Array.from({ length: 260 }, (_, index) => `Overflow line ${String(index + 1).padStart(3, '0')} with wrapped content that must continue exactly once.`).join('\n')}\nLONG_CELL_END`;
  const stressRequest = {
    ...request,
    outputFileName: 'docx-overflow-stress',
    datasets: {
      Sales: [
        { Name: longCell, Amount: 1 },
        ...Array.from({ length: 45 }, (_, index) => ({ Name: `UNIQUE_ROW_${String(index + 1).padStart(3, '0')}`, Amount: index + 2 })),
      ],
    },
  };
  const geometry = computeDocxTableGeometry(stressModel, tablix);
  assert.equal(geometry.tableTwips + geometry.indentTwips <= geometry.availableTwips, true);
  assert.equal(geometry.gridTwips.reduce((sum, width) => sum + width, 0), geometry.tableTwips);
  assert.equal(cellGridWidth(geometry.gridTwips, 0, 2), geometry.tableTwips);
  assert.equal(geometry.scaledToFit, true);

  const result = await renderEditableDocx(stressModel, stressRequest);
  const zip = await JSZip.loadAsync(result.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  assert.equal((documentXml.match(/<w:tblHeader\/>/g) || []).length, 1);
  assert.equal((documentXml.match(/<w:cantSplit\/>/g) || []).length, 1);
  assert.equal((documentXml.match(/LONG_CELL_START/g) || []).length, 1);
  assert.equal((documentXml.match(/LONG_CELL_END/g) || []).length, 1);
  for (let index = 1; index <= 45; index += 1) {
    const marker = `UNIQUE_ROW_${String(index).padStart(3, '0')}`;
    assert.equal((documentXml.match(new RegExp(marker, 'g')) || []).length, 1);
  }
});

test('renders visual DOCX with one page image per PDF page', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-visual-test-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pagedModel = structuredClone(model);
  const secondPage = structuredClone(pagedModel.body.items.find((item) => item.type === 'Textbox'));
  secondPage.name = 'VisualSecondPage';
  secondPage.value = 'VISUAL_SECOND_PAGE';
  secondPage.paragraphs = [['VISUAL_SECOND_PAGE']];
  secondPage.pageBreak = { location: 'Start', disabled: 'false' };
  pagedModel.body.items.push(secondPage);
  const result = await renderVisualDocx(pagedModel, request, config, tempDir);
  assert.equal(result.buffer.subarray(0, 2).toString(), 'PK');
  assert.equal(result.pageCount, 2);
  const zip = await JSZip.loadAsync(result.buffer);
  const mediaFiles = Object.keys(zip.files).filter((name) => /^word\/media\/[^/]+\.png$/.test(name));
  const documentXml = await zip.file('word/document.xml').async('string');
  assert.equal(mediaFiles.length, result.pageCount);
  assert.equal((documentXml.match(/<a:blip r:embed=/g) || []).length, result.pageCount);
  assert.equal((documentXml.match(/<wp:anchor/g) || []).length, result.pageCount);
  assert.equal((documentXml.match(/w:type="page"/g) || []).length, result.pageCount - 1);
});

test('editable DOCX preserves multi-line cell text as breaks and normalizes tabs', async () => {
  const TAB = String.fromCharCode(9);
  const multilineRequest = {
    ...request,
    outputFileName: 'multiline',
    datasets: { Sales: [{ Name: `First line${TAB}indented\nSecond line`, Amount: 5 }] },
  };
  const result = await renderEditableDocx(model, multilineRequest);
  const zip = await JSZip.loadAsync(result.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  // The "\n" between the two lines becomes a Word break rather than collapsing onto one line.
  assert.ok((documentXml.match(/<w:br\/>/g) || []).length >= 1);
  assert.match(documentXml, /First line indented/); // tab expanded to a space
  assert.match(documentXml, /Second line/);
  assert.equal(documentXml.includes(TAB), false); // no raw control character leaks into the XML text
});
