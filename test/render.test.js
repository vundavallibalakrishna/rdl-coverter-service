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
import {
  analyzeWindowsWordCompatibility,
  validateWindowsWordRequest,
} from '../src/render/windowsWordCompatibility.js';

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
    pagination: { continuationMarkers: true },
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
  // This row's own text is cut by the break, so the label belongs on every page it resumes on. Nothing
  // announces a continuation before it happens.
  assert.equal((extracted.stdout.match(/Continued from previous page/g) || []).length > 0, true);
  assert.equal((extracted.stdout.match(/Continued on next page/g) || []).length, 0);
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
  const settingsXml = await zip.file('word/settings.xml').async('string');
  assert.match(documentXml, /<w:tbl>/);
  assert.match(documentXml, /North/);
  assert.doesNotMatch(settingsXml, /<w:(?:documentProtection|writeProtection)\b/);
  assert.doesNotMatch(documentXml, /<wps:wsp>/);
});

test('preserves textbox background and foreground colors in page-specific editable DOCX header bands', async () => {
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
  const documentXml = await zip.file('word/document.xml').async('string');
  assert.match(documentXml, /VISIBLE_HEADER_TITLE/);
  assert.match(documentXml, /<w:shd[^>]*w:fill="000080"/);
  assert.match(documentXml, /<w:color w:val="FFFFFF"/i);
  assert.equal(zip.file('word/header1.xml'), null);
});

test('materializes page-dependent RDL header visibility independently on each canonical page', async () => {
  const headerModel = structuredClone(model);
  const repeated = structuredClone(model.body.items.find((item) => item.type === 'Textbox'));
  repeated.name = 'NonFinalPageHeader';
  repeated.value = 'REPEATED_NON_FINAL_HEADER';
  repeated.paragraphs = [['REPEATED_NON_FINAL_HEADER']];
  repeated.hidden = '=IIF(Globals!PageNumber = Globals!TotalPages, True, False)';
  const alwaysHidden = structuredClone(repeated);
  alwaysHidden.name = 'AlwaysHiddenHeader';
  alwaysHidden.value = 'ALWAYS_HIDDEN_HEADER';
  alwaysHidden.paragraphs = [['ALWAYS_HIDDEN_HEADER']];
  alwaysHidden.hidden = 'true';
  headerModel.page.header = {
    height: 40,
    printOnFirstPage: true,
    printOnLastPage: true,
    items: [repeated, alwaysHidden],
  };
  const secondPage = structuredClone(headerModel.body.items.find((item) => item.type === 'Textbox'));
  secondPage.name = 'SecondPageBody';
  secondPage.value = 'SECOND_PAGE_BODY';
  secondPage.paragraphs = [['SECOND_PAGE_BODY']];
  secondPage.pageBreak = { location: 'Start', disabled: 'false' };
  headerModel.body.items.push(secondPage);

  const result = await renderEditableDocx(headerModel, request);
  const zip = await JSZip.loadAsync(result.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  assert.equal(result.pageCount, 2);
  assert.equal((documentXml.match(/REPEATED_NON_FINAL_HEADER/g) || []).length, 1);
  assert.doesNotMatch(documentXml, /ALWAYS_HIDDEN_HEADER/);
});

test('respects page-dependent and literal-hidden nested footer content', async () => {
  const footerModel = structuredClone(model);
  const pageDependent = structuredClone(model.body.items.find((item) => item.type === 'Textbox'));
  pageDependent.name = 'PageDependentFooterText';
  pageDependent.value = 'PAGE_DEPENDENT_FOOTER_DATA';
  pageDependent.paragraphs = [['PAGE_DEPENDENT_FOOTER_DATA']];
  pageDependent.hidden = false;
  const alwaysHidden = structuredClone(pageDependent);
  alwaysHidden.name = 'AlwaysHiddenFooterText';
  alwaysHidden.value = 'ALWAYS_HIDDEN_FOOTER_DATA';
  alwaysHidden.paragraphs = [['ALWAYS_HIDDEN_FOOTER_DATA']];
  alwaysHidden.hidden = true;
  footerModel.page.footer = {
    height: 30,
    printOnFirstPage: true,
    printOnLastPage: true,
    items: [{
      type: 'Rectangle', name: 'PageConditionalFooterContainer', left: 0, top: 0, width: 500, height: 25,
      zIndex: 0, hidden: '=IIF(Globals!PageNumber < 4, True, False)', style: {},
      items: [pageDependent, alwaysHidden],
    }],
  };

  const result = await renderEditableDocx(footerModel, request);
  const zip = await JSZip.loadAsync(result.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  const footerXml = await zip.file('word/footer1.xml').async('string');
  assert.doesNotMatch(documentXml, /PAGE_DEPENDENT_FOOTER_DATA/);
  assert.doesNotMatch(documentXml, /ALWAYS_HIDDEN_FOOTER_DATA/);
  assert.doesNotMatch(footerXml, /PAGE_DEPENDENT_FOOTER_DATA/);
  assert.doesNotMatch(footerXml, /ALWAYS_HIDDEN_FOOTER_DATA/);
});

test('preserves tablix-level outer borders in editable DOCX', async () => {
  const borderedModel = structuredClone(model);
  const tablix = borderedModel.body.items.find((item) => item.type === 'Tablix');
  tablix.style.borders.bottom = { style: 'Solid', color: '#112233', width: 1 };
  const result = await renderEditableDocx(borderedModel, request);
  const zip = await JSZip.loadAsync(result.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  assert.match(documentXml, /<w:tcBorders>[\s\S]*?<w:bottom w:val="single" w:color="112233" w:sz="8"/);
});

test('page-locks a multi-page native table to the canonical PDF fragments without duplicating text', async () => {
  const stressModel = structuredClone(model);
  const tablix = stressModel.body.items.find((item) => item.type === 'Tablix');
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

  const [canonical, result] = await Promise.all([
    renderPdf(stressModel, stressRequest, config),
    renderEditableDocx(stressModel, stressRequest, config),
  ]);
  const zip = await JSZip.loadAsync(result.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  const nativeTables = (documentXml.match(/<w:tbl>/g) || []).length;
  const nativeTableXml = [...documentXml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
  const nativeRows = nativeTableXml.flatMap((table) => [...table.matchAll(/<w:tr>[\s\S]*?<\/w:tr>/g)].map((match) => match[0]));
  assert.equal(result.pageCount, canonical.pageCount);
  assert.equal(nativeTables, canonical.pageCount);
  assert.equal((documentXml.match(/<wps:wsp>/g) || []).length, 0);
  assert.equal(nativeRows.every((row) => /<w:cantSplit\/>/.test(row)), true, 'every physical row remains atomic');
  assert.equal(nativeRows.every((row) => /<w:trHeight\b/.test(row)), true, 'every physical row has a bounded measured height');
  assert.equal((documentXml.match(/LONG_CELL_START/g) || []).length, 1);
  assert.equal((documentXml.match(/LONG_CELL_END/g) || []).length, 1);
  for (let index = 1; index <= 45; index += 1) {
    const marker = `UNIQUE_ROW_${String(index).padStart(3, '0')}`;
    assert.equal((documentXml.match(new RegExp(marker, 'g')) || []).length, 1);
  }
});

test('Windows Word analysis reports the fixed platform contract and obsolete options fail closed', () => {
  const analysis = analyzeWindowsWordCompatibility(model, config);
  assert.equal(analysis.platform, 'Microsoft Word for Windows');
  assert.equal(analysis.layoutMode, 'windows-paged-editable');
  assert.equal(analysis.limits.tableColumns, 63);
  assert.throws(
    () => validateWindowsWordRequest({ docx: { nativePageFragments: false } }),
    (error) => error.code === 'RDL_INVALID',
  );
  assert.throws(
    () => validateWindowsWordRequest({ docxProfile: 'removed-profile' }),
    (error) => error.code === 'RDL_INVALID',
  );
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

test('page-specific footer materializes canonical page numbers as native literal text without fields', async () => {
  const footerModel = structuredClone(model);
  const footerText = structuredClone(model.body.items.find((item) => item.type === 'Textbox'));
  footerText.name = 'PageFooter';
  footerText.value = '="Page " & Globals!PageNumber & " of " & Globals!TotalPages';
  footerText.paragraphs = [[{ value: footerText.value, markupType: 'None' }]];
  footerText.style = { ...footerText.style, fontFamily: 'Arial', fontSize: 8, fontWeight: 'Normal' };
  footerText.left = 400;
  footerText.top = 0;
  footerText.width = 100;
  footerText.height = 20;
  footerModel.page.footer = { height: 24, printOnFirstPage: true, printOnLastPage: true, items: [footerText] };
  const result = await renderEditableDocx(footerModel, request);
  const zip = await JSZip.loadAsync(result.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  const footerXml = await zip.file('word/footer1.xml').async('string');
  assert.doesNotMatch(documentXml, /Page 1 of 1/);
  assert.match(footerXml.replace(/<[^>]+>/g, ''), /Page 1 of 1/);
  assert.doesNotMatch(footerXml, /<w:instrText|<w:fldSimple|w:dirty=/);
});

test('each page-number text run inherits its own resolved RDL style', async () => {
  const footerModel = structuredClone(model);
  const footerText = structuredClone(model.body.items.find((item) => item.type === 'Textbox'));
  const arial = { ...footerText.style, fontFamily: 'Arial', fontSize: 8 };
  const times = { ...footerText.style, fontFamily: 'Times New Roman', fontSize: 11, fontWeight: 'Bold' };
  footerText.name = 'MixedStylePageFooter';
  footerText.paragraphs = [[
    { value: '=Globals!PageNumber', markupType: 'None', style: arial },
    { value: '=" of "', markupType: 'None', style: arial },
    { value: '=Globals!TotalPages', markupType: 'None', style: times },
  ]];
  footerModel.page.footer = { height: 24, printOnFirstPage: true, printOnLastPage: true, items: [footerText] };

  const zip = await JSZip.loadAsync((await renderEditableDocx(footerModel, request)).buffer);
  const footerXml = await zip.file('word/footer1.xml').async('string');
  const runs = [...footerXml.matchAll(/<w:r>([\s\S]*?)<\/w:r>/g)].map((match) => match[0]);
  const pageRun = runs.find((run) => /<w:t[^>]*>1<\/w:t>/.test(run) && /w:ascii="Arial"/.test(run)) || '';
  const totalRun = runs.find((run) => /<w:t[^>]*>1<\/w:t>/.test(run) && /w:ascii="Times New Roman"/.test(run)) || '';
  assert.match(pageRun, /<w:sz w:val="16"/);
  assert.match(totalRun, /<w:sz w:val="22"/);
  assert.match(totalRun, /<w:b\/>/);
  assert.doesNotMatch(footerXml, /<w:instrText/);
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
  // The PDF-measured lines become distinct native Word lines (paragraphs or explicit breaks).
  assert.ok(
    (documentXml.match(/<w:br\/>/g) || []).length >= 1
      || (documentXml.match(/<w:p>/g) || []).length >= 2,
  );
  const nativeText = [...documentXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
    .map((match) => match[1])
    .join('');
  assert.match(nativeText, /First line indented/); // tab expanded to a space
  assert.match(nativeText, /Second line/);
  assert.equal(documentXml.includes(TAB), false); // no raw control character leaks into the XML text
});
