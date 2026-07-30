import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { promisify } from 'node:util';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import {
  editableFontEmbeddingPermission,
  fontEmbeddingEligibility,
  resolveFontFile,
} from '../src/render/fonts.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderPdf } from '../src/render/pdf.js';

const execFileAsync = promisify(execFile);
const fixture = await fs.readFile(new URL('./fixtures/basic.rdl', import.meta.url));
const baseModel = parseRdl(fixture);
const request = {
  outputFileName: 'windows-paged-contract',
  parameters: { Title: 'Sales', Choice: 'A' },
  datasets: {
    Sales: [
      { Name: 'North wrapped evidence', Amount: 1234.5 },
      { Name: 'South', Amount: 99 },
    ],
  },
};
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

test('PDF layout recording is visually non-invasive and captures grounded page/table/text geometry', async (context) => {
  const [ordinary, captured] = await Promise.all([
    renderPdf(baseModel, request, config),
    renderPdf(baseModel, request, config, { captureLayoutTrace: true }),
  ]);
  assert.equal(captured.pageCount, ordinary.pageCount);
  assert.equal(captured.layoutTrace.pageCount, ordinary.pageCount);
  const page = captured.layoutTrace.pages[0];
  assert.equal(page.width, baseModel.page.width);
  assert.equal(page.height, baseModel.page.height);
  assert.deepEqual(Object.keys(page.regions), ['header', 'body', 'footer']);
  assert.ok(page.items.length > 0);
  assert.ok(page.tablixFragments.length > 0);
  const fragment = page.tablixFragments[0];
  assert.ok(fragment.columnWidths.length > 0);
  assert.ok(fragment.rowHeights.length > 0);
  assert.ok(fragment.cells.some((cell) => cell.text.includes('North')));
  const tracedLine = page.items.flatMap((item) => item.lines || []).find((line) => line.runs?.length);
  assert.ok(Number.isFinite(tracedLine.x));
  assert.ok(Number.isFinite(tracedLine.y));
  assert.ok(Number.isFinite(tracedLine.baseline));
  assert.ok(Number.isFinite(tracedLine.runs[0].x));
  assert.ok(Number.isFinite(tracedLine.runs[0].baseline));
  assert.equal(typeof tracedLine.runs[0].font.file, 'string');

  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-trace-proof-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const ordinaryPath = path.join(tempDir, 'ordinary.pdf');
  const capturedPath = path.join(tempDir, 'captured.pdf');
  await Promise.all([
    fs.writeFile(ordinaryPath, ordinary.buffer),
    fs.writeFile(capturedPath, captured.buffer),
  ]);
  const [ordinaryText, capturedText] = await Promise.all([
    execFileAsync('pdftotext', ['-bbox-layout', ordinaryPath, '-']),
    execFileAsync('pdftotext', ['-bbox-layout', capturedPath, '-']),
  ]);
  const stableBboxXml = (value) => value.replace(
    /<meta name="CreationDate" content="[^"]*"\/>\s*/g,
    '',
  );
  assert.equal(
    stableBboxXml(capturedText.stdout),
    stableBboxXml(ordinaryText.stdout),
    'trace capture must not move or alter PDF text',
  );
  await Promise.all([
    execFileAsync(config.pdftoppmPath, ['-f', '1', '-singlefile', '-png', '-r', '144', ordinaryPath, path.join(tempDir, 'ordinary')]),
    execFileAsync(config.pdftoppmPath, ['-f', '1', '-singlefile', '-png', '-r', '144', capturedPath, path.join(tempDir, 'captured')]),
  ]);
  assert.deepEqual(
    await fs.readFile(path.join(tempDir, 'captured.png')),
    await fs.readFile(path.join(tempDir, 'ordinary.png')),
    'trace capture must not change the 144-DPI PDF raster',
  );
});

test('page-locked DOCX contains one native fixed page grid per PDF page and all four embedded font faces', async () => {
  const rendered = await renderEditableDocx(baseModel, request, config);
  const zip = await JSZip.loadAsync(rendered.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  const fontTableXml = await zip.file('word/fontTable.xml').async('string');
  const relationshipParts = await Promise.all(
    Object.keys(zip.files)
      .filter((name) => /(?:^|\/)_rels\/.+\.rels$/i.test(name))
      .map((name) => zip.file(name).async('string')),
  );
  assert.equal(rendered.layoutMode, 'windows-paged-editable');
  assert.equal(rendered.editableTextRatio, 1);
  assert.match(rendered.canonicalPdfSha256, /^[a-f0-9]{64}$/);
  assert.equal((documentXml.match(/<w:tbl>/g) || []).length, rendered.pageCount);
  assert.equal((documentXml.match(/<w:sectPr(?:\s|>)/g) || []).length, rendered.pageCount);
  assert.equal((documentXml.match(/<w:tblLayout w:type="fixed"\/>/g) || []).length, rendered.pageCount);
  assert.ok((documentXml.match(/<w:trHeight[^>]*w:hRule="exact"/g) || []).length > 0);
  assert.doesNotMatch(documentXml, /<w:docGrid\b/);
  const nativeText = [...documentXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
    .map((match) => match[1])
    .join('');
  assert.match(nativeText, /North wrapped evidence/);
  assert.doesNotMatch(documentXml, /<wps:wsp>|<v:shape(?:\s|>)/);
  assert.equal(relationshipParts.some((xml) => /TargetMode="External"/.test(xml)), false);

  const fontParts = Object.keys(zip.files).filter((name) => /^word\/fonts\/.+\.odttf$/i.test(name));
  assert.equal((fontTableXml.match(/<w:embedRegular\b/g) || []).length, 1);
  assert.equal((fontTableXml.match(/<w:embedBold\b/g) || []).length, 1);
  assert.equal((fontTableXml.match(/<w:embedItalic\b/g) || []).length, 1);
  assert.equal((fontTableXml.match(/<w:embedBoldItalic\b/g) || []).length, 1);
  assert.equal((fontTableXml.match(/\bw:subsetted="0"/g) || []).length, 4);
  assert.equal(fontParts.length, 4);
});

test('section anchors and footer terminators cannot snap page-locked geometry to Word’s document grid', async () => {
  const paged = structuredClone(baseModel);
  const secondPage = structuredClone(paged.body.items.find((item) => item.type === 'Textbox'));
  secondPage.name = 'SecondPageAnchorProbe';
  secondPage.value = 'SECOND_PAGE_ANCHOR_PROBE';
  secondPage.paragraphs = [['SECOND_PAGE_ANCHOR_PROBE']];
  secondPage.pageBreak = { location: 'Start', disabled: 'false' };
  paged.body.items.push(secondPage);
  const footerText = structuredClone(paged.body.items.find((item) => item.type === 'Textbox'));
  footerText.name = 'FooterTerminatorProbe';
  footerText.value = 'FOOTER_TERMINATOR_PROBE';
  footerText.paragraphs = [['FOOTER_TERMINATOR_PROBE']];
  footerText.left = 0;
  footerText.top = 0;
  footerText.width = 160;
  footerText.height = 16;
  paged.page.footer = {
    height: 20,
    printOnFirstPage: true,
    printOnLastPage: true,
    items: [footerText],
  };

  const rendered = await renderEditableDocx(paged, request, config);
  const zip = await JSZip.loadAsync(rendered.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  const footerXml = await zip.file('word/footer1.xml').async('string');

  assert.doesNotMatch(documentXml, /<w:docGrid\b/);
  assert.match(documentXml, /<w:pgMar\b(?=[^>]*w:bottom="-40")/);
  assert.match(
    documentXml,
    /<w:p><w:pPr><w:spacing w:after="0" w:before="0" w:line="1" w:lineRule="exact"\/><w:sectPr\b/,
  );
  assert.match(
    footerXml,
    /<\/w:tbl><w:p><w:pPr><w:spacing w:after="0" w:before="0" w:line="1" w:lineRule="exact"\/><\/w:pPr>/,
  );
});

test('a near-full PDF page preserves traced row heights and uses only the non-visible section flow allowance', async () => {
  const nearFull = structuredClone(baseModel);
  nearFull.page.width = 300;
  nearFull.page.height = 200;
  nearFull.page.marginTop = 0;
  nearFull.page.marginRight = 0;
  nearFull.page.marginBottom = 0;
  nearFull.page.marginLeft = 0;
  nearFull.page.header = {
    height: 10,
    printOnFirstPage: true,
    printOnLastPage: true,
    items: [],
  };
  const body = structuredClone(baseModel.body.items.find((item) => item.type === 'Textbox'));
  body.name = 'NearFullBody';
  body.value = 'NEAR_FULL_BODY';
  body.paragraphs = [['NEAR_FULL_BODY']];
  body.left = 0;
  body.top = 0;
  body.width = 100;
  body.height = 170;
  body.canGrow = false;
  nearFull.body.items = [body];
  const footer = structuredClone(body);
  footer.name = 'NearFullFooter';
  footer.value = 'NEAR_FULL_FOOTER';
  footer.paragraphs = [['NEAR_FULL_FOOTER']];
  footer.top = 0;
  footer.height = 16;
  nearFull.page.footer = {
    height: 20,
    printOnFirstPage: true,
    printOnLastPage: true,
    items: [footer],
  };

  const rendered = await renderEditableDocx(nearFull, request, config);
  const zip = await JSZip.loadAsync(rendered.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');

  // The 10pt PDF spacer remains exactly 200 twips; pagination capacity comes only from the section margin.
  assert.match(documentXml, /<w:trHeight w:val="200" w:hRule="exact"\/>/);
  assert.match(documentXml, /<w:pgMar\b(?=[^>]*w:bottom="-40")/);
});

test('multi-row PDF footer content is isolated in one native footer part outside body pagination', async () => {
  const footerModel = structuredClone(baseModel);
  const source = baseModel.body.items.find((item) => item.type === 'Textbox');
  const first = {
    ...structuredClone(source),
    name: 'FooterPrimaryRow',
    value: 'FOOTER_PRIMARY_ROW',
    paragraphs: [['FOOTER_PRIMARY_ROW']],
    left: 12,
    top: 0,
    width: 250,
    height: 16,
    canGrow: false,
  };
  const second = {
    ...structuredClone(source),
    name: 'FooterSecondaryRow',
    value: 'FOOTER_SECONDARY_ROW',
    paragraphs: [['FOOTER_SECONDARY_ROW']],
    left: 12,
    top: 16,
    width: 250,
    height: 16,
    canGrow: false,
  };
  footerModel.page.footer = {
    height: 32,
    printOnFirstPage: true,
    printOnLastPage: true,
    items: [first, second],
  };

  const rendered = await renderEditableDocx(footerModel, request, config);
  const zip = await JSZip.loadAsync(rendered.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  const documentRelationships = await zip.file('word/_rels/document.xml.rels').async('string');
  const footerXml = await zip.file('word/footer1.xml').async('string');

  assert.doesNotMatch(documentXml, /FOOTER_PRIMARY_ROW|FOOTER_SECONDARY_ROW/);
  assert.match(documentXml, /<w:footerReference w:type="default" r:id="[^"]+"\/>/);
  assert.match(documentRelationships, /Type="[^"]*\/footer" Target="footer1\.xml"/);
  assert.equal((footerXml.match(/<w:tbl>/g) || []).length, 1);
  assert.match(footerXml, /FOOTER_PRIMARY_ROW[\s\S]*FOOTER_SECONDARY_ROW/);
  assert.equal((footerXml.match(/<w:trHeight[^>]*w:hRule="exact"/g) || []).length >= 2, true);
});

test('Windows page, grid-column, and editable-overlap limits fail closed generically', async () => {
  const oversizedPage = structuredClone(baseModel);
  oversizedPage.page.width = 23 * 72;
  await assert.rejects(
    renderEditableDocx(oversizedPage, request, config),
    (error) => error.code === 'UNSUPPORTED_FEATURE' && /22-by-22-inch/.test(error.message),
  );

  const textbox = structuredClone(baseModel.body.items.find((item) => item.type === 'Textbox'));
  const tooWideGrid = structuredClone(baseModel);
  tooWideGrid.page.header = null;
  tooWideGrid.page.footer = null;
  tooWideGrid.page.marginLeft = 0;
  tooWideGrid.page.marginRight = 0;
  tooWideGrid.body.items = Array.from({ length: 64 }, (_, index) => ({
    type: 'Line',
    name: `GridBoundary${index}`,
    left: index * 8,
    top: 0,
    width: 0,
    height: 18,
    zIndex: 0,
    hidden: false,
    style: { border: { style: 'Solid', color: '#000000', width: 1 } },
  }));
  await assert.rejects(
    renderEditableDocx(tooWideGrid, request, config),
    (error) => error.code === 'UNSUPPORTED_FEATURE' && /63 table columns/.test(error.message),
  );

  const overlap = structuredClone(baseModel);
  overlap.page.header = null;
  overlap.page.footer = null;
  const first = {
    ...structuredClone(textbox), name: 'OverlapA', value: 'A', paragraphs: [['A']], canGrow: false,
  };
  const second = {
    ...structuredClone(textbox), name: 'OverlapB', value: 'B', paragraphs: [['B']], canGrow: false,
  };
  overlap.body.items = [first, second];
  await assert.rejects(
    renderEditableDocx(overlap, request, config),
    (error) => error.code === 'UNSUPPORTED_FEATURE' && /Overlapping editable PDF regions/.test(error.message),
  );
});

test('coincident PDF edges remain coincident after the quarter-point Word grid conversion', async () => {
  const adjacent = structuredClone(baseModel);
  adjacent.page.header = null;
  adjacent.page.footer = null;
  adjacent.page.marginLeft = 0;
  adjacent.page.marginRight = 0;
  const textbox = structuredClone(baseModel.body.items.find((item) => item.type === 'Textbox'));
  adjacent.body.items = [
    {
      ...structuredClone(textbox),
      name: 'LeftCell',
      value: 'Left',
      paragraphs: [['Left']],
      left: 0.167,
      top: 0,
      width: 100.169,
      height: 18,
      canGrow: false,
    },
    {
      ...structuredClone(textbox),
      name: 'RightCell',
      value: 'Right',
      paragraphs: [['Right']],
      left: 100.336,
      top: 0,
      width: 100,
      height: 18,
      canGrow: false,
    },
  ];
  const rendered = await renderEditableDocx(adjacent, request, config);
  assert.equal(rendered.layoutMode, 'windows-paged-editable');
  assert.equal(rendered.pageCount, 1);
});

test('exact Word rows compensate for the largest bottom cell padding without counting top padding twice', async () => {
  const padded = structuredClone(baseModel);
  padded.page.header = null;
  padded.page.footer = null;
  padded.page.marginLeft = 0;
  padded.page.marginRight = 0;
  const textbox = structuredClone(baseModel.body.items.find((item) => item.type === 'Textbox'));
  padded.body.items = [{
    ...textbox,
    name: 'PaddedExactRow',
    value: 'Padded row',
    paragraphs: [['Padded row']],
    left: 0,
    top: 0,
    width: 100,
    height: 20,
    canGrow: false,
    style: {
      ...textbox.style,
      paddingTop: 5,
      paddingBottom: 2,
    },
  }];

  const rendered = await renderEditableDocx(padded, request, config);
  const zip = await JSZip.loadAsync(rendered.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');

  // The PDF row is 20pt/400 twips. Word adds the largest 2pt/40-twip bottom cell margin even when
  // hRule="exact", so the serialized row height must be 360 twips. Top padding does not receive this
  // Word-specific addition and therefore must not be subtracted.
  assert.match(documentXml, /<w:trHeight w:val="360" w:hRule="exact"\/>/);
  assert.match(
    documentXml,
    /<w:tcMar>[\s\S]*?<w:top w:type="dxa" w:w="100"\/>[\s\S]*?<w:bottom w:type="dxa" w:w="40"\/>[\s\S]*?<\/w:tcMar>/,
  );

  const noBottomPadding = structuredClone(padded);
  noBottomPadding.body.items[0].style.paddingBottom = 0;
  const unadjusted = await renderEditableDocx(noBottomPadding, request, config);
  const unadjustedZip = await JSZip.loadAsync(unadjusted.buffer);
  const unadjustedXml = await unadjustedZip.file('word/document.xml').async('string');
  assert.match(unadjustedXml, /<w:trHeight w:val="400" w:hRule="exact"\/>/);

  const impossible = structuredClone(padded);
  impossible.body.items[0].height = 1;
  await assert.rejects(
    renderEditableDocx(impossible, request, config),
    (error) => error.code === 'UNSUPPORTED_FEATURE'
      && /too short to preserve its bottom cell padding/.test(error.message),
  );
});

test('OS/2 restricted embedding metadata returns FONT_EMBEDDING_FORBIDDEN', async (context) => {
  const source = resolveFontFile(config.fontDir, 'Arial', false, false);
  assert.ok(source, 'the renderer test environment must provide Arial');
  const data = Buffer.from(await fs.readFile(source));
  const tableCount = data.readUInt16BE(4);
  let os2Offset = null;
  for (let index = 0; index < tableCount; index += 1) {
    const record = 12 + index * 16;
    if (data.toString('ascii', record, record + 4) === 'OS/2') {
      os2Offset = data.readUInt32BE(record + 8);
      break;
    }
  }
  assert.notEqual(os2Offset, null);
  data.writeUInt16BE(0x0002, os2Offset + 8);
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-restricted-font-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const restrictedPath = path.join(tempDir, 'Arial.ttf');
  await fs.writeFile(restrictedPath, data);
  assert.throws(
    () => editableFontEmbeddingPermission(restrictedPath, 'Arial', 'regular'),
    (error) => error.code === 'FONT_EMBEDDING_FORBIDDEN',
  );
  const eligibility = fontEmbeddingEligibility(
    loadConfig({ ...process.env, RDL_FONT_DIR: tempDir, RDL_STRICT_FONTS: 'true' }),
    ['Arial'],
  );
  assert.equal(eligibility[0].eligible, false);
  assert.equal(eligibility[0].variants.regular.reason, 'license-restricted');
  assert.equal(eligibility[0].blocksWindowsPagedEditable, true);
});

test('page-locked DOCX rejects a missing consumed font even when legacy PDF strict mode is disabled', async () => {
  const missingFontFixture = Buffer.from(
    fixture.toString('utf8').replaceAll('Arial', 'Unavailable Certification Font'),
  );
  const missingFontModel = parseRdl(missingFontFixture);
  const nonStrictPdfConfig = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
  await assert.rejects(
    renderEditableDocx(missingFontModel, request, nonStrictPdfConfig),
    (error) => error.code === 'FONT_MISSING'
      && /Unavailable Certification Font:regular/.test(error.message),
  );
  const [eligibility] = fontEmbeddingEligibility(nonStrictPdfConfig, ['Unavailable Certification Font']);
  assert.equal(eligibility.eligible, false);
  assert.equal(eligibility.blocksWindowsPagedEditable, true);
});
