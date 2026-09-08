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
import { renderEditableDocx, renderReflowableDocx } from '../src/render/docx.js';
import { renderPdf } from '../src/render/pdf.js';
import { analyzeWindowsWordCompatibility } from '../src/render/windowsWordCompatibility.js';

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

// Reads the body tables of a Word document into the quantities Word's row arithmetic consumes: the
// published row height and rule, and per cell its grid position, merge state, margins, horizontal border
// thickness, and the paragraph content height (spacing before/after plus exact line pitch per line).
function wordTableRows(documentXml) {
  const attribute = (source, name) => Number((new RegExp(`w:${name}="(-?\\d+)"`).exec(source) || [0, 0])[1]);
  const thickness = (borders, side) => {
    const match = new RegExp(`<w:${side} w:val="(\\w+)"[^>]*w:sz="(\\d+)"`).exec(borders || '');
    if (!match || match[1] === 'none' || match[1] === 'nil') return 0;
    return Math.round((Number(match[2]) / 8) * 20 * (match[1] === 'double' ? 3 : 1));
  };
  return documentXml.split('<w:tbl>').slice(1).map((table) => (
    table.split('</w:tbl>')[0].split(/<w:tr\b/).slice(1).map((row) => {
      const height = /<w:trHeight w:val="(\d+)" w:hRule="(\w+)"/.exec(row);
      let gridStart = 0;
      const cells = row.split(/<w:tc>/).slice(1).map((cell) => {
        const properties = cell.split('</w:tcPr>')[0];
        const gridSpan = attribute(/<w:gridSpan[^>]*>/.exec(properties)?.[0] || '', 'val') || 1;
        const margins = /<w:tcMar>.*?<\/w:tcMar>/.exec(properties)?.[0] || '';
        const borders = /<w:tcBorders>.*?<\/w:tcBorders>/.exec(properties)?.[0] || '';
        const content = [...cell.matchAll(/<w:p>(.*?)<\/w:p>/gs)].reduce((sum, paragraph) => {
          const spacing = /<w:spacing [^>]*\/>/.exec(paragraph[1])?.[0] || '';
          const lines = (paragraph[1].match(/<w:br\/>/g) || []).length + 1;
          return sum + attribute(spacing, 'before') + attribute(spacing, 'after') + attribute(spacing, 'line') * lines;
        }, 0);
        const parsed = {
          gridStart,
          gridSpan,
          vMerge: /<w:vMerge w:val="restart"\/>/.test(properties) ? 'restart' : /<w:vMerge\/>/.test(properties) ? 'continue' : null,
          marginTop: attribute(/<w:top [^>]*>/.exec(margins)?.[0] || '', 'w'),
          marginBottom: attribute(/<w:bottom [^>]*>/.exec(margins)?.[0] || '', 'w'),
          topBorder: thickness(borders, 'top'),
          bottomBorder: thickness(borders, 'bottom'),
          content,
          text: [...cell.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((match) => match[1]).join(''),
        };
        gridStart += gridSpan;
        return parsed;
      });
      return { value: Number(height[1]), rule: height[2], cells };
    })
  ));
}

// Asserts the measured Microsoft Word row arithmetic on every reflowable body row: an `atLeast` row
// renders at value + largest margin pair + thickest rule on the edge above it (plus the edge below it
// for the last row), an `exact` row renders at its value, and both must equal the page-locked row that
// the same canonical trace produced. Content must fit the budget Word compares it with, so that an
// unedited row can never grow.
function assertReflowableRowsMatchPageLocked(reflowableXml, pageLockedXml) {
  const reflowableTables = wordTableRows(reflowableXml);
  const pageLockedTables = wordTableRows(pageLockedXml);
  assert.equal(reflowableTables.length, pageLockedTables.length);
  let atLeastRows = 0;
  reflowableTables.forEach((rows, tableIndex) => {
    const pageLocked = pageLockedTables[tableIndex];
    assert.equal(rows.length, pageLocked.length);
    rows.forEach((row, rowIndex) => {
      assert.equal(pageLocked[rowIndex].rule, 'exact');
      // The page-locked last row gives up the bottom rule Word draws below it, so its traced height is
      // the published exact value plus that rule.
      const traced = pageLocked[rowIndex].value + (rowIndex === rows.length - 1
        ? Math.max(...pageLocked[rowIndex].cells.map((cell) => cell.bottomBorder))
        : 0);
      const edgeAbove = Math.max(
        ...row.cells.map((cell) => cell.topBorder),
        rowIndex > 0 ? Math.max(...rows[rowIndex - 1].cells.map((cell) => cell.bottomBorder)) : 0,
      );
      const edgeBelow = rowIndex === rows.length - 1 ? Math.max(...row.cells.map((cell) => cell.bottomBorder)) : 0;
      const margins = Math.max(...row.cells.map((cell) => cell.marginTop + cell.marginBottom));
      if (row.rule === 'exact') {
        assert.equal(row.value, traced, `table ${tableIndex} row ${rowIndex} exact height`);
        return;
      }
      atLeastRows += 1;
      assert.equal(row.rule, 'atLeast');
      const rendered = row.value + margins + edgeAbove + edgeBelow;
      assert.ok(
        rendered <= traced && rendered >= traced - 1,
        `table ${tableIndex} row ${rowIndex}: Word renders ${rendered} twips for a ${traced} twip canonical row`,
      );
      for (const cell of row.cells) {
        if (cell.vMerge === 'continue') continue;
        let budget = row.rule === 'exact' ? traced - margins - edgeAbove : row.value;
        if (cell.vMerge === 'restart') {
          for (let next = rowIndex + 1; next < rows.length; next += 1) {
            const continuation = rows[next].cells.find((candidate) => candidate.gridStart === cell.gridStart);
            if (continuation?.vMerge !== 'continue') break;
            budget += rows[next].rule === 'exact' ? rows[next].value : rows[next].value;
          }
        }
        assert.ok(
          cell.content <= budget,
          `table ${tableIndex} row ${rowIndex} cell "${cell.text}" needs ${cell.content} twips of ${budget}`,
        );
      }
    });
  });
  assert.ok(atLeastRows > 0, 'the reflowable profile must publish growing rows');
}

async function documentXmlOf(rendered) {
  return (await JSZip.loadAsync(rendered.buffer)).file('word/document.xml').async('string');
}

test('reflowable DOCX publishes Word row arithmetic so unedited rows keep their canonical height', async () => {
  const rendered = await renderReflowableDocx(baseModel, { ...request, output: 'DOCX_REFLOWABLE' }, config);
  const pageLocked = await renderEditableDocx(baseModel, { ...request, output: 'DOCX_EDITABLE' }, config);
  const documentXml = await documentXmlOf(rendered);

  assert.equal(rendered.layoutMode, 'windows-reflowable-editable');
  assert.equal(rendered.pageCount, pageLocked.pageCount);
  assert.doesNotMatch(documentXml, /<w:fitText\b/);
  assert.match(documentXml, /<w:trHeight[^>]*w:hRule="atLeast"/);
  // Both profiles keep the canonical column grid; only row growth differs.
  assert.equal((documentXml.match(/<w:tblLayout w:type="fixed"\/>/g) || []).length, rendered.pageCount);
  assert.match(documentXml, /<w:sz w:val="20"\/>/);
  assertReflowableRowsMatchPageLocked(documentXml, await documentXmlOf(pageLocked));
});

function syntheticTextbox(overrides) {
  const source = baseModel.body.items.find((item) => item.type === 'Textbox');
  const item = structuredClone(source);
  Object.assign(item, overrides);
  item.paragraphs = [[{ ...item.paragraphs[0][0], value: overrides.value }]];
  return item;
}

function withBorders(item, width = 1) {
  for (const side of ['top', 'right', 'bottom', 'left']) {
    item.style.borders[side] = { style: 'Solid', color: '#000000', width };
  }
  item.style.border = { style: 'Solid', color: '#000000', width };
  return item;
}

test('reflowable DOCX compresses content SSRS clips rather than letting Word grow the canonical row', async () => {
  // A non-growing textbox shorter than its own wrapped text: SSRS clips it at the box edge, and the
  // canonical trace records two 14pt lines inside an 8pt box. Word must keep the 8pt row.
  const clipped = structuredClone(baseModel);
  clipped.body.items.push(syntheticTextbox({
    name: 'ClippedBox', value: 'CLIPPED TEXT THAT IS TALLER THAN ITS BOX', left: 7.2, top: 120, width: 300, height: 8, canGrow: false,
  }));
  const rendered = await renderReflowableDocx(clipped, { ...request, output: 'DOCX_REFLOWABLE' }, config);
  const documentXml = await documentXmlOf(rendered);
  const pageLockedXml = await documentXmlOf(await renderEditableDocx(clipped, { ...request, output: 'DOCX_EDITABLE' }, config));
  assertReflowableRowsMatchPageLocked(documentXml, pageLockedXml);

  const row = wordTableRows(documentXml)[0].find((candidate) => candidate.cells.some((cell) => cell.text.includes('CLIPPED')));
  const cell = row.cells.find((candidate) => candidate.text.includes('CLIPPED'));
  assert.equal(row.rule, 'atLeast');
  assert.equal(row.value + cell.marginTop, 160, 'the 8pt canonical row is preserved');
  assert.ok(cell.content <= row.value, 'the clipped text fits the row budget');
  // Every traced word is still present; only the line pitch was reduced.
  assert.equal(cell.text, 'CLIPPED TEXT THAT IS TALLER THAN ITS BOX');
  const clippedCell = documentXml.split('<w:tc>').find((chunk) => chunk.includes('>CLIPPED<'));
  const pitch = Number(/<w:spacing[^>]*w:line="(\d+)" w:lineRule="exact"\/>/.exec(clippedCell)[1]);
  assert.ok(pitch > 1 && pitch < Math.round(15.640625 * 20), `line pitch ${pitch} is compressed below the traced 14pt pitch`);
});

test('reflowable DOCX publishes bands thinner than their own rules as exact rows and gives empty cells an editable line', async () => {
  const variant = structuredClone(baseModel);
  // Two bordered boxes offset by half a point create a half-point band that carries a one-point rule;
  // Word cannot publish that as `atLeast` because the rule alone is taller than the band.
  variant.body.items.push(
    withBorders(syntheticTextbox({ name: 'LeftBox', value: 'LeftBox', left: 7.2, top: 120, width: 150, height: 20, canGrow: false })),
    withBorders(syntheticTextbox({ name: 'RightBox', value: 'RightBox', left: 200, top: 120.5, width: 150, height: 20, canGrow: false })),
    withBorders(syntheticTextbox({ name: 'EmptyBox', value: '', left: 7.2, top: 160, width: 150, height: 20, canGrow: false })),
  );
  const rendered = await renderReflowableDocx(variant, { ...request, output: 'DOCX_REFLOWABLE' }, config);
  const documentXml = await documentXmlOf(rendered);
  const pageLockedXml = await documentXmlOf(await renderEditableDocx(variant, { ...request, output: 'DOCX_EDITABLE' }, config));
  assertReflowableRowsMatchPageLocked(documentXml, pageLockedXml);

  const rows = wordTableRows(documentXml)[0];
  const thin = rows.filter((row) => row.value === 10);
  assert.equal(thin.length, 2);
  assert.ok(thin.every((row) => row.rule === 'exact'));
  assert.ok(rows.filter((row) => row.rule === 'atLeast').length >= rows.length - 2);

  // The empty textbox keeps its canonical row but offers a real line to type into, unlike the one-twip
  // paragraph the page-locked profile uses for empty cells.
  const emptyRow = rows.find((row) => row.cells.some((cell) => cell.vMerge !== 'continue' && cell.content > 1 && cell.text === '' && cell.topBorder > 0));
  assert.ok(emptyRow, 'empty bordered textbox row is present');
  const emptyCell = emptyRow.cells.find((cell) => cell.text === '' && cell.topBorder > 0 && cell.content > 1);
  assert.ok(emptyCell.content >= 200 && emptyCell.content <= emptyRow.value, `empty cell content ${emptyCell.content} within ${emptyRow.value}`);
});

test('reflowable DOCX extends tight lines away from their alignment so Word cannot wrap a traced line', async () => {
  // Measure the traced width of one line first, then size textboxes so that the line has almost no
  // horizontal slack. Word may measure such a line a little wider than the PDF did; without extra
  // editable width it would wrap the last word and grow the row.
  const probeModel = structuredClone(baseModel);
  probeModel.body.items.push(syntheticTextbox({
    name: 'ProbeBox', value: 'Tight measurement line', left: 7.2, top: 150, width: 400, height: 20, canGrow: false,
  }));
  const probe = await renderPdf(probeModel, request, config, { captureLayoutTrace: true });
  const probeLine = probe.layoutTrace.pages[0].items.find((item) => item.itemName === 'ProbeBox').lines[0];
  const tightWidth = probeLine.width + 4 + 0.5; // 2pt padding either side plus half a point of slack

  const variant = structuredClone(baseModel);
  const tight = (name, textAlign, top) => {
    const item = syntheticTextbox({ name, value: 'Tight measurement line', left: 7.2, top, width: tightWidth, height: 20, canGrow: false });
    item.style.textAlign = textAlign;
    item.paragraphs[0][0].style.textAlign = textAlign;
    // The paragraph-level RDL style wins over the textbox style for alignment.
    if (Array.isArray(item.paragraphStyles)) {
      item.paragraphStyles = item.paragraphStyles.map((style) => ({ ...style, textAlign }));
    }
    return item;
  };
  variant.body.items.push(
    tight('TightLeft', 'Left', 150),
    tight('TightRight', 'Right', 175),
    tight('TightCenter', 'Center', 200),
    syntheticTextbox({ name: 'LooseBox', value: 'Tight measurement line', left: 7.2, top: 225, width: tightWidth + 40, height: 20, canGrow: false }),
  );
  const rendered = await renderReflowableDocx(variant, { ...request, output: 'DOCX_REFLOWABLE' }, config);
  const documentXml = await documentXmlOf(rendered);
  const paragraphFor = (text, xml = documentXml) => xml.split('<w:tc>').find((chunk) => chunk.includes(`>${text}<`)).split('</w:pPr>')[0];
  const minimum = Math.ceil((3 + 0.0025 * probeLine.width - 0.5) * 20);

  const left = /<w:ind w:right="(-\d+)"\/>/.exec(paragraphFor('Tight'))?.[1];
  assert.ok(left && -Number(left) >= minimum, `left-aligned tight line extends to the right by at least ${minimum} twips`);
  const tightCells = documentXml.split('<w:tc>').filter((chunk) => chunk.includes('>Tight<'));
  assert.equal(tightCells.length, 4);
  const indents = tightCells.map((chunk) => /<w:ind [^>]*\/>/.exec(chunk.split('</w:pPr>')[0])?.[0] || null);
  assert.match(indents[0], /w:right="-\d+"/);
  assert.doesNotMatch(indents[0], /w:left/);
  assert.match(indents[1], /w:left="-\d+"/);
  assert.doesNotMatch(indents[1], /w:right/);
  assert.match(indents[2], /w:left="-\d+"/);
  assert.match(indents[2], /w:right="-\d+"/);
  assert.equal(indents[3], null, 'a line with generous slack keeps its natural text area');

  // The page-locked profile pins lines with FitText and never needs the allowance.
  const pageLockedXml = await documentXmlOf(await renderEditableDocx(variant, { ...request, output: 'DOCX_EDITABLE' }, config));
  assert.doesNotMatch(pageLockedXml, /<w:ind /);
  assert.match(pageLockedXml, /<w:fitText\b/);
});

test('both Word profiles carry the RDL page margins as section margins without moving any content', async () => {
  // basic.rdl declares 0.5in margins on every side: Word must see them as real margins (its print check
  // flags zero or negative margins), the page grid must span only the text area, and every item must
  // keep the physical position a page-origin grid gave it.
  const canonical = await renderPdf(baseModel, request, config, { captureLayoutTrace: true });
  const page = canonical.layoutTrace.pages[0];
  const marginTwips = Math.round(page.regions.body.x * 20);
  assert.equal(marginTwips, 720);
  for (const render of [renderEditableDocx, renderReflowableDocx]) {
    const documentXml = await documentXmlOf(await render(baseModel, request, config));
    const margins = /<w:pgMar [^>]*>/.exec(documentXml)[0];
    assert.match(margins, new RegExp(`w:left="${marginTwips}"`));
    assert.match(margins, new RegExp(`w:right="${Math.round((page.width - page.regions.body.x - page.regions.body.width) * 20)}"`));
    assert.match(margins, new RegExp(`w:top="${Math.round(page.regions.body.y * 20)}"`));
    assert.match(margins, new RegExp(`w:bottom="${Math.round((page.height - page.bodyBottom) * 20)}"`));
    assert.doesNotMatch(margins, /w:(?:left|right|top|bottom)="(?:0|-\d+)"/);
    const columns = [...documentXml.matchAll(/<w:gridCol w:w="(\d+)"\/>/g)].map((match) => Number(match[1]));
    assert.equal(columns.reduce((sum, width) => sum + width, 0), Math.round(page.regions.body.width * 20));
    assert.match(documentXml, new RegExp(`<w:tblW w:type="dxa" w:w="${columns.reduce((sum, width) => sum + width, 0)}"`));
    // The title box starts 7.2pt inside the body: its column offset plus the margin is its page position
    // on the quarter-point Word grid, exactly as a page-origin grid placed it.
    const title = page.items.find((item) => item.itemName === 'TitleBox');
    const gap = columns[0];
    assert.equal(gap + marginTwips, Math.round(Math.round(title.x / 0.25) * 0.25 * 20));
  }
});

test('page-locked DOCX lets its last row give up the bottom rule Word draws below the table', async () => {
  // Word renders an exact-row table at the sum of its rows plus the bottom rule of its last row and
  // nothing else, so a bordered last row publishes its traced height minus that rule and the story still
  // renders at exactly its traced height.
  const variant = structuredClone(baseModel);
  variant.body.items.push(withBorders(syntheticTextbox({
    name: 'LastBordered', value: 'LAST', left: 7.2, top: 200, width: 150, height: 20, canGrow: false,
  }), 2));
  const documentXml = await documentXmlOf(await renderEditableDocx(variant, { ...request, output: 'DOCX_EDITABLE' }, config));
  const rows = wordTableRows(documentXml)[0];
  const last = rows[rows.length - 1];
  assert.ok(last.cells.some((cell) => cell.text === 'LAST'));
  assert.equal(last.rule, 'exact');
  assert.equal(Math.max(...last.cells.map((cell) => cell.bottomBorder)), 40);
  assert.equal(last.value, 400 - 40);
  // Inner rows keep their full traced height: Word draws their rules inside the exact height.
  assert.ok(rows.slice(0, -1).every((row) => row.value > 0));
});

test('reflowable DOCX reserves Word rule rounding when a ruled grid fills the body band', async () => {
  // Twenty bordered rows stacked to the body boundary: Word renders each 1pt rule a fraction of a twip
  // tall, so the flush grid must give up at least one twip per ruled row at its last row while every
  // other row keeps its exact traced arithmetic.
  const ruled = structuredClone(baseModel);
  ruled.body.items = [];
  const bodyHeight = baseModel.page.height - baseModel.page.marginTop - baseModel.page.marginBottom;
  const rowsCount = 20;
  const rowHeight = bodyHeight / rowsCount;
  for (let index = 0; index < rowsCount; index += 1) {
    ruled.body.items.push(withBorders(syntheticTextbox({
      name: `Ruled${index}`, value: `Row ${index}`, left: 7.2, top: index * rowHeight, width: 200, height: rowHeight, canGrow: false,
    })));
  }
  const ruledRequest = { ...request, datasets: { Sales: [{ Name: 'Only', Amount: 1 }] } };
  const canonical = await renderPdf(ruled, ruledRequest, config, { captureLayoutTrace: true });
  assert.equal(canonical.pageCount, 1);
  const page = canonical.layoutTrace.pages[0];
  const bandTwips = Math.round((page.bodyBottom - page.regions.body.y) * 20);
  const reflowableXml = await documentXmlOf(await renderReflowableDocx(ruled, { ...ruledRequest, output: 'DOCX_REFLOWABLE' }, config));
  const rows = wordTableRows(reflowableXml)[0];
  const rendered = rows.reduce((sum, row, index) => {
    if (row.rule === 'exact') return sum + row.value;
    const edgeAbove = Math.max(...row.cells.map((cell) => cell.topBorder), index > 0 ? Math.max(...rows[index - 1].cells.map((cell) => cell.bottomBorder)) : 0);
    const edgeBelow = index === rows.length - 1 ? Math.max(...row.cells.map((cell) => cell.bottomBorder)) : 0;
    return sum + row.value + Math.max(...row.cells.map((cell) => cell.marginTop + cell.marginBottom)) + edgeAbove + edgeBelow;
  }, 0);
  const ruledRows = rows.filter((row) => row.cells.some((cell) => cell.topBorder > 0 || cell.bottomBorder > 0)).length;
  assert.ok(ruledRows >= rowsCount);
  assert.ok(rendered <= bandTwips - ruledRows, `Word renders ${rendered} twips; the band is ${bandTwips} with ${ruledRows} ruled rows`);
  assert.ok(rendered >= bandTwips - ruledRows - 2, 'no more than the rounding allowance is given up');
});

test('reflowable DOCX renders a body grid closing flush on the body boundary at exactly the band height', async () => {
  // A non-growing textbox whose bottom edge lands exactly on the traced body boundary. The terminal
  // paragraphs are hidden, so the grid may fill the whole band and must not exceed it by a twip.
  const flush = structuredClone(baseModel);
  const bodyHeight = baseModel.page.height - baseModel.page.marginTop - baseModel.page.marginBottom;
  flush.body.items.push(syntheticTextbox({
    name: 'FlushBox', value: 'FLUSH', left: 7.2, top: bodyHeight - 20, width: 200, height: 20, canGrow: false,
  }));
  // One detail row keeps the tablix at its declared height, so nothing below it is displaced.
  const flushRequest = { ...request, datasets: { Sales: [{ Name: 'Only', Amount: 1 }] } };
  const canonical = await renderPdf(flush, flushRequest, config, { captureLayoutTrace: true });
  assert.equal(canonical.pageCount, 1);
  const page = canonical.layoutTrace.pages[0];
  const bandTwips = Math.round((page.bodyBottom - page.regions.body.y) * 20);
  const reflowableXml = await documentXmlOf(await renderReflowableDocx(flush, { ...flushRequest, output: 'DOCX_REFLOWABLE' }, config));
  const pageLockedXml = await documentXmlOf(await renderEditableDocx(flush, { ...flushRequest, output: 'DOCX_EDITABLE' }, config));
  const pageLockedRows = wordTableRows(pageLockedXml)[0];
  assert.equal(pageLockedRows.reduce((sum, row) => sum + row.value, 0), bandTwips, 'the canonical grid fills the body band');

  const rows = wordTableRows(reflowableXml)[0];
  const rendered = rows.reduce((sum, row, index) => {
    if (row.rule === 'exact') return sum + row.value;
    const edgeAbove = Math.max(
      ...row.cells.map((cell) => cell.topBorder),
      index > 0 ? Math.max(...rows[index - 1].cells.map((cell) => cell.bottomBorder)) : 0,
    );
    const edgeBelow = index === rows.length - 1 ? Math.max(...row.cells.map((cell) => cell.bottomBorder)) : 0;
    return sum + row.value + Math.max(...row.cells.map((cell) => cell.marginTop + cell.marginBottom)) + edgeAbove + edgeBelow;
  }, 0);
  // The grid fills the band up to the one-twip rounding allowance of each ruled row (the tablix rows).
  const ruledRows = rows.filter((row) => row.cells.some((cell) => cell.topBorder > 0 || cell.bottomBorder > 0)).length;
  assert.ok(
    rendered <= bandTwips && rendered >= bandTwips - ruledRows - 1,
    `Word renders ${rendered} twips inside a ${bandTwips} twip band with ${ruledRows} ruled rows`,
  );
  const last = rows[rows.length - 1];
  const flushCell = last.cells.find((cell) => cell.text === 'FLUSH');
  assert.ok(flushCell.content <= last.value, 'the flush textbox content fits its row');
});

test('reflowable DOCX reserves the canonical footer band as a Word body boundary', async () => {
  const footerModel = structuredClone(baseModel);
  const source = baseModel.body.items.find((item) => item.type === 'Textbox');
  footerModel.page.footer = {
    height: 24,
    printOnFirstPage: true,
    printOnLastPage: true,
    items: [{
      ...structuredClone(source),
      name: 'ReflowableFooter',
      value: 'REFLOWABLE_FOOTER',
      paragraphs: [['REFLOWABLE_FOOTER']],
      left: 0,
      top: 0,
      width: 200,
      height: 18,
      canGrow: false,
    }],
  };
  const canonical = await renderPdf(footerModel, request, config, { captureLayoutTrace: true });
  const footer = canonical.layoutTrace.pages[0].regions.footer;
  const expectedBottomMargin = Math.round((canonical.layoutTrace.pages[0].height - footer.y) * 20);

  const rendered = await renderReflowableDocx(footerModel, { ...request, output: 'DOCX_REFLOWABLE' }, config);
  const documentXml = await (await JSZip.loadAsync(rendered.buffer))
    .file('word/document.xml').async('string');

  assert.match(
    documentXml,
    new RegExp(`<w:pgMar\\b(?=[^>]*w:bottom="${expectedBottomMargin}")`),
  );
  assert.doesNotMatch(documentXml, /<w:pgMar\b(?=[^>]*w:bottom="-40")/);
});

test('page-locked DOCX preserves mixed PDF line pitches without inflating every wrapped Word line', async () => {
  const mixed = structuredClone(baseModel);
  mixed.page.header = null;
  mixed.page.footer = null;
  const item = structuredClone(mixed.body.items.find((candidate) => candidate.type === 'Textbox'));
  const valueStyle = { ...item.style, fontSize: 9, fontWeight: 'Normal' };
  item.top = 0;
  item.left = 0;
  item.width = 180;
  item.height = 18;
  item.value = 'Division : alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau';
  item.style = valueStyle;
  item.paragraphs = [[
    {
      value: 'Division : ',
      evaluationMode: 'Auto',
      markupType: 'None',
      style: { ...valueStyle, fontSize: 10, fontWeight: 'Bold' },
    },
    {
      value: 'alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau',
      evaluationMode: 'Auto',
      markupType: 'None',
      style: valueStyle,
    },
  ]];
  item.paragraphStyles = [{ ...valueStyle, spaceBefore: 0, spaceAfter: 0 }];
  mixed.body.items = [item];

  const mixedRequest = { outputFileName: 'mixed-word-line-pitch', parameters: {}, datasets: {} };
  const canonical = await renderPdf(mixed, mixedRequest, config, { captureLayoutTrace: true });
  const traced = canonical.layoutTrace.pages[0].items.find((candidate) => candidate.text?.startsWith('Division'));
  assert.ok(traced);
  assert.ok(traced.lines.length >= 3);
  const tracedPitches = traced.lines.map((line) => Math.max(1, Math.round(line.contentHeight * 20)));
  assert.ok(new Set(tracedPitches).size >= 2, 'fixture must contain physically different line pitches');

  const rendered = await renderEditableDocx(mixed, mixedRequest, config);
  const documentXml = await (await JSZip.loadAsync(rendered.buffer))
    .file('word/document.xml').async('string');
  const textParagraphs = [...documentXml.matchAll(/<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g)]
    .map((match) => match[1])
    .filter((paragraph) => /<w:t(?:\s[^>]*)?>[^<]+<\/w:t>/.test(paragraph));

  assert.equal(textParagraphs.length, new Set(tracedPitches).size);
  for (const pitch of new Set(tracedPitches)) {
    assert.ok(textParagraphs.some((paragraph) => new RegExp(`<w:spacing\\b[^>]*w:line="${pitch}"`).test(paragraph)));
  }
  const explicitBoundaries = (textParagraphs.length - 1)
    + textParagraphs.reduce((total, paragraph) => total + (paragraph.match(/<w:br\/>/g) || []).length, 0);
  assert.equal(explicitBoundaries, traced.lines.length - 1, 'every canonical wrap point must remain explicit');
  const fitTextElements = documentXml.match(/<w:fitText\b[^>]*\/>/g) || [];
  const fitTextIds = new Set(fitTextElements.map((element) => element.match(/w:id="(\d+)"/)?.[1]).filter(Boolean));
  assert.equal(fitTextIds.size, traced.lines.length, 'each physical PDF line must own one manual Word width');
  for (const line of traced.lines) {
    const expectedWidth = Math.max(1, Math.round(line.width * 20));
    assert.ok(
      fitTextElements.some((element) => new RegExp(`w:val="${expectedWidth}"`).test(element)),
      `missing manual Word width for traced ${expectedWidth}-twip line`,
    );
  }
  assert.equal(rendered.pageCount, canonical.pageCount);
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
  // The footer band and RDL bottom margin are a real (positive) Word bottom margin: 20pt band + 36pt.
  assert.match(documentXml, /<w:pgMar\b(?=[^>]*w:bottom="1120")/);
  assert.doesNotMatch(documentXml, /<w:pgMar\b(?=[^>]*w:bottom="-)/);
  // Every section paragraph is one exact twip with a hidden mark, so it needs no vertical room.
  assert.match(
    documentXml,
    /<w:p><w:pPr><w:spacing w:after="0" w:before="0" w:line="1" w:lineRule="exact"\/><w:rPr><w:vanish\/><\/w:rPr><w:sectPr\b/,
  );
  assert.match(
    documentXml,
    /<\/w:tbl><w:p><w:pPr><w:spacing w:after="0" w:before="0" w:line="1" w:lineRule="exact"\/><w:rPr><w:vanish\/><\/w:rPr><\/w:pPr><\/w:p><w:sectPr\b/,
  );
  // The footer story terminator is hidden too, so a footer table filling its band cannot grow the story.
  assert.match(
    footerXml,
    /<\/w:tbl><w:p><w:pPr><w:spacing w:after="0" w:before="0" w:line="1" w:lineRule="exact"\/><w:rPr><w:vanish\/><\/w:rPr><\/w:pPr><w:r><w:rPr><w:vanish\/><\/w:rPr>/,
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

  // Native headers own the 10pt header band, so the body grid starts at the canonical body origin
  // instead of adding an opaque spacer row that can cover the Word header story.
  assert.match(documentXml, /<w:trHeight w:val="3400" w:hRule="exact"\/>/);
  assert.match(documentXml, /<w:pgMar\b(?=[^>]*w:top="200")/);
  // The footer band and RDL bottom margin are a real positive Word margin in both profiles.
  const tracedPage = (await renderPdf(nearFull, request, config, { captureLayoutTrace: true })).layoutTrace.pages[0];
  assert.match(documentXml, new RegExp(`<w:pgMar\\b(?=[^>]*w:bottom="${Math.round((tracedPage.height - tracedPage.bodyBottom) * 20)}")`));
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

test('PDF page headers are emitted in native Word header stories rather than the body grid', async () => {
  const headerModel = structuredClone(baseModel);
  const source = baseModel.body.items.find((item) => item.type === 'Textbox');
  const headerItem = {
    ...structuredClone(source),
    name: 'NativeHeaderStory',
    value: 'NATIVE_HEADER_STORY_MARKER',
    paragraphs: [['NATIVE_HEADER_STORY_MARKER']],
    left: 12,
    top: 6,
    width: 250,
    height: 18,
    canGrow: false,
  };
  headerModel.page.header = {
    height: 42,
    printOnFirstPage: true,
    printOnLastPage: true,
    items: [headerItem],
  };
  headerModel.page.footer = null;

  const canonical = await renderPdf(headerModel, request, config, { captureLayoutTrace: true });
  const tracedHeader = canonical.layoutTrace.pages[0].regions.header;
  assert.equal(tracedHeader.y, headerModel.page.marginTop);
  assert.equal(tracedHeader.height, headerModel.page.header.height);

  const rendered = await renderEditableDocx(headerModel, request, config);
  const zip = await JSZip.loadAsync(rendered.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  const documentRelationships = await zip.file('word/_rels/document.xml.rels').async('string');
  const headerXml = await zip.file('word/header1.xml').async('string');

  assert.doesNotMatch(documentXml, /NATIVE_HEADER_STORY_MARKER/);
  assert.match(documentXml, /<w:headerReference w:type="default" r:id="[^"]+"\/>/);
  assert.match(
    documentXml,
    new RegExp(`<w:pgMar\\b(?=[^>]*w:header="${Math.round(tracedHeader.y * 20)}")`),
  );
  assert.match(documentRelationships, /Type="[^"]*\/header" Target="header1\.xml"/);
  assert.match(headerXml, /NATIVE_HEADER_STORY_MARKER/);
  assert.equal((headerXml.match(/<w:tbl>/g) || []).length, 1);
});

test('a tablix top edge on the body/header boundary remains in the body grid only', async () => {
  // A resolved fragment border has a zero-height box on the first body coordinate. The trace's
  // inclusive region classification calls that coordinate "header", but SSRS paints the rule as the
  // tablix's body edge. Native Word must not also add it to the header story.
  const boundaryModel = structuredClone(baseModel);
  const bodyTextbox = boundaryModel.body.items.find((item) => item.type === 'Textbox');
  const tablix = boundaryModel.body.items.find((item) => item.type === 'Tablix');
  boundaryModel.body.items = [tablix];
  tablix.top = 0;
  tablix.style.borders = {
    top: { style: 'Solid', color: '#000000', width: 1 },
    right: { style: 'Solid', color: '#000000', width: 1 },
    bottom: { style: 'Solid', color: '#000000', width: 1 },
    left: { style: 'Solid', color: '#000000', width: 1 },
  };
  boundaryModel.page.marginTop = 36;
  boundaryModel.page.header = {
    height: 30,
    printOnFirstPage: true,
    printOnLastPage: true,
    items: [{
      ...bodyTextbox,
      name: 'BoundaryHeader',
      value: 'HEADER_MARKER',
      paragraphs: [['HEADER_MARKER']],
      left: 0,
      top: 0,
      width: 200,
      height: 20,
      canGrow: false,
    }],
  };

  const canonical = await renderPdf(boundaryModel, request, config, { captureLayoutTrace: true });
  const boundaryLine = canonical.layoutTrace.pages[0].items.find((item) => (
    item.traceRole === 'resolvedTablixFragmentBorder' && item.fragmentSide === 'top'
  ));
  assert.equal(boundaryLine?.region, 'header');
  assert.equal(boundaryLine?.y, canonical.layoutTrace.pages[0].regions.body.y);

  for (const [render, output] of [
    [renderEditableDocx, 'DOCX_EDITABLE'],
    [renderReflowableDocx, 'DOCX_REFLOWABLE'],
  ]) {
    const rendered = await render(boundaryModel, { ...request, output }, config);
    const zip = await JSZip.loadAsync(rendered.buffer);
    const headerXml = await zip.file('word/header1.xml').async('string');
    const documentXml = await zip.file('word/document.xml').async('string');
    assert.match(headerXml, /HEADER_MARKER/);
    assert.equal((headerXml.match(/<w:tr>/g) || []).length, 1);
    assert.doesNotMatch(headerXml, /<w:top w:val="single"/);
    assert.match(documentXml, /<w:top w:val="single"/);
  }
});

test('Windows page, grid-column, and editable-overlap limits fail closed generically', async () => {
  const oversizedPage = structuredClone(baseModel);
  oversizedPage.page.width = 23 * 72;
  await assert.rejects(
    renderEditableDocx(oversizedPage, request, config),
    (error) => error.code === 'UNSUPPORTED_FEATURE'
      && /22-by-22-inch/.test(error.message)
      && error.details?.widthIn === 23
      && error.details?.maximumIn === 22,
  );
  const pageAnalysis = analyzeWindowsWordCompatibility(oversizedPage, config);
  assert.equal(pageAnalysis.page.widthIn, 23);
  assert.equal(pageAnalysis.page.maximumCm, 55.88);
  assert.equal(
    pageAnalysis.unsupported.find((entry) => entry.code === 'WORD_PAGE_SIZE_LIMIT')
      ?.details?.exactPageLockedOutputAvailable,
    false,
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

test('safe shared-edge overlaps coalesce without permitting genuine content crossings', async () => {
  const adjacent = structuredClone(baseModel);
  adjacent.page.header = null;
  adjacent.page.footer = null;
  adjacent.page.marginLeft = 0;
  adjacent.page.marginRight = 0;
  const source = structuredClone(baseModel.body.items.find((item) => item.type === 'Textbox'));
  const textBox = (name, value, left, top, width, height) => ({
    ...structuredClone(source),
    name,
    value,
    paragraphs: [[value]],
    left,
    top,
    width,
    height,
    canGrow: false,
    style: {
      ...structuredClone(source.style),
      border: { style: 'Solid', width: 1, color: '#000000' },
      borders: {
        top: { style: 'Solid', width: 1, color: '#000000' },
        right: { style: 'Solid', width: 1, color: '#000000' },
        bottom: { style: 'Solid', width: 1, color: '#000000' },
        left: { style: 'Solid', width: 1, color: '#000000' },
      },
    },
  });
  const icon = textBox('ClippedIcon', 'X', 0, 80, 22, 34);
  const iconLabel = textBox('IconLabel', 'ICON_LABEL', 20, 80, 80, 34);
  for (const [item, textAlign] of [[icon, 'Center'], [iconLabel, 'Left']]) {
    item.style.backgroundColor = null;
    item.style.textAlign = textAlign;
    item.style.border = { style: 'None', width: 0, color: '#000000' };
    item.style.borders = Object.fromEntries(
      ['top', 'right', 'bottom', 'left']
        .map((side) => [side, { style: 'None', width: 0, color: '#000000' }]),
    );
  }
  adjacent.body.items = [{
    type: 'Rectangle',
    name: 'AdjacentEdgeContainer',
    left: 0,
    top: 0,
    width: 200,
    height: 114,
    zIndex: 0,
    hidden: false,
    style: {},
    items: [
      textBox('Upper', 'UPPER', 0, 0, 100, 20),
      textBox('Lower', 'LOWER', 0, 19.25, 100, 20),
      textBox('Left', 'LEFT', 0, 50, 100, 20),
      textBox('Right', 'RIGHT', 99.25, 50, 100, 20),
      icon,
      iconLabel,
    ],
  }];

  const rendered = await renderEditableDocx(adjacent, request, config);
  const documentXml = await (
    await JSZip.loadAsync(rendered.buffer)
  ).file('word/document.xml').async('string');
  const nativeText = [...documentXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
    .map((match) => match[1])
    .join('');

  for (const marker of ['UPPER', 'LOWER', 'LEFT', 'RIGHT', 'ICON_LABEL']) {
    assert.match(nativeText, new RegExp(marker));
  }
  assert.match(nativeText, /X/);
  assert.match(documentXml, /<w:trHeight w:val="395" w:hRule="exact"\/>/);
  assert.match(documentXml, /<w:gridCol w:w="1985"\/>/);
  assert.doesNotMatch(documentXml, /<wps:wsp>|<v:shape(?:\s|>)/);
});

test('a trace-rounded or sub-half-point textbox/image edge coalesces without allowing real overlap', async () => {
  const { PNG } = await import('pngjs');
  const png = new PNG({ width: 4, height: 4 });
  png.data.fill(0x80);
  for (let index = 3; index < png.data.length; index += 4) png.data[index] = 0xFF;

  const adjacent = structuredClone(baseModel);
  adjacent.page.header = null;
  adjacent.page.footer = null;
  adjacent.page.marginLeft = 0;
  adjacent.page.marginRight = 0;
  adjacent.embeddedImages = {
    ...(adjacent.embeddedImages || {}),
    HeaderLogo: { data: PNG.sync.write(png).toString('base64'), mimeType: 'image/png' },
  };
  const title = structuredClone(baseModel.body.items.find((item) => item.type === 'Textbox'));
  Object.assign(title, {
    name: 'HeaderTitle',
    value: 'Risk Register Report',
    paragraphs: [['Risk Register Report']],
    left: 10.123505,
    top: 10,
    width: 100.000507,
    height: 30,
    canGrow: false,
  });
  title.style = { ...title.style, backgroundColor: '#000080' };
  const logo = {
    type: 'Image',
    name: 'HeaderLogo',
    source: 'Embedded',
    value: 'HeaderLogo',
    sizing: 'FitProportional',
    left: title.left + title.width,
    top: 10,
    width: 30,
    height: 30,
    zIndex: 1,
    hidden: 'false',
    style: {},
  };
  adjacent.body.items = [title, logo];

  const canonical = await renderPdf(adjacent, request, config, { captureLayoutTrace: true });
  const tracedTitle = canonical.layoutTrace.pages[0].items
    .find((item) => item.itemName === 'HeaderTitle');
  const tracedLogo = canonical.layoutTrace.pages[0].items
    .find((item) => item.itemName === 'HeaderLogo');
  assert.equal(tracedTitle.x + tracedTitle.width, 110.125);
  assert.equal(tracedLogo.x, 110.124);
  assert.equal(Math.round((tracedTitle.x + tracedTitle.width) / 0.25) * 0.25, 110.25);
  assert.equal(Math.round(tracedLogo.x / 0.25) * 0.25, 110);

  const rendered = await renderEditableDocx(adjacent, request, config);
  const documentXml = await (await JSZip.loadAsync(rendered.buffer))
    .file('word/document.xml').async('string');
  const nativeText = [...documentXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
    .map((match) => match[1])
    .join('');
  assert.match(nativeText, /Risk Register Report/);
  assert.match(documentXml, /<a:blip\b/);

  const declaredShallowOverlap = structuredClone(adjacent);
  declaredShallowOverlap.body.items.find((item) => item.name === 'HeaderLogo').left -= (0.01 / 2.54) * 72;
  const shallowCanonical = await renderPdf(
    declaredShallowOverlap,
    request,
    config,
    { captureLayoutTrace: true },
  );
  const shallowTitle = shallowCanonical.layoutTrace.pages[0].items
    .find((item) => item.itemName === 'HeaderTitle');
  const shallowLogo = shallowCanonical.layoutTrace.pages[0].items
    .find((item) => item.itemName === 'HeaderLogo');
  assert.ok(
    shallowTitle.x + shallowTitle.width - shallowLogo.x > 0.28,
    'the canonical PDF retains the explicitly declared 0.01cm image/title overlap',
  );
  const shallowRendered = await renderEditableDocx(declaredShallowOverlap, request, config);
  const shallowXml = await (await JSZip.loadAsync(shallowRendered.buffer))
    .file('word/document.xml').async('string');
  const shallowNativeText = [...shallowXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
    .map((match) => match[1])
    .join('');
  assert.match(shallowNativeText, /Risk Register Report/);
  assert.match(shallowXml, /<a:blip\b/);

  const genuineOverlap = structuredClone(adjacent);
  genuineOverlap.body.items.find((item) => item.name === 'HeaderLogo').left -= 2;
  await assert.rejects(
    renderEditableDocx(genuineOverlap, request, config),
    (error) => error.code === 'UNSUPPORTED_FEATURE'
      && /Overlapping editable PDF regions/.test(error.message)
      && error.details?.first === 'HeaderTitle'
      && error.details?.second === 'HeaderLogo',
  );
});

test('page-locked DOCX coalesces transparent horizontal neighbours when their smaller overlap is vertical', async () => {
  const adjacent = structuredClone(baseModel);
  adjacent.page.header = null;
  adjacent.page.footer = null;
  adjacent.page.marginLeft = 0;
  adjacent.page.marginRight = 0;
  const source = structuredClone(baseModel.body.items.find((item) => item.type === 'Textbox'));
  const textBox = (name, value, left) => ({
    ...structuredClone(source),
    name,
    value,
    paragraphs: [[value]],
    left,
    top: 0,
    width: 80,
    height: 18,
    canGrow: false,
    style: {
      ...structuredClone(source.style),
      backgroundColor: null,
      border: { style: 'None', width: 0, color: '#000000' },
      borders: Object.fromEntries(
        ['top', 'right', 'bottom', 'left']
          .map((side) => [side, { style: 'None', width: 0, color: '#000000' }]),
      ),
    },
  });
  // The 20pt horizontal overlap is visually empty. Its 18pt vertical overlap is smaller, so the
  // previous one-axis heuristic tried an impossible vertical split and rejected the report.
  adjacent.body.items = [textBox('LeftLabel', 'A', 0), textBox('RightValue', 'B', 60)];

  const rendered = await renderEditableDocx(adjacent, request, config);
  const documentXml = await (await JSZip.loadAsync(rendered.buffer))
    .file('word/document.xml').async('string');
  assert.match(documentXml, />A<\/w:t>/);
  assert.match(documentXml, />B<\/w:t>/);
});

test('page-locked DOCX coalesces a centered icon into its declared right padding beside a label', async () => {
  const iconFixture = Buffer.from(fixture.toString('utf8').replace(
    '          <Tablix Name="SalesTable">',
    `          <Textbox Name="CenteredTrendIcon">
            <CanGrow>false</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun><Value>V</Value><Style><FontSize>20pt</FontSize><FontWeight>Bold</FontWeight></Style></TextRun></TextRuns><Style><TextAlign>Center</TextAlign></Style></Paragraph></Paragraphs>
            <Top>1.3in</Top><Left>5in</Left><Height>0.4774in</Height><Width>0.3063in</Width>
            <Style><Border><Style>None</Style></Border><VerticalAlign>Top</VerticalAlign><PaddingLeft>2pt</PaddingLeft><PaddingRight>2pt</PaddingRight><PaddingTop>2pt</PaddingTop><PaddingBottom>2pt</PaddingBottom></Style>
          </Textbox>
          <Textbox Name="TrendLabel">
            <CanGrow>false</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun><Value>Downward Trend</Value></TextRun></TextRuns><Style><TextAlign>Left</TextAlign></Style></Paragraph></Paragraphs>
            <Top>1.3in</Top><Left>5.278in</Left><Height>0.4774in</Height><Width>1.142in</Width>
            <Style><Border><Style>None</Style></Border><VerticalAlign>Middle</VerticalAlign><PaddingLeft>2pt</PaddingLeft><PaddingRight>2pt</PaddingRight><PaddingTop>2pt</PaddingTop><PaddingBottom>2pt</PaddingBottom></Style>
          </Textbox>
          <Tablix Name="SalesTable">`,
  ));
  // This is the generic construct in Risk Trend Report: a centered icon overlaps the transparent
  // left padding of its left-aligned label. The icon's right padding preserves its text rectangle.
  const adjacent = parseRdl(iconFixture);

  const rendered = await renderEditableDocx(adjacent, request, config);
  const documentXml = await (await JSZip.loadAsync(rendered.buffer))
    .file('word/document.xml').async('string');
  assert.match(documentXml, />V<\/w:t>/);
  assert.match(documentXml, /Downward[\s\S]*Trend/);
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

test('exact Word rows preserve bottom padding as trailing content space without Word height inflation', async () => {
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

  // Word adds the largest tcMar/bottom value to hRule="exact". Keep the canonical 20pt/400-twip row
  // untouched, emit zero bottom cell margin, and preserve the declared 2pt/40-twip padding as trailing
  // paragraph space inside that exact box.
  assert.match(documentXml, /<w:trHeight w:val="400" w:hRule="exact"\/>/);
  assert.match(
    documentXml,
    /<w:tcMar>[\s\S]*?<w:top w:type="dxa" w:w="100"\/>[\s\S]*?<w:bottom w:type="dxa" w:w="0"\/>[\s\S]*?<\/w:tcMar>/,
  );
  assert.match(documentXml, /<w:spacing\b(?=[^>]*w:after="40")[^>]*\/>/);

  const noBottomPadding = structuredClone(padded);
  noBottomPadding.body.items[0].style.paddingBottom = 0;
  const unadjusted = await renderEditableDocx(noBottomPadding, request, config);
  const unadjustedZip = await JSZip.loadAsync(unadjusted.buffer);
  const unadjustedXml = await unadjustedZip.file('word/document.xml').async('string');
  assert.match(unadjustedXml, /<w:trHeight w:val="400" w:hRule="exact"\/>/);

  const splitGrid = structuredClone(padded);
  const spanning = structuredClone(splitGrid.body.items[0]);
  spanning.top = 0;
  const peer = {
    ...structuredClone(spanning),
    name: 'OffsetPeerCreatesTinyFirstGridRow',
    value: 'Peer',
    paragraphs: [['Peer']],
    left: 100,
    top: 1.5,
    width: 100,
    height: 10,
    style: {
      ...splitGrid.body.items[0].style,
      paddingBottom: 0,
    },
  };
  splitGrid.body.items = [{
    type: 'Rectangle',
    name: 'SplitGridContainer',
    left: 0,
    top: 0,
    width: 200,
    height: 20,
    style: {},
    items: [spanning, peer],
  }];
  const splitZip = await JSZip.loadAsync((await renderEditableDocx(splitGrid, request, config)).buffer);
  const splitXml = await splitZip.file('word/document.xml').async('string');
  assert.match(
    splitXml,
    /<w:trHeight w:val="30" w:hRule="exact"\/>/,
    'a spanning padded cell must not make its 1.5pt first trace-grid row unrepresentable',
  );
});

test('standalone page-band lines are traced and materialized as native Word borders', async () => {
  const lined = structuredClone(baseModel);
  lined.page.footer = {
    height: 24,
    printOnFirstPage: true,
    printOnLastPage: true,
    items: [{
      type: 'Line',
      name: 'FooterDivider',
      left: 0,
      top: 2,
      width: lined.page.width - lined.page.marginLeft - lined.page.marginRight,
      height: 0,
      style: {
        border: {
          color: '#123456',
          width: 1.5,
        },
      },
    }],
  };

  const canonical = await renderPdf(lined, request, config, { captureLayoutTrace: true });
  const traced = canonical.layoutTrace.pages[0].items.find((item) => item.itemName === 'FooterDivider');
  assert.deepEqual(traced.line, { style: 'Solid', width: 1.5, color: '#123456' });

  const zip = await JSZip.loadAsync((await renderEditableDocx(lined, request, config)).buffer);
  const footerXml = await zip.file('word/footer1.xml').async('string');
  const dividerBorders = footerXml.match(
    /<w:(?:top|bottom) w:val="single" w:color="123456" w:sz="12"\/>/g,
  ) || [];
  assert.ok(dividerBorders.length > 0, 'the standalone line must become a native Word border');
  assert.doesNotMatch(
    footerXml,
    /<w:top w:val="single" w:color="123456" w:sz="12"\/>/,
    'an interior horizontal line must have one owner instead of competing top and bottom borders',
  );
});

test('footer divider within tolerance of the next content edge avoids an empty sub-point Word row', async () => {
  const lined = structuredClone(baseModel);
  lined.page.footer = {
    height: 24,
    printOnFirstPage: true,
    printOnLastPage: true,
    items: [{
      type: 'Line',
      name: 'FooterDivider',
      left: 0,
      top: 3.75,
      width: lined.page.width - lined.page.marginLeft - lined.page.marginRight,
      height: 0,
      style: { border: { style: 'Solid', color: '#000000', width: 1 } },
    }, {
      type: 'Textbox',
      name: 'FooterText',
      left: 0,
      top: 4,
      width: 120,
      height: 10,
      value: 'Footer text',
      style: {},
    }],
  };

  const zip = await JSZip.loadAsync((await renderEditableDocx(lined, request, config)).buffer);
  const footerXml = await zip.file('word/footer1.xml').async('string');
  const bottomBorders = footerXml.match(
    /<w:bottom w:val="single" w:color="000000" w:sz="8"\/>/g,
  ) || [];
  const topBorders = footerXml.match(
    /<w:top w:val="single" w:color="000000" w:sz="8"\/>/g,
  ) || [];
  assert.ok(bottomBorders.length > 0, 'the row above must own the footer divider');
  assert.ok(topBorders.length > 0, 'the first content row must retain the same shared divider edge in Word');
  assert.doesNotMatch(footerXml, /<w:trHeight w:val="5" w:hRule="exact"\/>/);
});

test('page-locked DOCX accepts line endpoints that meet editable content without crossing it', async () => {
  const endpointFixture = Buffer.from(fixture.toString('utf8').replace(
    '          <Tablix Name="SalesTable">',
    `          <Line Name="EndpointLine">
            <Top>0.4in</Top><Left>0.1in</Left><Height>0.2in</Height><Width>0in</Width>
            <Style><Border><Style>Solid</Style><Width>1pt</Width><Color>#123456</Color></Border></Style>
          </Line>
          <Tablix Name="SalesTable">`,
  ));
  const endpointModel = parseRdl(endpointFixture);
  const canonical = await renderPdf(endpointModel, request, config, { captureLayoutTrace: true });
  const title = canonical.layoutTrace.pages[0].items.find((item) => item.itemName === 'TitleBox');
  const endpoint = canonical.layoutTrace.pages[0].items.find((item) => item.itemName === 'EndpointLine');
  assert.ok(Math.abs(endpoint.y - (title.y + title.height)) <= 0.001,
    'the line must begin exactly at the textbox endpoint within trace precision');
  const rendered = await renderEditableDocx(endpointModel, request, config);
  const zip = await JSZip.loadAsync(rendered.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  assert.match(documentXml, /<w:(?:left|right) w:val="single" w:color="123456" w:sz="8"\/>/);
});

test('page-locked DOCX coalesces a trace-rounded tablix fragment closure with its cell edge', async () => {
  const fragmentFixture = Buffer.from(fixture.toString('utf8')
    .replace(
      '<Top>0.1in</Top><Left>0.1in</Left><Height>0.3in</Height>',
      '<Top>0in</Top><Left>0.1in</Left><Height>0.3in</Height>',
    )
    .replace(
      '<DataSetName>Sales</DataSetName><Top>0.6in</Top>',
      '<DataSetName>Sales</DataSetName><Top>0.382274in</Top>',
    )
    .replace(
      '<TablixRow><Height>0.25in</Height>',
      '<TablixRow><Height>0.2361182in</Height>',
    ));
  const fragmentModel = parseRdl(fragmentFixture);
  const fragmentRequest = {
    ...request,
    outputFileName: 'trace-rounded-fragment-closure',
    datasets: { Sales: [{ Name: 'North', Amount: 1 }] },
  };
  const canonical = await renderPdf(fragmentModel, fragmentRequest, config, { captureLayoutTrace: true });
  const page = canonical.layoutTrace.pages[0];
  const cell = page.items.find((item) => item.kind === 'tablixCell'
    && item.tablixName === 'SalesTable'
    && item.rowIndex === 1);
  const closure = page.items.find((item) => item.traceRole === 'resolvedTablixFragmentBorder'
    && item.tablixName === 'SalesTable'
    && item.fragmentSide === 'bottom');
  assert.ok(cell);
  assert.ok(closure);
  assert.equal(cell.y + cell.height, 102.125);
  assert.equal(closure.y, 102.124);
  assert.equal(Math.round((cell.y + cell.height) / 0.25) * 0.25, 102.25);
  assert.equal(Math.round(closure.y / 0.25) * 0.25, 102);

  const rendered = await renderEditableDocx(fragmentModel, fragmentRequest, config);
  const documentXml = await (await JSZip.loadAsync(rendered.buffer))
    .file('word/document.xml').async('string');
  assert.equal(rendered.pageCount, canonical.pageCount);
  assert.match(documentXml, /<w:t(?:\s[^>]*)?>North<\/w:t>/);
  assert.match(documentXml, /<w:bottom w:val="single" w:color="000000" w:sz="8"\/>/);
});

test('page-locked DOCX keeps a tablix closing border on the body/footer boundary', async () => {
  const boundaryModel = structuredClone(baseModel);
  const tablix = boundaryModel.body.items.find((item) => item.type === 'Tablix');
  const boundaryRequest = {
    ...request,
    outputFileName: 'body-footer-tablix-closure',
    datasets: {
      Sales: Array.from({ length: 100 }, (_, index) => ({ Name: `Boundary row ${index + 1}`, Amount: index + 1 })),
    },
  };
  const initial = await renderPdf(boundaryModel, boundaryRequest, config, { captureLayoutTrace: true });
  const initialClosure = initial.layoutTrace.pages[0].items.find((item) => (
    item.traceRole === 'resolvedTablixFragmentBorder'
    && item.tablixName === tablix.name
    && item.fragmentSide === 'bottom'
  ));
  assert.ok(initialClosure, 'the fixture must draw a first-page tablix closure');
  tablix.top += initial.layoutTrace.pages[0].bodyBottom - initialClosure.y;
  const canonical = await renderPdf(boundaryModel, boundaryRequest, config, { captureLayoutTrace: true });
  const closure = canonical.layoutTrace.pages[0].items.find((item) => (
    item.traceRole === 'resolvedTablixFragmentBorder'
    && item.tablixName === tablix.name
    && item.fragmentSide === 'bottom'
  ));
  assert.ok(closure);
  assert.equal(canonical.pageCount > 1, true, 'the growable detail row must split across pages');
  assert.equal(closure.region, 'footer', 'the physical boundary is classified with the footer region');
  assert.ok(Math.abs(closure.y - canonical.layoutTrace.pages[0].bodyBottom) <= 0.5);

  const rendered = await renderEditableDocx(boundaryModel, boundaryRequest, config);
  const documentXml = await (await JSZip.loadAsync(rendered.buffer))
    .file('word/document.xml').async('string');
  const bodyTable = [...documentXml.matchAll(/<w:tbl>([\s\S]*?)<\/w:tbl>/g)][0]?.[1] || '';
  const finalRow = [...bodyTable.matchAll(/<w:tr>([\s\S]*?)<\/w:tr>/g)].at(-1)?.[1] || '';
  assert.match(finalRow, /<w:bottom w:val="single" w:color="000000" w:sz="8"\/>/,
    'the canonical closure must belong to the last native body-table row, not the footer story');
});

test('page-locked DOCX still rejects a line that penetrates editable content', async () => {
  const crossingFixture = Buffer.from(fixture.toString('utf8').replace(
    '          <Tablix Name="SalesTable">',
    `          <Line Name="CrossingLine">
            <Top>0.1in</Top><Left>1in</Left><Height>0.2in</Height><Width>0in</Width>
            <Style><Border><Style>Solid</Style><Width>1pt</Width><Color>#123456</Color></Border></Style>
          </Line>
          <Tablix Name="SalesTable">`,
  ));
  const crossingModel = parseRdl(crossingFixture);
  await assert.rejects(
    renderEditableDocx(crossingModel, request, config),
    (error) => error.code === 'UNSUPPORTED_FEATURE'
      && /line crosses editable content/.test(error.message)
      && error.details?.line === 'CrossingLine'
      && error.details?.item === 'TitleBox',
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
