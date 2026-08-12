// A CanGrow textbox is a flowing block of text, not an atomic unit. When its resolved height exceeds the
// current page remainder it must fill that remainder and continue on the next page, instead of deferring
// the whole block to a fresh page and leaving a blank band its own height behind. The block only defers
// when the remainder cannot carry the typographic orphan minimum.
//
// These tests drive a synthetic minimal RDL. The construct — a heading textbox followed by a growable
// prose textbox, both declaring KeepTogether — is exercised at the top level of the body and nested in a
// Rectangle, so the rule is proven at the shared pagination layer rather than for one container.
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
import { renderPdf } from '../src/render/pdf.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';

const execFileAsync = promisify(execFile);
const fixture = await fs.readFile(new URL('./fixtures/basic.rdl', import.meta.url));
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '30000' });

// Explicit metrics keep the arithmetic below independent of the fixture's fonts: line height 12pt with no
// vertical padding puts the orphan minimum (2 lines) at exactly 24pt.
const LINE_HEIGHT_PT = 12;
const ORPHAN_MINIMUM_PT = 2 * LINE_HEIGHT_PT;
const HEADING_HEIGHT_PT = 14;
const PAGE_HEIGHT_PT = 320;
const MARGIN_PT = 20;
const BODY_TOP_PT = MARGIN_PT;
const BODY_BOTTOM_PT = PAGE_HEIGHT_PT - MARGIN_PT;

const itemMarker = (index) => `ITEM_${String(index).padStart(3, '0')}`;

function applyMetrics(textbox) {
  textbox.style = {
    ...textbox.style,
    fontSize: 10,
    lineHeight: LINE_HEIGHT_PT,
    paddingTop: 0,
    paddingBottom: 0,
    paddingLeft: 0,
    paddingRight: 0,
  };
  textbox.paragraphStyles = [{ ...textbox.style, spaceBefore: 0, spaceAfter: 0 }];
  return textbox;
}

function textboxFrom(template, { name, top, height, text, canGrow }) {
  const textbox = applyMetrics(structuredClone(template));
  Object.assign(textbox, {
    name,
    top,
    left: 0,
    width: 400,
    height,
    canGrow,
    // Report designers emit KeepTogether on essentially every textbox. It is a best-effort hint, so it
    // must not stop a growable block from flowing across a page boundary.
    keepTogether: true,
    value: text,
    paragraphs: [[text]],
  });
  return textbox;
}

// leadingHeight controls how much of the page is consumed before the section starts, which is what
// determines the remainder the growable block has to work with.
function sectionScenario({ leadingHeight, bodyLines, nested = false, bodyHeight = HEADING_HEIGHT_PT }) {
  const model = parseRdl(fixture);
  model.page.height = PAGE_HEIGHT_PT;
  model.page.marginTop = MARGIN_PT;
  model.page.marginBottom = MARGIN_PT;
  model.page.header = null;
  model.page.footer = null;

  const template = model.body.items.find((item) => item.type === 'Textbox');
  const leading = textboxFrom(template, {
    name: 'LeadingBlock',
    top: 0,
    height: leadingHeight,
    text: 'LEADING_BLOCK',
    canGrow: false,
  });
  const heading = textboxFrom(template, {
    name: 'SectionHeading',
    top: 0,
    height: HEADING_HEIGHT_PT,
    text: 'SECTION_HEADING',
    canGrow: true,
  });
  const body = textboxFrom(template, {
    name: 'SectionBody',
    top: HEADING_HEIGHT_PT,
    height: bodyHeight,
    text: Array.from({ length: bodyLines }, (unused, index) => itemMarker(index + 1)).join('\n'),
    canGrow: true,
  });

  if (nested) {
    const container = {
      type: 'Rectangle',
      name: 'SectionContainer',
      top: leadingHeight,
      left: 0,
      width: 400,
      height: HEADING_HEIGHT_PT * 2,
      zIndex: 0,
      keepTogether: true,
      hidden: 'false',
      style: structuredClone(model.body.items[0].style),
      items: [heading, body],
    };
    container.style.border = { style: 'None' };
    container.style.borders = {
      top: { style: 'None' }, right: { style: 'None' }, bottom: { style: 'None' }, left: { style: 'None' },
    };
    container.style.backgroundColor = null;
    model.body.items = [leading, container];
  } else {
    heading.top = leadingHeight;
    body.top = leadingHeight + HEADING_HEIGHT_PT;
    model.body.items = [leading, heading, body];
  }

  const request = {
    outputFileName: 'growable-textbox-flow',
    parameters: { Title: 'Flow', Choice: 'A' },
    datasets: { Sales: [{ Name: 'Row', Amount: 1 }] },
  };
  return { model, request };
}

async function renderPages(context, scenario, name) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-growable-flow-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, `${name}.pdf`);
  const result = await renderPdf(scenario.model, scenario.request, config);
  await fs.writeFile(pdfPath, result.buffer);
  const extracted = await execFileAsync('pdftotext', ['-layout', pdfPath, '-']);
  return { result, tempDir, pages: extracted.stdout.split('\f') };
}

function pageContaining(pages, marker) {
  return pages.findIndex((page) => page.includes(marker)) + 1;
}

function assertNoContentLost(pages, bodyLines) {
  const all = pages.join('\n');
  assert.equal(
    (all.match(/SECTION_HEADING/g) || []).length,
    1,
    'the section heading must never be repeated on a continuation page',
  );
  for (let index = 1; index <= bodyLines; index += 1) {
    assert.equal(
      (all.match(new RegExp(itemMarker(index), 'g')) || []).length,
      1,
      `${itemMarker(index)} must appear exactly once across the document`,
    );
  }
}

test('a growable textbox fills the current page remainder and continues on the next page', async (context) => {
  // Leading block 200pt: the heading ends at 234pt, leaving 66pt (5 lines) before the 300pt body bottom.
  // The block needs 20 lines = 240pt, which fits on an empty page — the case that used to defer wholesale.
  const bodyLines = 20;
  const scenario = sectionScenario({ leadingHeight: 200, bodyLines });
  const rendered = await renderPages(context, scenario, 'flow-top-level');

  const headingPage = pageContaining(rendered.pages, 'SECTION_HEADING');
  assert.equal(headingPage, 1);
  assert.equal(
    pageContaining(rendered.pages, itemMarker(1)),
    headingPage,
    'the block must start on the page that still has usable space',
  );
  assert.ok(
    pageContaining(rendered.pages, itemMarker(bodyLines)) > headingPage,
    'the overflow must continue on a later page',
  );
  assertNoContentLost(rendered.pages, bodyLines);
});

test('a growable textbox nested in a rectangle flows the same way', async (context) => {
  const bodyLines = 20;
  const scenario = sectionScenario({ leadingHeight: 200, bodyLines, nested: true });
  const rendered = await renderPages(context, scenario, 'flow-nested');

  const headingPage = pageContaining(rendered.pages, 'SECTION_HEADING');
  assert.equal(pageContaining(rendered.pages, itemMarker(1)), headingPage);
  assert.ok(pageContaining(rendered.pages, itemMarker(bodyLines)) > headingPage);
  assertNoContentLost(rendered.pages, bodyLines);
});

test('a growable textbox defers when the remainder cannot hold the orphan minimum', async (context) => {
  // Leading block 250pt: the heading ends at 284pt, leaving 16pt — less than the 24pt orphan minimum.
  const bodyLines = 20;
  const scenario = sectionScenario({ leadingHeight: 250, bodyLines });
  const rendered = await renderPages(context, scenario, 'flow-orphan');
  assert.ok(BODY_BOTTOM_PT - (250 + BODY_TOP_PT + HEADING_HEIGHT_PT) < ORPHAN_MINIMUM_PT);

  const headingPage = pageContaining(rendered.pages, 'SECTION_HEADING');
  assert.equal(headingPage, 1);
  assert.equal(
    pageContaining(rendered.pages, itemMarker(1)),
    headingPage + 1,
    'a remainder too small for the orphan minimum must not strand a single line',
  );
  assertNoContentLost(rendered.pages, bodyLines);
});

test('a growable textbox whose declared height alone overflows moves as a unit', async (context) => {
  // The heading ends at 234pt, leaving 66pt. The block declares 100pt but holds only 2 lines (24pt) of
  // text: the declared height is a reservation, not content, so there is nothing to flow and the block
  // must move whole rather than be squashed into the remainder.
  const bodyLines = 2;
  const scenario = sectionScenario({ leadingHeight: 200, bodyLines, bodyHeight: 100 });
  const rendered = await renderPages(context, scenario, 'flow-declared-height');

  const headingPage = pageContaining(rendered.pages, 'SECTION_HEADING');
  assert.equal(headingPage, 1);
  assert.equal(
    pageContaining(rendered.pages, itemMarker(1)),
    headingPage + 1,
    'a reserved-height block whose text fits must not be split into the remainder',
  );
  assert.equal(
    pageContaining(rendered.pages, itemMarker(bodyLines)),
    headingPage + 1,
    'the reserved block stays whole on its new page',
  );
  assertNoContentLost(rendered.pages, bodyLines);
});

test('a growable textbox taller than a whole page still splits across pages', async (context) => {
  const bodyLines = 80;
  const scenario = sectionScenario({ leadingHeight: 200, bodyLines });
  const rendered = await renderPages(context, scenario, 'flow-oversized');

  assert.ok(rendered.result.pageCount > 3);
  assert.ok(rendered.result.pageCount < 20);
  assertNoContentLost(rendered.pages, bodyLines);
});

// DOCX_EDITABLE builds one Word section per canonical PDF page from the layout trace, so it inherits the
// PDF's flow decisions. Assert that inheritance directly: the heading and the block's first line must land
// in the same Word section, and the overflow in a later one.
test('DOCX_EDITABLE inherits the flowed pagination from the canonical PDF trace', async (context) => {
  const bodyLines = 20;
  const scenario = sectionScenario({ leadingHeight: 200, bodyLines });
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-growable-flow-docx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));

  const pdf = await renderPdf(scenario.model, scenario.request, config);
  const docx = await renderEditableDocx(scenario.model, scenario.request, config, tempDir);
  assert.equal(docx.pageCount, pdf.pageCount);

  const zip = await JSZip.loadAsync(docx.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  assert.equal((documentXml.match(/<w:sectPr(?:\s|>)/g) || []).length, docx.pageCount);
  // Each section terminates its page, so splitting on the section break yields per-page content.
  const sections = documentXml.split(/<w:sectPr(?:\s|>)/);
  const sectionContaining = (marker) => sections.findIndex((section) => section.includes(marker));

  const headingSection = sectionContaining('SECTION_HEADING');
  assert.ok(headingSection >= 0);
  assert.equal(
    sectionContaining(itemMarker(1)),
    headingSection,
    'Word must start the block on the same page the PDF did',
  );
  assert.ok(
    sectionContaining(itemMarker(bodyLines)) > headingSection,
    'Word must continue the overflow on a later page',
  );
  assert.equal((documentXml.match(/SECTION_HEADING/g) || []).length, 1);
  assert.match(documentXml, /<w:t[^>]*>[^<]*ITEM_001/);
});

// XLSX is not page-paginated: a worksheet is one continuous grid, so a growable textbox becomes a single
// styled cell block and there is no page remainder to fill. The construct must therefore stay whole and
// intact, which is the format-appropriate equivalent of "nothing is lost at a boundary".
test('XLSX keeps the growable block intact because a worksheet has no page remainder', async (context) => {
  const bodyLines = 20;
  const scenario = sectionScenario({ leadingHeight: 200, bodyLines });
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-growable-flow-xlsx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));

  const workbook = await renderExcel(scenario.model, scenario.request, config, tempDir);
  const zip = await JSZip.loadAsync(workbook.buffer);
  const parts = zip.file(/^xl\/(worksheets\/sheet\d+|sharedStrings)\.xml$/);
  assert.ok(parts.length > 0, 'the workbook must contain worksheet and shared-string parts');
  const sheetXml = (await Promise.all(parts.map((part) => part.async('string')))).join('');

  assert.equal((sheetXml.match(/SECTION_HEADING/g) || []).length, 1);
  for (let index = 1; index <= bodyLines; index += 1) {
    assert.equal((sheetXml.match(new RegExp(itemMarker(index), 'g')) || []).length, 1);
  }
});
