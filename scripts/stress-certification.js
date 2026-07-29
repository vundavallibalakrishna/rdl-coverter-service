import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { promisify } from 'node:util';
import { PDFDocument } from 'pdf-lib';
import JSZip from 'jszip';
import { PNG } from 'pngjs';
import { loadConfig } from '../src/config.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderPdf } from '../src/render/pdf.js';
import { createNestedStressScenario } from './lib/nestedStressScenario.js';
import { STRESS_OVERFLOW_LINES, STRESS_ROW_COUNT, createStressScenario } from './lib/stressScenario.js';

const execFileAsync = promisify(execFile);
const serviceRoot = path.resolve(new URL('..', import.meta.url).pathname);
const nestedMode = process.argv.includes('--nested');
const scenario = nestedMode ? createNestedStressScenario() : createStressScenario();
const { model, request } = scenario;
const certification = scenario.certification || {
  id: 'rdl-table-stress',
  targetPages: { minimum: 30, maximum: 40 },
  exactMarkers: Array.from({ length: STRESS_ROW_COUNT }, (_, index) => `UNIQUE_ROW_${String(index + 1).padStart(4, '0')}`),
  overflowSpecs: [{
    id: 'default', lines: STRESS_OVERFLOW_LINES, start: 'GIANT_CELL_START', end: 'GIANT_CELL_END', linePrefix: 'GIANT_LINE_',
  }],
  endMarker: 'STRESS_DOCUMENT_END',
  expectedHeaderRows: 3,
  expectedFonts: model.fonts,
  expectedRowCount: STRESS_ROW_COUNT,
};
// Repository QA artifacts must be direct children of tmp/. Runtime-only conversion and raster files use
// the security-controlled temporary directory below and are completely removed in finally.
const outputRoot = path.join(serviceRoot, 'tmp');
const reportPath = path.join(outputRoot, `${certification.id}-certification.json`);
const directPdfPath = path.join(outputRoot, `${certification.id}-direct.pdf`);
const editableDocxPath = path.join(outputRoot, `${certification.id}-editable.docx`);
const editablePdfPath = path.join(outputRoot, `${certification.id}-editable-rendered.pdf`);
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-stress-certification-'));

function count(text, pattern) {
  return (text.match(pattern) || []).length;
}

function visibleBottomBorder(xml, container) {
  const section = xml.match(new RegExp(`<w:${container}>[\\s\\S]*?<\\/w:${container}>`))?.[0] || '';
  return /<w:bottom\b[^>]*w:val="(?!none\b|nil\b)[^"]+"[^>]*\/>/i.test(section);
}

function visibleTopBorder(xml, container) {
  const section = xml.match(new RegExp(`<w:${container}>[\\s\\S]*?<\\/w:${container}>`))?.[0] || '';
  return /<w:top\b[^>]*w:val="(?!none\b|nil\b)[^"]+"[^>]*\/>/i.test(section);
}

function tableBorderAudit(documentXml) {
  const tables = [...documentXml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
  const fragments = tables.map((table, index) => {
    const rows = [...table.matchAll(/<w:tr>[\s\S]*?<\/w:tr>/g)].map((match) => match[0]);
    const lastRow = rows.at(-1) || '';
    const cells = [...lastRow.matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map((match) => match[0]);
    const openCells = cells.map((cell, cellIndex) => ({
      cellIndex,
      closed: visibleBottomBorder(cell, 'tcBorders') || visibleTopBorder(cell, 'tcBorders'),
    }))
      .filter(({ closed }) => !closed).map(({ cellIndex }) => cellIndex);
    const tableBottomClosed = visibleBottomBorder(table, 'tblBorders');
    const closureTopClosed = cells.length > 0 && cells.every((cell) => visibleTopBorder(cell, 'tcBorders'));
    return {
      fragment: index + 1,
      rows: rows.length,
      lastRowCells: cells.length,
      tableBottomClosed,
      closureTopClosed,
      openLastRowCells: openCells,
      passed: (tableBottomClosed || closureTopClosed) && cells.length > 0 && openCells.length === 0,
    };
  });
  return { fragments, passed: fragments.length > 0 && fragments.every((fragment) => fragment.passed) };
}

function tableGridAudit(documentXml) {
  const tables = [...documentXml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
  const invalidRows = [];
  let rowCount = 0;
  tables.forEach((table, tableIndex) => {
    const columnCount = count(table.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/)?.[0] || '', /<w:gridCol\b/g);
    const rows = [...table.matchAll(/<w:tr>[\s\S]*?<\/w:tr>/g)].map((match) => match[0]);
    rows.forEach((row, rowIndex) => {
      rowCount += 1;
      const occupiedColumns = [...row.matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)]
        .reduce((sum, match) => sum + Number(match[0].match(/<w:gridSpan w:val="(\d+)"\/>/)?.[1] || 1), 0);
      if (occupiedColumns !== columnCount) invalidRows.push({
        table: tableIndex + 1, row: rowIndex + 1, occupiedColumns, declaredColumns: columnCount,
      });
    });
  });
  return { tables: tables.length, rows: rowCount, invalidRows, passed: tables.length > 0 && invalidRows.length === 0 };
}

function fontGroundingAudit(parts, expectedFamilies) {
  const runs = [];
  for (const [part, xml] of parts) {
    for (const match of xml.matchAll(/<w:(?:t|instrText)\b[^>]*>[\s\S]*?<\/w:(?:t|instrText)>/g)) {
      // Empty structural cells and exact-height closure rows may serialize an empty w:t. They have no
      // displayed glyph whose font could be substituted, so font grounding applies only to visible text
      // and field instructions.
      const displayed = match[0].replace(/<[^>]+>/g, '').replace(/&(?:#\d+|#x[\da-f]+|\w+);/gi, ' ').trim();
      if (!displayed) continue;
      const start = xml.lastIndexOf('<w:r>', match.index);
      const end = xml.indexOf('</w:r>', match.index) + '</w:r>'.length;
      const run = start >= 0 && end > start ? xml.slice(start, end) : '';
      const family = run.match(/<w:rFonts\b[^>]*w:ascii="([^"]+)"/)?.[1] || null;
      runs.push({ part, family, grounded: Boolean(family) });
    }
  }
  const families = [...new Set(runs.map((run) => run.family).filter(Boolean))].sort();
  const missingFamilies = expectedFamilies.filter((family) => !families.includes(family));
  const ungroundedRuns = runs.filter((run) => !run.grounded);
  return {
    textRuns: runs.length,
    groundedRuns: runs.length - ungroundedRuns.length,
    ungroundedRuns: ungroundedRuns.length,
    families,
    missingFamilies,
    passed: runs.length > 0 && ungroundedRuns.length === 0 && missingFamilies.length === 0,
  };
}

function isDark(png, x, y) {
  if (x < 0 || y < 0 || x >= png.width || y >= png.height) return false;
  const offset = (y * png.width + x) * 4;
  return png.data[offset + 3] > 0 && (png.data[offset] + png.data[offset + 1] + png.data[offset + 2]) / 3 < 205;
}

function hasDarkNear(png, x, y, radius = 3) {
  for (let dx = -radius; dx <= radius; dx += 1) if (isDark(png, x + dx, y)) return true;
  return false;
}

function horizontalCoverage(png, left, right, y) {
  let covered = 0;
  for (let x = left; x <= right; x += 1) {
    let dark = false;
    for (let dy = -2; dy <= 2 && !dark; dy += 1) dark = isDark(png, x, y + dy);
    if (dark) covered += 1;
  }
  return covered / Math.max(1, right - left + 1);
}

async function rasterBorderAudit(pdfPath, label, tablePages, model) {
  const rasterRoot = path.join(tempDir, `raster-${label}`);
  await fs.mkdir(rasterRoot);
  const prefix = path.join(rasterRoot, 'page');
  await execFileAsync(config.pdftoppmPath, ['-png', '-r', '144', pdfPath, prefix], { maxBuffer: 64 * 1024 * 1024 });
  const files = (await fs.readdir(rasterRoot)).filter((name) => name.endsWith('.png')).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  const scale = 2;
  const table = model.body.items.find((item) => item.type === 'Tablix');
  const left = Math.round((model.page.marginLeft + table.left) * scale);
  const right = Math.round((model.page.marginLeft + table.left + table.width) * scale);
  const top = Math.round((model.page.marginTop + (model.page.header?.height || 0) + table.top) * scale);
  const bottomLimit = Math.round((model.page.height - model.page.marginBottom - (model.page.footer?.height || 0)) * scale);
  const pages = [];
  for (let index = 0; index < Math.min(tablePages, files.length); index += 1) {
    const png = PNG.sync.read(await fs.readFile(path.join(rasterRoot, files[index])));
    let bottomY = -1;
    for (let y = top; y < Math.min(bottomLimit, png.height); y += 1) {
      if (hasDarkNear(png, left, y) && hasDarkNear(png, right, y)) bottomY = y;
    }
    const coverage = bottomY >= 0 ? horizontalCoverage(png, left, right, bottomY) : 0;
    pages.push({ page: index + 1, bottomY, horizontalCoverage: Number(coverage.toFixed(4)), passed: bottomY >= 0 && coverage >= 0.9 });
  }
  return { pages, passed: pages.length === tablePages && pages.every((page) => page.passed) };
}

function markerAudit(text, markers) {
  const missing = [];
  const duplicated = [];
  for (const marker of markers) {
    const occurrences = count(text, new RegExp(marker, 'g'));
    if (occurrences === 0) missing.push(marker);
    else if (occurrences !== 1) duplicated.push({ marker, occurrences });
  }
  return { missing, duplicated, passed: missing.length === 0 && duplicated.length === 0 };
}

function overflowAudit(text, specs) {
  const missing = [];
  const duplicated = [];
  const cells = [];
  for (const spec of specs) {
    for (let index = 1; index <= spec.lines; index += 1) {
      const marker = `${spec.linePrefix}${String(index).padStart(3, '0')}`;
      const occurrences = count(text, new RegExp(marker, 'g'));
      if (occurrences === 0) missing.push(marker);
      else if (occurrences !== 1) duplicated.push({ marker, occurrences });
    }
    cells.push({
      id: spec.id,
      startCount: count(text, new RegExp(spec.start, 'g')),
      endCount: count(text, new RegExp(spec.end, 'g')),
    });
  }
  return {
    cells,
    missing,
    duplicated,
    passed: missing.length === 0 && duplicated.length === 0
      && cells.every((cell) => cell.startCount === 1 && cell.endCount === 1),
  };
}

function wordRowSafetyAudit(documentXml, overflowSpecs) {
  const rows = [...documentXml.matchAll(/<w:tr>[\s\S]*?<\/w:tr>/g)].map((match) => match[0]);
  const oversized = overflowSpecs.map((spec) => {
    const sourceRows = rows.filter((candidate) => candidate.includes(spec.start)
      || candidate.includes(spec.end)
      || candidate.includes(`${spec.linePrefix}${String(1).padStart(3, '0')}`));
    const startRow = rows.find((candidate) => candidate.includes(spec.start)) || '';
    const endRow = rows.find((candidate) => candidate.includes(spec.end)) || '';
    return {
      id: spec.id,
      found: Boolean(startRow) && Boolean(endRow),
      splitAcrossNativeRows: startRow !== endRow,
      sourceRows: sourceRows.length,
      allAtomic: sourceRows.every((row) => /<w:cantSplit\/>/.test(row)),
      allBounded: sourceRows.every((row) => /<w:trHeight\b/.test(row)),
    };
  });
  const rowsWithoutCantSplit = rows.filter((row) => !/<w:cantSplit\/>/.test(row)).length;
  return {
    oversized,
    rowsWithoutCantSplit,
    passed: oversized.every((entry) => entry.found && entry.splitAcrossNativeRows
      && entry.sourceRows > 1 && entry.allAtomic && entry.allBounded)
      && rowsWithoutCantSplit === 0,
  };
}

async function textFromPdf(pdfPath, outputName) {
  const outputPath = path.join(tempDir, outputName);
  await execFileAsync('pdftotext', ['-layout', pdfPath, outputPath]);
  return fs.readFile(outputPath, 'utf8');
}

async function convertDocxToPdf(docxPath, pdfPath, label) {
  const profile = path.join(tempDir, `libreoffice-profile-${label}`);
  const conversionDir = path.join(tempDir, `converted-${label}`);
  await Promise.all([fs.mkdir(profile), fs.mkdir(conversionDir)]);
  await execFileAsync(process.env.RDL_SOFFICE_PATH || 'soffice', [
    `-env:UserInstallation=file://${profile}`,
    '--invisible', '--headless', '--norestore', '--convert-to', 'pdf', '--outdir', conversionDir, docxPath,
  ], { maxBuffer: 16 * 1024 * 1024 });
  const converted = path.join(conversionDir, `${path.basename(docxPath, '.docx')}.pdf`);
  await fs.copyFile(converted, pdfPath);
}

try {
  await fs.mkdir(outputRoot, { recursive: true });
  const [directPdf, editableDocx] = await Promise.all([
    renderPdf(model, request, config),
    renderEditableDocx(model, request),
  ]);
  await Promise.all([
    fs.writeFile(directPdfPath, directPdf.buffer),
    fs.writeFile(editableDocxPath, editableDocx.buffer),
  ]);
  await convertDocxToPdf(editableDocxPath, editablePdfPath, 'structured');

  const [directPdfDocument, editablePdfDocument, directText, editableText, zip] = await Promise.all([
    PDFDocument.load(await fs.readFile(directPdfPath)),
    PDFDocument.load(await fs.readFile(editablePdfPath)),
    textFromPdf(directPdfPath, 'direct.txt'),
    textFromPdf(editablePdfPath, 'editable.txt'),
    JSZip.loadAsync(await fs.readFile(editableDocxPath)),
  ]);
  const documentXml = await zip.file('word/document.xml').async('string');
  const textPartNames = Object.keys(zip.files).filter((name) => /^word\/(?:document|header\d+|footer\d+)\.xml$/.test(name));
  const textParts = await Promise.all(textPartNames.map(async (name) => [name, await zip.file(name).async('string')]));
  const relationshipPartNames = Object.keys(zip.files).filter((name) => name.endsWith('.rels'));
  const relationshipParts = await Promise.all(relationshipPartNames.map((name) => zip.file(name).async('string')));
  const directPageCount = directPdfDocument.getPageCount();
  const editablePageCount = editablePdfDocument.getPageCount();
  const openXmlBorders = tableBorderAudit(documentXml);
  const openXmlGrid = tableGridAudit(documentXml);
  const fontGrounding = fontGroundingAudit(textParts, certification.expectedFonts);
  const rowSafety = wordRowSafetyAudit(documentXml, certification.overflowSpecs);
  const externalRelationships = relationshipParts.reduce((sum, xml) => sum + count(xml, /TargetMode="External"/g), 0);
  const directRasterBorders = await rasterBorderAudit(directPdfPath, 'direct', directPageCount - 1, model);
  // The exceptional single row taller than a page is allowed to occupy multiple physical pages. Audit all
  // table-bearing pages, not merely one page per explicit fragment, so those implicit pages cannot escape
  // the visible-border gate.
  const editableRasterBorders = await rasterBorderAudit(editablePdfPath, 'editable', editablePageCount - 1, model);
  const report = {
    passed: false,
    targetPages: { ...certification.targetPages, structuredDocxAdvisoryOnly: true },
    inputs: {
      rows: certification.expectedRowCount,
      giantOverflowCells: certification.overflowSpecs.map(({ id, lines }) => ({ id, lines })),
      hierarchyDepth: nestedMode ? 5 : 2,
    },
    artifacts: { directPdfPath, editableDocxPath, editablePdfPath },
    directPdf: {
      pages: directPageCount,
      pageRangePassed: directPageCount >= certification.targetPages.minimum && directPageCount <= certification.targetPages.maximum,
      rowMarkers: markerAudit(directText, certification.exactMarkers),
      overflow: overflowAudit(directText, certification.overflowSpecs),
      endMarkerCount: count(directText, new RegExp(certification.endMarker, 'g')),
      rasterBottomBorders: directRasterBorders,
    },
    editableDocx: {
      renderedPages: editablePageCount,
      pageRangeAdvisoryPassed: editablePageCount >= certification.targetPages.minimum && editablePageCount <= certification.targetPages.maximum,
      rowMarkers: markerAudit(editableText, certification.exactMarkers),
      overflow: overflowAudit(editableText, certification.overflowSpecs),
      endMarkerCount: count(editableText, new RegExp(certification.endMarker, 'g')),
      rasterBottomBorders: editableRasterBorders,
      openXml: {
        nativeTables: count(documentXml, /<w:tbl>/g),
        repeatingHeaderRows: count(documentXml, /<w:tblHeader\/>/g),
        atomicHeaderRows: count(documentXml, /<w:cantSplit\/>/g),
        horizontalMergedHeaderCells: count(documentXml, /<w:gridSpan w:val="2"\/>/g),
        mergedCellSpans: count(documentXml, /<w:gridSpan w:val="(?:[2-9]|\d{2,})"\/>/g),
        verticalMergedHeaderStarts: count(documentXml, /<w:vMerge w:val="restart"\/>/g),
        verticalMergedHeaderContinuations: count(documentXml, /<w:vMerge w:val="continue"\/>/g),
        bottomBorders: openXmlBorders,
        gridOccupancy: openXmlGrid,
        fontGrounding,
        rowSafety,
        externalRelationships,
      },
    },
    combinations: {
      repeatedThreeRowHeader: certification.expectedHeaderRows === 3,
      horizontalMergedHeaderCells: true,
      verticalMergedHeaderCells: true,
      nestedGroupAndSubgroupSorting: true,
      hierarchyDepth: nestedMode ? 5 : 2,
      groupHeadersAndFooters: nestedMode,
      multipleOversizedCells: certification.overflowSpecs.length > 1,
      conditionalVisibility: nestedMode,
      conditionalCellBackgrounds: true,
      perSideAndDashedBorders: true,
      deliberatelyMissingSourceBottomBorders: true,
      fontFamilies: model.fonts,
      boldItalicUnderlineAndColor: true,
      wrappedMultilineCells: true,
      cellTallerThanPage: true,
      naturalPagination: true,
      explicitFinalPageBreak: true,
      uniqueRowTraceability: true,
    },
  };
  report.passed = report.directPdf.pageRangePassed
    && report.directPdf.rowMarkers.passed
    && report.directPdf.overflow.passed
    && report.directPdf.endMarkerCount === 1
    && report.directPdf.rasterBottomBorders.passed
    && report.editableDocx.rowMarkers.passed
    && report.editableDocx.overflow.passed
    && report.editableDocx.endMarkerCount === 1
    && report.editableDocx.rasterBottomBorders.passed
    && report.editableDocx.openXml.nativeTables === report.editableDocx.openXml.bottomBorders.fragments.length
    && report.editableDocx.openXml.repeatingHeaderRows === report.editableDocx.openXml.nativeTables * certification.expectedHeaderRows
    && report.editableDocx.openXml.atomicHeaderRows >= report.editableDocx.openXml.nativeTables * certification.expectedHeaderRows
    && (nestedMode
      ? report.editableDocx.openXml.mergedCellSpans >= report.editableDocx.openXml.nativeTables * 4
        && report.editableDocx.openXml.verticalMergedHeaderStarts > 0
        && report.editableDocx.openXml.verticalMergedHeaderContinuations > 0
      : report.editableDocx.openXml.horizontalMergedHeaderCells === report.editableDocx.openXml.nativeTables * 5
        && report.editableDocx.openXml.verticalMergedHeaderStarts === report.editableDocx.openXml.nativeTables * 6
        && report.editableDocx.openXml.verticalMergedHeaderContinuations === report.editableDocx.openXml.nativeTables * 6)
    && report.editableDocx.openXml.bottomBorders.passed
    && report.editableDocx.openXml.gridOccupancy.passed
    && report.editableDocx.openXml.fontGrounding.passed
    && report.editableDocx.openXml.rowSafety.passed
    && report.editableDocx.openXml.externalRelationships === 0;
  await fs.writeFile(reportPath, `${JSON.stringify(report, null, 2)}\n`);
  process.stdout.write(`${JSON.stringify({ reportPath, ...report }, null, 2)}\n`);
  if (process.argv.includes('--require-pass') && !report.passed) process.exitCode = 1;
} finally {
  await fs.rm(tempDir, { recursive: true, force: true });
}
