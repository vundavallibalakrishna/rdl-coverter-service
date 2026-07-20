import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { promisify } from 'node:util';
import { PDFDocument } from 'pdf-lib';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderFixedEditableDocx } from '../src/render/fixedDocx.js';
import { renderPdf } from '../src/render/pdf.js';
import { STRESS_OVERFLOW_LINES, STRESS_ROW_COUNT, createStressScenario } from './lib/stressScenario.js';

const execFileAsync = promisify(execFile);
const serviceRoot = path.resolve(new URL('..', import.meta.url).pathname);
// Generated artifacts live under tmp/ so they stay out of version control (see .gitignore).
const outputRoot = path.join(serviceRoot, 'tmp', 'output');
const pdfOutputDir = path.join(outputRoot, 'pdf');
const docxOutputDir = path.join(outputRoot, 'docx');
const reportPath = path.join(outputRoot, 'stress-certification.json');
const directPdfPath = path.join(pdfOutputDir, 'rdl-table-stress-direct.pdf');
const editableDocxPath = path.join(docxOutputDir, 'rdl-table-stress-editable.docx');
const editablePdfPath = path.join(pdfOutputDir, 'rdl-table-stress-editable-rendered.pdf');
const fixedDocxPath = path.join(docxOutputDir, 'rdl-table-stress-fixed-editable.docx');
const fixedPdfPath = path.join(pdfOutputDir, 'rdl-table-stress-fixed-editable-rendered.pdf');
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const { model, request } = createStressScenario();
const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-stress-certification-'));

function count(text, pattern) {
  return (text.match(pattern) || []).length;
}

function markerAudit(text) {
  const missing = [];
  const duplicated = [];
  for (let index = 1; index <= STRESS_ROW_COUNT; index += 1) {
    const marker = `UNIQUE_ROW_${String(index).padStart(4, '0')}`;
    const occurrences = count(text, new RegExp(marker, 'g'));
    if (occurrences === 0) missing.push(marker);
    else if (occurrences !== 1) duplicated.push({ marker, occurrences });
  }
  return { missing, duplicated, passed: missing.length === 0 && duplicated.length === 0 };
}

function overflowAudit(text) {
  const missing = [];
  const duplicated = [];
  for (let index = 1; index <= STRESS_OVERFLOW_LINES; index += 1) {
    const marker = `GIANT_LINE_${String(index).padStart(3, '0')}`;
    const occurrences = count(text, new RegExp(marker, 'g'));
    if (occurrences === 0) missing.push(marker);
    else if (occurrences !== 1) duplicated.push({ marker, occurrences });
  }
  return {
    startCount: count(text, /GIANT_CELL_START/g),
    endCount: count(text, /GIANT_CELL_END/g),
    missing,
    duplicated,
    passed: missing.length === 0 && duplicated.length === 0
      && count(text, /GIANT_CELL_START/g) === 1 && count(text, /GIANT_CELL_END/g) === 1,
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

function pageDimensions(document) {
  return document.getPages().map((page) => ({ width: page.getWidth(), height: page.getHeight() }));
}

// OOXML page dimensions are integer twips (1/20 pt). A PDF can encode finer values, so one twip is the
// strictest portable comparison without changing the canonical PDF (for A4 the observed delta is 0.028 pt).
function dimensionsMatch(left, right, tolerance = 0.05) {
  return left.length === right.length && left.every((page, index) => (
    Math.abs(page.width - right[index].width) <= tolerance
    && Math.abs(page.height - right[index].height) <= tolerance
  ));
}

try {
  await Promise.all([
    fs.mkdir(pdfOutputDir, { recursive: true }),
    fs.mkdir(docxOutputDir, { recursive: true }),
  ]);
  const [directPdf, editableDocx, fixedDocx] = await Promise.all([
    renderPdf(model, request, config),
    renderEditableDocx(model, request),
    renderFixedEditableDocx(model, request, config),
  ]);
  await Promise.all([
    fs.writeFile(directPdfPath, directPdf.buffer),
    fs.writeFile(editableDocxPath, editableDocx.buffer),
    fs.writeFile(fixedDocxPath, fixedDocx.buffer),
  ]);
  await convertDocxToPdf(editableDocxPath, editablePdfPath, 'structured');
  await convertDocxToPdf(fixedDocxPath, fixedPdfPath, 'fixed');

  const [directPdfDocument, editablePdfDocument, fixedPdfDocument, directText, editableText, fixedText, zip, fixedZip] = await Promise.all([
    PDFDocument.load(await fs.readFile(directPdfPath)),
    PDFDocument.load(await fs.readFile(editablePdfPath)),
    PDFDocument.load(await fs.readFile(fixedPdfPath)),
    textFromPdf(directPdfPath, 'direct.txt'),
    textFromPdf(editablePdfPath, 'editable.txt'),
    textFromPdf(fixedPdfPath, 'fixed.txt'),
    JSZip.loadAsync(await fs.readFile(editableDocxPath)),
    JSZip.loadAsync(await fs.readFile(fixedDocxPath)),
  ]);
  const documentXml = await zip.file('word/document.xml').async('string');
  const fixedDocumentXml = await fixedZip.file('word/document.xml').async('string');
  const directPageCount = directPdfDocument.getPageCount();
  const editablePageCount = editablePdfDocument.getPageCount();
  const fixedPageCount = fixedPdfDocument.getPageCount();
  const directDimensions = pageDimensions(directPdfDocument);
  const fixedDimensions = pageDimensions(fixedPdfDocument);
  const report = {
    passed: false,
    targetPages: { minimum: 30, maximum: 40, structuredDocxAdvisoryOnly: true },
    inputs: { rows: STRESS_ROW_COUNT, giantOverflowLines: STRESS_OVERFLOW_LINES, groups: 13, subgroupsPerGroup: 4 },
    artifacts: { directPdfPath, editableDocxPath, editablePdfPath, fixedDocxPath, fixedPdfPath },
    directPdf: {
      pages: directPageCount,
      pageRangePassed: directPageCount >= 30 && directPageCount <= 40,
      rowMarkers: markerAudit(directText),
      overflow: overflowAudit(directText),
      endMarkerCount: count(directText, /STRESS_DOCUMENT_END/g),
    },
    editableDocx: {
      renderedPages: editablePageCount,
      pageRangeAdvisoryPassed: editablePageCount >= 30 && editablePageCount <= 40,
      rowMarkers: markerAudit(editableText),
      overflow: overflowAudit(editableText),
      endMarkerCount: count(editableText, /STRESS_DOCUMENT_END/g),
      openXml: {
        nativeTables: count(documentXml, /<w:tbl>/g),
        repeatingHeaderRows: count(documentXml, /<w:tblHeader\/>/g),
        atomicHeaderRows: count(documentXml, /<w:cantSplit\/>/g),
        horizontalMergedHeaderCells: count(documentXml, /<w:gridSpan w:val="2"\/>/g),
        verticalMergedHeaderStarts: count(documentXml, /<w:vMerge w:val="restart"\/>/g),
        verticalMergedHeaderContinuations: count(documentXml, /<w:vMerge w:val="continue"\/>/g),
      },
    },
    fixedEditableDocx: {
      sourcePages: fixedDocx.pageCount,
      renderedPages: fixedPageCount,
      pageCountMatchesPdf: fixedDocx.pageCount === directPageCount && fixedPageCount === directPageCount,
      pageDimensionsMatchPdf: dimensionsMatch(directDimensions, fixedDimensions),
      pageDimensionTolerancePt: 0.05,
      editableTextRatio: fixedDocx.editableTextRatio,
      rowMarkers: markerAudit(fixedText),
      overflow: overflowAudit(fixedText),
      endMarkerCount: count(fixedText, /STRESS_DOCUMENT_END/g),
      openXml: {
        nativeTables: count(fixedDocumentXml, /<w:tbl>/g),
        positionedShapes: count(fixedDocumentXml, /<wps:wsp>/g),
        forcedPageBreaks: count(fixedDocumentXml, /w:type="page"/g),
      },
    },
    combinations: {
      repeatedThreeRowHeader: true,
      horizontalMergedHeaderCells: true,
      verticalMergedHeaderCells: true,
      nestedGroupAndSubgroupSorting: true,
      conditionalCellBackgrounds: true,
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
    && report.editableDocx.rowMarkers.passed
    && report.editableDocx.overflow.passed
    && report.editableDocx.endMarkerCount === 1
    && report.editableDocx.openXml.nativeTables === 1
    && report.editableDocx.openXml.repeatingHeaderRows === 3
    && report.editableDocx.openXml.atomicHeaderRows === 3
    && report.editableDocx.openXml.horizontalMergedHeaderCells === 5
    && report.editableDocx.openXml.verticalMergedHeaderStarts === 6
    && report.editableDocx.openXml.verticalMergedHeaderContinuations === 6
    && report.fixedEditableDocx.pageCountMatchesPdf
    && report.fixedEditableDocx.pageDimensionsMatchPdf
    && report.fixedEditableDocx.editableTextRatio === 1
    && report.fixedEditableDocx.rowMarkers.passed
    && report.fixedEditableDocx.overflow.passed
    && report.fixedEditableDocx.endMarkerCount === 1
    && report.fixedEditableDocx.openXml.nativeTables === 0
    && report.fixedEditableDocx.openXml.positionedShapes > 0
    && report.fixedEditableDocx.openXml.forcedPageBreaks === directPageCount - 1;
  await fs.writeFile(reportPath, `${JSON.stringify(report, null, 2)}\n`);
  process.stdout.write(`${JSON.stringify({ reportPath, ...report }, null, 2)}\n`);
  if (process.argv.includes('--require-pass') && !report.passed) process.exitCode = 1;
} finally {
  await fs.rm(tempDir, { recursive: true, force: true });
}
