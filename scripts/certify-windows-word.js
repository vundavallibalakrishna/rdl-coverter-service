#!/usr/bin/env node
import fs from 'node:fs/promises';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';

const execFileAsync = promisify(execFile);
const serviceRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const tmpRoot = path.join(serviceRoot, 'tmp');
const [docxArgument, pdfArgument, ...options] = process.argv.slice(2);

if (!docxArgument || !pdfArgument || options.includes('--help') || options.includes('-h')) {
  process.stdout.write(`Usage:
  npm run certify:windows-word -- <page-locked.docx> <canonical.pdf> [--prefix <name>]

Runs only on Windows with desktop Microsoft Word installed. Word opens and repaginates the DOCX without
updating fields or external links, exports it to PDF, and produces OOXML/text/geometry/144-DPI evidence
directly under repository tmp/.
`);
  process.exit(docxArgument && pdfArgument ? 0 : 2);
}
if (process.platform !== 'win32') {
  throw new Error('Windows Word certification requires Windows with desktop Microsoft Word installed');
}

const prefixIndex = options.indexOf('--prefix');
const prefix = String(prefixIndex >= 0 ? options[prefixIndex + 1] : path.basename(docxArgument, path.extname(docxArgument)))
  .replace(/[^A-Za-z0-9._-]/g, '-')
  .slice(0, 100);
if (!prefix) throw new Error('Certification prefix is empty');

const docxPath = path.resolve(docxArgument);
const canonicalPdfPath = path.resolve(pdfArgument);
const wordPdfPath = path.join(tmpRoot, `${prefix}-word-export.pdf`);
const automationPath = path.join(tmpRoot, `${prefix}-word-automation.json`);
const canonicalManifestPath = path.join(tmpRoot, `${prefix}-canonical-pdf-manifest.json`);
const wordManifestPath = path.join(tmpRoot, `${prefix}-word-pdf-manifest.json`);
const comparisonPath = path.join(tmpRoot, `${prefix}-word-pdf-comparison.json`);
const ooxmlAuditPath = path.join(tmpRoot, `${prefix}-ooxml-audit.json`);
const certificationPath = path.join(tmpRoot, `${prefix}-windows-word-certification.json`);
const powershellScript = path.join(
  serviceRoot,
  'skills',
  'rdl-windows-word-fidelity',
  'scripts',
  'export_word_pdf.ps1',
);
const captureScript = path.join(
  serviceRoot,
  'skills',
  'rdl-pdf-layout-certification',
  'scripts',
  'capture_pdf_manifest.mjs',
);
const compareScript = path.join(
  serviceRoot,
  'skills',
  'rdl-pdf-layout-certification',
  'scripts',
  'compare_pdf_layout.mjs',
);
const auditScript = path.join(
  serviceRoot,
  'skills',
  'rdl-windows-word-fidelity',
  'scripts',
  'audit_docx_ooxml.mjs',
);

await fs.mkdir(tmpRoot, { recursive: true });
await execFileAsync('powershell.exe', [
  '-NoProfile',
  '-NonInteractive',
  '-ExecutionPolicy',
  'Bypass',
  '-File',
  powershellScript,
  '-InputDocx',
  docxPath,
  '-OutputPdf',
  wordPdfPath,
  '-ResultJson',
  automationPath,
], { maxBuffer: 32 * 1024 * 1024 });

await Promise.all([
  execFileAsync(process.execPath, [captureScript, canonicalPdfPath, '--out', canonicalManifestPath]),
  execFileAsync(process.execPath, [captureScript, wordPdfPath, '--out', wordManifestPath]),
  execFileAsync(process.execPath, [auditScript, docxPath, '--out', ooxmlAuditPath]),
]);
let comparisonProcessPassed = true;
try {
  await execFileAsync(process.execPath, [
    compareScript,
    canonicalPdfPath,
    wordPdfPath,
    '--out',
    comparisonPath,
  ], { maxBuffer: 32 * 1024 * 1024 });
} catch {
  comparisonProcessPassed = false;
}

const readJson = async (file) => JSON.parse((await fs.readFile(file, 'utf8')).replace(/^\uFEFF/, ''));
const [automation, canonical, word, comparison, ooxml] = await Promise.all([
  readJson(automationPath),
  readJson(canonicalManifestPath),
  readJson(wordManifestPath),
  readJson(comparisonPath),
  readJson(ooxmlAuditPath),
]);
const textMatches = canonical.normalizedText === word.normalizedText;
const certification = {
  schemaVersion: 1,
  renderer: 'DOCX_EDITABLE',
  layoutMode: 'windows-paged-editable',
  authority: 'Microsoft Word for Windows',
  input: {
    docx: path.basename(docxPath),
    canonicalPdf: path.basename(canonicalPdfPath),
  },
  word: automation,
  ooxml,
  comparison,
  text: {
    matches: textMatches,
    canonicalLength: canonical.normalizedText.length,
    wordLength: word.normalizedText.length,
  },
  gates: {
    wordAutomation: automation.passed === true,
    package: ooxml.passed === true,
    exactPageCount: comparison.pageCountMatches === true && automation.pageCount === canonical.pageCount,
    exactPageDimensions: comparison.pageDimensionsMatch === true,
    identicalDisplayedText: textMatches && comparison.textMatches === true,
    geometryWithinHalfPoint: comparison.geometryGate === true,
    pixelDifference144Dpi: comparison.pixelGate === true && comparisonProcessPassed,
  },
};
certification.passed = Object.values(certification.gates).every(Boolean);
await fs.writeFile(certificationPath, `${JSON.stringify(certification, null, 2)}\n`, { mode: 0o600 });
process.stdout.write(`${certificationPath}\n`);
if (!certification.passed) process.exitCode = 1;
