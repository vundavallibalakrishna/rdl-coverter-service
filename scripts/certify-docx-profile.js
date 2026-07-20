import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import { PDFDocument } from 'pdf-lib';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { validateRenderInput } from '../src/rdl/validation.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderPdf } from '../src/render/pdf.js';
import { analyzeStructuredEditableCompatibility } from '../src/render/structuredCompatibility.js';

const execFileAsync = promisify(execFile);
const serviceRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);

if (args.includes('--help') || args.includes('-h')) {
  process.stdout.write(`Usage: node scripts/certify-docx-profile.js <report.rdl> <request.json> [ssrs-reference.pdf] [--renderer=auto|word|libreoffice] [--exact-inputs] [--require-match] [--out=dir]

Renders the canonical PDF plus structured DOCX variants, exports DOCX through Microsoft Word for Mac or
LibreOffice, checks page count/dimensions/text/OpenXML editability, and writes a certification report plus a
candidate DOCX profile. Use --renderer=word --exact-inputs --require-match for release certification.
`);
  process.exit(0);
}

function valueAfter(name) {
  const prefixed = args.find((argument) => argument.startsWith(`${name}=`));
  if (prefixed) return prefixed.slice(name.length + 1);
  const index = args.indexOf(name);
  return index >= 0 ? args[index + 1] : null;
}

function hasFlag(name) {
  return args.includes(name);
}

const positional = args.filter((argument, index) => {
  if (argument.startsWith('--')) return false;
  const previous = args[index - 1];
  return !previous?.startsWith('--') || previous.includes('=');
});

const rdlPath = positional[0] ? path.resolve(positional[0]) : null;
const requestPath = positional[1] ? path.resolve(positional[1]) : null;
const referencePath = positional[2] ? path.resolve(positional[2]) : null;
const renderer = valueAfter('--renderer') || 'auto';
const requireMatch = hasFlag('--require-match');
const exactInputs = hasFlag('--exact-inputs');
const outputRoot = path.resolve(valueAfter('--out') || path.join(serviceRoot, 'tmp', 'output', 'docx-certification'));

if (!rdlPath || !requestPath) {
  process.stderr.write('Usage: node scripts/certify-docx-profile.js <report.rdl> <request.json> [ssrs-reference.pdf] [--renderer=auto|word|libreoffice] [--exact-inputs] [--require-match] [--out=dir]\n');
  process.exit(2);
}

function safeSlug(value) {
  return String(value || 'report').replace(/\.[A-Za-z0-9]+$/, '').replace(/[^A-Za-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 100) || 'report';
}

function normalizeText(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function textCoverage(referenceText, candidateText) {
  const referenceTokens = [...new Set(normalizeText(referenceText)
    .split(/[^A-Za-z0-9_.:/-]+/)
    .filter((token) => token.length >= 3))];
  if (referenceTokens.length === 0) return { ratio: 1, matched: 0, total: 0, passed: true };
  const candidate = normalizeText(candidateText);
  const matched = referenceTokens.filter((token) => candidate.includes(token)).length;
  return {
    ratio: matched / referenceTokens.length,
    matched,
    total: referenceTokens.length,
    passed: matched / referenceTokens.length >= 0.95,
  };
}

function pageDimensions(document) {
  return document.getPages().map((page) => ({ width: page.getWidth(), height: page.getHeight() }));
}

function dimensionsMatch(left, right, tolerance = 0.05) {
  return left.length === right.length && left.every((page, index) => (
    Math.abs(page.width - right[index].width) <= tolerance
    && Math.abs(page.height - right[index].height) <= tolerance
  ));
}

async function commandExists(command) {
  try {
    await execFileAsync('which', [command]);
    return true;
  } catch {
    return false;
  }
}

async function wordAvailable() {
  if (process.platform !== 'darwin') return false;
  try {
    const result = await execFileAsync('osascript', ['-e', 'tell application "System Events" to return exists application process "Microsoft Word"']);
    if (/true|false/i.test(result.stdout)) return true;
  } catch {
    // System Events can fail when Word is not running; fall back to app lookup.
  }
  try {
    await execFileAsync('mdfind', ['kMDItemCFBundleIdentifier == "com.microsoft.Word"']);
    return true;
  } catch {
    return false;
  }
}

async function selectedRenderer() {
  if (renderer === 'word') {
    if (!(await wordAvailable())) throw new Error('Microsoft Word for Mac is not available');
    return 'word';
  }
  if (renderer === 'libreoffice') {
    if (!(await commandExists(process.env.RDL_SOFFICE_PATH || 'soffice'))) throw new Error('LibreOffice/soffice is not available');
    return 'libreoffice';
  }
  if (renderer !== 'auto') throw new Error(`Unknown renderer: ${renderer}`);
  if (await wordAvailable()) return 'word';
  if (await commandExists(process.env.RDL_SOFFICE_PATH || 'soffice')) return 'libreoffice';
  throw new Error('No DOCX-to-PDF renderer is available. Install Microsoft Word for Mac or LibreOffice.');
}

async function exportWithLibreOffice(docxPath, pdfPath, tempDir, label) {
  const profile = path.join(tempDir, `lo-profile-${label}`);
  const conversionDir = path.join(tempDir, `lo-out-${label}`);
  await Promise.all([fs.mkdir(profile), fs.mkdir(conversionDir)]);
  await execFileAsync(process.env.RDL_SOFFICE_PATH || 'soffice', [
    `-env:UserInstallation=file://${profile}`,
    '--invisible', '--headless', '--norestore', '--convert-to', 'pdf', '--outdir', conversionDir, docxPath,
  ], { maxBuffer: 16 * 1024 * 1024 });
  await fs.copyFile(path.join(conversionDir, `${path.basename(docxPath, '.docx')}.pdf`), pdfPath);
}

function escapeAppleScript(value) {
  return String(value).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}

async function exportWithWord(docxPath, pdfPath) {
  const script = `
set inputPath to POSIX file "${escapeAppleScript(docxPath)}"
set outputPath to POSIX file "${escapeAppleScript(pdfPath)}"
tell application "Microsoft Word"
  activate
  open inputPath
  set theDoc to active document
  save as theDoc file name outputPath file format format PDF
  close theDoc saving no
end tell
`;
  await execFileAsync('osascript', ['-e', script], { timeout: 600_000, maxBuffer: 16 * 1024 * 1024 });
}

async function exportDocx(docxPath, pdfPath, tempDir, label, chosenRenderer) {
  if (chosenRenderer === 'word') return exportWithWord(docxPath, pdfPath);
  return exportWithLibreOffice(docxPath, pdfPath, tempDir, label);
}

async function textFromPdf(pdfPath, tempDir, label) {
  const textPath = path.join(tempDir, `${label}.txt`);
  await execFileAsync('pdftotext', ['-layout', pdfPath, textPath], { maxBuffer: 16 * 1024 * 1024 });
  return fs.readFile(textPath, 'utf8');
}

async function inspectDocx(docxPath) {
  const zip = await JSZip.loadAsync(await fs.readFile(docxPath));
  const documentXml = await zip.file('word/document.xml').async('string');
  const settingsXml = await zip.file('word/settings.xml')?.async('string') || '';
  return {
    nativeTables: (documentXml.match(/<w:tbl>/g) || []).length,
    bodyPositionedShapes: (documentXml.match(/<wps:wsp>/g) || []).length,
    repeatingHeaderRows: (documentXml.match(/<w:tblHeader\/>/g) || []).length,
    protected: /<w:(?:documentProtection|writeProtection)\b/.test(settingsXml),
  };
}

function scoreVariant(variant, canonicalPages) {
  const pageDelta = Math.abs(variant.pages - canonicalPages);
  let score = pageDelta * 100;
  if (!variant.dimensionsMatchPdf) score += 20;
  if (!variant.textCoverage.passed) score += Math.round((1 - variant.textCoverage.ratio) * 100);
  if (!variant.openXml.nativeTables) score += 100;
  if (variant.openXml.bodyPositionedShapes > 0) score += 100;
  if (variant.openXml.protected) score += 100;
  return score;
}

function candidateProfile(model, best, baseline) {
  if (!best) return null;
  const idPrefix = `${safeSlug(model.identity?.name || model.name)}-${model.identity.definitionSha256.slice(0, 12)}`;
  if (best.name === 'structured-default') return {
    id: `${idPrefix}-structured-default`,
    certified: false,
    match: { definitionSha256: model.identity.definitionSha256 },
    docx: { nativePageFragments: false },
    note: 'Candidate only. Mark certified=true after Word-export page review against the exact SSRS reference run.',
  };
  if (best.name === 'structured-native-fragments' && best.score < baseline.score) return {
    id: `${idPrefix}-native-fragments`,
    certified: false,
    match: { definitionSha256: model.identity.definitionSha256 },
    docx: { nativePageFragments: true },
    note: 'Candidate only. Mark certified=true after Word-export page review against the exact SSRS reference run.',
  };
  return null;
}

const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-docx-cert-'));
try {
  const [rdlBuffer, requestText] = await Promise.all([fs.readFile(rdlPath), fs.readFile(requestPath, 'utf8')]);
  const request = JSON.parse(requestText);
  const model = parseRdl(rdlBuffer);
  const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: process.env.RDL_STRICT_FONTS || 'false' });
  const validation = validateRenderInput(model, request, config);
  const analysis = analyzeStructuredEditableCompatibility(model, config, request);
  const chosenRenderer = await selectedRenderer();
  const slug = safeSlug(request.outputFileName || path.basename(rdlPath));
  const outputDir = path.join(outputRoot, `${slug}-${model.identity.definitionSha256.slice(0, 12)}`);
  await fs.mkdir(outputDir, { recursive: true });

  const canonicalPdf = await renderPdf(model, { ...request, output: 'PDF' }, config);
  const canonicalPdfPath = path.join(outputDir, 'canonical.pdf');
  await fs.writeFile(canonicalPdfPath, canonicalPdf.buffer);
  const variants = [
    { name: 'structured-default', request: { ...request, output: 'DOCX_EDITABLE', docx: { ...(request.docx || {}), nativePageFragments: false } } },
    { name: 'structured-native-fragments', request: { ...request, output: 'DOCX_EDITABLE', docx: { ...(request.docx || {}), nativePageFragments: true } } },
  ];

  const canonicalDocument = await PDFDocument.load(canonicalPdf.buffer);
  const canonicalDimensions = pageDimensions(canonicalDocument);
  const canonicalText = normalizeText(await textFromPdf(canonicalPdfPath, tempDir, 'canonical'));
  const variantReports = [];

  for (const variant of variants) {
    const rendered = await renderEditableDocx(model, variant.request, config, tempDir);
    const docxPath = path.join(outputDir, `${variant.name}.docx`);
    const pdfPath = path.join(outputDir, `${variant.name}.word-export.pdf`);
    await fs.writeFile(docxPath, rendered.buffer);
    await exportDocx(docxPath, pdfPath, tempDir, variant.name, chosenRenderer);
    const exportedDocument = await PDFDocument.load(await fs.readFile(pdfPath));
    const exportedText = normalizeText(await textFromPdf(pdfPath, tempDir, variant.name));
    const report = {
      name: variant.name,
      docxPath,
      exportedPdfPath: pdfPath,
      pages: exportedDocument.getPageCount(),
      pageCountMatchesPdf: exportedDocument.getPageCount() === canonicalDocument.getPageCount(),
      dimensionsMatchPdf: dimensionsMatch(canonicalDimensions, pageDimensions(exportedDocument)),
      textCoverage: { ...textCoverage(canonicalText, exportedText), exportedTextBytes: Buffer.byteLength(exportedText) },
      openXml: await inspectDocx(docxPath),
    };
    report.score = scoreVariant(report, canonicalDocument.getPageCount());
    variantReports.push(report);
  }

  const baseline = variantReports.find((variant) => variant.name === 'structured-default');
  const best = [...variantReports].sort((left, right) => left.score - right.score)[0];
  const profileCandidate = candidateProfile(model, best, baseline);
  const reference = referencePath ? await PDFDocument.load(await fs.readFile(referencePath)) : null;
  const referenceText = referencePath ? normalizeText(await textFromPdf(referencePath, tempDir, 'reference')) : null;
  const referenceChecks = reference ? {
    path: referencePath,
    pages: reference.getPageCount(),
    pageCountMatchesCanonicalPdf: reference.getPageCount() === canonicalDocument.getPageCount(),
    dimensionsMatchCanonicalPdf: dimensionsMatch(pageDimensions(reference), canonicalDimensions),
    textAnchorOverlap: referenceText.length > 0 && canonicalText.includes(referenceText.slice(0, Math.min(1000, referenceText.length))),
  } : null;
  const exactReferenceSatisfied = Boolean(exactInputs && referenceChecks?.pageCountMatchesCanonicalPdf && referenceChecks?.dimensionsMatchCanonicalPdf);
  const report = {
    certified: false,
    certificationBlocked: !exactReferenceSatisfied,
    renderer: chosenRenderer,
    inputs: {
      rdlPath,
      requestPath,
      referencePath,
      exactInputs,
      rows: validation.totalRows,
    },
    identity: model.identity,
    analysis,
    canonicalPdf: {
      path: canonicalPdfPath,
      pages: canonicalDocument.getPageCount(),
      dimensions: canonicalDimensions[0],
    },
    reference: referenceChecks,
    variants: variantReports,
    bestVariant: best.name,
    profileCandidate,
    blockingDifferences: [
      ...(!referencePath ? ['No SSRS reference PDF was supplied.'] : []),
      ...(!exactInputs ? ['--exact-inputs was not supplied; certification requires exact SSRS parameters, rows, and font versions.'] : []),
      ...(referenceChecks && !referenceChecks.pageCountMatchesCanonicalPdf ? ['Canonical PDF page count differs from the SSRS reference PDF.'] : []),
      ...(referenceChecks && !referenceChecks.dimensionsMatchCanonicalPdf ? ['Canonical PDF dimensions differ from the SSRS reference PDF.'] : []),
      ...(profileCandidate ? [] : ['No structured DOCX profile candidate improved or preserved the native editable layout score.']),
    ],
  };
  report.certified = exactReferenceSatisfied
    && profileCandidate
    && best.pageCountMatchesPdf
    && best.dimensionsMatchPdf
    && best.openXml.nativeTables > 0
    && best.openXml.bodyPositionedShapes === 0
    && !best.openXml.protected;
  const reportPath = path.join(outputDir, 'docx-certification-report.json');
  const profilePath = path.join(outputDir, 'docx-profile-candidate.json');
  await fs.writeFile(reportPath, `${JSON.stringify(report, null, 2)}\n`);
  if (profileCandidate) await fs.writeFile(profilePath, `${JSON.stringify({ profiles: [profileCandidate] }, null, 2)}\n`);
  process.stdout.write(`${JSON.stringify({ reportPath, profilePath: profileCandidate ? profilePath : null, ...report }, null, 2)}\n`);
  if (requireMatch && !report.certified) process.exitCode = 1;
} finally {
  await fs.rm(tempDir, { recursive: true, force: true });
}
