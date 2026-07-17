import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import { PDFDocument } from 'pdf-lib';
import { PNG } from 'pngjs';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { validateRenderInput } from '../src/rdl/validation.js';
import { RenderRunner } from '../src/worker/runner.js';
import { samplePath } from './lib/samples.js';

const execFileAsync = promisify(execFile);
const serviceRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const positionalArguments = process.argv.slice(2).filter((argument) => !argument.startsWith('--'));
const requireMatch = process.argv.includes('--require-match');
const exactInputs = process.argv.includes('--exact-inputs');
const rdlPath = path.resolve(positionalArguments[0] || path.join(serviceRoot, '..', 'lumina', 'Combined Assurance Reports (2).rdl'));
const referencePath = path.resolve(positionalArguments[1] || path.join(serviceRoot, 'Combined_Assurance_Reports_modifiedV4.pdf'));
const hydrationPath = path.resolve(positionalArguments[2] || samplePath('combined-assurance-hydration.json'));
// Generated artifacts live under tmp/ so they stay out of version control (see .gitignore).
const outputDir = path.join(serviceRoot, 'tmp', 'output', 'pdf');
const generatedPath = path.join(outputDir, 'combined-assurance-hydrated.pdf');
const comparisonPath = path.join(outputDir, 'combined-assurance-hydrated.comparison.json');

function pageGeometry(pdf) {
  return Array.from({ length: pdf.getPageCount() }, (_, index) => {
    const page = pdf.getPage(index);
    return { width: page.getWidth(), height: page.getHeight() };
  });
}

function dimensionsMatch(reference, generated, tolerance = 0.01) {
  if (reference.length === 0 || generated.length === 0) return false;
  return Math.abs(reference[0].width - generated[0].width) <= tolerance
    && Math.abs(reference[0].height - generated[0].height) <= tolerance;
}

async function extractText(pdfPath, textPath) {
  await execFileAsync('pdftotext', ['-layout', pdfPath, textPath]);
  return fs.readFile(textPath, 'utf8');
}

async function rasterPage(pdftoppmPath, pdfPath, pageNumber, outputPrefix) {
  await execFileAsync(pdftoppmPath, [
    '-f', String(pageNumber), '-l', String(pageNumber), '-singlefile', '-png', '-r', '144', pdfPath, outputPrefix,
  ]);
  return PNG.sync.read(await fs.readFile(`${outputPrefix}.png`));
}

function pixelDifference(reference, generated) {
  if (reference.width !== generated.width || reference.height !== generated.height) {
    return { comparable: false, reference: { width: reference.width, height: reference.height }, generated: { width: generated.width, height: generated.height } };
  }
  let exceeding = 0;
  const pixels = reference.width * reference.height;
  for (let offset = 0; offset < reference.data.length; offset += 4) {
    if (Math.max(
      Math.abs(reference.data[offset] - generated.data[offset]),
      Math.abs(reference.data[offset + 1] - generated.data[offset + 1]),
      Math.abs(reference.data[offset + 2] - generated.data[offset + 2]),
    ) > 16) exceeding += 1;
  }
  return { comparable: true, pixels, exceeding, ratio: exceeding / pixels };
}

function normalizedText(value) {
  return String(value).replace(/\s+/g, ' ').trim();
}

const [rdlBuffer, referenceBuffer, hydrationText] = await Promise.all([
  fs.readFile(rdlPath),
  fs.readFile(referencePath),
  fs.readFile(hydrationPath, 'utf8'),
]);
const request = JSON.parse(hydrationText);
const model = parseRdl(rdlBuffer);
const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-reference-verification-'));
const config = loadConfig({ ...process.env, RDL_TEMP_ROOT: tempRoot, RDL_STRICT_FONTS: 'false' });
const validation = validateRenderInput(model, request, config);
const runner = new RenderRunner(config);

try {
  await fs.mkdir(outputDir, { recursive: true });
  const rendered = await runner.render({ rdlBuffer, request });
  await fs.writeFile(generatedPath, rendered.buffer);

  const [referencePdf, generatedPdf] = await Promise.all([
    PDFDocument.load(referenceBuffer),
    PDFDocument.load(rendered.buffer),
  ]);
  const referencePages = pageGeometry(referencePdf);
  const generatedPages = pageGeometry(generatedPdf);
  const [referenceText, generatedText] = await Promise.all([
    extractText(referencePath, path.join(tempRoot, 'reference.txt')),
    extractText(generatedPath, path.join(tempRoot, 'generated.txt')),
  ]);
  const normalizedReferenceText = normalizedText(referenceText);
  const normalizedGeneratedText = normalizedText(generatedText);
  const anchors = [
    'Introduction based on the UAT refinements for this report.',
    'Extensive Assurance',
    'Describing',
    '2nd line',
    'System Administrator',
    'Risk Manager : Jerry Mathevhula',
  ].map((text) => ({
    text,
    reference: normalizedReferenceText.includes(normalizedText(text)),
    generated: normalizedGeneratedText.includes(normalizedText(text)),
  }));

  const tablixHierarchyGaps = model.body.items.filter((item) => item.type === 'Tablix').map((item) => {
    const bodyColumnWidth = (item.columns || []).reduce((sum, width) => sum + width, 0);
    return {
      name: item.name,
      totalWidth: item.width,
      bodyColumnWidth,
      rowHierarchyWidth: Math.max(0, item.width - bodyColumnWidth),
    };
  }).filter((item) => item.rowHierarchyWidth > 0.5);

  const rasterComparisons = [];
  for (let pageNumber = 1; pageNumber <= Math.min(3, referencePdf.getPageCount(), generatedPdf.getPageCount()); pageNumber += 1) {
    const [referenceRaster, generatedRaster] = await Promise.all([
      rasterPage(config.pdftoppmPath, referencePath, pageNumber, path.join(tempRoot, `reference-${pageNumber}`)),
      rasterPage(config.pdftoppmPath, generatedPath, pageNumber, path.join(tempRoot, `generated-${pageNumber}`)),
    ]);
    rasterComparisons.push({ pageNumber, ...pixelDifference(referenceRaster, generatedRaster) });
  }

  const pageCountMatches = referencePdf.getPageCount() === generatedPdf.getPageCount();
  const pageDimensionsMatch = dimensionsMatch(referencePages, generatedPages);
  const anchorsMatch = anchors.every((anchor) => !anchor.reference || anchor.generated);
  const fixedPagesMeetPixelTolerance = rasterComparisons.every((page) => page.comparable && page.ratio <= 0.005);
  const certified = exactInputs && pageCountMatches && pageDimensionsMatch && anchorsMatch && fixedPagesMeetPixelTolerance;
  const blockingDifferences = [];
  if (!pageCountMatches) blockingDifferences.push('Generated page count differs from the SSRS reference.');
  if (!pageDimensionsMatch) blockingDifferences.push('Generated page dimensions differ from the SSRS reference.');
  if (!anchorsMatch) blockingDifferences.push('One or more reference text anchors are missing from the generated PDF.');
  if (!fixedPagesMeetPixelTolerance) blockingDifferences.push('The first three fixed-layout pages exceed the 0.5% pixel-difference tolerance.');
  if (tablixHierarchyGaps.length > 0) blockingDifferences.push('The current normalized tablix cells do not include the RDL row-header hierarchy width.');
  if (!exactInputs) blockingDifferences.push('The hydration is reconstructed from visible PDF values; pass --exact-inputs only with the exact SSRS parameters, rows, and font versions.');

  const report = {
    certified,
    inputs: {
      rdl: path.relative(serviceRoot, rdlPath),
      referencePdf: path.relative(serviceRoot, referencePath),
      hydration: path.relative(serviceRoot, hydrationPath),
      hydratedRows: validation.totalRows,
    },
    reference: { pages: referencePdf.getPageCount(), firstPage: referencePages[0] },
    generated: { pages: generatedPdf.getPageCount(), firstPage: generatedPages[0], bytes: rendered.buffer.length },
    checks: { exactInputs, pageCountMatches, pageDimensionsMatch, anchors, rasterComparisons, tablixHierarchyGaps },
    blockingDifferences,
  };
  await fs.writeFile(comparisonPath, `${JSON.stringify(report, null, 2)}\n`);
  process.stdout.write(`${JSON.stringify({ generatedPath, comparisonPath, ...report }, null, 2)}\n`);
  if (requireMatch && !certified) process.exitCode = 1;
} finally {
  await runner.shutdown();
  await fs.rm(tempRoot, { recursive: true, force: true });
}
