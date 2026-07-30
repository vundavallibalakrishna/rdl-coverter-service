#!/usr/bin/env node
import fs from 'node:fs/promises';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { PDFDocument } from 'pdf-lib';
import { PNG } from 'pngjs';

const execFileAsync = promisify(execFile);
const args = process.argv.slice(2);
const referencePath = args[0];
const candidatePath = args[1];
const outIndex = args.indexOf('--out');
const outputPath = outIndex >= 0 ? args[outIndex + 1] : null;
if (!referencePath || !candidatePath || !outputPath) {
  console.error('Usage: compare_pdf_layout.mjs <reference.pdf> <candidate.pdf> --out tmp/<diff>.json');
  process.exit(2);
}

const serviceRoot = path.resolve(path.dirname(new URL(import.meta.url).pathname), '../../..');
const tmpRoot = path.join(serviceRoot, 'tmp');
const resolvedOutput = path.resolve(outputPath);
if (path.dirname(resolvedOutput) !== tmpRoot) throw new Error('Comparison must be written directly under repository tmp/');

const [referenceBytes, candidateBytes] = await Promise.all([fs.readFile(referencePath), fs.readFile(candidatePath)]);
const [reference, candidate] = await Promise.all([PDFDocument.load(referenceBytes), PDFDocument.load(candidateBytes)]);
const pageCountMatches = reference.getPageCount() === candidate.getPageCount();
const referencePages = reference.getPages();
const candidatePages = candidate.getPages();
const pageDimensionsMatch = pageCountMatches && referencePages.every((page, index) => (
  Math.abs(page.getWidth() - candidatePages[index].getWidth()) <= 0.01
  && Math.abs(page.getHeight() - candidatePages[index].getHeight()) <= 0.01
));

function decodeText(value) {
  return String(value)
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function wordsByPage(xml) {
  return [...xml.matchAll(/<page\b[^>]*>([\s\S]*?)<\/page>/g)].map((page, pageIndex) => (
    [...page[1].matchAll(/<word xMin="([^"]+)" yMin="([^"]+)" xMax="([^"]+)" yMax="([^"]+)">([\s\S]*?)<\/word>/g)]
      .map((word, wordIndex) => ({
        pageNumber: pageIndex + 1,
        wordIndex,
        xMin: Number(word[1]),
        yMin: Number(word[2]),
        xMax: Number(word[3]),
        yMax: Number(word[4]),
        text: decodeText(word[5]),
      }))
  ));
}

const [{ stdout: referenceBbox }, { stdout: candidateBbox }] = await Promise.all([
  execFileAsync('pdftotext', ['-bbox-layout', referencePath, '-'], { maxBuffer: 256 * 1024 * 1024 }),
  execFileAsync('pdftotext', ['-bbox-layout', candidatePath, '-'], { maxBuffer: 256 * 1024 * 1024 }),
]);
const referenceWords = wordsByPage(referenceBbox);
const candidateWords = wordsByPage(candidateBbox);
const textMatches = pageCountMatches && referenceWords.every((words, pageIndex) => {
  const compared = candidateWords[pageIndex] || [];
  return words.length === compared.length && words.every((word, index) => word.text === compared[index].text);
});
const geometryPages = referenceWords.map((words, pageIndex) => {
  const compared = candidateWords[pageIndex] || [];
  const deltas = words.map((word, wordIndex) => {
    const candidateWord = compared[wordIndex];
    if (!candidateWord || candidateWord.text !== word.text) {
      return { wordIndex, text: word.text, comparable: false, maximumDeltaPt: null };
    }
    const maximumDeltaPt = Math.max(
      Math.abs(word.xMin - candidateWord.xMin),
      Math.abs(word.yMin - candidateWord.yMin),
      Math.abs(word.xMax - candidateWord.xMax),
      Math.abs(word.yMax - candidateWord.yMax),
    );
    return {
      wordIndex,
      text: word.text,
      comparable: true,
      maximumDeltaPt,
    };
  });
  const comparable = words.length === compared.length && deltas.every((word) => word.comparable);
  return {
    pageNumber: pageIndex + 1,
    comparable,
    maximumDeltaPt: comparable ? Math.max(0, ...deltas.map((word) => word.maximumDeltaPt)) : null,
    withinHalfPoint: comparable && deltas.every((word) => word.maximumDeltaPt <= 0.5),
    differences: deltas.filter((word) => !word.comparable || word.maximumDeltaPt > 0.5),
  };
});
const geometryGate = pageCountMatches
  && geometryPages.length === candidateWords.length
  && geometryPages.every((page) => page.withinHalfPoint);

const prefixBase = path.basename(outputPath, path.extname(outputPath)).replace(/[^A-Za-z0-9._-]/g, '-');
const comparisons = [];
for (let pageNumber = 1; pageNumber <= Math.min(referencePages.length, candidatePages.length); pageNumber += 1) {
  const refPrefix = path.join(tmpRoot, `${prefixBase}-reference-${pageNumber}`);
  const candidatePrefix = path.join(tmpRoot, `${prefixBase}-candidate-${pageNumber}`);
  await Promise.all([
    execFileAsync('pdftoppm', ['-f', String(pageNumber), '-l', String(pageNumber), '-singlefile', '-png', '-r', '144', referencePath, refPrefix]),
    execFileAsync('pdftoppm', ['-f', String(pageNumber), '-l', String(pageNumber), '-singlefile', '-png', '-r', '144', candidatePath, candidatePrefix]),
  ]);
  const refFile = `${refPrefix}.png`;
  const candidateFile = `${candidatePrefix}.png`;
  const [ref, rendered] = await Promise.all([
    fs.readFile(refFile).then((buffer) => PNG.sync.read(buffer)),
    fs.readFile(candidateFile).then((buffer) => PNG.sync.read(buffer)),
  ]);
  let exceeding = 0;
  const comparable = ref.width === rendered.width && ref.height === rendered.height;
  if (comparable) {
    for (let offset = 0; offset < ref.data.length; offset += 4) {
      if (Math.max(
        Math.abs(ref.data[offset] - rendered.data[offset]),
        Math.abs(ref.data[offset + 1] - rendered.data[offset + 1]),
        Math.abs(ref.data[offset + 2] - rendered.data[offset + 2]),
      ) > 16) exceeding += 1;
    }
  }
  const pixels = ref.width * ref.height;
  comparisons.push({ pageNumber, comparable, pixels, exceeding, ratio: comparable ? exceeding / pixels : null });
  await Promise.all([fs.unlink(refFile), fs.unlink(candidateFile)]);
}
const pixelGate = comparisons.every((page) => page.comparable && page.ratio <= 0.005);
const report = {
  schemaVersion: 1,
  reference: path.basename(referencePath),
  candidate: path.basename(candidatePath),
  pageCountMatches,
  pageDimensionsMatch,
  textMatches,
  geometryGate,
  geometryTolerancePt: 0.5,
  geometryPages,
  pixelGate,
  passed: pageCountMatches && pageDimensionsMatch && textMatches && geometryGate && pixelGate,
  pages: comparisons,
};
await fs.writeFile(resolvedOutput, `${JSON.stringify(report, null, 2)}\n`, { mode: 0o600 });
console.log(resolvedOutput);
if (!report.passed) process.exitCode = 1;
