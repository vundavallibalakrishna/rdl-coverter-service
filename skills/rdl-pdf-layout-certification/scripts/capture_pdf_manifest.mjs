#!/usr/bin/env node
import crypto from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { PDFDocument } from 'pdf-lib';

const execFileAsync = promisify(execFile);
const args = process.argv.slice(2);
const pdfPath = args[0];
const outIndex = args.indexOf('--out');
const outputPath = outIndex >= 0 ? args[outIndex + 1] : null;
if (!pdfPath || !outputPath) {
  console.error('Usage: capture_pdf_manifest.mjs <pdf> --out tmp/<manifest>.json');
  process.exit(2);
}

const serviceRoot = path.resolve(path.dirname(new URL(import.meta.url).pathname), '../../..');
const tmpRoot = path.join(serviceRoot, 'tmp');
const resolvedOutput = path.resolve(outputPath);
if (path.dirname(resolvedOutput) !== tmpRoot) throw new Error('Manifest must be written directly under repository tmp/');

const pdfBytes = await fs.readFile(pdfPath);
const pdf = await PDFDocument.load(pdfBytes);
const [{ stdout: fontsText }, { stdout: bboxText }] = await Promise.all([
  execFileAsync('pdffonts', [pdfPath], { maxBuffer: 16 * 1024 * 1024 }),
  execFileAsync('pdftotext', ['-bbox-layout', pdfPath, '-'], { maxBuffer: 256 * 1024 * 1024 }),
]);

const fontLines = fontsText.split(/\r?\n/).slice(2).filter((line) => line.trim());
const fonts = fontLines.map((line) => {
  const parts = line.trim().split(/\s+/);
  return {
    name: parts[0] || '',
    type: parts[1] || '',
    encoding: parts[2] || '',
    embedded: parts[3] === 'yes',
    subset: parts[4] === 'yes',
    unicode: parts[5] === 'yes',
  };
});
const words = [...bboxText.matchAll(/<word xMin="([^"]+)" yMin="([^"]+)" xMax="([^"]+)" yMax="([^"]+)">([\s\S]*?)<\/word>/g)]
  .map((match) => ({
    xMin: Number(match[1]),
    yMin: Number(match[2]),
    xMax: Number(match[3]),
    yMax: Number(match[4]),
    text: match[5].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>'),
  }));
const manifest = {
  schemaVersion: 1,
  file: path.basename(pdfPath),
  sha256: crypto.createHash('sha256').update(pdfBytes).digest('hex'),
  pageCount: pdf.getPageCount(),
  pages: pdf.getPages().map((page, index) => ({
    number: index + 1,
    widthPt: page.getWidth(),
    heightPt: page.getHeight(),
  })),
  fonts,
  allFontsEmbedded: fonts.every((font) => font.embedded),
  normalizedText: words.map((word) => word.text).join(' ').replace(/\s+/g, ' ').trim(),
  words,
};
await fs.writeFile(resolvedOutput, `${JSON.stringify(manifest, null, 2)}\n`, { mode: 0o600 });
console.log(resolvedOutput);
