import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import path from 'node:path';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';
import { parseRdl } from '../src/rdl/parser.js';

const run = promisify(execFile);
const image = process.env.RDL_DOCKER_IMAGE || 'rdl-converter-service:smoke';
const fontHostDir = process.env.RDL_FONT_HOST_DIR || '/System/Library/Fonts/Supplemental';
const serviceRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const sample = path.resolve(process.argv[2] || path.join(serviceRoot, '..', 'lumina', 'Combined Assurance Reports (2).rdl'));
const rdl = await fs.readFile(sample);
const model = parseRdl(rdl);
const requiredVariants = {
  Arial: [['Arial.ttf'], ['Arial Bold.ttf'], ['Arial Italic.ttf'], ['Arial Bold Italic.ttf']],
  'Times New Roman': [['Times New Roman.ttf'], ['Times New Roman Bold.ttf'], ['Times New Roman Italic.ttf'], ['Times New Roman Bold Italic.ttf']],
  'Segoe UI': [['Segoe UI.ttf', 'segoeui.ttf'], ['Segoe UI Bold.ttf', 'segoeuib.ttf'], ['Segoe UI Italic.ttf', 'segoeuii.ttf'], ['Segoe UI Bold Italic.ttf', 'segoeuiz.ttf']],
  'Segoe UI Symbol': [['Segoe UI Symbol.ttf', 'seguisym.ttf', 'Arial Unicode.ttf', 'Apple Symbols.ttf']],
};
const datasets = Object.fromEntries(model.datasets.filter((dataset) => model.renderingDatasets.includes(dataset.name)).map((dataset) => [
  dataset.name,
  [Object.fromEntries(dataset.fields.map((field, index) => [field.dataField, /Int|Decimal|Double/i.test(field.typeName) ? index + 1 : `Sample ${field.dataField}`]))],
]));
const parameters = Object.fromEntries(model.parameters.map((parameter) => {
  if (parameter.multiValue) return [parameter.name, ['Sample']];
  if (parameter.dataType === 'Integer') return [parameter.name, 1];
  if (parameter.dataType === 'Float') return [parameter.name, 1.5];
  if (parameter.dataType === 'Boolean') return [parameter.name, true];
  if (parameter.dataType === 'DateTime') return [parameter.name, '2026-01-01T00:00:00Z'];
  return [parameter.name, 'Sample'];
}));

let container;
let stagedFonts;
try {
  await run('docker', ['build', '-t', image, '.'], { maxBuffer: 16 * 1024 * 1024 });
  stagedFonts = await fs.mkdtemp(path.join(process.cwd(), '.docker-smoke-fonts-'));
  const available = await fs.readdir(fontHostDir);
  const families = [...new Set(['Arial', 'Times New Roman', ...model.fonts])];
  const expected = families.flatMap((family) => requiredVariants[family] || []);
  const unsupportedFamilies = families.filter((family) => !requiredVariants[family]);
  if (unsupportedFamilies.length) throw new Error(`Docker smoke has no licensed-font filename mapping for: ${unsupportedFamilies.join(', ')}`);
  const selected = expected.map((alternatives) => available.find((name) => alternatives.some((expectedName) => name.toLowerCase() === expectedName.toLowerCase()))).filter(Boolean);
  const missing = expected.filter((alternatives) => !available.some((name) => alternatives.some((expectedName) => name.toLowerCase() === expectedName.toLowerCase())));
  if (missing.length) throw new Error(`The font directory is missing required licensed variants: ${missing.map((alternatives) => alternatives.join(' or ')).join(', ')}`);
  await Promise.all(selected.map((name) => fs.copyFile(path.join(fontHostDir, name), path.join(stagedFonts, name))));
  const { stdout } = await run('docker', [
    'run', '--rm', '-d', '-p', '127.0.0.1::7070',
    '-v', `${stagedFonts}:/app/fonts:ro`, image,
  ]);
  container = stdout.trim();
  const portResult = await run('docker', ['port', container, '7070/tcp']);
  const port = portResult.stdout.trim().match(/:(\d+)$/)?.[1];
  if (!port) throw new Error('Docker did not publish the service port');
  const baseUrl = `http://127.0.0.1:${port}`;
  let ready = false;
  for (let attempt = 0; attempt < 40; attempt += 1) {
    try {
      const response = await fetch(`${baseUrl}/readyz`);
      if (response.ok) { ready = true; break; }
    } catch { /* container is still starting */ }
    await new Promise((resolve) => setTimeout(resolve, 250));
  }
  if (!ready) throw new Error('Container did not become ready');

  const results = [];
  for (const output of ['PDF', 'DOCX_EDITABLE', 'DOCX_VISUAL']) {
    const response = await fetch(`${baseUrl}/v1/render`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ fileName: path.basename(sample), rdlBase64: rdl.toString('base64'), output, outputFileName: 'docker-smoke', parameters, datasets }),
    });
    const artifact = Buffer.from(await response.arrayBuffer());
    if (!response.ok) throw new Error(`${output} failed: ${artifact.toString('utf8')}`);
    const expectedMagic = output === 'PDF' ? '%PDF' : 'PK';
    if (artifact.subarray(0, expectedMagic.length).toString() !== expectedMagic) throw new Error(`${output} returned an invalid artifact`);
    results.push({
      output,
      bytes: artifact.length,
      pageCount: response.headers.get('x-page-count'),
      layoutMode: response.headers.get('x-docx-layout-mode'),
      editableTextRatio: response.headers.get('x-docx-editable-text-ratio'),
    });
  }
  const leftovers = await run('docker', ['exec', container, 'find', '/tmp/rdl-converter', '-mindepth', '1', '-print']);
  if (leftovers.stdout.trim()) throw new Error(`Container temporary directory is not empty: ${leftovers.stdout.trim()}`);
  console.log(JSON.stringify({ ready: true, results, temporaryDirectoryEmpty: true }, null, 2));
} finally {
  if (container) await run('docker', ['rm', '-f', container]).catch(() => {});
  if (stagedFonts) await fs.rm(stagedFonts, { recursive: true, force: true });
}
