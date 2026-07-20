import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { analyzeRdl, parseRdl } from '../src/rdl/parser.js';
import { loadConfig } from '../src/config.js';
import { RenderRunner } from '../src/worker/runner.js';

const serviceRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const sample = process.argv[2] || path.join(serviceRoot, '..', 'lumina', 'Combined Assurance Reports (2).rdl');
const buffer = await fs.readFile(path.resolve(sample));
const analysis = analyzeRdl(buffer);
const model = parseRdl(buffer);
const rowFor = (dataset) => Object.fromEntries(dataset.fields.map((field, index) => [
  field.dataField,
  /Int|Decimal|Double/i.test(field.typeName) ? index + 1 : `Sample ${field.dataField}`,
]));
const datasets = Object.fromEntries(model.datasets
  .filter((dataset) => model.renderingDatasets.includes(dataset.name))
  .map((dataset) => [dataset.name, [rowFor(dataset)]]));
const parameters = Object.fromEntries(model.parameters.map((parameter) => {
  if (parameter.multiValue) return [parameter.name, ['Sample']];
  if (parameter.dataType === 'Integer') return [parameter.name, 1];
  if (parameter.dataType === 'Float') return [parameter.name, 1.5];
  if (parameter.dataType === 'Boolean') return [parameter.name, true];
  if (parameter.dataType === 'DateTime') return [parameter.name, '2026-01-01T00:00:00Z'];
  return [parameter.name, 'Sample'];
}));
const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-sample-smoke-'));
const config = loadConfig({ ...process.env, RDL_TEMP_ROOT: tempRoot });
const runner = new RenderRunner(config);
const outputs = [];
try {
  for (const output of ['PDF', 'DOCX_EDITABLE', 'DOCX_VISUAL']) {
    const rendered = await runner.render({
      rdlBuffer: buffer,
      request: { output, outputFileName: 'combined-assurance-smoke', parameters, datasets },
    });
    outputs.push({
      output,
      bytes: rendered.buffer.length,
      pageCount: rendered.pageCount,
      layoutMode: rendered.layoutMode,
      editableTextRatio: rendered.editableTextRatio,
    });
  }
  const leftovers = await fs.readdir(tempRoot);
  if (leftovers.length > 0) throw new Error(`Temporary files were not removed: ${leftovers.join(', ')}`);
} finally {
  await runner.shutdown();
  await fs.rm(tempRoot, { recursive: true, force: true });
}

console.log(JSON.stringify({
  namespace: analysis.namespace,
  page: {
    width: analysis.page.width,
    height: analysis.page.height,
    marginTop: analysis.page.marginTop,
    marginRight: analysis.page.marginRight,
    marginBottom: analysis.page.marginBottom,
    marginLeft: analysis.page.marginLeft,
    headerHeight: analysis.page.header?.height || 0,
    footerHeight: analysis.page.footer?.height || 0,
  },
  parameters: analysis.parameters.map(({ name, dataType, multiValue }) => ({ name, dataType, multiValue })),
  datasets: analysis.datasets.map(({ name, fields, requiredForRendering, parameterOnly }) => ({ name, fields: fields.map((field) => field.dataField), requiredForRendering, parameterOnly })),
  fonts: analysis.fonts,
  features: analysis.features,
  compatible: analysis.compatible,
  unsupported: model.unsupported,
  outputs,
}, null, 2));
