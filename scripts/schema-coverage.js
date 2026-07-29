import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { rdl2016SchemaCatalogue } from '../src/rdl/capabilities.js';

const serviceRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const outputPath = path.join(serviceRoot, 'tmp', 'rdl-2016-capability-catalogue.json');
const catalogue = rdl2016SchemaCatalogue();

await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.writeFile(outputPath, `${JSON.stringify(catalogue, null, 2)}\n`, { mode: 0o600 });
process.stdout.write(`${JSON.stringify({ outputPath, ...catalogue.summary, total: catalogue.entries.length }, null, 2)}\n`);
