#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import process, { loadEnvFile } from 'node:process';
import { fileURLToPath } from 'node:url';

const serviceRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
process.chdir(serviceRoot);

// Host/service environment is authoritative. Node's loadEnvFile preserves already-defined process values,
// so an optional file fills the next precedence level without replacing settings injected by IIS, systemd,
// Docker, or an operator's shell. Defaults below apply only after both sources have been considered.
const envFile = path.resolve(process.env.RDL_ENV_FILE || path.join(serviceRoot, '.env.production'));
let envFileLoaded = false;
if (fs.existsSync(envFile)) {
  try {
    loadEnvFile(envFile);
    envFileLoaded = true;
  } catch (error) {
    console.error(`Unable to load production environment file ${envFile}: ${error.message}`);
    process.exit(1);
  }
}

const defaults = {
  NODE_ENV: 'production',
  LOG_LEVEL: 'info',
  HOST: '0.0.0.0',
  PORT: '7070',
  RDL_TEMP_ROOT: path.join(process.env.TMPDIR || process.env.TEMP || process.env.TMP || '/tmp', 'rdl-converter'),
  RDL_FONT_DIR: path.join(serviceRoot, 'fonts'),
  RDL_STRICT_FONTS: 'true',
  RDL_ALLOW_COMPATIBLE_FONT_FALLBACKS: 'false',
  RDL_MAX_RDL_BYTES: '10485760',
  RDL_MAX_REQUEST_BYTES: '26214400',
  RDL_MAX_ROWS: '100000',
  RDL_MAX_CONCURRENCY: '1',
  RDL_RENDER_TIMEOUT_MS: '300000',
  RDL_WORKER_MEMORY_MB: '512',
  RDL_WORKER_MEMORY_MAX_MB: '8192',
  RDL_MAX_XML_NODES: '250000',
  RDL_MAX_XML_DEPTH: '256',
  RDL_PDFTOPPM_PATH: 'pdftoppm',
  RDL_PDF_LAYOUT_OPTIMIZATIONS: 'true',
  RDL_EXPRESSION_PLAN_CACHE: 'true',
  RDL_PDF_FONT_SELECTION_CACHE: 'true',
  RDL_BORDER_WIDTH_FLOOR_PT: '0',
};

for (const [name, value] of Object.entries(defaults)) {
  if (process.env[name] === undefined || process.env[name] === '') process.env[name] = value;
}

console.log('Starting RDL Converter Service in production mode');
console.log(`EnvironmentFile=${envFileLoaded ? envFile : 'not found; using environment and production defaults'}`);
console.log(`Host=${process.env.HOST} Port=${process.env.PORT} Concurrency=${process.env.RDL_MAX_CONCURRENCY} WorkerHeap=${process.env.RDL_WORKER_MEMORY_MB}-${process.env.RDL_WORKER_MEMORY_MAX_MB}MB Timeout=${process.env.RDL_RENDER_TIMEOUT_MS}ms`);
console.log(`Fonts=${process.env.RDL_FONT_DIR} Poppler=${process.env.RDL_PDFTOPPM_PATH} Temp=${process.env.RDL_TEMP_ROOT} LogLevel=${process.env.LOG_LEVEL}`);

// Render workers receive their own workload-sized --max-old-space-size in src/worker/runner.js. A heap
// argument on this main server process would not raise the isolated worker limit.
await import('../src/server.js');
