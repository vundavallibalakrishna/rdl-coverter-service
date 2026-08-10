import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import test from 'node:test';

const shell = await fs.readFile(new URL('../start-production.sh', import.meta.url), 'utf8');
const batch = await fs.readFile(new URL('../start-production.bat', import.meta.url), 'utf8');
const bootstrap = await fs.readFile(new URL('../scripts/start-production.js', import.meta.url), 'utf8');
const envExample = await fs.readFile(new URL('../.env.example', import.meta.url), 'utf8');

const requiredProductionSettings = [
  'NODE_ENV',
  'LOG_LEVEL',
  'HOST',
  'PORT',
  'RDL_TEMP_ROOT',
  'RDL_FONT_DIR',
  'RDL_STRICT_FONTS',
  'RDL_ALLOW_COMPATIBLE_FONT_FALLBACKS',
  'RDL_MAX_RDL_BYTES',
  'RDL_MAX_REQUEST_BYTES',
  'RDL_MAX_ROWS',
  'RDL_MAX_CONCURRENCY',
  'RDL_RENDER_TIMEOUT_MS',
  'RDL_WORKER_MEMORY_MB',
  'RDL_WORKER_MEMORY_MAX_MB',
  'RDL_MAX_XML_NODES',
  'RDL_MAX_XML_DEPTH',
  'RDL_PDFTOPPM_PATH',
  'RDL_PDF_LAYOUT_OPTIMIZATIONS',
  'RDL_EXPRESSION_PLAN_CACHE',
  'RDL_PDF_FONT_SELECTION_CACHE',
  'RDL_BORDER_WIDTH_FLOOR_PT',
];

test('production bootstrap exposes every supported production runtime setting', () => {
  for (const setting of requiredProductionSettings) {
    assert.match(bootstrap, new RegExp(`\\b${setting}\\b`), `${setting} is missing from the production bootstrap`);
  }
});

test('both platform launchers invoke the same production bootstrap', () => {
  assert.match(shell, /node scripts\/start-production\.js/);
  assert.match(batch, /node scripts\\start-production\.js/i);
});

test('production bootstrap configures isolated workers rather than inflating the server heap', () => {
  assert.match(bootstrap, /RDL_WORKER_MEMORY_MB:\s*'512'/);
  assert.match(bootstrap, /RDL_WORKER_MEMORY_MAX_MB:\s*'8192'/);
  assert.match(bootstrap, /RDL_MAX_CONCURRENCY:\s*'1'/);
  assert.match(bootstrap, /RDL_RENDER_TIMEOUT_MS:\s*'300000'/);
  assert.doesNotMatch(`${shell}\n${batch}\n${bootstrap}`, /node\s+--max-old-space-size/i);
});

test('production bootstrap loads an optional environment file before applying missing defaults', () => {
  const loadIndex = bootstrap.indexOf('loadEnvFile(envFile)');
  const defaultIndex = bootstrap.indexOf('for (const [name, value] of Object.entries(defaults))');
  assert.ok(loadIndex >= 0);
  assert.ok(defaultIndex > loadIndex);
  assert.match(bootstrap, /process\.env\.RDL_ENV_FILE/);
  assert.match(bootstrap, /\.env\.production/);
  assert.match(bootstrap, /process\.env\[name\] === undefined/);
});

test('the copyable environment example retains the large-report production safety profile', () => {
  assert.match(envExample, /^NODE_ENV=production$/m);
  assert.match(envExample, /^RDL_MAX_CONCURRENCY=1$/m);
  assert.match(envExample, /^RDL_RENDER_TIMEOUT_MS=300000$/m);
  assert.match(envExample, /^RDL_WORKER_MEMORY_MB=512$/m);
  assert.match(envExample, /^RDL_WORKER_MEMORY_MAX_MB=8192$/m);
  assert.match(envExample, /^RDL_STRICT_FONTS=true$/m);
});
