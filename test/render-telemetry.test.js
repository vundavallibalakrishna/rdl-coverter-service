import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { loadConfig } from '../src/config.js';
import { sanitizeTelemetryValue } from '../src/telemetry.js';
import { RenderRunner } from '../src/worker/runner.js';

const fixture = await fs.readFile(new URL('./fixtures/basic.rdl', import.meta.url));

function renderRequest(datasets) {
  return {
    fileName: 'telemetry.rdl',
    outputFileName: 'telemetry-output',
    output: 'PDF',
    parameters: { Title: 'DO_NOT_LOG_PARAMETER', Choice: 'A' },
    datasets,
  };
}

async function runnerForTest(context) {
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-telemetry-test-'));
  const config = loadConfig({
    ...process.env,
    RDL_TEMP_ROOT: tempRoot,
    RDL_STRICT_FONTS: 'false',
    RDL_RENDER_TIMEOUT_MS: '30000',
  });
  const runner = new RenderRunner(config);
  context.after(async () => {
    await runner.shutdown();
    await fs.rm(tempRoot, { recursive: true, force: true });
  });
  return { runner, tempRoot };
}

test('render telemetry reports bounded structural phases without report or dataset content', async (context) => {
  const { runner, tempRoot } = await runnerForTest(context);
  const events = [];
  const rendered = await runner.render({
    rdlBuffer: fixture,
    request: renderRequest({ Sales: [{ Name: 'DO_NOT_LOG_DATASET_VALUE', Amount: 10 }] }),
    onTelemetry: (event) => events.push(event),
  });

  assert.equal(rendered.buffer.subarray(0, 4).toString(), '%PDF');
  assert.ok(events.some(({ source, phase }) => source === 'runner' && phase === 'memory-estimated'));
  assert.ok(events.some(({ source, phase }) => source === 'worker' && phase === 'rdl-parsed'));
  assert.ok(events.some(({ source, phase }) => source === 'worker' && phase === 'input-validated'));
  assert.ok(events.some(({ source, phase }) => source === 'worker' && phase === 'pdf.body-layout-completed'));
  assert.ok(events.some(({ source, phase }) => source === 'worker' && phase === 'pdf.page-bands-completed'));
  assert.ok(events.some(({ source, phase }) => source === 'worker' && phase === 'pdf.serialized'));
  assert.ok(events.some(({ source, phase }) => source === 'worker' && phase === 'pdf.validated'));
  assert.ok(events.some(({ source, phase }) => source === 'worker' && phase === 'renderer-completed'));
  assert.ok(events.some(({ source, phase }) => source === 'runner' && phase === 'cleanup-completed'));
  const pdfBody = events.find(({ source, phase }) => source === 'worker' && phase === 'pdf.body-layout-completed');
  assert.equal(pdfBody.metrics.optimizationsEnabled, true);
  assert.equal(Number.isInteger(pdfBody.metrics.rowMeasurementRequests), true);
  assert.equal(Number.isInteger(pdfBody.metrics.rowMeasurementsComputed), true);
  assert.equal(Number.isInteger(pdfBody.metrics.rowMeasurementCacheHits), true);
  const rendererCompleted = events.find(({ source, phase }) => source === 'worker' && phase === 'renderer-completed');
  assert.equal(rendererCompleted.metrics.expressionPlanCache.enabled, true);
  assert.equal(Number.isInteger(rendererCompleted.metrics.expressionPlanCache.entries), true);
  assert.equal(Number.isInteger(rendererCompleted.metrics.expressionPlanCache.hits), true);
  assert.equal(Number.isInteger(rendererCompleted.metrics.expressionPlanCache.misses), true);
  assert.equal(rendererCompleted.metrics.pdfFontSelectionCache.enabled, true);
  assert.equal(Number.isInteger(rendererCompleted.metrics.pdfFontSelectionCache.entries), true);
  assert.equal(Number.isInteger(rendererCompleted.metrics.pdfFontSelectionCache.hits), true);
  assert.equal(Number.isInteger(rendererCompleted.metrics.pdfFontSelectionCache.misses), true);
  for (const event of events) {
    assert.equal(Number.isFinite(event.phaseDurationMs), true);
    assert.equal(Number.isFinite(event.totalDurationMs), true);
    assert.equal(Number.isFinite(event.rssMb), true);
    assert.equal(Number.isFinite(event.cpuUserMs), true);
  }
  const serialized = JSON.stringify(events);
  assert.equal(serialized.includes('DO_NOT_LOG_PARAMETER'), false);
  assert.equal(serialized.includes('DO_NOT_LOG_DATASET_VALUE'), false);
  assert.equal(serialized.includes('<Report'), false);
  assert.deepEqual(await fs.readdir(tempRoot), []);
});

test('render telemetry records sanitized failure and cleanup phases', async (context) => {
  const { runner, tempRoot } = await runnerForTest(context);
  const events = [];
  await assert.rejects(
    runner.render({
      rdlBuffer: fixture,
      request: renderRequest({}),
      onTelemetry: (event) => events.push(event),
    }),
    (error) => error.code === 'DATASET_MISSING',
  );
  assert.ok(events.some(({ source, phase, status }) => source === 'worker' && phase === 'failed' && status === 'failed'));
  assert.ok(events.some(({ source, phase, status }) => source === 'runner' && phase === 'render-failed' && status === 'failed'));
  assert.ok(events.some(({ source, phase }) => source === 'runner' && phase === 'cleanup-completed'));
  assert.deepEqual(await fs.readdir(tempRoot), []);
});

test('telemetry sanitizer bounds strings and removes unsupported values', () => {
  const safe = sanitizeTelemetryValue({
    text: `line\n${'x'.repeat(200)}`,
    valid: 12,
    invalid: Number.NaN,
    callback: () => {},
  });
  assert.equal(safe.text.includes('\n'), false);
  assert.equal(safe.text.length, 80);
  assert.equal(safe.valid, 12);
  assert.equal('invalid' in safe, false);
  assert.equal('callback' in safe, false);
});
