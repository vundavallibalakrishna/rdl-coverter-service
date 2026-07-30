import assert from 'node:assert/strict';
import test from 'node:test';
import { loadConfig } from '../src/config.js';
import { estimateWorkerMemory } from '../src/worker/runner.js';

const config = loadConfig({
  RDL_WORKER_MEMORY_MB: '512',
  RDL_WORKER_MEMORY_MAX_MB: '2048',
});

function rdl({ width = '11.69in', columns = 2 } = {}) {
  return Buffer.from(`<Report><PageWidth>${width}</PageWidth><TablixColumns>${
    '<TablixColumn><Width>1in</Width></TablixColumn>'.repeat(columns)
  }</TablixColumns></Report>`);
}

test('ordinary reports retain the baseline worker heap', () => {
  const estimate = estimateWorkerMemory(rdl(), {
    output: 'PDF',
    datasets: { D: [{ Value: 'ordinary' }] },
  }, config);
  assert.equal(estimate.memoryMb, 512);
  assert.equal(estimate.capped, false);
});

test('exceptionally wide PDF workloads scale deterministically up to the configured hard cap', () => {
  const estimate = estimateWorkerMemory(rdl({ width: '57.0866in', columns: 50 }), {
    output: 'PDF',
    datasets: { D: [{ Value: 'x'.repeat(650_000) }] },
  }, config);
  assert.equal(estimate.memoryMb, 2048);
  assert.equal(estimate.capped, true);
  assert.ok(estimate.uncappedMb > estimate.memoryMb);
  assert.ok(estimate.metrics.pageWidthPt > 4_000);
  assert.equal(estimate.metrics.maxTablixColumns, 50);
});

test('continuous Excel output does not scale merely because the print page is wide', () => {
  const estimate = estimateWorkerMemory(rdl({ width: '57.0866in', columns: 50 }), {
    output: 'XLSX',
    datasets: { D: [{ Value: 'ordinary' }] },
  }, config);
  assert.equal(estimate.memoryMb, 512);
});

test('the configured maximum is never exceeded and cannot be lower than the baseline', () => {
  const capped = loadConfig({
    RDL_WORKER_MEMORY_MB: '768',
    RDL_WORKER_MEMORY_MAX_MB: '256',
  });
  assert.equal(capped.workerMemoryMaxMb, 768);
  const estimate = estimateWorkerMemory(rdl({ width: '100in', columns: 100 }), {
    output: 'PDF',
    datasets: { D: [{ Value: 'x'.repeat(1_000_000) }] },
  }, capped);
  assert.equal(estimate.memoryMb, 768);
  assert.equal(estimate.capped, true);
});

test('page-locked DOCX estimates font, embedded-image, chart, and bundled-subreport working data', () => {
  const imageData = Buffer.alloc(6 * 1024 * 1024, 0x2A).toString('base64');
  const childRdl = Buffer.alloc(3 * 1024 * 1024, 0x3B).toString('base64');
  const source = Buffer.from(`<Report>
    <PageWidth>11.69in</PageWidth>
    <FontFamily>Arial</FontFamily>
    <EmbeddedImage><ImageData>${imageData}</ImageData></EmbeddedImage>
    <Image Name="Logo" />
    <Chart Name="Chart1" />
    <Chart Name="Chart2" />
  </Report>`);
  const estimate = estimateWorkerMemory(source, {
    output: 'DOCX_EDITABLE',
    datasets: {},
    subreports: {
      Child: { rdlBase64: childRdl, instances: [] },
    },
  }, config);
  assert.equal(estimate.metrics.internalCanonicalPdf, true);
  assert.equal(estimate.metrics.declaredFontFamilies, 1);
  assert.equal(estimate.metrics.embeddedFontVariants, 4);
  assert.ok(estimate.metrics.embeddedFontBytes > 0);
  assert.equal(estimate.metrics.embeddedImageBytes, 6 * 1024 * 1024);
  assert.equal(estimate.metrics.declaredImageItems, 1);
  assert.equal(estimate.metrics.declaredChartItems, 2);
  assert.equal(estimate.metrics.bundledRdlBytes, 3 * 1024 * 1024);
  assert.ok(estimate.metrics.binaryWorkingSetBytes > estimate.metrics.embeddedImageBytes);
  assert.ok(estimate.memoryMb > config.workerMemoryMb);
  assert.ok(estimate.memoryMb <= config.workerMemoryMaxMb);
});

test('page-locked DOCX estimates repeated native cell topology independently of text bytes', () => {
  const emptyRecord = Object.fromEntries(
    Array.from({ length: 20 }, (_, index) => [`Field${index + 1}`, '']),
  );
  const estimate = estimateWorkerMemory(rdl({ columns: 12 }), {
    output: 'DOCX_EDITABLE',
    datasets: {
      LongReport: Array.from({ length: 1_000 }, () => ({ ...emptyRecord })),
    },
  }, config);

  assert.equal(estimate.metrics.datasetTextBytes, 0);
  assert.equal(estimate.metrics.datasetValueCount, 20_000);
  assert.equal(estimate.memoryMb, 1792);
  assert.equal(estimate.capped, false);
});
