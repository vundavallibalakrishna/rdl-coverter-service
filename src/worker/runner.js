import fs from 'node:fs/promises';
import { statSync } from 'node:fs';
import path from 'node:path';
import { fork } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { ServiceError } from '../errors.js';
import { toPoints } from '../units.js';
import { resolveFontFile } from '../render/fonts.js';

const workerPath = fileURLToPath(new URL('./renderWorker.js', import.meta.url));
const MEMORY_STEP_MB = 128;
const REFERENCE_PAGE_WIDTH_PT = 11.69 * 72; // landscape A4/Letter-class report width
const REFERENCE_COLUMNS = 20;
const REFERENCE_TEXT_BYTES = 384 * 1024;
const REFERENCE_DATA_VALUES = 10_000;
const REFERENCE_BINARY_WORKING_SET_BYTES = 32 * 1024 * 1024;
const FALLBACK_FONT_VARIANT_BYTES = 2 * 1024 * 1024;
const CHART_RASTER_WORKING_BYTES = 8 * 1024 * 1024;
const IMAGE_RASTER_WORKING_BYTES = 4 * 1024 * 1024;
const FONT_VARIANTS = [
  [false, false],
  [true, false],
  [false, true],
  [true, true],
];

function datasetTextBytes(request) {
  let bytes = 0;
  const measureDatasets = (datasets) => {
    for (const rows of Object.values(datasets || {})) {
      if (!Array.isArray(rows)) continue;
      for (const row of rows) {
        if (!row || typeof row !== 'object') continue;
        for (const value of Object.values(row)) {
          if (value === null || value === undefined) continue;
          bytes += Buffer.byteLength(typeof value === 'string' ? value : JSON.stringify(value));
        }
      }
    }
  };
  measureDatasets(request?.datasets);
  for (const definition of Object.values(request?.subreports || {})) {
    for (const instance of definition?.instances || []) measureDatasets(instance?.datasets);
  }
  return bytes;
}

function datasetValueCount(request) {
  let count = 0;
  const measureDatasets = (datasets) => {
    for (const rows of Object.values(datasets || {})) {
      if (!Array.isArray(rows)) continue;
      for (const row of rows) {
        if (!row || typeof row !== 'object') continue;
        // Every supplied field participates in validation/materialization and may become repeated native
        // PDF-trace and OOXML cell state even when its display value is empty. Text bytes alone therefore
        // cannot predict the structural working set of a long editable Word report.
        count += Object.keys(row).length;
      }
    }
  };
  measureDatasets(request?.datasets);
  for (const definition of Object.values(request?.subreports || {})) {
    for (const instance of definition?.instances || []) measureDatasets(instance?.datasets);
  }
  return count;
}

function maximumTablixColumns(xml) {
  let maximum = 0;
  for (const match of xml.matchAll(/<TablixColumns\b[^>]*>([\s\S]*?)<\/TablixColumns>/gi)) {
    maximum = Math.max(maximum, (match[1].match(/<TablixColumn\b/g) || []).length);
  }
  return maximum;
}

function declaredFontFamilies(xml) {
  return new Set(
    [...xml.matchAll(/<FontFamily\b[^>]*>\s*([^<=][^<]*?)\s*<\/FontFamily>/gi)]
      .map((match) => match[1].trim())
      .filter(Boolean),
  );
}

function decodedBase64Bytes(value) {
  const source = String(value || '').replace(/\s+/g, '');
  if (!source) return 0;
  const padding = source.endsWith('==') ? 2 : source.endsWith('=') ? 1 : 0;
  return Math.max(0, Math.floor(source.length * 3 / 4) - padding);
}

function embeddedImageBytes(xml) {
  let bytes = 0;
  for (const match of xml.matchAll(/<ImageData\b[^>]*>([\s\S]*?)<\/ImageData>/gi)) {
    bytes += decodedBase64Bytes(match[1]);
  }
  return bytes;
}

function bundledRdlBytes(request) {
  let bytes = 0;
  for (const definition of Object.values(request?.subreports || {})) {
    bytes += decodedBase64Bytes(definition?.rdlBase64);
  }
  return bytes;
}

function declaredItemCount(xml, localName) {
  return (xml.match(new RegExp(`<${localName}\\b`, 'gi')) || []).length;
}

function embeddedFontBytes(config, families) {
  let bytes = 0;
  for (const family of families) {
    for (const [bold, italic] of FONT_VARIANTS) {
      const file = resolveFontFile(config.fontDir, family, bold, italic);
      if (!file) {
        bytes += FALLBACK_FONT_VARIANT_BYTES;
        continue;
      }
      try {
        bytes += statSync(file).size;
      } catch {
        bytes += FALLBACK_FONT_VARIANT_BYTES;
      }
    }
  }
  return bytes;
}

// Selects a heap before the worker starts, without parsing/executing the RDL in the main process. The
// workload factors are format geometry, maximum native grid width, caller-supplied text volume, and the
// extra canonical-PDF/layout-trace/OOXML/font buffers held by Windows page-locked editable DOCX. The result
// is always capped by the configured per-worker maximum, preserving a hard memory bound.
export function estimateWorkerMemory(rdlBuffer, request, config) {
  const xml = Buffer.isBuffer(rdlBuffer) ? rdlBuffer.toString('utf8') : String(rdlBuffer || '');
  const pageWidthText = xml.match(/<PageWidth\b[^>]*>\s*([^<]+?)\s*<\/PageWidth>/i)?.[1] || '11.69in';
  const pageWidthPt = toPoints(pageWidthText, REFERENCE_PAGE_WIDTH_PT);
  const maxTablixColumns = maximumTablixColumns(xml);
  const textBytes = datasetTextBytes(request);
  const dataValues = datasetValueCount(request);
  const output = String(request?.output || '').toUpperCase();
  const pageLockedDocx = output === 'DOCX_EDITABLE';
  const producesPdf = output === 'PDF' || output === 'DOCX_VISUAL' || pageLockedDocx;
  const fontFamilies = declaredFontFamilies(xml);
  const imageBytes = embeddedImageBytes(xml);
  const bundledBytes = bundledRdlBytes(request);
  const imageItems = declaredItemCount(xml, 'Image');
  const chartItems = declaredItemCount(xml, 'Chart');
  const fontBytes = pageLockedDocx ? embeddedFontBytes(config, fontFamilies) : 0;
  // DOCX creation holds decoded image bytes, raster working buffers, font parts, canonical PDF bytes,
  // layout-trace nodes, and the compressed OOXML package concurrently. Count the binary sources
  // independently instead of assuming dataset text predicts them.
  const binaryWorkingSetBytes = pageLockedDocx
    ? imageBytes
      + bundledBytes
      + fontBytes
      + imageItems * IMAGE_RASTER_WORKING_BYTES
      + chartItems * CHART_RASTER_WORKING_BYTES
    : 0;
  const baseScale = Math.max(
    1,
    textBytes / REFERENCE_TEXT_BYTES,
    producesPdf ? pageWidthPt / REFERENCE_PAGE_WIDTH_PT : 1,
    producesPdf ? maxTablixColumns / REFERENCE_COLUMNS : 1,
    pageLockedDocx ? dataValues / REFERENCE_DATA_VALUES : 1,
    pageLockedDocx ? binaryWorkingSetBytes / REFERENCE_BINARY_WORKING_SET_BYTES : 1,
  );
  // PDF, trace nodes, the native Word page grid, JSZip, and full font variants coexist during packaging.
  // This is a workload-class multiplier, not a report identity/profile and therefore applies uniformly.
  const docxPackagingScale = pageLockedDocx ? 1.75 + Math.min(0.75, fontFamilies.size * 0.05) : 1;
  const scale = baseScale * docxPackagingScale;
  const uncappedMb = Math.ceil((config.workerMemoryMb * scale) / MEMORY_STEP_MB) * MEMORY_STEP_MB;
  const memoryMb = Math.min(config.workerMemoryMaxMb || config.workerMemoryMb, Math.max(config.workerMemoryMb, uncappedMb));
  return {
    memoryMb,
    uncappedMb,
    capped: memoryMb < uncappedMb,
    metrics: {
      output,
      pageWidthPt,
      maxTablixColumns,
      datasetTextBytes: textBytes,
      datasetValueCount: dataValues,
      declaredFontFamilies: fontFamilies.size,
      embeddedFontVariants: pageLockedDocx ? fontFamilies.size * 4 : 0,
      embeddedFontBytes: fontBytes,
      embeddedImageBytes: imageBytes,
      declaredImageItems: imageItems,
      declaredChartItems: chartItems,
      chartRasterWorkingBytes: pageLockedDocx ? chartItems * CHART_RASTER_WORKING_BYTES : 0,
      imageRasterWorkingBytes: pageLockedDocx ? imageItems * IMAGE_RASTER_WORKING_BYTES : 0,
      bundledRdlBytes: bundledBytes,
      binaryWorkingSetBytes,
      internalCanonicalPdf: pageLockedDocx,
      docxPackagingScale,
      scale,
    },
  };
}

export class RenderRunner {
  constructor(config) {
    this.config = config;
    this.active = new Set();
    this.inFlight = 0;
    this.shuttingDown = false;
  }

  async render({ rdlBuffer, request, signal }) {
    if (this.shuttingDown) throw new ServiceError('BUSY', 'Service is shutting down', 503);
    if (this.inFlight >= this.config.maxConcurrency) throw new ServiceError('BUSY', 'Render capacity is currently full', 503, { retryAfterSeconds: 5 });
    this.inFlight += 1;
    let tempDir;
    let child;
    try {
      tempDir = await fs.mkdtemp(path.join(this.config.tempRoot, 'request-'));
      await fs.chmod(tempDir, 0o700);
      const rdlPath = path.join(tempDir, 'input.rdl');
      const requestPath = path.join(tempDir, 'request.json');
      await Promise.all([
        fs.writeFile(rdlPath, rdlBuffer, { mode: 0o600 }),
        fs.writeFile(requestPath, JSON.stringify(request), { mode: 0o600 }),
      ]);

      const memoryEstimate = estimateWorkerMemory(rdlBuffer, request, this.config);
      child = fork(workerPath, [], {
        stdio: ['ignore', 'ignore', 'ignore', 'ipc'],
        env: process.env,
        execArgv: [`--max-old-space-size=${memoryEstimate.memoryMb}`],
      });
      this.active.add(child);
      const metadata = await new Promise((resolve, reject) => {
        let settled = false;
        const finish = (callback, value) => {
          if (settled) return;
          settled = true;
          clearTimeout(timeout);
          signal?.removeEventListener('abort', abort);
          callback(value);
        };
        const timeout = setTimeout(() => {
          child.kill('SIGKILL');
          finish(reject, new ServiceError('RENDER_TIMEOUT', 'Rendering exceeded the configured timeout', 504));
        }, this.config.renderTimeoutMs);
        const abort = () => {
          child.kill('SIGKILL');
          finish(reject, new ServiceError('RENDER_FAILED', 'Client disconnected before rendering completed', 499));
        };
        signal?.addEventListener('abort', abort, { once: true });
        if (signal?.aborted) abort();
        child.once('error', () => finish(reject, new ServiceError('RENDER_FAILED', 'Render worker could not be started', 500)));
        child.once('exit', (code, exitSignal) => {
          if (!settled && code !== 0) finish(reject, new ServiceError('RENDER_FAILED', `Render worker stopped unexpectedly (${exitSignal || code})`, 500));
        });
        child.on('message', (message) => {
          if (message?.type === 'failed') finish(reject, new ServiceError(message.error.code, message.error.message, message.error.statusCode, message.error.details));
          if (message?.type === 'completed') finish(resolve, message);
        });
        child.send({ type: 'render', tempDir, rdlPath, requestPath, config: this.config });
      });
      const buffer = await fs.readFile(metadata.outputPath);
      return { ...metadata, buffer, workerMemoryMb: memoryEstimate.memoryMb, workerMemoryEstimate: memoryEstimate };
    } finally {
      if (child?.connected) child.disconnect();
      if (child && !child.killed) child.kill('SIGTERM');
      if (child) this.active.delete(child);
      this.inFlight -= 1;
      if (tempDir) await fs.rm(tempDir, { recursive: true, force: true });
    }
  }

  async shutdown() {
    this.shuttingDown = true;
    for (const child of this.active) child.kill('SIGTERM');
    await Promise.allSettled([...this.active].map((child) => new Promise((resolve) => child.once('exit', resolve))));
    this.active.clear();
  }
}
