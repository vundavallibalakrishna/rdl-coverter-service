import fs from 'node:fs/promises';
import { statSync } from 'node:fs';
import path from 'node:path';
import { fork } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { ServiceError } from '../errors.js';
import { toPoints } from '../units.js';
import { resolveFontFile } from '../render/fonts.js';
import { createTelemetryClock } from '../telemetry.js';

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
const MAX_WORKER_STDERR_BYTES = 32 * 1024;
const FONT_VARIANTS = [
  [false, false],
  [true, false],
  [false, true],
  [true, true],
];

// A render worker can die below JavaScript's catch boundary (V8 heap exhaustion, a native assertion, or
// process.abort()). Its stderr may contain file paths and native stack text, so never forward or log it.
// Reduce the bounded tail to a stable category that is safe for telemetry and operational diagnostics.
export function classifyWorkerFatalStderr(stderr, exitCode = null, exitSignal = null) {
  const source = String(stderr || '');
  if (/heap out of memory|reached heap limit|allocation failed[^\r\n]*javascript heap/i.test(source)) {
    return 'V8_HEAP_OUT_OF_MEMORY';
  }
  if (/fatal process out of memory|process out of memory/i.test(source)) return 'V8_HEAP_OUT_OF_MEMORY';
  if (/check failed|assertion failed|fatal error/i.test(source)) return 'NATIVE_RUNTIME_ABORT';
  if (Number(exitCode) === 134 || /ABRT/i.test(String(exitSignal || ''))) return 'PROCESS_ABORT';
  return 'WORKER_EXIT';
}

function appendWorkerStderr(state, chunk) {
  const bytes = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
  state.bytes += bytes.length;
  state.tail = Buffer.concat([state.tail, bytes]);
  if (state.tail.length > MAX_WORKER_STDERR_BYTES) {
    state.tail = state.tail.subarray(state.tail.length - MAX_WORKER_STDERR_BYTES);
    state.truncated = true;
  }
}

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

  async render({ rdlBuffer, request, signal, onTelemetry }) {
    if (this.shuttingDown) throw new ServiceError('BUSY', 'Service is shutting down', 503);
    if (this.inFlight >= this.config.maxConcurrency) throw new ServiceError('BUSY', 'Render capacity is currently full', 503, { retryAfterSeconds: 5 });
    this.inFlight += 1;
    const telemetry = createTelemetryClock('runner', onTelemetry);
    const output = String(request?.output || '').toUpperCase().slice(0, 32);
    const inputMetrics = {
      output,
      rdlBytes: rdlBuffer.length,
      datasetCount: Object.keys(request?.datasets || {}).length,
      datasetTextBytes: datasetTextBytes(request),
      datasetValueCount: datasetValueCount(request),
      bundledSubreportCount: Object.keys(request?.subreports || {}).length,
      inFlight: this.inFlight,
      maxConcurrency: this.config.maxConcurrency,
    };
    telemetry.mark('admitted', 'completed', inputMetrics);
    let tempDir;
    let child;
    let outcome = 'failed';
    try {
      tempDir = await fs.mkdtemp(path.join(this.config.tempRoot, 'request-'));
      await fs.chmod(tempDir, 0o700);
      telemetry.mark('temporary-storage-prepared');
      const rdlPath = path.join(tempDir, 'input.rdl');
      const requestPath = path.join(tempDir, 'request.json');
      const requestJson = JSON.stringify(request);
      await Promise.all([
        fs.writeFile(rdlPath, rdlBuffer, { mode: 0o600 }),
        fs.writeFile(requestPath, requestJson, { mode: 0o600 }),
      ]);
      telemetry.mark('inputs-written', 'completed', { requestBytes: Buffer.byteLength(requestJson) });

      const memoryEstimate = estimateWorkerMemory(rdlBuffer, request, this.config);
      telemetry.mark('memory-estimated', 'completed', {
        workerMemoryMb: memoryEstimate.memoryMb,
        uncappedWorkerMemoryMb: memoryEstimate.uncappedMb,
        memoryEstimateCapped: memoryEstimate.capped,
        ...memoryEstimate.metrics,
      });
      child = fork(workerPath, [], {
        // stdout remains discarded because render code must never log report content. stderr is retained
        // only as a bounded private tail and reduced to a fixed fatal category; raw bytes never leave this
        // function or enter application logs.
        stdio: ['ignore', 'ignore', 'pipe', 'ipc'],
        env: process.env,
        execArgv: [`--max-old-space-size=${memoryEstimate.memoryMb}`],
      });
      const workerStderr = { tail: Buffer.alloc(0), bytes: 0, truncated: false };
      child.stderr?.on('data', (chunk) => appendWorkerStderr(workerStderr, chunk));
      this.active.add(child);
      telemetry.mark('worker-created', 'completed', { workerPid: child.pid, workerMemoryMb: memoryEstimate.memoryMb });
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
          telemetry.mark('worker-timeout', 'failed', { timeoutMs: this.config.renderTimeoutMs });
          finish(reject, new ServiceError('RENDER_TIMEOUT', 'Rendering exceeded the configured timeout', 504));
        }, this.config.renderTimeoutMs);
        const abort = () => {
          child.kill('SIGKILL');
          telemetry.mark('client-aborted', 'failed');
          finish(reject, new ServiceError('RENDER_FAILED', 'Client disconnected before rendering completed', 499));
        };
        signal?.addEventListener('abort', abort, { once: true });
        if (signal?.aborted) abort();
        child.once('error', () => {
          telemetry.mark('worker-start-error', 'failed');
          finish(reject, new ServiceError('RENDER_FAILED', 'Render worker could not be started', 500));
        });
        // `close` follows stderr closure, unlike `exit`, so the fatal classifier sees V8's complete final
        // message on Windows as well as Unix-like hosts.
        child.once('close', (code, exitSignal) => {
          if (!settled && code !== 0) {
            const fatalCategory = classifyWorkerFatalStderr(workerStderr.tail.toString('utf8'), code, exitSignal);
            telemetry.mark('worker-exited', 'failed', {
              exitCode: code,
              exitSignal: exitSignal || undefined,
              fatalCategory,
              stderrBytes: workerStderr.bytes,
              stderrTruncated: workerStderr.truncated,
              workerMemoryMb: memoryEstimate.memoryMb,
              uncappedWorkerMemoryMb: memoryEstimate.uncappedMb,
              memoryEstimateCapped: memoryEstimate.capped,
            });
            const failure = new ServiceError('RENDER_FAILED', `Render worker stopped unexpectedly (${exitSignal || code})`, 500);
            failure.diagnostic = {
              name: 'RenderWorkerExit',
              message: fatalCategory,
              exitCode: code,
              exitSignal: exitSignal || undefined,
              workerMemoryMb: memoryEstimate.memoryMb,
              uncappedWorkerMemoryMb: memoryEstimate.uncappedMb,
              memoryEstimateCapped: memoryEstimate.capped,
              stderrBytes: workerStderr.bytes,
              stderrTruncated: workerStderr.truncated,
            };
            finish(reject, failure);
          }
        });
        child.on('message', (message) => {
          if (message?.type === 'telemetry') {
            try { onTelemetry?.(message.event); } catch { /* Logging cannot affect rendering. */ }
            return;
          }
          if (message?.type === 'failed') {
            const failure = new ServiceError(message.error.code, message.error.message, message.error.statusCode, message.error.details);
            // Server-side only: the underlying exception behind a scrubbed RENDER_FAILED (see renderWorker).
            if (message.diagnostic) failure.diagnostic = message.diagnostic;
            finish(reject, failure);
          }
          if (message?.type === 'completed') finish(resolve, message);
        });
        child.send({ type: 'render', tempDir, rdlPath, requestPath, config: this.config });
      });
      telemetry.mark('worker-completed', 'completed', {
        pageCount: metadata.pageCount,
        outputBytes: metadata.size,
        totalRows: metadata.totalRows,
        sheetCount: metadata.sheetCount,
        workbookRowCount: metadata.rowCount,
      });
      const buffer = await fs.readFile(metadata.outputPath);
      telemetry.mark('artifact-read', 'completed', { outputBytes: buffer.length });
      outcome = 'completed';
      return { ...metadata, buffer, workerMemoryMb: memoryEstimate.memoryMb, workerMemoryEstimate: memoryEstimate };
    } catch (error) {
      telemetry.mark('render-failed', 'failed', {
        errorCode: error?.code || 'RENDER_FAILED',
        statusCode: error?.statusCode || 500,
      });
      throw error;
    } finally {
      if (child?.connected) child.disconnect();
      if (child && !child.killed) child.kill('SIGTERM');
      if (child) this.active.delete(child);
      this.inFlight -= 1;
      if (tempDir) await fs.rm(tempDir, { recursive: true, force: true });
      telemetry.mark('cleanup-completed', outcome, { inFlight: this.inFlight });
    }
  }

  async shutdown() {
    this.shuttingDown = true;
    for (const child of this.active) child.kill('SIGTERM');
    await Promise.allSettled([...this.active].map((child) => new Promise((resolve) => child.once('exit', resolve))));
    this.active.clear();
  }
}
