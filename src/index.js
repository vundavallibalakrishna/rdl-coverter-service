// Public library entry point. Everything the HTTP service does is available here so that an in-process
// caller gets the identical pipeline — parse, font check, validate, render — with the identical isolation
// guarantees. The HTTP layer in app.js is a transport over this same surface, not a separate code path.
import fs from 'node:fs/promises';
import { loadConfig } from './config.js';
import { ServiceError } from './errors.js';
import { analyzeParsedRdl, analyzeRdl, parseRdl } from './rdl/parser.js';
import { analyzeWindowsWordCompatibility } from './render/windowsWordCompatibility.js';
import { RenderRunner } from './worker/runner.js';

export { buildApp } from './app.js';
export { loadConfig } from './config.js';
export { ServiceError, toServiceError } from './errors.js';
export { analyzeParsedRdl, analyzeRdl, parseRdl } from './rdl/parser.js';
export { OUTPUTS } from './render/index.js';
export { analyzeWindowsWordCompatibility } from './render/windowsWordCompatibility.js';
export { RenderRunner } from './worker/runner.js';
export { checkFonts } from './render/fonts.js';
export { readiness } from './readiness.js';
export { sanitizedFilename } from './request.js';

function toBuffer(rdl) {
  if (Buffer.isBuffer(rdl)) return rdl;
  if (rdl instanceof Uint8Array) return Buffer.from(rdl);
  if (typeof rdl === 'string') return Buffer.from(rdl, 'utf8');
  throw new ServiceError('RDL_INVALID', 'rdl must be a Buffer, Uint8Array, or string');
}

/**
 * Creates an in-process converter.
 *
 * Prefer this over constructing a RenderRunner directly: the runner assumes `config.tempRoot` already
 * exists with 0700 permissions (the HTTP app creates it at boot), so a bare `new RenderRunner(config)`
 * fails on first render. This performs that setup and returns a handle that owns it.
 *
 * Each render still runs in a forked worker with the configured timeout, heap cap, concurrency admission,
 * and guaranteed temp-directory cleanup — the isolation is a property of the pipeline, not of the server.
 *
 * @param {{ config?: object, env?: object }} [options] `config` overrides individual resolved fields;
 *   `env` supplies the environment to resolve defaults from (defaults to `process.env`).
 * @returns {Promise<{ config: object, analyze: Function, render: Function, close: Function }>}
 */
export async function createConverter(options = {}) {
  const config = Object.freeze({ ...loadConfig(options.env), ...options.config });
  await fs.mkdir(config.tempRoot, { recursive: true, mode: 0o700 });
  await fs.chmod(config.tempRoot, 0o700);
  const runner = new RenderRunner(config);

  return {
    config,

    /**
     * Reports what an RDL uses and whether this service can render it, without rendering.
     * Pure and synchronous: no worker, no temp files. Check `compatible` before calling `render`.
     */
    analyze(rdl) {
      const model = parseRdl(toBuffer(rdl), {
        maxRdlBytes: config.maxRdlBytes,
        maxXmlNodes: config.maxXmlNodes,
        maxXmlDepth: config.maxXmlDepth,
      });
      const result = analyzeParsedRdl(model);
      result.windowsWordEditable = analyzeWindowsWordCompatibility(model, config);
      return result;
    },

    /**
     * Renders one artifact. Rejects with a ServiceError carrying a stable `code` and `statusCode`.
     *
     * @param {{ rdl: Buffer|Uint8Array|string, output: 'PDF'|'DOCX_EDITABLE'|'DOCX_VISUAL'|'XLSX',
     *   parameters?: object, datasets?: object, subreports?: object, signal?: AbortSignal }} request
     *   `datasets` values are arrays of row objects keyed by exact RDL `DataField` names.
     *   `subreports` maps child ReportName to its base64 RDL and invocation-scoped rows.
     * @returns {Promise<{ buffer: Buffer, mimeType: string, extension: string,
     *   pageCount: number|null, size: number, totalRows: number, layoutMode?: string,
     *   editableTextRatio?: number }>}
     */
    async render({ rdl, signal, ...request }) {
      return runner.render({ rdlBuffer: toBuffer(rdl), request, signal });
    },

    /** Terminates any in-flight workers. Call on shutdown; the handle is unusable afterwards. */
    async close() {
      await runner.shutdown();
    },
  };
}
