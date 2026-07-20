import os from 'node:os';
import path from 'node:path';

function positiveInteger(value, fallback) {
  const parsed = Number.parseInt(value, 10);
  return Number.isInteger(parsed) && parsed > 0 ? parsed : fallback;
}

function nonNegativeNumber(value, fallback) {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) && parsed >= 0 ? parsed : fallback;
}

function booleanFlag(value, fallback = false) {
  if (value === undefined || value === null || value === '') return fallback;
  return /^(1|true|yes|on)$/i.test(String(value));
}

export function loadConfig(env = process.env) {
  return Object.freeze({
    port: positiveInteger(env.PORT, 7070),
    host: env.HOST || '0.0.0.0',
    tempRoot: path.resolve(env.RDL_TEMP_ROOT || path.join(os.tmpdir(), 'rdl-converter')),
    fontDir: path.resolve(env.RDL_FONT_DIR || path.join(process.cwd(), 'fonts')),
    strictFonts: env.RDL_STRICT_FONTS !== 'false',
    maxRdlBytes: positiveInteger(env.RDL_MAX_RDL_BYTES, 10 * 1024 * 1024),
    maxRequestBytes: positiveInteger(env.RDL_MAX_REQUEST_BYTES, 25 * 1024 * 1024),
    maxRows: positiveInteger(env.RDL_MAX_ROWS, 100_000),
    maxConcurrency: positiveInteger(env.RDL_MAX_CONCURRENCY, 2),
    renderTimeoutMs: positiveInteger(env.RDL_RENDER_TIMEOUT_MS, 120_000),
    workerMemoryMb: positiveInteger(env.RDL_WORKER_MEMORY_MB, 512),
    maxXmlNodes: positiveInteger(env.RDL_MAX_XML_NODES, 250_000),
    maxXmlDepth: positiveInteger(env.RDL_MAX_XML_DEPTH, 256),
    docxNativePageFragments: booleanFlag(env.RDL_DOCX_NATIVE_PAGE_FRAGMENTS, false),
    docxProfilePath: env.RDL_DOCX_PROFILE_PATH ? path.resolve(env.RDL_DOCX_PROFILE_PATH) : null,
    docxProfileAuto: booleanFlag(env.RDL_DOCX_PROFILE_AUTO, false),
    pdftoppmPath: env.RDL_PDFTOPPM_PATH || 'pdftoppm',
    // Minimum PDF border stroke width in points. 0 = honour the RDL width exactly (default). A floor only
    // makes borders heavier, not more uniform (the unevenness is a viewer sub-pixel artifact, not width),
    // so it is opt-in and off by default.
    borderWidthFloorPt: nonNegativeNumber(env.RDL_BORDER_WIDTH_FLOOR_PT, 0),
  });
}
