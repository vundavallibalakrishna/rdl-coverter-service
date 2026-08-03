import fs from 'node:fs/promises';
import path from 'node:path';
import { loadConfig } from '../config.js';
import { ServiceError, toServiceError } from '../errors.js';
import { parseRdl } from '../rdl/parser.js';
import { resolveBundledSubreports } from '../rdl/subreports.js';
import { validateRenderInput } from '../rdl/validation.js';
import { renderDocument } from '../render/index.js';
import { checkFonts, takeFontSubstitutions } from '../render/fonts.js';
import { createTelemetryClock } from '../telemetry.js';

process.on('message', async (message) => {
  if (message?.type !== 'render') return;
  const config = Object.freeze({ ...loadConfig(), ...(message.config || {}) });
  const sendTelemetry = (event) => {
    if (process.connected) process.send?.({ type: 'telemetry', event });
  };
  const telemetry = createTelemetryClock('worker', sendTelemetry);
  telemetry.mark('started', 'completed', { workerMemoryLimitMb: Number(process.execArgv.join(' ').match(/max-old-space-size=(\d+)/)?.[1]) || undefined });
  try {
    const [rdl, requestText] = await Promise.all([
      fs.readFile(message.rdlPath),
      fs.readFile(message.requestPath, 'utf8'),
    ]);
    telemetry.mark('inputs-read', 'completed', { rdlBytes: rdl.length, requestBytes: Buffer.byteLength(requestText) });
    const request = JSON.parse(requestText);
    telemetry.mark('request-decoded', 'completed', {
      output: String(request.output || '').toUpperCase().slice(0, 32),
      datasetCount: Object.keys(request.datasets || {}).length,
      bundledSubreportCount: Object.keys(request.subreports || {}).length,
    });
    const model = parseRdl(rdl, { maxRdlBytes: config.maxRdlBytes, maxXmlNodes: config.maxXmlNodes, maxXmlDepth: config.maxXmlDepth });
    telemetry.mark('rdl-parsed', 'completed', {
      bodyItemCount: model.body?.items?.length || 0,
      declaredDatasetCount: Array.isArray(model.datasets)
        ? model.datasets.length
        : Object.keys(model.datasets || {}).length,
      declaredFontFamilyCount: model.fonts?.length || 0,
      pageWidthPt: model.page?.width,
      pageHeightPt: model.page?.height,
    });
    const subreports = resolveBundledSubreports(model, request, config);
    telemetry.mark('subreports-resolved', 'completed', {
      bundledRows: subreports.bundledRows,
      consumedFontFamilyCount: subreports.fonts?.length || 0,
    });
    const fontCheck = checkFonts(config, subreports.fonts.length ? subreports.fonts : ['Arial']);
    if (!fontCheck.ready) throw new ServiceError('FONT_MISSING', `Required fonts are unavailable: ${fontCheck.missing.join(', ')}`, 503);
    telemetry.mark('fonts-validated', 'completed', {
      strictFonts: config.strictFonts,
      checkedFontVariantCount: fontCheck.checked?.length,
      missingFontVariantCount: fontCheck.missing?.length || 0,
    });
    const validation = validateRenderInput(model, request, config);
    if (validation.totalRows + subreports.bundledRows > config.maxRows) {
      throw new ServiceError('RDL_INVALID', `Dataset rows exceed the ${config.maxRows} row limit`, 413);
    }
    request.parameters = validation.parameters;
    telemetry.mark('input-validated', 'completed', {
      totalRows: validation.totalRows + subreports.bundledRows,
      maxRows: config.maxRows,
    });
    telemetry.mark('renderer-started', 'started', { output: String(request.output || '').toUpperCase().slice(0, 32) });
    const rendered = await renderDocument(
      model,
      request,
      config,
      message.tempDir,
      (phase, metrics) => telemetry.mark(phase, 'completed', metrics),
    );
    telemetry.mark('renderer-completed', 'completed', {
      outputBytes: rendered.buffer.length,
      pageCount: rendered.pageCount,
      layoutMode: rendered.layoutMode,
      sheetCount: rendered.sheetCount,
      workbookRowCount: rendered.rowCount,
      editableTextRatio: rendered.editableTextRatio,
    });
    const outputPath = path.join(message.tempDir, `artifact.${rendered.extension}`);
    await fs.writeFile(outputPath, rendered.buffer, { mode: 0o600 });
    telemetry.mark('artifact-written', 'completed', { outputBytes: rendered.buffer.length });
    process.send?.({
      type: 'completed',
      outputPath,
      pageCount: rendered.pageCount,
      mimeType: rendered.mimeType,
      extension: rendered.extension,
      size: rendered.buffer.length,
      totalRows: validation.totalRows + subreports.bundledRows,
      layoutMode: rendered.layoutMode,
      editableTextRatio: rendered.editableTextRatio,
      sceneStats: rendered.sceneStats,
      sheetCount: rendered.sheetCount,
      rowCount: rendered.rowCount,
      // Which declared families could not draw some run and what drew it instead. A coverage substitution
      // is deliberate but still a deviation from the report's declared styling, so it is reported rather
      // than silent — the worker's stdio is discarded, making the completion message the only way out.
      fontSubstitutions: takeFontSubstitutions(),
    });
  } catch (error) {
    const safe = toServiceError(error);
    telemetry.mark('failed', 'failed', { errorCode: safe.code, statusCode: safe.statusCode });
    process.send?.({
      type: 'failed',
      error: { code: safe.code, message: safe.message, statusCode: safe.statusCode, details: safe.details },
      // RENDER_FAILED deliberately scrubs the message before it reaches the caller, and this worker's
      // stdio is discarded, so an unexpected exception would otherwise leave no trace anywhere at all.
      // Carry the real one back for the server-side log only; it never enters the HTTP response.
      diagnostic: safe === error ? undefined : {
        name: error?.name,
        message: error?.message,
        stack: typeof error?.stack === 'string' ? error.stack.split('\n').slice(0, 12).join('\n') : undefined,
      },
    });
  }
});
