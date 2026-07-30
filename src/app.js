import fs from 'node:fs/promises';
import Fastify from 'fastify';
import multipart from '@fastify/multipart';
import { loadConfig } from './config.js';
import { ServiceError, toServiceError } from './errors.js';
import { analyzeParsedRdl, parseRdl } from './rdl/parser.js';
import { readInput, sanitizedFilename } from './request.js';
import { readiness } from './readiness.js';
import { RenderRunner } from './worker/runner.js';
import { analyzeWindowsWordCompatibility } from './render/windowsWordCompatibility.js';
import { fontAvailability } from './render/fonts.js';
import { testUiPage } from './testUi.js';

function safeHeaderValue(value) {
  return String(value ?? '').replace(/[\r\n\t]/g, ' ').replace(/[^\x20-\x7E]/g, '?').trim().slice(0, 200);
}

export async function buildApp(options = {}) {
  const config = options.config || loadConfig();
  await fs.mkdir(config.tempRoot, { recursive: true, mode: 0o700 });
  await fs.chmod(config.tempRoot, 0o700);
  const runner = options.runner || new RenderRunner(config);
  const app = Fastify({
    logger: options.logger ?? { level: process.env.LOG_LEVEL || 'info', redact: ['req.headers.authorization', 'req.body', 'res.body'] },
    bodyLimit: config.maxRequestBytes,
    requestIdHeader: 'x-request-id',
  });
  await app.register(multipart, {
    limits: { fileSize: config.maxRdlBytes, files: 1, fields: 8, fieldSize: Math.max(1, config.maxRequestBytes - config.maxRdlBytes), parts: 10 },
  });

  const sendTestUi = async (_request, reply) => {
    const page = testUiPage();
    return reply
      .header('Cache-Control', 'no-store')
      .header('Content-Security-Policy', page.contentSecurityPolicy)
      .header('Referrer-Policy', 'no-referrer')
      .header('X-Content-Type-Options', 'nosniff')
      .type('text/html; charset=utf-8')
      .send(page.html);
  };
  app.get('/', sendTestUi);
  app.get('/test-ui', sendTestUi);

  app.get('/healthz', async () => ({ status: 'ok', uptimeSeconds: Math.floor(process.uptime()) }));
  app.get('/readyz', async (_request, reply) => {
    const result = await readiness(config);
    return reply.code(result.ready ? 200 : 503).send(result);
  });

  app.post('/v1/analyze', async (request, reply) => {
    const startedAt = Date.now();
    const { rdlBuffer, options } = await readInput(request, config);
    const model = parseRdl(rdlBuffer, { maxRdlBytes: config.maxRdlBytes, maxXmlNodes: config.maxXmlNodes, maxXmlDepth: config.maxXmlDepth });
    const result = analyzeParsedRdl(model);
    result.windowsWordEditable = analyzeWindowsWordCompatibility(model, config, options);
    // Surface which of the report's declared consumed fonts are actually present on this render host, so an
    // absent licensed face (e.g. Segoe UI) is visible instead of silently substituted at render time.
    result.fontAvailability = fontAvailability(config, model.fonts);
    request.log.info({ requestId: request.id, rdlBytes: rdlBuffer.length, durationMs: Date.now() - startedAt, compatible: result.compatible }, 'RDL analysis completed');
    return reply.header('X-Request-Id', request.id).send(result);
  });

  app.post('/v1/render', async (request, reply) => {
    const startedAt = Date.now();
    const abortController = new AbortController();
    const abort = () => abortController.abort();
    request.raw.once('aborted', abort);
    reply.raw.once('close', abort);
    const { rdlBuffer, options: renderRequest } = await readInput(request, config);
    let rendered;
    try {
      rendered = await runner.render({ rdlBuffer, request: renderRequest, signal: abortController.signal });
    } finally {
      request.raw.removeListener('aborted', abort);
      reply.raw.removeListener('close', abort);
    }
    const durationMs = Date.now() - startedAt;
    const filename = sanitizedFilename(renderRequest.outputFileName || renderRequest.fileName, rendered.extension);
    request.log.info({
      requestId: request.id,
      rdlBytes: rdlBuffer.length,
      datasetCount: Object.keys(renderRequest.datasets || {}).length,
      totalRows: rendered.totalRows,
      output: renderRequest.output,
      outputBytes: rendered.buffer.length,
      pageCount: rendered.pageCount,
      layoutMode: rendered.layoutMode,
      editableTextRatio: rendered.editableTextRatio,
      durationMs,
    }, 'RDL rendering completed');
    reply
      .header('Content-Type', rendered.mimeType)
      .header('Content-Disposition', `attachment; filename="${filename}"`)
      .header('Content-Length', rendered.buffer.length)
      .header('X-Request-Id', request.id)
      .header('X-Page-Count', rendered.pageCount ?? 'unknown')
      .header('X-Render-Duration-Ms', durationMs);
    if (rendered.layoutMode && renderRequest.output === 'XLSX') reply.header('X-Xlsx-Layout-Mode', safeHeaderValue(rendered.layoutMode));
    else if (rendered.layoutMode && /^DOCX_/.test(renderRequest.output)) reply.header('X-Docx-Layout-Mode', safeHeaderValue(rendered.layoutMode));
    if (rendered.editableTextRatio !== undefined) reply.header('X-Docx-Editable-Text-Ratio', rendered.editableTextRatio);
    return reply.send(rendered.buffer);
  });

  app.setErrorHandler((error, request, reply) => {
    let safe;
    if (error instanceof ServiceError) safe = error;
    else if (error?.code === 'FST_ERR_CTP_BODY_TOO_LARGE' || error?.statusCode === 413) safe = new ServiceError('RDL_INVALID', 'Request exceeds the configured size limit', 413);
    else safe = toServiceError(error);
    request.log.error({ requestId: request.id, code: safe.code, statusCode: safe.statusCode }, 'RDL request failed');
    if (safe.code === 'BUSY') reply.header('Retry-After', safe.details?.retryAfterSeconds || 5);
    return reply.code(safe.statusCode).header('X-Request-Id', request.id).send({ error: { code: safe.code, message: safe.message, details: safe.details } });
  });

  app.addHook('onClose', async () => runner.shutdown());
  app.decorate('converterConfig', config);
  app.decorate('renderRunner', runner);
  return app;
}
