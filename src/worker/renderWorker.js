import fs from 'node:fs/promises';
import path from 'node:path';
import { loadConfig } from '../config.js';
import { ServiceError, toServiceError } from '../errors.js';
import { parseRdl } from '../rdl/parser.js';
import { validateRenderInput } from '../rdl/validation.js';
import { renderDocument } from '../render/index.js';
import { checkFonts } from '../render/fonts.js';

process.on('message', async (message) => {
  if (message?.type !== 'render') return;
  const config = Object.freeze({ ...loadConfig(), ...(message.config || {}) });
  try {
    const [rdl, requestText] = await Promise.all([
      fs.readFile(message.rdlPath),
      fs.readFile(message.requestPath, 'utf8'),
    ]);
    const request = JSON.parse(requestText);
    const model = parseRdl(rdl, { maxRdlBytes: config.maxRdlBytes, maxXmlNodes: config.maxXmlNodes, maxXmlDepth: config.maxXmlDepth });
    const fontCheck = checkFonts(config, model.fonts.length ? model.fonts : ['Arial']);
    if (!fontCheck.ready) throw new ServiceError('FONT_MISSING', `Required fonts are unavailable: ${fontCheck.missing.join(', ')}`, 503);
    const validation = validateRenderInput(model, request, config);
    request.parameters = validation.parameters;
    const rendered = await renderDocument(model, request, config, message.tempDir);
    const outputPath = path.join(message.tempDir, `artifact.${rendered.extension}`);
    await fs.writeFile(outputPath, rendered.buffer, { mode: 0o600 });
    process.send?.({
      type: 'completed',
      outputPath,
      pageCount: rendered.pageCount,
      mimeType: rendered.mimeType,
      extension: rendered.extension,
      size: rendered.buffer.length,
      totalRows: validation.totalRows,
    });
  } catch (error) {
    const safe = toServiceError(error);
    process.send?.({ type: 'failed', error: { code: safe.code, message: safe.message, statusCode: safe.statusCode, details: safe.details } });
  }
});
