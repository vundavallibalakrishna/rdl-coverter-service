import { ServiceError } from '../errors.js';

export const OUTPUTS = new Set(['PDF', 'DOCX_EDITABLE', 'DOCX_VISUAL', 'XLSX']);

export async function renderDocument(model, request, config, tempDir, telemetry) {
  const output = String(request.output || '').toUpperCase();
  if (!OUTPUTS.has(output)) throw new ServiceError('RDL_INVALID', `Unsupported output: ${output || 'missing'}`);
  const reportTelemetry = (phase, metrics) => {
    try { telemetry?.(phase, metrics); } catch { /* Observability cannot affect rendering. */ }
  };
  // Each request runs in a fresh isolated process. Eagerly importing every renderer made a PDF worker load
  // ExcelJS, docx, JSZip and the visual raster pipeline before it could emit its first telemetry event. On
  // Windows that cold module graph costs several seconds per request. Selective dynamic imports preserve
  // the one-request worker/security boundary while loading only the format that was actually requested.
  if (output === 'PDF') {
    const { renderPdf } = await import('./pdf.js');
    reportTelemetry('renderer-module-loaded', { output });
    return renderPdf(model, request, config, { telemetry });
  }
  if (output === 'DOCX_EDITABLE') {
    const { renderEditableDocx } = await import('./docx.js');
    reportTelemetry('renderer-module-loaded', { output });
    return renderEditableDocx(model, request, config, tempDir, telemetry);
  }
  if (output === 'XLSX') {
    const { renderExcel } = await import('./excel.js');
    reportTelemetry('renderer-module-loaded', { output });
    return renderExcel(model, request, config, tempDir);
  }
  const { renderVisualDocx } = await import('./visualDocx.js');
  reportTelemetry('renderer-module-loaded', { output });
  return renderVisualDocx(model, request, config, tempDir);
}
