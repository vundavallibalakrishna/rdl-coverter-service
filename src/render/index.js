import { ServiceError } from '../errors.js';
import { renderPdf } from './pdf.js';
import { renderEditableDocx } from './docx.js';
import { renderFixedEditableDocx } from './fixedDocx.js';
import { renderVisualDocx } from './visualDocx.js';
import { renderExcel } from './excel.js';

export const OUTPUTS = new Set(['PDF', 'DOCX_EDITABLE', 'DOCX_FIXED_EDITABLE', 'DOCX_VISUAL', 'XLSX']);

export async function renderDocument(model, request, config, tempDir) {
  const output = String(request.output || '').toUpperCase();
  if (!OUTPUTS.has(output)) throw new ServiceError('RDL_INVALID', `Unsupported output: ${output || 'missing'}`);
  if (output === 'PDF') return renderPdf(model, request, config);
  if (output === 'DOCX_EDITABLE') return renderEditableDocx(model, request, config, tempDir);
  if (output === 'DOCX_FIXED_EDITABLE') return renderFixedEditableDocx(model, request, config, tempDir);
  if (output === 'XLSX') return renderExcel(model, request, config, tempDir);
  return renderVisualDocx(model, request, config, tempDir);
}
