import { ServiceError } from './errors.js';

/**
 * Resolves the public Excel layout contract without depending on the renderer.
 * Request validation and pre-render capability checks must use the same rule so
 * the default REPORT mode cannot be mistaken for the legacy DATA renderer.
 */
export function resolveExcelLayoutMode(request = {}) {
  const requested = request.excel?.layoutMode;
  if (requested !== undefined && requested !== null && String(requested).trim() !== '') {
    const mode = String(requested).trim().toUpperCase();
    if (!['REPORT', 'DATA'].includes(mode)) {
      throw new ServiceError('RDL_INVALID', `Unsupported Excel layoutMode: ${requested}`);
    }
    if (mode === 'REPORT' && request.excel?.sheetPerTablix === true) {
      throw new ServiceError('RDL_INVALID', 'excel.sheetPerTablix is only valid with excel.layoutMode DATA');
    }
    return mode;
  }
  // Backward compatibility for callers that already explicitly selected the old per-tablix workbook.
  if (request.excel?.sheetPerTablix === true) return 'DATA';
  return 'REPORT';
}
