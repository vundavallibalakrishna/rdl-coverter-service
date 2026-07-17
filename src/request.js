import { ServiceError } from './errors.js';

function decodeBase64(value) {
  if (typeof value !== 'string' || value.length === 0) throw new ServiceError('RDL_INVALID', 'rdlBase64 is required');
  const normalized = value.replace(/^data:[^;]+;base64,/, '').replace(/\s+/g, '');
  if (!/^[A-Za-z0-9+/]*={0,2}$/.test(normalized)) throw new ServiceError('RDL_INVALID', 'rdlBase64 is invalid');
  const buffer = Buffer.from(normalized, 'base64');
  if (buffer.length === 0) throw new ServiceError('RDL_INVALID', 'Uploaded RDL is empty');
  return buffer;
}

function ensureRdlLimit(buffer, config) {
  if (buffer.length > config.maxRdlBytes) throw new ServiceError('RDL_INVALID', 'RDL file exceeds the configured size limit', 413);
  return buffer;
}

export async function readInput(request, config) {
  if (request.isMultipart()) {
    let rdlBuffer = null;
    let options = null;
    for await (const part of request.parts()) {
      if (part.type === 'file') {
        if (part.fieldname !== 'rdl') {
          part.file.resume();
          continue;
        }
        try {
          rdlBuffer = await part.toBuffer();
        } catch {
          throw new ServiceError('RDL_INVALID', 'RDL file exceeds the configured size limit', 413);
        }
      } else if (part.fieldname === 'request') {
        try {
          options = typeof part.value === 'string' ? JSON.parse(part.value) : part.value;
          if (!options || typeof options !== 'object' || Array.isArray(options)) throw new Error('Invalid request object');
        } catch {
          throw new ServiceError('RDL_INVALID', 'Multipart request field must contain valid JSON');
        }
      }
    }
    if (!rdlBuffer) throw new ServiceError('RDL_INVALID', 'Multipart field rdl is required');
    return { rdlBuffer: ensureRdlLimit(rdlBuffer, config), options: options || {} };
  }

  if (!request.body || typeof request.body !== 'object' || Array.isArray(request.body)) throw new ServiceError('RDL_INVALID', 'JSON request body is required');
  const { rdlBase64, ...options } = request.body;
  return { rdlBuffer: ensureRdlLimit(decodeBase64(rdlBase64), config), options };
}

export function sanitizedFilename(value, extension) {
  const base = String(value || 'report').replace(/\.[A-Za-z0-9]+$/, '').replace(/[^A-Za-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 120) || 'report';
  return `${base}.${extension}`;
}
