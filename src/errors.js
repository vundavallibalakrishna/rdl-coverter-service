export class ServiceError extends Error {
  constructor(code, message, statusCode = 400, details = undefined) {
    super(message);
    this.name = 'ServiceError';
    this.code = code;
    this.statusCode = statusCode;
    this.details = details;
  }
}

export function toServiceError(error) {
  if (error instanceof ServiceError) return error;
  const wrapped = new ServiceError('RENDER_FAILED', 'Document rendering failed', 500);
  wrapped.cause = error;
  return wrapped;
}
