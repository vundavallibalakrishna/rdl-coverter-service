import { ServiceError } from '../../errors.js';
import { parseDateValue } from '../dateValue.js';

// Shared, dependency-light coercion helpers for the expression-function registry. These must NOT import
// from expression.js (that module imports the registry, so importing back would create a cycle). Each
// registry function receives ALREADY-EVALUATED argument values (numbers/strings/Dates/null) plus the
// evaluation context, and returns a plain value.

export { ServiceError };

export function isBlank(value) {
  return value === null || value === undefined || value === '';
}

// Number coercion matching VB semantics closely enough for reports: blanks/invalid -> NaN so callers can
// decide a fallback. Booleans map to 0/1; strings are parsed leniently.
export function toNumber(value) {
  if (typeof value === 'number') return value;
  if (typeof value === 'boolean') return value ? 1 : 0;
  if (isBlank(value)) return NaN;
  const n = Number(value);
  return Number.isNaN(n) ? NaN : n;
}

// String coercion: null/undefined -> '' (VB Nothing renders blank), Dates -> ISO-ish, everything else String().
export function toText(value) {
  if (isBlank(value)) return '';
  if (value instanceof Date) return value.toISOString();
  return String(value);
}

// Parse a value to a Date (UTC-safe) or null. Mirrors expression.js toDate so date functions agree.
export function toDate(value) {
  if (isBlank(value)) return null;
  return parseDateValue(value);
}
