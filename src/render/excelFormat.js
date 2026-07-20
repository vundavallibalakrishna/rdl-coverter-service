// Translates an RDL/.NET format string into an Excel number-format code, and coerces cell values into the
// right Excel type. The point of an Excel export (over PDF/DOCX) is that a number stays a real number the
// caller can sum, filter, and pivot — so a numeric value is written as a number with a format code, not as
// a pre-formatted string. Anything we cannot translate confidently is written as its display TEXT instead,
// which is always safe and never wrong, just not live.

// Coerces a cell value to a plain string. Formula injection is NOT a risk here the way it is for CSV:
// exceljs writes a string value as a typed string cell (t="s"), and Excel never evaluates a string-typed
// cell as a formula regardless of a leading =, +, -, or @. Prefixing with an apostrophe (the CSV-era
// mitigation) would instead corrupt legitimate values like "-N/A" or "@owner" with a visible leading quote.
// The invariant that keeps this safe is simply that we never assign a { formula } object to a cell.
export function cellString(value) {
  return String(value ?? '');
}

function digits(specifier, fallback) {
  const match = /\d+/.exec(specifier);
  return match ? Number(match[0]) : fallback;
}

// RDL/.NET format string -> Excel number-format code, or null for "no numeric format" (write as General or
// as text). Handles the standard single-letter specifiers and passes through Excel-compatible custom
// numeric patterns (#,0,.,%). Date/time and anything with letters we don't recognize returns null so the
// caller falls back to the display string rather than risk a wrong format.
export function excelNumberFormat(format) {
  if (format === null || format === undefined) return null;
  const raw = String(format).trim();
  if (!raw || raw.startsWith('=')) return null;

  const zeros = (n) => (n > 0 ? `.${'0'.repeat(n)}` : '');
  if (/^C\d*$/i.test(raw)) return `$#,##0${zeros(digits(raw, 2))}`;
  if (/^N\d*$/i.test(raw)) return `#,##0${zeros(digits(raw, 2))}`;
  if (/^F\d*$/i.test(raw)) return `0${zeros(digits(raw, 2))}`;
  if (/^P\d*$/i.test(raw)) return `0${zeros(digits(raw, 2))}%`;
  if (/^D\d*$/i.test(raw)) return '0';
  if (/^[EG]\d*$/i.test(raw)) return null; // scientific/general: let Excel default

  // Custom numeric pattern already in Excel's vocabulary (#, 0, comma, dot, %, currency, spacing).
  if (/^[#0.,%$€£\s()-]+$/.test(raw)) return raw;

  // Dates and other letter-bearing patterns: not safely translatable here.
  return null;
}

// A general-purpose Excel date format used when a real Date value carries a format we don't translate.
export const DEFAULT_EXCEL_DATE_FORMAT = 'yyyy-mm-dd';
