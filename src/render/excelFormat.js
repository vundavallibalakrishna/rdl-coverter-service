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

// A general-purpose Excel date format used when a real Date value carries an explicit format we don't
// translate. Explicit-format Excel date translation is a separate known deviation.
export const DEFAULT_EXCEL_DATE_FORMAT = 'yyyy-mm-dd';

// The Excel display for a DateTime with NO format: general date/time (date and time), matching .NET/SSRS
// default DateTime rendering so Excel stays consistent with the PDF/DOCX text renderers.
export const DEFAULT_EXCEL_DATETIME_FORMAT = 'yyyy-mm-dd hh:mm:ss';

// Standard single-letter .NET date specifiers mapped to an Excel number format that matches what the text
// renderer (formatValue) produces for the same specifier. Only the specifiers formatValue special-cases
// are mapped; anything else returns null and the caller writes the already-formatted string instead.
const STANDARD_DATE_NUMFMT = {
  y: 'mmmm yyyy', Y: 'mmmm yyyy', d: 'dd/mm/yyyy', D: 'd mmmm yyyy',
  g: 'dd/mm/yyyy hh:mm', G: 'dd/mm/yyyy hh:mm:ss',
};

// .NET custom date tokens → Excel number-format tokens. Month (`M`) and minute (`m`) both become `m`/`mm`
// in Excel, which disambiguates by position (an `mm` following `h`/`hh` is minutes). Longest-first.
const DATE_TOKEN_TRANSLATION = {
  yyyy: 'yyyy', yyy: 'yyyy', yy: 'yy', y: 'yy',
  MMMM: 'mmmm', MMM: 'mmm', MM: 'mm', M: 'm',
  dddd: 'dddd', ddd: 'ddd', dd: 'dd', d: 'd',
  HH: 'hh', H: 'h', hh: 'hh', h: 'h',
  mm: 'mm', m: 'm', ss: 'ss', s: 's', tt: 'AM/PM',
};
const DATE_TOKEN_RE = /(yyyy|yyy|yy|y|MMMM|MMM|MM|M|dddd|ddd|dd|d|HH|H|hh|h|mm|m|ss|s|tt)/g;

// Translate an RDL/.NET date `Format` into an Excel number format so an explicitly-formatted date stays a
// live typed Excel date that displays the same as the PDF/DOCX text. Returns null when the format is a
// specifier we do not map or carries quoted literals — the caller then writes the exact formatted string.
export function excelDateFormat(format) {
  const value = String(format || '').trim();
  if (!value) return null;
  if (value.length === 1) return STANDARD_DATE_NUMFMT[value] ?? null;
  if (/['"\\]/.test(value)) return null;
  if (!/[yMdHhst]/.test(value)) return null;
  let matched = false;
  const translated = value.replace(DATE_TOKEN_RE, (token) => {
    matched = true;
    return DATE_TOKEN_TRANSLATION[token];
  });
  return matched ? translated : null;
}
