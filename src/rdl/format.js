// Self-contained .NET / SSRS style format-string engine.
//
// The RDL `Format` property carries .NET format strings. This module renders a
// JS value (number, string, Date, boolean, null) into the string a .NET
// `String.Format`/`ToString(format)` call would produce, without executing any
// RDL expressions or external code. Everything is pure string/number work.
//
// Design invariants:
// - Never throw. On an unknown/blank format, fall back to `String(value ?? '')`.
// - A null/undefined value always renders as ''.
// - Dates are formatted in UTC so output is deterministic regardless of host
//   timezone.

const DATE_TOKEN_RE = /(yyyy|yyy|yy|y|MMMM|MMM|MM|M|dddd|ddd|dd|d|HH|H|hh|h|mm|m|ss|s|tt)/;
const NUMERIC_HINT_RE = /[#0]|^[CcNnPpFfDdEeGgRrXx]\d*$/;

/**
 * Format a value using a .NET / SSRS format string.
 *
 * @param {number|string|Date|boolean|null|undefined} value
 * @param {string} formatString
 * @param {string} [culture='en-US']
 * @returns {string}
 */
export function formatNet(value, formatString, culture = 'en-US') {
  if (value === null || value === undefined) return '';

  const format = typeof formatString === 'string' ? formatString : '';
  if (format.trim() === '') return String(value);

  try {
    const looksDate = formatLooksLikeDate(format);
    const looksNumeric = formatLooksLikeNumeric(format);

    const asDate = coerceDate(value);
    // Single-letter specifiers such as `D`, `M`, `g`, `F` are ambiguous between
    // numeric and date meanings. When the value is an actual Date, the date
    // interpretation wins; otherwise a numeric format string wins.
    if (asDate && looksDate && (value instanceof Date || !looksNumeric)) {
      return formatDate(asDate, format, culture);
    }

    const asNumber = coerceNumber(value);
    if (asNumber !== null && (looksNumeric || !looksDate)) {
      return formatNumber(asNumber, format, culture);
    }

    // Date-ish format but the value isn't a date, or numeric format but value
    // isn't numeric: fall through to the raw string.
    return String(value);
  } catch {
    return String(value ?? '');
  }
}

// ---------------------------------------------------------------------------
// Format classification
// ---------------------------------------------------------------------------

function formatLooksLikeDate(format) {
  // Standard single-letter date specifiers.
  if (/^[dDtTgGfFmMoOrRsuUyY]$/.test(format)) return true;
  // Custom formats containing date/time tokens.
  return DATE_TOKEN_RE.test(format);
}

function formatLooksLikeNumeric(format) {
  if (/^[CcNnPpFfDdEeGgRrXx]\d*$/.test(format)) return true;
  return /[#0]/.test(format);
}

// ---------------------------------------------------------------------------
// Coercion
// ---------------------------------------------------------------------------

function coerceNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  if (typeof value === 'boolean') return value ? 1 : 0;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (trimmed === '') return null;
    const n = Number(trimmed);
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

function coerceDate(value) {
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (trimmed === '') return null;
    // Avoid treating bare numeric strings as dates.
    if (/^-?\d+(\.\d+)?$/.test(trimmed)) return null;
    const d = new Date(trimmed);
    return Number.isNaN(d.getTime()) ? null : d;
  }
  return null;
}

// ---------------------------------------------------------------------------
// Numeric formatting
// ---------------------------------------------------------------------------

function formatNumber(value, format, culture) {
  const standard = tryStandardNumeric(value, format, culture);
  if (standard !== null) return standard;
  return formatCustomNumeric(value, format, culture);
}

function tryStandardNumeric(value, format, culture) {
  const match = /^([CcNnPpFfDdEeGgRrXx])(\d*)$/.exec(format);
  if (!match) return null;

  const specifier = match[1];
  const hasPrecision = match[2] !== '';
  const precision = hasPrecision ? parseInt(match[2], 10) : undefined;

  switch (specifier) {
    case 'C':
    case 'c': {
      const digits = precision ?? 2;
      return new Intl.NumberFormat(culture, {
        style: 'currency',
        currency: 'USD',
        minimumFractionDigits: digits,
        maximumFractionDigits: digits,
      }).format(value);
    }
    case 'N':
    case 'n': {
      const digits = precision ?? 2;
      return new Intl.NumberFormat(culture, {
        style: 'decimal',
        useGrouping: true,
        minimumFractionDigits: digits,
        maximumFractionDigits: digits,
      }).format(value);
    }
    case 'P':
    case 'p': {
      const digits = precision ?? 2;
      // .NET multiplies by 100 and appends the percent sign.
      const scaled = value * 100;
      const body = new Intl.NumberFormat(culture, {
        style: 'decimal',
        useGrouping: true,
        minimumFractionDigits: digits,
        maximumFractionDigits: digits,
      }).format(scaled);
      return `${body}%`;
    }
    case 'F':
    case 'f': {
      const digits = precision ?? 2;
      return new Intl.NumberFormat(culture, {
        style: 'decimal',
        useGrouping: false,
        minimumFractionDigits: digits,
        maximumFractionDigits: digits,
      }).format(value);
    }
    case 'D':
    case 'd': {
      // Integer with zero padding to `precision` digits.
      const rounded = Math.round(value);
      const negative = rounded < 0;
      let digits = Math.abs(rounded).toString();
      if (precision !== undefined && digits.length < precision) {
        digits = digits.padStart(precision, '0');
      }
      return negative ? `-${digits}` : digits;
    }
    case 'E':
    case 'e': {
      const digits = precision ?? 6;
      let out = value.toExponential(digits);
      // .NET uses a signed 3-digit exponent and the case of the specifier.
      out = out.replace(/e([+-])(\d+)/i, (_, sign, exp) => {
        const padded = exp.padStart(3, '0');
        return `${specifier}${sign}${padded}`;
      });
      return out;
    }
    case 'G':
    case 'g': {
      if (precision !== undefined && precision > 0) {
        return trimNumberString(value.toPrecision(precision));
      }
      return value.toString();
    }
    case 'R':
    case 'r':
      return value.toString();
    case 'X':
    case 'x': {
      const rounded = Math.round(value);
      let hex = (rounded >>> 0).toString(16);
      if (specifier === 'X') hex = hex.toUpperCase();
      if (precision !== undefined && hex.length < precision) {
        hex = hex.padStart(precision, '0');
      }
      return hex;
    }
    default:
      return null;
  }
}

function trimNumberString(str) {
  if (str.indexOf('.') === -1) return str;
  return str.replace(/\.?0+$/, '');
}

/**
 * Handle custom numeric format strings, including `;`-separated sections for
 * positive / negative / zero (/ null) values.
 */
function formatCustomNumeric(value, format, culture) {
  const sections = splitSections(format);

  let pattern;
  let forceNegativeSign = false;

  if (sections.length === 1) {
    pattern = sections[0];
    forceNegativeSign = value < 0;
  } else if (sections.length === 2) {
    if (value < 0) {
      pattern = sections[1];
    } else {
      pattern = sections[0];
    }
  } else {
    // 3+ sections: positive ; negative ; zero
    if (value > 0) {
      pattern = sections[0];
    } else if (value < 0) {
      pattern = sections[1];
    } else {
      pattern = sections[2];
    }
  }

  // For a multi-section format the negative section already contains the desired
  // negative representation, so operate on the magnitude. For a single section
  // .NET prefixes a '-' automatically.
  const magnitude = sections.length >= 2 ? Math.abs(value) : Math.abs(value);
  const rendered = applyCustomPattern(magnitude, pattern, culture);
  return forceNegativeSign ? `-${rendered}` : rendered;
}

/**
 * Split on literal `;` while respecting quoted literals and escapes.
 */
function splitSections(format) {
  const sections = [];
  let current = '';
  for (let i = 0; i < format.length; i += 1) {
    const ch = format[i];
    if (ch === '\\') {
      current += ch + (format[i + 1] ?? '');
      i += 1;
      continue;
    }
    if (ch === '"' || ch === "'") {
      const quote = ch;
      current += ch;
      i += 1;
      while (i < format.length && format[i] !== quote) {
        current += format[i];
        i += 1;
      }
      if (i < format.length) current += format[i];
      continue;
    }
    if (ch === ';') {
      sections.push(current);
      current = '';
      continue;
    }
    current += ch;
  }
  sections.push(current);
  return sections;
}

/**
 * Render a single custom numeric pattern (no section separators) against a
 * non-negative value.
 */
function applyCustomPattern(value, pattern, culture) {
  // A section with no digit placeholders is a pure literal (e.g. the zero
  // section `"-"`); render it verbatim without injecting the number.
  if (!/[#0]/.test(pattern)) {
    return renderTemplate(pattern, '');
  }

  // Detect percent scaling: each literal '%' multiplies the value by 100.
  const percentCount = countUnescaped(pattern, '%');
  let scaled = value;
  for (let i = 0; i < percentCount; i += 1) scaled *= 100;

  // Detect grouping (a ',' that sits between digit placeholders).
  const useGrouping = hasGroupingComma(pattern);

  // Determine decimal placeholder counts.
  const { intZeros, fracZeros, fracHashes } = analyzeDigits(pattern);
  const minFrac = fracZeros;
  const maxFrac = fracZeros + fracHashes;

  const rounded = roundTo(scaled, maxFrac);
  const negative = rounded < 0;
  const absVal = Math.abs(rounded);

  // Build the numeric portion.
  const fixed = absVal.toFixed(maxFrac);
  let [intPart, fracPart = ''] = fixed.split('.');

  // Trim optional trailing fraction digits (the `#` placeholders) but keep the
  // required `0` placeholders.
  if (fracPart.length > 0) {
    let keep = fracPart.length;
    while (keep > minFrac && fracPart[keep - 1] === '0') keep -= 1;
    fracPart = fracPart.slice(0, keep);
  }

  // Pad integer part up to the required leading zeros.
  if (intPart.length < intZeros) {
    intPart = intPart.padStart(intZeros, '0');
  }

  if (useGrouping) {
    intPart = groupThousands(intPart, culture);
  }

  let numberText = intPart;
  if (fracPart.length > 0 || minFrac > 0) {
    const sep = decimalSeparator(culture);
    numberText += sep + fracPart;
  }
  if (negative) numberText = `-${numberText}`;

  // Substitute the numeric text back into the literal template.
  return renderTemplate(pattern, numberText);
}

function roundTo(value, digits) {
  if (!Number.isFinite(value)) return value;
  const factor = 10 ** digits;
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

function countUnescaped(pattern, target) {
  let count = 0;
  for (let i = 0; i < pattern.length; i += 1) {
    const ch = pattern[i];
    if (ch === '\\') {
      i += 1;
      continue;
    }
    if (ch === '"' || ch === "'") {
      const quote = ch;
      i += 1;
      while (i < pattern.length && pattern[i] !== quote) i += 1;
      continue;
    }
    if (ch === target) count += 1;
  }
  return count;
}

function hasGroupingComma(pattern) {
  // A comma is a grouping indicator when it appears between digit placeholders
  // in the integer part of the (unquoted) pattern.
  let inFraction = false;
  for (let i = 0; i < pattern.length; i += 1) {
    const ch = pattern[i];
    if (ch === '\\') {
      i += 1;
      continue;
    }
    if (ch === '"' || ch === "'") {
      const quote = ch;
      i += 1;
      while (i < pattern.length && pattern[i] !== quote) i += 1;
      continue;
    }
    if (ch === '.') {
      inFraction = true;
      continue;
    }
    if (ch === ',' && !inFraction) {
      // Must have a digit placeholder before and after somewhere.
      const before = /[#0]/.test(pattern.slice(0, i));
      const after = /[#0]/.test(pattern.slice(i + 1));
      if (before && after) return true;
    }
  }
  return false;
}

function analyzeDigits(pattern) {
  let intZeros = 0;
  let fracZeros = 0;
  let fracHashes = 0;
  let inFraction = false;
  for (let i = 0; i < pattern.length; i += 1) {
    const ch = pattern[i];
    if (ch === '\\') {
      i += 1;
      continue;
    }
    if (ch === '"' || ch === "'") {
      const quote = ch;
      i += 1;
      while (i < pattern.length && pattern[i] !== quote) i += 1;
      continue;
    }
    if (ch === '.') {
      inFraction = true;
      continue;
    }
    if (ch === '0') {
      if (inFraction) fracZeros += 1;
      else intZeros += 1;
    } else if (ch === '#') {
      if (inFraction) fracHashes += 1;
    }
  }
  return { intZeros, fracZeros, fracHashes };
}

function groupThousands(intDigits, culture) {
  const sep = groupSeparator(culture);
  return intDigits.replace(/\B(?=(\d{3})+(?!\d))/g, sep);
}

/**
 * Walk the pattern, emitting literal characters as-is and replacing the digit
 * placeholder region with the already-formatted number text.
 */
function renderTemplate(pattern, numberText) {
  let out = '';
  let injected = false;
  for (let i = 0; i < pattern.length; i += 1) {
    const ch = pattern[i];
    if (ch === '\\') {
      out += pattern[i + 1] ?? '';
      i += 1;
      continue;
    }
    if (ch === '"' || ch === "'") {
      const quote = ch;
      i += 1;
      while (i < pattern.length && pattern[i] !== quote) {
        out += pattern[i];
        i += 1;
      }
      continue;
    }
    if (ch === '#' || ch === '0' || ch === '.' || ch === ',') {
      if (!injected) {
        out += numberText;
        injected = true;
      }
      continue;
    }
    if (ch === '%') {
      out += '%';
      continue;
    }
    out += ch;
  }
  if (!injected) out += numberText;
  return out;
}

// ---------------------------------------------------------------------------
// Culture-derived separators (via Intl)
// ---------------------------------------------------------------------------

function separatorsFor(culture) {
  const parts = new Intl.NumberFormat(culture, {
    minimumFractionDigits: 1,
    useGrouping: true,
  }).formatToParts(1234.5);
  let group = ',';
  let decimal = '.';
  for (const part of parts) {
    if (part.type === 'group') group = part.value;
    if (part.type === 'decimal') decimal = part.value;
  }
  return { group, decimal };
}

function groupSeparator(culture) {
  return separatorsFor(culture).group;
}

function decimalSeparator(culture) {
  return separatorsFor(culture).decimal;
}

// ---------------------------------------------------------------------------
// Date formatting
// ---------------------------------------------------------------------------

function formatDate(date, format, culture) {
  const standard = tryStandardDate(date, format, culture);
  if (standard !== null) return standard;
  return formatCustomDate(date, format, culture);
}

function tryStandardDate(date, format, culture) {
  if (format.length !== 1) return null;
  switch (format) {
    case 'd':
      return intlDate(date, culture, { dateStyle: 'short' });
    case 'D':
      return intlDate(date, culture, { dateStyle: 'long' });
    case 't':
      return intlDate(date, culture, { timeStyle: 'short' });
    case 'T':
      return intlDate(date, culture, { timeStyle: 'medium' });
    case 'g':
      return (
        intlDate(date, culture, { dateStyle: 'short' }) +
        ' ' +
        intlDate(date, culture, { timeStyle: 'short' })
      );
    case 'G':
      return (
        intlDate(date, culture, { dateStyle: 'short' }) +
        ' ' +
        intlDate(date, culture, { timeStyle: 'medium' })
      );
    case 'f':
      return (
        intlDate(date, culture, { dateStyle: 'long' }) +
        ' ' +
        intlDate(date, culture, { timeStyle: 'short' })
      );
    case 'F':
      return (
        intlDate(date, culture, { dateStyle: 'long' }) +
        ' ' +
        intlDate(date, culture, { timeStyle: 'medium' })
      );
    default:
      // M/m (month day), y/Y (year month), etc. handled via custom tokens.
      return null;
  }
}

function intlDate(date, culture, options) {
  return new Intl.DateTimeFormat(culture, { timeZone: 'UTC', ...options }).format(date);
}

function formatCustomDate(date, format, culture) {
  // Extract UTC components.
  const year = date.getUTCFullYear();
  const month = date.getUTCMonth(); // 0-based
  const day = date.getUTCDate();
  const hours24 = date.getUTCHours();
  const minutes = date.getUTCMinutes();
  const seconds = date.getUTCSeconds();
  const hours12 = hours24 % 12 === 0 ? 12 : hours24 % 12;
  const isPM = hours24 >= 12;

  const monthLong = new Intl.DateTimeFormat(culture, {
    timeZone: 'UTC',
    month: 'long',
  }).format(date);
  const monthShort = new Intl.DateTimeFormat(culture, {
    timeZone: 'UTC',
    month: 'short',
  }).format(date);
  const weekdayLong = new Intl.DateTimeFormat(culture, {
    timeZone: 'UTC',
    weekday: 'long',
  }).format(date);
  const weekdayShort = new Intl.DateTimeFormat(culture, {
    timeZone: 'UTC',
    weekday: 'short',
  }).format(date);
  const ampm = new Intl.DateTimeFormat(culture, {
    timeZone: 'UTC',
    hour: 'numeric',
    hour12: true,
  })
    .formatToParts(date)
    .find((p) => p.type === 'dayPeriod')?.value ?? (isPM ? 'PM' : 'AM');

  const tokenValue = (token) => {
    switch (token) {
      case 'yyyy':
        return String(year).padStart(4, '0');
      case 'yyy':
        return String(year).padStart(3, '0');
      case 'yy':
        return String(year % 100).padStart(2, '0');
      case 'y':
        return String(year % 100);
      case 'MMMM':
        return monthLong;
      case 'MMM':
        return monthShort;
      case 'MM':
        return String(month + 1).padStart(2, '0');
      case 'M':
        return String(month + 1);
      case 'dddd':
        return weekdayLong;
      case 'ddd':
        return weekdayShort;
      case 'dd':
        return String(day).padStart(2, '0');
      case 'd':
        return String(day);
      case 'HH':
        return String(hours24).padStart(2, '0');
      case 'H':
        return String(hours24);
      case 'hh':
        return String(hours12).padStart(2, '0');
      case 'h':
        return String(hours12);
      case 'mm':
        return String(minutes).padStart(2, '0');
      case 'm':
        return String(minutes);
      case 'ss':
        return String(seconds).padStart(2, '0');
      case 's':
        return String(seconds);
      case 'tt':
        return ampm;
      default:
        return token;
    }
  };

  let out = '';
  let i = 0;
  const tokenGlobal = new RegExp(DATE_TOKEN_RE.source, 'y');
  while (i < format.length) {
    const ch = format[i];
    if (ch === '\\') {
      out += format[i + 1] ?? '';
      i += 2;
      continue;
    }
    if (ch === '"' || ch === "'") {
      const quote = ch;
      i += 1;
      while (i < format.length && format[i] !== quote) {
        out += format[i];
        i += 1;
      }
      i += 1;
      continue;
    }
    tokenGlobal.lastIndex = i;
    const m = tokenGlobal.exec(format);
    if (m && m.index === i) {
      out += tokenValue(m[0]);
      i += m[0].length;
      continue;
    }
    out += ch;
    i += 1;
  }
  return out;
}
