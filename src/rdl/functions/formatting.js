// RDL/SSRS formatting functions (VB.NET semantics). Args are already-evaluated; each returns a STRING.
// Invalid/blank input returns '' (never throws). Uses Intl with locale en-US and timeZone UTC for dates.
import { toNumber, toDate, isBlank } from './shared.js';

function numberFormat(value, decimals, options = {}) {
  const number = toNumber(value);
  if (!Number.isFinite(number)) return '';
  const digits = isBlank(decimals) ? 2 : Math.trunc(toNumber(decimals));
  if (!Number.isFinite(digits) || digits < 0) return '';
  return new Intl.NumberFormat('en-US', {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
    ...options,
  }).format(number);
}

function dateTimeFormat(value, options) {
  return new Intl.DateTimeFormat('en-US', { timeZone: 'UTC', ...options }).format(value);
}

// VB DateFormat enum: accept the numeric code or the enum name (case-insensitive).
const DATE_FORMATS = { 0: 'generaldate', 1: 'longdate', 2: 'shortdate', 3: 'longtime', 4: 'shorttime' };

export const formattingFunctions = {
  FormatNumber(args) {
    return numberFormat(args[0], args[1]);
  },

  FormatCurrency(args) {
    return numberFormat(args[0], args[1], { style: 'currency', currency: 'USD' });
  },

  FormatPercent(args) {
    const number = toNumber(args[0]);
    if (!Number.isFinite(number)) return '';
    const formatted = numberFormat(number * 100, args[1]);
    return formatted === '' ? '' : `${formatted}%`;
  },

  FormatDateTime(args) {
    const date = toDate(args[0]);
    if (!date) return '';
    let named = 'generaldate';
    if (!isBlank(args[1])) {
      const raw = args[1];
      const numeric = Number(raw);
      named = Number.isInteger(numeric) && DATE_FORMATS[numeric]
        ? DATE_FORMATS[numeric]
        : String(raw).toLowerCase();
    }
    switch (named) {
      case 'longdate':
        return dateTimeFormat(date, { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
      case 'shortdate':
        return dateTimeFormat(date, { year: 'numeric', month: 'numeric', day: 'numeric' });
      case 'longtime':
        return dateTimeFormat(date, { hour: 'numeric', minute: '2-digit', second: '2-digit', hour12: true });
      case 'shorttime':
        return dateTimeFormat(date, { hour: 'numeric', minute: '2-digit', hour12: true });
      case 'generaldate':
      default: {
        const hasTime = date.getUTCHours() || date.getUTCMinutes() || date.getUTCSeconds() || date.getUTCMilliseconds();
        return hasTime
          ? dateTimeFormat(date, { year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: '2-digit', second: '2-digit', hour12: true })
          : dateTimeFormat(date, { year: 'numeric', month: 'numeric', day: 'numeric' });
      }
    }
  },
};
