// An RDL DateTime is a wall-clock value. SSRS renders exactly the date and time it was given and never
// converts time zones, so `2026-05-15T00:00:00` must render as 15 May everywhere.
//
// Every formatter in this service reads a Date through its UTC accessors (`getUTCDate`, `timeZone: 'UTC'`).
// JavaScript, however, parses a date-time string that carries no time-zone designator as LOCAL time, so on
// any host east of UTC that midnight becomes the previous day in UTC and renders as 14 May. That also makes
// the output depend on the server's time zone, breaking the determinism invariant.
//
// Parsing an unzoned value as UTC keeps the wall clock intact end to end. A value that DOES carry `Z` or a
// `±HH:MM` offset names a real instant, so it keeps the standard parse and is rendered in UTC.
const UNZONED = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2})(?:[.,](\d{1,7}))?)?)?$/;

export function parseDateValue(value) {
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
  if (typeof value !== 'string') return null;
  const trimmed = value.trim();
  if (trimmed === '') return null;
  const match = UNZONED.exec(trimmed);
  if (!match) {
    const parsed = new Date(trimmed);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }
  const [, year, month, day, hours = '0', minutes = '0', seconds = '0', fraction = ''] = match;
  const milliseconds = fraction ? Math.round(Number(`0.${fraction}`) * 1000) : 0;
  const parsed = new Date(Date.UTC(
    Number(year), Number(month) - 1, Number(day),
    Number(hours), Number(minutes), Number(seconds), milliseconds,
  ));
  // Date.UTC maps a two-digit year into the 1900s; a four-digit year below 100 must stay literal.
  if (Number(year) < 100) parsed.setUTCFullYear(Number(year));
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}
