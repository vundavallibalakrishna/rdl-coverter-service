// RDL/SSRS date & time expression functions (VB.NET semantics, UTC). Registry entries receive
// already-evaluated argument values as an array plus the evaluation context; blank/invalid input
// returns null (or '' for name lookups) rather than throwing.
import { toDate, toNumber } from './shared.js';

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

// Resolve "now" deterministically from the report's ExecutionTime global when present.
function now(context) {
  const execution = context?.globals?.ExecutionTime;
  if (execution instanceof Date && !Number.isNaN(execution.getTime())) return new Date(execution.getTime());
  return new Date();
}

function midnightUtc(date) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
}

// UTC datePart helper: null if the value is not a valid date.
function part(value, getter) {
  const date = toDate(value);
  return date ? getter(date) : null;
}

export const datetimeFunctions = {
  Now: (_args, context) => now(context),
  Today: (_args, context) => midnightUtc(now(context)),
  Year: ([value]) => part(value, (d) => d.getUTCFullYear()),
  Month: ([value]) => part(value, (d) => d.getUTCMonth() + 1),
  Day: ([value]) => part(value, (d) => d.getUTCDate()),
  Hour: ([value]) => part(value, (d) => d.getUTCHours()),
  Minute: ([value]) => part(value, (d) => d.getUTCMinutes()),
  Second: ([value]) => part(value, (d) => d.getUTCSeconds()),

  // VB Weekday: 1-7. firstDayOfWeek (1=Sunday..7=Saturday) designates which weekday counts as 1.
  Weekday: ([value, firstDayOfWeek]) => {
    const date = toDate(value);
    if (!date) return null;
    const first = Number.isFinite(toNumber(firstDayOfWeek)) && toNumber(firstDayOfWeek) >= 1 ? toNumber(firstDayOfWeek) : 1;
    return ((date.getUTCDay() - (first - 1) + 7) % 7) + 1;
  },

  MonthName: ([n, abbreviate]) => {
    const index = toNumber(n) - 1;
    if (!Number.isInteger(index) || index < 0 || index > 11) return '';
    const name = MONTHS[index];
    return abbreviate ? name.slice(0, 3) : name;
  },

  // Weekday number -> English day name. n=1 means Sunday under the default firstDayOfWeek (1=Sunday).
  WeekdayName: ([n, abbreviate, firstDayOfWeek]) => {
    const weekday = toNumber(n);
    if (!Number.isInteger(weekday) || weekday < 1 || weekday > 7) return '';
    const first = Number.isFinite(toNumber(firstDayOfWeek)) && toNumber(firstDayOfWeek) >= 1 ? toNumber(firstDayOfWeek) : 1;
    const name = DAYS[((first - 1) + (weekday - 1)) % 7];
    return abbreviate ? name.slice(0, 3) : name;
  },

  // Date.UTC gives VB-style rollover for out-of-range month/day (month is 1-12 here).
  DateSerial: ([year, month, day]) => {
    const y = toNumber(year);
    const m = toNumber(month);
    const d = toNumber(day);
    if (!Number.isFinite(y) || !Number.isFinite(m) || !Number.isFinite(d)) return null;
    return new Date(Date.UTC(y, m - 1, d));
  },

  DateValue: ([value]) => {
    const date = toDate(value);
    return date ? midnightUtc(date) : null;
  },

  // Time-of-day only, carried on a fixed reference day (1899-12-30 UTC) so Hour/Minute/Second round-trip.
  TimeValue: ([value]) => {
    const date = toDate(value);
    if (!date) return null;
    return new Date(Date.UTC(1899, 11, 30, date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds(), date.getUTCMilliseconds()));
  },
};
