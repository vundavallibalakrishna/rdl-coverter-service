// VB.NET-style conversion functions (CInt/CDbl/CStr/CBool/CDec/Val). Args arrive already evaluated;
// bad/blank input yields a VB fallback (null or 0/'') rather than throwing.
import { toNumber, isBlank } from './shared.js';

// Banker's rounding (round half to even), matching VB CInt.
function bankersRound(n) {
  const floor = Math.floor(n);
  const diff = n - floor;
  if (diff < 0.5) return floor;
  if (diff > 0.5) return floor + 1;
  return floor % 2 === 0 ? floor : floor + 1; // exactly .5 -> nearest even
}

export const conversionFunctions = {
  CInt(args) {
    const n = toNumber(args[0]);
    return Number.isNaN(n) ? null : bankersRound(n);
  },

  CDbl(args) {
    const n = toNumber(args[0]);
    return Number.isNaN(n) ? null : n;
  },

  // No native decimal in JS; behaves as CDbl.
  CDec(args) {
    const n = toNumber(args[0]);
    return Number.isNaN(n) ? null : n;
  },

  CStr(args) {
    const v = args[0];
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'True' : 'False';
    if (v instanceof Date) return v.toLocaleString('en-US');
    return String(v);
  },

  // VB Str reserves a leading character for the sign of non-negative numbers, uses "." as its decimal
  // separator, and emits a minus sign (without the leading space) for negative numbers. This is deliberately
  // distinct from CStr: SSRS reports sometimes concatenate Str(...) and rely on that sign-space.
  Str(args) {
    const n = toNumber(args[0]);
    if (!Number.isFinite(n)) return null;
    const rendered = String(Object.is(n, -0) ? 0 : n);
    return n < 0 ? rendered : ` ${rendered}`;
  },

  CBool(args) {
    const v = args[0];
    if (isBlank(v)) return false;
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    const s = String(v).trim();
    if (/^true$/i.test(s)) return true;
    if (/^false$/i.test(s)) return false;
    const n = toNumber(s);
    return Number.isNaN(n) ? false : n !== 0;
  },

  // 64-bit-style integer; JS numbers can't hold full Long range, but banker's rounding matches CInt.
  CLng(args) {
    const n = toNumber(args[0]);
    return Number.isNaN(n) ? null : bankersRound(n);
  },

  // No distinct single-precision type in JS; behaves as CDbl.
  CSng(args) {
    const n = toNumber(args[0]);
    return Number.isNaN(n) ? null : n;
  },

  // Banker's-rounded integer clamped to the Byte range 0..255; out of range or invalid -> null.
  CByte(args) {
    const n = toNumber(args[0]);
    if (Number.isNaN(n)) return null;
    const rounded = bankersRound(n);
    if (rounded < 0 || rounded > 255) return null;
    return rounded;
  },

  // Uppercase hexadecimal string of the integer value; invalid -> ''.
  Hex(args) {
    const n = toNumber(args[0]);
    if (Number.isNaN(n)) return '';
    return Math.trunc(n).toString(16).toUpperCase();
  },

  // Octal string of the integer value; invalid -> ''.
  Oct(args) {
    const n = toNumber(args[0]);
    if (Number.isNaN(n)) return '';
    return Math.trunc(n).toString(8);
  },

  // Parse the leading numeric portion, ignoring whitespace; 0 when no leading number.
  Val(args) {
    const v = args[0];
    if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
    if (v === null || v === undefined) return 0;
    const s = String(v).replace(/\s+/g, '');
    const match = s.match(/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/);
    if (!match) return 0;
    const n = Number(match[0]);
    return Number.isNaN(n) ? 0 : n;
  },
};
