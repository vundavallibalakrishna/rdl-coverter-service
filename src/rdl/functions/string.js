// String expression functions (VB.NET/SSRS semantics). Args arrive already evaluated; never throw on
// bad/blank input — return a VB-style fallback ('' or the input). Canonical casing; registry lowercases.
import { toNumber, toText } from './shared.js';

export const stringFunctions = {
  Left: ([s, n]) => {
    const text = toText(s);
    const count = toNumber(n);
    if (Number.isNaN(count) || count < 0) return '';
    return text.slice(0, count);
  },

  Right: ([s, n]) => {
    const text = toText(s);
    const count = toNumber(n);
    if (Number.isNaN(count) || count < 0) return '';
    if (count >= text.length) return text;
    return text.slice(text.length - count);
  },

  // VB Mid is 1-indexed. start<1 clamps to 1; missing length -> rest of string.
  Mid: ([s, start, length]) => {
    const text = toText(s);
    let begin = toNumber(start);
    if (Number.isNaN(begin)) return '';
    if (begin < 1) begin = 1;
    const from = begin - 1;
    if (from >= text.length) return '';
    if (length === undefined || length === null) return text.slice(from);
    const len = toNumber(length);
    if (Number.isNaN(len) || len <= 0) return '';
    return text.slice(from, from + len);
  },

  Len: ([s]) => toText(s).length,
  Trim: ([s]) => toText(s).trim(),
  LTrim: ([s]) => toText(s).replace(/^\s+/, ''),
  RTrim: ([s]) => toText(s).replace(/\s+$/, ''),
  UCase: ([s]) => toText(s).toUpperCase(),
  LCase: ([s]) => toText(s).toLowerCase(),

  // Replace all literal occurrences of find. Empty find -> string unchanged.
  Replace: ([s, find, replacement]) => {
    const text = toText(s);
    const target = toText(find);
    if (target === '') return text;
    return text.split(target).join(toText(replacement));
  },

  // VB InStr. Two forms: InStr(string, substr) or InStr(start, string, substr) with 1-based start.
  // Returns the 1-based index of the first occurrence, or 0 if not found.
  InStr: (args) => {
    let start = 1;
    let text;
    let search;
    if (args.length >= 3 && typeof args[0] === 'number') {
      start = toNumber(args[0]);
      if (Number.isNaN(start) || start < 1) start = 1;
      text = toText(args[1]);
      search = toText(args[2]);
    } else {
      text = toText(args[0]);
      search = toText(args[1]);
    }
    const index = text.indexOf(search, start - 1);
    return index === -1 ? 0 : index + 1;
  },

  StrReverse: ([s]) => [...toText(s)].reverse().join(''),

  // VB InStrRev. 1-based index of the LAST occurrence at/before start (default -1 = from end); 0 if none.
  InStrRev: ([s, substr, start]) => {
    const text = toText(s);
    const search = toText(substr);
    let from = start === undefined || start === null ? -1 : toNumber(start);
    if (Number.isNaN(from)) from = -1;
    // VB start is 1-based over the string; lastIndexOf takes a 0-based upper bound.
    const upper = from < 0 ? text.length - 1 : from - 1;
    const index = text.lastIndexOf(search, upper);
    return index === -1 ? 0 : index + 1;
  },

  // A string of n spaces. n<0 or NaN -> ''.
  Space: ([n]) => {
    const count = toNumber(n);
    if (Number.isNaN(count) || count < 0) return '';
    return ' '.repeat(Math.floor(count));
  },

  // VB StrComp. Returns -1, 0, or 1. compareMode=1 -> case-insensitive; default binary/ordinal.
  StrComp: ([a, b, compareMode]) => {
    let left = toText(a);
    let right = toText(b);
    if (toNumber(compareMode) === 1) {
      left = left.toLowerCase();
      right = right.toLowerCase();
    }
    if (left < right) return -1;
    if (left > right) return 1;
    return 0;
  },

  // .NET String.StartsWith -> boolean.
  StartsWith: ([s, prefix]) => toText(s).startsWith(toText(prefix)),

  // .NET String.EndsWith -> boolean.
  EndsWith: ([s, suffix]) => toText(s).endsWith(toText(suffix)),

  // .NET String.Contains -> boolean ('' substr -> true).
  Contains: ([s, substr]) => toText(s).includes(toText(substr)),

  // Return an array of substrings; default delimiter is a single space.
  Split: ([s, delimiter]) => {
    const text = toText(s);
    const sep = delimiter === undefined || delimiter === null ? ' ' : toText(delimiter);
    return text.split(sep);
  },
};
