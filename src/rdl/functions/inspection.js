// Inspection / validation expression functions. Args are already-evaluated values; context carries
// fields/parameters/globals/datasets/scopes. Never throws; returns plain values. RDL names are
// case-insensitive (the registry lower-cases keys).
import { toNumber, toDate, toText, isBlank } from './shared.js';

export const inspectionFunctions = {
  // True when the value parses as a number. Blank/'' is explicitly not numeric.
  IsNumeric(args) {
    const value = args[0];
    if (isBlank(value)) return false;
    return !Number.isNaN(toNumber(value));
  },

  // True when the value parses to a valid Date. Design choice: blanks are false, and bare numbers
  // (typeof 'number') are treated as NOT dates so report counts/measures like 2026 aren't mistaken
  // for dates. Everything else defers to toDate() returning a non-null Date.
  IsDate(args) {
    const value = args[0];
    if (isBlank(value)) return false;
    if (typeof value === 'number') return false;
    return toDate(value) !== null;
  },

  // SSRS InScope(scope): is the current cell/row within the named scope? Best-effort approximation —
  // true when the current row belongs to that named group scope (i.e. context.scopes has the key).
  InScope(args, context) {
    return Boolean(
      context
        && context.scopes
        && Object.prototype.hasOwnProperty.call(context.scopes, toText(args[0])),
    );
  },

  // SSRS Level(scope): recursive/nesting depth of the current scope. We do not track recursive parent
  // hierarchies, so this is a best-effort approximation: the count of active group scopes minus one,
  // clamped to >= 0 (0 when no scope info is available).
  Level(args, context) {
    return Math.max(0, Object.keys(context?.scopes || {}).length - 1);
  },

  // True when the value is a JS array (a multi-value parameter/field), else false.
  IsArray(args) {
    return Array.isArray(args[0]);
  },

  // Best-effort error test: true when the value is an Error instance or a NaN number, else false.
  // NOTE: arguments are pre-evaluated before this function runs, so IsError cannot catch errors thrown
  // during downstream evaluation of its argument (a known limitation of the eager-eval registry).
  IsError(args) {
    const value = args[0];
    if (value instanceof Error) return true;
    if (typeof value === 'number' && Number.isNaN(value)) return true;
    return false;
  },

  // SSRS RGB(r, g, b): build a '#RRGGBB' CSS hex color from three 0..255 components. Each component is
  // rounded and clamped to 0..255. e.g. RGB(255, 0, 0) -> '#FF0000'.
  RGB(args) {
    const clamp = (v) => {
      const n = Math.round(toNumber(v));
      if (Number.isNaN(n)) return 0;
      return Math.max(0, Math.min(255, n));
    };
    const hex = (v) => clamp(v).toString(16).padStart(2, '0').toUpperCase();
    return `#${hex(args[0])}${hex(args[1])}${hex(args[2])}`;
  },
};
