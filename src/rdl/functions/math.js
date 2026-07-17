// RDL/SSRS math functions (System.Math / VB semantics). Args arrive already-evaluated; blanks/non-numeric
// coerce to NaN and are returned as null so the evaluator never throws on bad input.
import { toNumber } from './shared.js';

// VB Math.Round: banker's rounding (round half to even). Symmetric across zero.
function bankersRound(value, digits = 0) {
  const sign = value < 0 ? -1 : 1;
  const factor = 10 ** digits;
  const x = Math.abs(value) * factor;
  const floor = Math.floor(x);
  const diff = x - floor;
  let rounded;
  if (Math.abs(diff - 0.5) < 1e-9) {
    rounded = floor % 2 === 0 ? floor : floor + 1; // tie -> nearest even
  } else {
    rounded = Math.round(x);
  }
  return (sign * rounded) / factor;
}

// Return null when a coerced value is not a finite/usable number.
const nn = (n) => (Number.isNaN(n) ? null : n);

export const mathFunctions = {
  Round(args) {
    const n = toNumber(args[0]);
    if (Number.isNaN(n)) return null;
    const digits = args.length > 1 ? toNumber(args[1]) : 0;
    const d = Number.isNaN(digits) ? 0 : Math.trunc(digits);
    return bankersRound(n, d);
  },
  Ceiling(args) {
    return nn(Math.ceil(toNumber(args[0])));
  },
  Floor(args) {
    return nn(Math.floor(toNumber(args[0])));
  },
  Abs(args) {
    return nn(Math.abs(toNumber(args[0])));
  },
  // VB Int: floor toward negative infinity (Int(-2.5) = -3).
  Int(args) {
    return nn(Math.floor(toNumber(args[0])));
  },
  // VB Fix: truncate toward zero (Fix(-2.5) = -2).
  Fix(args) {
    return nn(Math.trunc(toNumber(args[0])));
  },
  Sqrt(args) {
    const n = toNumber(args[0]);
    if (Number.isNaN(n) || n < 0) return null; // negative -> null, not NaN
    return Math.sqrt(n);
  },
  Sign(args) {
    const n = toNumber(args[0]);
    return Number.isNaN(n) ? null : Math.sign(n);
  },
  Power(args) {
    const base = toNumber(args[0]);
    const exp = toNumber(args[1]);
    if (Number.isNaN(base) || Number.isNaN(exp)) return null;
    return nn(Math.pow(base, exp));
  },
  // Natural log by default; with a second arg, log to that base. Domain: n > 0 (and base > 0, base != 1).
  Log(args) {
    const n = toNumber(args[0]);
    if (Number.isNaN(n) || n <= 0) return null;
    if (args.length > 1) {
      const base = toNumber(args[1]);
      if (Number.isNaN(base) || base <= 0 || base === 1) return null;
      return nn(Math.log(n) / Math.log(base));
    }
    return nn(Math.log(n));
  },
  Log10(args) {
    const n = toNumber(args[0]);
    if (Number.isNaN(n) || n <= 0) return null; // domain: n > 0
    return nn(Math.log10(n));
  },
  Exp(args) {
    const n = toNumber(args[0]);
    if (Number.isNaN(n)) return null;
    return nn(Math.exp(n));
  },
  // Math.Truncate: toward zero (Truncate(-2.7) = -2).
  Truncate(args) {
    return nn(Math.trunc(toNumber(args[0])));
  },
  Atan(args) {
    return nn(Math.atan(toNumber(args[0])));
  },
  Sin(args) {
    return nn(Math.sin(toNumber(args[0])));
  },
  Cos(args) {
    return nn(Math.cos(toNumber(args[0])));
  },
  Tan(args) {
    return nn(Math.tan(toNumber(args[0])));
  },
};
