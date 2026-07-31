// Safe, service-owned replacements for explicitly allowlisted `Code.*` calls.
//
// The RDL's embedded VB is never compiled, interpreted, or inspected by this registry. Each entry is a
// fixed native implementation with a documented contract. Unknown Code.* names remain fail-closed.
function integer(value) {
  const number = Number(value);
  return Number.isFinite(number) ? Math.round(number) : 0;
}

function calculateColor(args) {
  const y = integer(args[0]);
  const x = integer(args[1]);

  if (x === 1) {
    if (y >= 1 && y <= 4) return 'Green';
    if (y === 5) return '#ffff00';
    return '#00b050';
  }
  if (x === 2) {
    if (y >= 1 && y <= 2) return 'Green';
    if (y >= 3 && y <= 4) return '#ffff00';
    if (y === 5) return '#FFA500';
    return '#00b050';
  }
  if (x === 3) {
    if (y === 1) return 'Green';
    if (y >= 2 && y <= 3) return '#ffff00';
    if (y >= 4 && y <= 5) return '#FFA500';
    return '#00b050';
  }
  if (x === 4) {
    if (y === 1) return 'Green';
    if (y === 2) return '#ffff00';
    if (y === 3) return '#FFA500';
    if (y >= 4 && y <= 5) return '#ff0000';
    return '#00b050';
  }
  if (x === 5) {
    if (y === 1) return '#ffff00';
    if (y >= 2 && y <= 3) return '#FFA500';
    if (y >= 4 && y <= 5) return '#ff0000';
    return '#00b050';
  }
  return '#ff0000';
}

// Native equivalent of the client RDL's GetPercentLevel(Double) helper. SSRS/VB coerces Nothing to the
// numeric default for a Double argument, so nullish input starts at zero. Non-numeric input cannot produce
// a meaningful maturity band and also falls back to that deterministic default instead of executing code.
function getPercentLevel(args) {
  const raw = args[0];
  const totalScore = raw === null || raw === undefined || raw === '' ? 0 : Number(raw);
  const pct = (Number.isFinite(totalScore) ? totalScore : 0) * 100;
  if (pct <= 20) return 'Level 1';
  if (pct <= 40) return 'Level 2';
  if (pct <= 60) return 'Level 3';
  if (pct <= 80) return 'Level 4';
  return 'Level 5';
}

// Native equivalent of GetDistinct(Object()). The VB implementation returns Nothing for a missing array
// and a de-duplicated object array otherwise. Preserve first occurrence order to keep service output
// deterministic; Hashtable key enumeration order is unspecified and must not leak runtime-dependent order.
function getDistinct(args) {
  const input = args[0];
  if (input === null || input === undefined) return null;
  const values = Array.isArray(input) ? input : [input];
  const seen = new Set();
  const result = [];
  for (const value of values) {
    if (seen.has(value)) continue;
    seen.add(value);
    result.push(value);
  }
  return result;
}

export const CUSTOM_CODE_FUNCTIONS = Object.freeze({
  'Code.CalculateColor': calculateColor,
  'Code.GetPercentLevel': getPercentLevel,
  'Code.GetDistinct': getDistinct,
});

export const CUSTOM_CODE_FUNCTION_NAMES = Object.freeze(Object.keys(CUSTOM_CODE_FUNCTIONS));

export const CUSTOM_CODE_FUNCTION_REGISTRY = Object.freeze(Object.fromEntries(
  Object.entries(CUSTOM_CODE_FUNCTIONS).map(([name, implementation]) => [name.toLowerCase(), implementation]),
));
