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

export const CUSTOM_CODE_FUNCTIONS = Object.freeze({
  'Code.CalculateColor': calculateColor,
});

export const CUSTOM_CODE_FUNCTION_NAMES = Object.freeze(Object.keys(CUSTOM_CODE_FUNCTIONS));

export const CUSTOM_CODE_FUNCTION_REGISTRY = Object.freeze(Object.fromEntries(
  Object.entries(CUSTOM_CODE_FUNCTIONS).map(([name, implementation]) => [name.toLowerCase(), implementation]),
));

