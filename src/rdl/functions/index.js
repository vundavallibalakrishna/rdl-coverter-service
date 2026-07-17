import { stringFunctions } from './string.js';
import { mathFunctions } from './math.js';
import { conversionFunctions } from './conversion.js';
import { datetimeFunctions } from './datetime.js';
import { formattingFunctions } from './formatting.js';
import { inspectionFunctions } from './inspection.js';

// Each category module exports a map of { CanonicalName: (args, context) => value }, where `args` are the
// already-evaluated argument values and `context` is the expression evaluation context. New expression
// functions are added purely by extending a category module — the parser's supported-function set and the
// fail-closed capability catalogue both derive their names from FUNCTION_NAMES below, so nothing else needs
// to change.
const CATEGORY_MAPS = [
  stringFunctions,
  mathFunctions,
  conversionFunctions,
  datetimeFunctions,
  formattingFunctions,
  inspectionFunctions,
];

// Canonical names (e.g. "FormatNumber") for the capability catalogue and analyze output.
export const FUNCTION_NAMES = CATEGORY_MAPS.flatMap((map) => Object.keys(map));

// Lower-cased lookup used by the evaluator (RDL function names are case-insensitive).
export const FUNCTION_REGISTRY = Object.fromEntries(
  CATEGORY_MAPS.flatMap((map) => Object.entries(map)).map(([name, impl]) => [name.toLowerCase(), impl]),
);
