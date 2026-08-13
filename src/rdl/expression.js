import { ServiceError } from '../errors.js';
import { FUNCTION_REGISTRY } from './functions/index.js';
import { CUSTOM_CODE_FUNCTION_REGISTRY } from './functions/customCode.js';
import { formatNet } from './format.js';

function splitArguments(value) {
  const args = [];
  let depth = 0;
  let quote = null;
  let start = 0;
  for (let index = 0; index < value.length; index += 1) {
    const character = value[index];
    if (quote) {
      if (character === quote && value[index - 1] !== '\\') quote = null;
      continue;
    }
    if (character === '"' || character === "'") quote = character;
    else if (character === '(') depth += 1;
    else if (character === ')') depth -= 1;
    else if (character === ',' && depth === 0) {
      args.push(value.slice(start, index).trim());
      start = index + 1;
    }
  }
  args.push(value.slice(start).trim());
  return args;
}

function findTopLevelOperator(value, operators) {
  let depth = 0;
  let quote = null;
  for (let index = value.length - 1; index >= 0; index -= 1) {
    const character = value[index];
    if (quote) {
      if (character === quote && value[index - 1] !== '\\') quote = null;
      continue;
    }
    if (character === '"' || character === "'") quote = character;
    else if (character === ')') depth += 1;
    else if (character === '(') depth -= 1;
    if (depth !== 0) continue;
    for (const operator of operators) {
      const start = index - operator.length + 1;
      if (start >= 0 && value.slice(start, index + 1).toLowerCase() === operator.toLowerCase()) {
        return { index: start, operator };
      }
    }
  }
  return null;
}

function unwrap(value) {
  let current = value.trim();
  while (current.startsWith('(') && current.endsWith(')')) {
    let depth = 0;
    let wraps = true;
    for (let index = 0; index < current.length; index += 1) {
      if (current[index] === '(') depth += 1;
      if (current[index] === ')') depth -= 1;
      if (depth === 0 && index < current.length - 1) {
        wraps = false;
        break;
      }
    }
    if (!wraps) break;
    current = current.slice(1, -1).trim();
  }
  return current;
}

export function formatValue(value, format) {
  if (value === null || value === undefined) return '';
  const normalized = String(format || '').toLowerCase();
  if (normalized.includes('%')) return `${(Number(value) * 100).toFixed(normalized.includes('.00') ? 2 : 0)}%`;
  if (/^[cC]\d?$/.test(String(format || ''))) {
    const digits = Number(String(format).slice(1) || 2);
    return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: digits }).format(Number(value));
  }
  if (/^[nN]\d?$/.test(String(format || ''))) {
    const digits = Number(String(format).slice(1) || 2);
    return new Intl.NumberFormat('en-US', { minimumFractionDigits: digits, maximumFractionDigits: digits }).format(Number(value));
  }
  if (value instanceof Date || /^\d{4}-\d{2}-\d{2}/.test(String(value))) {
    const date = new Date(value);
    if (!Number.isNaN(date.getTime())) {
      const pattern = String(format || 'dd/MM/yyyy');
      if (pattern === 'y' || pattern === 'Y') return new Intl.DateTimeFormat('en-GB', { month: 'long', year: 'numeric', timeZone: 'UTC' }).format(date);
      if (pattern === 'd') return new Intl.DateTimeFormat('en-GB', { dateStyle: 'short', timeZone: 'UTC' }).format(date);
      if (pattern === 'D') return new Intl.DateTimeFormat('en-GB', { dateStyle: 'long', timeZone: 'UTC' }).format(date);
      if (pattern === 'g' || pattern === 'G') return new Intl.DateTimeFormat('en-GB', { dateStyle: 'short', timeStyle: pattern === 'G' ? 'medium' : 'short', timeZone: 'UTC' }).format(date);
      // Custom date/time patterns resolve through the complete .NET format engine. The local token table
      // this replaced covered only {yyyy, MMM, MM, dd, HH, mm, ss} and substituted them with a single
      // first-match-wins alternation, so any longer token was cut short and its remainder emitted as a
      // literal: "MMMM yyyy" produced "AugM 2026" because MMM consumed three of the four Ms. The same
      // truncation hit dddd/ddd, and the single-character forms M/d/y/h/s plus tt were not tokens at all
      // and survived verbatim. Both engines resolve custom patterns in UTC, so every token that already
      // worked is unchanged.
      return formatNet(date, pattern);
    }
  }
  // Fall back to the full .NET format-string engine for everything the branches above don't handle
  // (custom numeric like "#,##0.00", standard "P1"/"F2", section formats, …). Additive: the cases handled
  // above already returned, so existing date/currency/percent output is unchanged.
  return formatNet(value, format);
}

function evaluateArgument(argument, context) {
  const value = String(argument ?? '').trim();
  return evaluateExpression(value.startsWith('=') ? value : `=${value}`, context);
}

const AGGREGATE_FUNCTIONS = new Set(['sum', 'count', 'countdistinct', 'avg', 'min', 'max', 'stdev', 'stdevp', 'var', 'varp']);

function aggregate(name, values) {
  const clean = values.filter((value) => value !== null && value !== undefined);
  if (name === 'count') return clean.length;
  if (name === 'countdistinct') return new Set(clean.map(String)).size;
  const numbers = clean.map(Number).filter(Number.isFinite);
  if (name === 'sum') return numbers.reduce((sum, item) => sum + item, 0);
  if (name === 'avg') return numbers.length ? numbers.reduce((sum, item) => sum + item, 0) / numbers.length : 0;
  if (name === 'min') return numbers.length ? Math.min(...numbers) : null;
  if (name === 'max') return numbers.length ? Math.max(...numbers) : null;
  if (name === 'stdev' || name === 'stdevp' || name === 'var' || name === 'varp') {
    if (numbers.length === 0) return null;
    const mean = numbers.reduce((sum, item) => sum + item, 0) / numbers.length;
    const squares = numbers.reduce((sum, item) => sum + (item - mean) ** 2, 0);
    const population = name === 'varp' || name === 'stdevp';
    const variance = squares / (population ? numbers.length : Math.max(1, numbers.length - 1));
    return name === 'stdev' || name === 'stdevp' ? Math.sqrt(variance) : variance;
  }
  return null;
}

// Resolves an aggregate scope argument to a row set. A named scope resolves against the current group
// instances (context.scopes) first — so a Sum() in a group header/footer aggregates only that group —
// then dataset-by-name, then the current scope's rows (context.dataset). No scope arg → current scope.
function scopeRows(scopeArgument, context) {
  const scope = scopeArgument ? evaluateArgument(scopeArgument, context) : null;
  if (scope) {
    if (context.scopes && Object.prototype.hasOwnProperty.call(context.scopes, scope)) return context.scopes[scope];
    if (context.datasets?.[scope]) return context.datasets[scope];
  }
  return context.dataset || [];
}

// RunningValue/RowNumber give Nothing a different meaning from an omitted aggregate scope: it is the
// outermost containing scope and therefore does not reset at each nested group. Materialization exposes
// that processed (filtered/sorted) data-region row sequence through the dataset/data-region scope map.
// Falling back to context.dataset preserves direct evaluator use where there is no tablix context.
function runningScopeRows(scopeArgument, context) {
  const scope = scopeArgument ? evaluateArgument(scopeArgument, context) : null;
  if (scope !== null && scope !== undefined && scope !== '') {
    if (context.scopes && Object.prototype.hasOwnProperty.call(context.scopes, scope)) return context.scopes[scope];
    if (context.datasets?.[scope]) return context.datasets[scope];
    return context.dataset || [];
  }
  if (
    context.tablixDatasetName
    && context.scopes
    && Object.prototype.hasOwnProperty.call(context.scopes, context.tablixDatasetName)
  ) return context.scopes[context.tablixDatasetName];
  if (
    context.tablixName
    && context.scopes
    && Object.prototype.hasOwnProperty.call(context.scopes, context.tablixName)
  ) return context.scopes[context.tablixName];
  return context.outermostDataset || context.dataset || [];
}

function toDate(value) {
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
  if (value === null || value === undefined || value === '') return null;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

// VB.NET DateAdd: add `count` of the given interval to a date. Uses UTC to match formatValue().
function dateAdd(interval, count, dateValue) {
  const date = toDate(dateValue);
  if (!date) return null;
  const amount = Number(count) || 0;
  const result = new Date(date.getTime());
  switch (String(interval).toLowerCase()) {
    case 'yyyy': result.setUTCFullYear(result.getUTCFullYear() + amount); break;
    case 'q': result.setUTCMonth(result.getUTCMonth() + amount * 3); break;
    case 'm': result.setUTCMonth(result.getUTCMonth() + amount); break;
    case 'd': case 'y': case 'w': result.setUTCDate(result.getUTCDate() + amount); break;
    case 'ww': result.setUTCDate(result.getUTCDate() + amount * 7); break;
    case 'h': result.setUTCHours(result.getUTCHours() + amount); break;
    case 'n': result.setUTCMinutes(result.getUTCMinutes() + amount); break;
    case 's': result.setUTCSeconds(result.getUTCSeconds() + amount); break;
    default: return null;
  }
  return result;
}

function datePart(interval, dateValue) {
  const date = toDate(dateValue);
  if (!date) return null;
  const dayOfYear = () => Math.floor((date.getTime() - Date.UTC(date.getUTCFullYear(), 0, 1)) / 86400000) + 1;
  switch (String(interval).toLowerCase()) {
    case 'yyyy': return date.getUTCFullYear();
    case 'q': return Math.floor(date.getUTCMonth() / 3) + 1;
    case 'm': return date.getUTCMonth() + 1;
    case 'y': return dayOfYear();
    case 'd': return date.getUTCDate();
    case 'w': return date.getUTCDay() + 1;
    case 'ww': return Math.ceil((dayOfYear() + new Date(Date.UTC(date.getUTCFullYear(), 0, 1)).getUTCDay()) / 7);
    case 'h': return date.getUTCHours();
    case 'n': return date.getUTCMinutes();
    case 's': return date.getUTCSeconds();
    default: return null;
  }
}

function dateDiff(interval, startValue, endValue) {
  const start = toDate(startValue);
  const end = toDate(endValue);
  if (!start || !end) return null;
  const ms = end.getTime() - start.getTime();
  switch (String(interval).toLowerCase()) {
    case 'yyyy': return end.getUTCFullYear() - start.getUTCFullYear();
    case 'q': return (end.getUTCFullYear() - start.getUTCFullYear()) * 4 + (Math.floor(end.getUTCMonth() / 3) - Math.floor(start.getUTCMonth() / 3));
    case 'm': return (end.getUTCFullYear() - start.getUTCFullYear()) * 12 + (end.getUTCMonth() - start.getUTCMonth());
    case 'd': case 'y': return Math.round(ms / 86400000);
    case 'w': case 'ww': return Math.trunc(Math.round(ms / 86400000) / 7);
    case 'h': return Math.round(ms / 3600000);
    case 'n': return Math.round(ms / 60000);
    case 's': return Math.round(ms / 1000);
    default: return null;
  }
}

// Lookup family: match the source value against a destination expression evaluated over the named
// dataset's rows, returning the result expression from matching rows. Datasets are the caller-supplied
// rows already in context; no data source is executed.
function lookupMatches(sourceValue, destinationArg, rows, context) {
  return rows.filter((row) => {
    const destination = evaluateArgument(destinationArg, { ...context, fields: row });
    return destination === sourceValue || String(destination) === String(sourceValue);
  });
}

function evaluateFunction(name, args, context) {
  const normalized = name.toLowerCase();
  const customCodeFunction = CUSTOM_CODE_FUNCTION_REGISTRY[normalized];
  if (customCodeFunction) return customCodeFunction(args.map((arg) => evaluateArgument(arg, context)), context);
  if (normalized === 'iif') {
    if (args.length !== 3) throw new ServiceError('UNSUPPORTED_FEATURE', 'IIF requires three arguments');
    return evaluateArgument(args[0], context) ? evaluateArgument(args[1], context) : evaluateArgument(args[2], context);
  }
  if (normalized === 'first') {
    const rows = scopeRows(args[1], context);
    return evaluateArgument(args[0], { ...context, fields: rows[0] || context.fields || {} });
  }
  if (normalized === 'last') {
    const rows = scopeRows(args[1], context);
    return evaluateArgument(args[0], { ...context, fields: rows[rows.length - 1] || context.fields || {} });
  }
  if (normalized === 'format') return formatValue(evaluateArgument(args[0], context), evaluateArgument(args[1], context));
  if (normalized === 'isnothing') {
    const value = evaluateArgument(args[0], context);
    return value === null || value === undefined;
  }
  if (normalized === 'switch') {
    // Switch(condition1, value1, condition2, value2, ...): first value whose condition is true.
    for (let index = 0; index + 1 < args.length; index += 2) {
      if (evaluateArgument(args[index], context)) return evaluateArgument(args[index + 1], context);
    }
    return null;
  }
  if (normalized === 'choose') {
    // Choose(index, value1, value2, ...): 1-based selection.
    const selected = Number(evaluateArgument(args[0], context));
    return Number.isInteger(selected) && selected >= 1 && selected < args.length ? evaluateArgument(args[selected], context) : null;
  }
  if (normalized === 'join') {
    const value = evaluateArgument(args[0], context);
    const separator = args.length > 1 ? evaluateArgument(args[1], context) : ', ';
    const list = Array.isArray(value) ? value : [value];
    return list.map((entry) => (entry === null || entry === undefined ? '' : String(entry))).join(String(separator ?? ''));
  }
  if (normalized === 'countrows') return scopeRows(args[0], context).length;
  if (normalized === 'rownumber') {
    const rows = runningScopeRows(args[0], context);
    const index = context.fields ? rows.indexOf(context.fields) : -1;
    return (index >= 0 ? index : rows.length - 1) + 1;
  }
  if (normalized === 'previous') {
    // SSRS defines Previous(expression, scope) as "the value of expression for the PREVIOUS INSTANCE of
    // that scope", returning Nothing for the first instance. The instance sequence is a rendering concept
    // the aggregate row set cannot express: in a group scope the previous instance is the previous GROUP,
    // and the current instance's representative row is always row 0 of its own rows, so stepping back one
    // row inside `scopeRows` resolved to Nothing on every instance (turning a conditional group-boundary
    // border into a rule on every row). Materialization records the sequence it emitted; use it whenever
    // present and keep the positional walk for direct evaluator use outside a data region.
    const scopeName = args[1] ? evaluateArgument(args[1], context) : null;
    const recorded = scopeName ? context.previousInstances?.[scopeName] : context.previousInstance;
    if (recorded !== undefined) {
      return recorded
        ? evaluateArgument(args[0], { ...context, fields: recorded.fields, dataset: recorded.dataset })
        : null;
    }
    const rows = scopeRows(args[1], context);
    const index = context.fields ? rows.indexOf(context.fields) : rows.length - 1;
    return index > 0 ? evaluateArgument(args[0], { ...context, fields: rows[index - 1] }) : null;
  }
  if (normalized === 'cdate') return toDate(evaluateArgument(args[0], context));
  if (normalized === 'dateadd') return dateAdd(evaluateArgument(args[0], context), evaluateArgument(args[1], context), evaluateArgument(args[2], context));
  if (normalized === 'datediff') return dateDiff(evaluateArgument(args[0], context), evaluateArgument(args[1], context), evaluateArgument(args[2], context));
  if (normalized === 'datepart') return datePart(evaluateArgument(args[0], context), evaluateArgument(args[1], context));
  if (normalized === 'lookup' || normalized === 'lookupset' || normalized === 'multilookup') {
    const rows = context.datasets?.[String(evaluateArgument(args[3], context) ?? '')] || [];
    const resultFor = (row) => evaluateArgument(args[2], { ...context, fields: row });
    if (normalized === 'lookup') {
      const match = lookupMatches(evaluateArgument(args[0], context), args[1], rows, context)[0];
      return match ? resultFor(match) : null;
    }
    if (normalized === 'lookupset') {
      return lookupMatches(evaluateArgument(args[0], context), args[1], rows, context).map(resultFor);
    }
    const sources = evaluateArgument(args[0], context);
    return (Array.isArray(sources) ? sources : [sources]).map((source) => {
      const match = lookupMatches(source, args[1], rows, context)[0];
      return match ? resultFor(match) : null;
    });
  }
  if (AGGREGATE_FUNCTIONS.has(normalized)) {
    const rows = scopeRows(args[1], context);
    return aggregate(normalized, rows.map((row) => evaluateArgument(args[0], { ...context, fields: row })));
  }
  if (normalized === 'runningvalue') {
    // RunningValue(expression, aggregate, scope): the aggregate accumulated in row order from the
    // first row of the scope through the CURRENT row. The aggregate name is a bare identifier, not
    // an expression. context.fields is the current row (a member of the scope rows), so its index
    // marks where accumulation stops; when it cannot be located the full scope is used.
    const aggregateName = String(args[1] ?? '').trim().toLowerCase();
    if (!AGGREGATE_FUNCTIONS.has(aggregateName)) {
      throw new ServiceError('UNSUPPORTED_FEATURE', `Unsupported RunningValue aggregate: ${args[1]}`);
    }
    const rows = runningScopeRows(args[2], context);
    const currentIndex = context.fields ? rows.indexOf(context.fields) : -1;
    const upto = currentIndex >= 0 ? rows.slice(0, currentIndex + 1) : rows;
    return aggregate(aggregateName, upto.map((row) => evaluateArgument(args[0], { ...context, fields: row })));
  }
  // Registry functions (string / math / conversion / date / formatting / inspection) receive their
  // arguments already evaluated. Kept last so the hand-written scope-aware functions above always win.
  const registered = FUNCTION_REGISTRY[normalized];
  if (registered) return registered(args.map((arg) => evaluateArgument(arg, context)), context);
  throw new ServiceError('UNSUPPORTED_FEATURE', `Unsupported RDL expression function: ${name}`);
}

// True when the parenthesis at `openIndex` closes on the last character of `value` — i.e. the whole
// expression is a single call like `Sum(x)`, rather than `Sum(x) + Sum(y)` where the first "(" closes
// mid-expression. Quotes are skipped so parentheses inside string literals never affect the depth.
function closesAtEnd(value, openIndex) {
  let depth = 0;
  let quote = null;
  for (let index = openIndex; index < value.length; index += 1) {
    const character = value[index];
    if (quote) {
      if (character === quote && value[index - 1] !== '\\') quote = null;
      continue;
    }
    if (character === '"' || character === "'") quote = character;
    else if (character === '(') depth += 1;
    else if (character === ')') {
      depth -= 1;
      if (depth === 0) return index === value.length - 1;
    }
  }
  return false;
}

// VB `Like` pattern → RegExp: * = any run, ? = one char, # = one digit, [list]/[!list] = char class.
function isVbNumericEqualityValue(value) {
  return typeof value !== 'boolean'
    && !(typeof value === 'string' && value.trim() === '')
    && Number.isFinite(Number(value));
}

// VB equality, which is NOT JavaScript equality where Nothing is involved. VB coerces Nothing to the other
// operand's default value, so `Nothing = 0` and `Nothing = ""` are both True — SSRS reports rely on this to
// drive conditional borders from a NULL row-number field (e.g. Combined Assurance draws its group-boundary
// rule from `=IIF(Fields!rn.Value = 0, "Solid", "None")`, and a NULL rn must still count as 0 or the rule
// silently disappears). Comparing via String() alone got this backwards AND matched `Nothing = "null"`,
// because String(null) is the literal text "null". Arithmetic here already coerces Nothing to 0
// (`Nothing + 1` is 1), so equality treating it as unequal to 0 was internally inconsistent too.
function vbEqual(left, right) {
  const leftNothing = left === null || left === undefined;
  const rightNothing = right === null || right === undefined;
  if (leftNothing || rightNothing) {
    if (leftNothing && rightNothing) return true;
    const other = leftNothing ? right : left;
    if (typeof other === 'number') return other === 0;
    if (typeof other === 'string') return other === '';
    if (typeof other === 'boolean') return other === false;
    return false;
  }
  // Dataset values can arrive through JSON as text even when the RDL field/query type is numeric. VB/SSRS
  // coerces that text when the other operand is numeric, so zero-padded row numbers such as "001" must
  // satisfy `Fields!Number.Value = 1`. Keep string-to-string equality lexical: only a genuinely numeric
  // operand opts the comparison into numeric coercion, and blank/non-numeric text is never coerced.
  const oneOperandIsNumeric = typeof left === 'number' || typeof right === 'number';
  if (oneOperandIsNumeric && isVbNumericEqualityValue(left) && isVbNumericEqualityValue(right)) {
    return Number(left) === Number(right);
  }
  return left === right || String(left) === String(right);
}

// VB relational comparison for `< > <= >=`. Two differences from raw JS comparison: (1) numeric-looking
// operands compare NUMERICALLY even when the caller supplied them as strings (a query field typed as a
// number but delivered as "80" must satisfy `>= 9`, not lose to lexical order), and (2) Nothing coerces to
// the other operand's default (0 for numbers), consistent with vbEqual and VB arithmetic. Returns -1/0/1.
function vbCompare(left, right) {
  const l = left === null || left === undefined ? 0 : left;
  const r = right === null || right === undefined ? 0 : right;
  const isNumeric = (v) => !(typeof v === 'string' && v.trim() === '') && Number.isFinite(Number(v));
  if (isNumeric(l) && isNumeric(r)) {
    const ln = Number(l);
    const rn = Number(r);
    return ln < rn ? -1 : ln > rn ? 1 : 0;
  }
  const ls = String(l);
  const rs = String(r);
  return ls < rs ? -1 : ls > rs ? 1 : 0;
}

function vbLike(text, pattern) {
  let regex = '^';
  for (let index = 0; index < pattern.length; index += 1) {
    const character = pattern[index];
    if (character === '*') regex += '.*';
    else if (character === '?') regex += '.';
    else if (character === '#') regex += '\\d';
    else if (character === '[') {
      let cursor = index + 1;
      let charClass = '[';
      if (pattern[cursor] === '!') { charClass += '^'; cursor += 1; }
      while (cursor < pattern.length && pattern[cursor] !== ']') {
        charClass += pattern[cursor].replace(/[\\^\]]/g, '\\$&');
        cursor += 1;
      }
      regex += `${charClass}]`;
      index = cursor;
    } else regex += character.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  try { return new RegExp(`${regex}$`).test(text); } catch { return false; }
}

const BINARY_OPERATOR_GROUPS = Object.freeze([
  Object.freeze([' OrElse ', ' Or ']),
  Object.freeze([' Xor ']),
  Object.freeze([' AndAlso ', ' And ']),
  Object.freeze(['<>', '>=', '<=', '=', '>', '<', ' Like ']),
  Object.freeze(['&']),
  Object.freeze(['+', '-']),
  Object.freeze([' Mod ']),
  Object.freeze(['\\']),
  Object.freeze(['*', '/']),
  Object.freeze(['^']),
]);

// Expressions are immutable report metadata but are evaluated thousands of times against different row
// contexts. Cache only their structural analysis—not their result—so every field, parameter, scope and
// page-dependent value is still resolved afresh. The cap keeps direct library use bounded; production
// workers are already request-scoped and exit after one render.
const EXPRESSION_PLAN_CACHE_LIMIT = 20_000;
const expressionPlanCache = new Map();
let expressionPlanCacheEnabled = true;
let expressionPlanCacheHits = 0;
let expressionPlanCacheMisses = 0;

function expressionPlan(source) {
  const cached = expressionPlanCacheEnabled ? expressionPlanCache.get(source) : undefined;
  if (cached !== undefined) {
    expressionPlanCacheHits += 1;
    return cached;
  }
  expressionPlanCacheMisses += 1;
  const value = unwrap(source);
  let plan;
  if (!value) plan = { kind: 'literal', value: '' };
  else {
    const functionMatch = value.match(/^((?:Code\.)?[A-Za-z_][A-Za-z0-9_]*)\((.*)\)$/is);
    if (functionMatch && closesAtEnd(value, functionMatch[1].length)) {
      plan = { kind: 'function', name: functionMatch[1], args: splitArguments(functionMatch[2]) };
    } else {
      const notMatch = value.match(/^Not\b\s*(.+)$/is);
      if (notMatch) plan = { kind: 'not', operand: notMatch[1] };
      else {
        for (const operators of BINARY_OPERATOR_GROUPS) {
          const found = findTopLevelOperator(value, operators);
          if (!found) continue;
          plan = {
            kind: 'binary',
            operator: found.operator.trim().toLowerCase(),
            left: value.slice(0, found.index),
            right: value.slice(found.index + found.operator.length),
          };
          break;
        }
      }
    }
  }
  if (!plan) {
    const member = value.match(/^(Fields|Parameters|Globals|User|ReportItems|Variables)!([A-Za-z_][A-Za-z0-9_]*)(?:\.([A-Za-z]+))?(?:\(\s*(\d+)\s*\))?$/i);
    if (member) {
      plan = {
        kind: 'member',
        collectionName: member[1].toLowerCase(),
        itemName: member[2],
        property: (member[3] || 'Value').toLowerCase(),
        index: member[4] === undefined ? null : Number(member[4]),
      };
    } else {
      const vbConstants = {
        vbcrlf: '\r\n', vbnewline: '\r\n', vblf: '\n', vbcr: '\r', vbtab: '\t',
      };
      const lower = value.toLowerCase();
      if (Object.hasOwn(vbConstants, lower)) plan = { kind: 'literal', value: vbConstants[lower] };
      else if (/^Nothing$/i.test(value)) plan = { kind: 'literal', value: null };
      else if (/^(True|False)$/i.test(value)) plan = { kind: 'literal', value: lower === 'true' };
      else if (/^-?\d+(?:\.\d+)?$/.test(value)) plan = { kind: 'literal', value: Number(value) };
      else if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
        plan = { kind: 'literal', value: value.slice(1, -1).replace(/""/g, '"') };
      } else plan = { kind: 'unsupported', value };
    }
  }
  if (expressionPlanCacheEnabled) {
    if (expressionPlanCache.size >= EXPRESSION_PLAN_CACHE_LIMIT) expressionPlanCache.clear();
    expressionPlanCache.set(source, plan);
  }
  return plan;
}

export function configureExpressionPlanCache(enabled = true) {
  expressionPlanCacheEnabled = enabled !== false;
  expressionPlanCache.clear();
  expressionPlanCacheHits = 0;
  expressionPlanCacheMisses = 0;
}

export function expressionPlanCacheStatistics() {
  return {
    enabled: expressionPlanCacheEnabled,
    entries: expressionPlanCache.size,
    hits: expressionPlanCacheHits,
    misses: expressionPlanCacheMisses,
  };
}

export function evaluateExpression(input, context = {}) {
  if (input === undefined || input === null) return '';
  const source = String(input).trim();
  if (!source.startsWith('=')) return input;
  const plan = expressionPlan(source.slice(1));
  if (plan.kind === 'literal') return plan.value;
  if (plan.kind === 'function') return evaluateFunction(plan.name, plan.args, context);
  if (plan.kind === 'not') return !Boolean(evaluateExpression(`=${plan.operand}`, context));
  if (plan.kind === 'binary') {
    const left = evaluateExpression(`=${plan.left}`, context);
    const right = evaluateExpression(`=${plan.right}`, context);
    switch (plan.operator) {
      case 'or': case 'orelse': return Boolean(left) || Boolean(right);
      case 'and': case 'andalso': return Boolean(left) && Boolean(right);
      case 'xor': return Boolean(left) !== Boolean(right);
      case 'like': return vbLike(`${left ?? ''}`, `${right ?? ''}`);
      case '<>': return !vbEqual(left, right);
      case '=': return vbEqual(left, right);
      case '>=': return vbCompare(left, right) >= 0;
      case '<=': return vbCompare(left, right) <= 0;
      case '>': return vbCompare(left, right) > 0;
      case '<': return vbCompare(left, right) < 0;
      case '&': return `${left ?? ''}${right ?? ''}`;
      case '+': return Number.isFinite(Number(left)) && Number.isFinite(Number(right)) ? Number(left) + Number(right) : `${left ?? ''}${right ?? ''}`;
      case '-': return Number(left) - Number(right);
      case 'mod': { const l = Number(left); const r = Number(right); return Number.isFinite(l) && Number.isFinite(r) && r !== 0 ? l % r : null; }
      case '\\': { const l = Number(left); const r = Number(right); return Number.isFinite(l) && Number.isFinite(r) && r !== 0 ? Math.trunc(l / r) : null; }
      case '*': return Number(left) * Number(right);
      case '/': return Number(left) / Number(right);
      case '^': return Number(left) ** Number(right);
      default: break;
    }
  }

  // Collection member access: Collection!Name[.Property][(index)]. `.Value` (default) keeps the original
  // behaviour; the extra properties cover the common header/footer/conditional idioms.
  if (plan.kind === 'member') {
    const { collectionName, itemName, property, index } = plan;
    if (collectionName === 'variables') {
      // Resolve a report/group Variable by evaluating its definition in the current scope, guarding against
      // a variable that references itself (directly or via another variable).
      const definition = context.globals?.variables?.[itemName];
      if (definition === undefined || context.resolvingVariables?.has(itemName)) return null;
      const resolvingVariables = new Set(context.resolvingVariables || []);
      resolvingVariables.add(itemName);
      return evaluateExpression(String(definition), { ...context, resolvingVariables });
    }
    const sources = {
      fields: context.fields || {}, parameters: context.parameters || {}, globals: context.globals || {},
      user: context.user || {}, reportitems: context.reportItems || {},
    };
    const raw = sources[collectionName]?.[itemName];
    if (collectionName === 'fields') {
      if (property === 'ismissing') return context.fields ? !(itemName in context.fields) : true;
      if (property === 'count') return Array.isArray(raw) ? raw.length : (raw == null ? 0 : 1);
      return raw ?? null; // Value (default), Label, Key, …
    }
    if (collectionName === 'parameters') {
      if (property === 'count') return Array.isArray(raw) ? raw.length : (raw == null ? 0 : 1);
      if (property === 'ismultivalue') return Array.isArray(raw);
      if (index !== null) return Array.isArray(raw) ? (raw[index] ?? null) : (index === 0 ? (raw ?? null) : null);
      if (property === 'label') return context.parameters?.__rdlParameterLabels?.[itemName] ?? raw ?? null;
      return raw ?? null;
    }
    return raw ?? null; // globals, user, reportitems
  }
  throw new ServiceError('UNSUPPORTED_FEATURE', `Unsupported RDL expression: ${plan.value.slice(0, 120)}`);
}
