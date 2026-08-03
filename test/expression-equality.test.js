// VB equality semantics, which are NOT JavaScript's where Nothing is involved.
//
// Why this matters: SSRS reports routinely drive conditional formatting from a query-computed row-number
// field, e.g. Combined Assurance draws its group-boundary rule with
//   <TopBorder><Style>=IIF(Fields!rn.Value = 0, "Solid", "None")</Style></TopBorder>
// and hides per-group cells with `=IIF(Fields!rn.Value = 0, False, True)`. When that field is NULL, VB
// coerces Nothing to 0, so the comparison is True and SSRS still draws the rule. Comparing with JS `===`
// plus a String() fallback made it False, so the rule silently vanished — a full-width table line that
// simply stopped halfway across, with no error anywhere.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  configureExpressionPlanCache,
  evaluateExpression,
  expressionPlanCacheStatistics,
} from '../src/rdl/expression.js';

const ev = (expr, fields = {}) => evaluateExpression(expr, { fields, parameters: {}, globals: {}, dataset: [], datasets: {} });

test('Nothing compares equal to zero, as VB coerces it to the numeric default', () => {
  assert.equal(ev('=Fields!rn.Value = 0', { rn: null }), true);
  assert.equal(ev('=Fields!rn.Value = 0', { rn: undefined }), true);
  assert.equal(ev('=Fields!rn.Value = 0', { rn: 0 }), true);
  assert.equal(ev('=Fields!rn.Value = 0', { rn: 1 }), false);
  // The border rule this exists for: a NULL row number must still open a group boundary.
  assert.equal(ev('=IIF(Fields!rn.Value = 0, "Solid", "None")', { rn: null }), 'Solid');
  assert.equal(ev('=IIF(Fields!rn.Value = 0, "Solid", "None")', { rn: 2 }), 'None');
});

test('Nothing compares equal to the empty string but not to other text', () => {
  assert.equal(ev('=Fields!x.Value = ""', { x: null }), true);
  assert.equal(ev('=Fields!x.Value = "abc"', { x: null }), false);
  // The RDL's own `Fields!Division.Value = Nothing` idiom, both directions.
  assert.equal(ev('=Fields!x.Value = Nothing', { x: null }), true);
  assert.equal(ev('=Fields!x.Value = Nothing', { x: '' }), true);
  assert.equal(ev('=Fields!x.Value = Nothing', { x: 'set' }), false);
});

test('a null field is not equal to the literal text "null"', () => {
  // Regression: String(null) is "null", so the old String()-based comparison matched this.
  assert.equal(ev('=Fields!x.Value = "null"', { x: null }), false);
  assert.equal(ev('=Fields!x.Value = "undefined"', { x: undefined }), false);
});

test('<> is the exact negation of =, including for Nothing', () => {
  for (const [expr, fields] of [
    ['=Fields!rn.Value', { rn: null }],
    ['=Fields!rn.Value', { rn: 0 }],
    ['=Fields!rn.Value', { rn: 5 }],
    ['=Fields!rn.Value', { rn: '0' }],
    ['=Fields!rn.Value', { rn: '' }],
  ]) {
    const equal = ev(`${expr} = 0`, fields);
    const notEqual = ev(`${expr} <> 0`, fields);
    assert.equal(notEqual, !equal, `= and <> disagreed for ${JSON.stringify(fields)}`);
  }
});

test('numeric text still compares equal to its number, and equality stays reflexive', () => {
  assert.equal(ev('=Fields!x.Value = 0', { x: '0' }), true);
  assert.equal(ev('=Fields!x.Value = "A"', { x: 'A' }), true);
  assert.equal(ev('=Fields!x.Value = Fields!x.Value', { x: null }), true);
  assert.equal(ev('=1 = 1'), true);
  assert.equal(ev('=1 = 2'), false);
});

// VB relational comparison semantics (Phase 4): numeric-looking operands compare numerically even as
// strings, and Nothing coerces to 0 — mirroring vbEqual. Raw JS comparison got both wrong.
test('relational operators compare numeric-looking strings numerically, not lexically', () => {
  assert.equal(ev('=Fields!a.Value >= 9', { a: '80' }), true); // "80" >= 9 numerically (lexical "8"<"9" would be false)
  assert.equal(ev('=Fields!a.Value < Fields!b.Value', { a: '9', b: '80' }), true);
  assert.equal(ev('="10" > "9"'), true); // both numeric strings -> 10 > 9
  assert.equal(ev('=Fields!score.Value >= 80', { score: 80 }), true);
  assert.equal(ev('=Fields!score.Value > 80', { score: 79 }), false);
});

test('relational operators treat Nothing as 0 and fall back to string order for non-numeric text', () => {
  assert.equal(ev('=Fields!x.Value < 5', { x: null }), true); // Nothing -> 0 < 5
  assert.equal(ev('=5 > Fields!x.Value', { x: null }), true);
  assert.equal(ev('=Fields!a.Value < Fields!b.Value', { a: 'apple', b: 'banana' }), true); // string order
  assert.equal(ev('=Fields!a.Value > Fields!b.Value', { a: 'zebra', b: 'apple' }), true);
});

test('cached expression plans resolve every row context afresh instead of caching values', () => {
  const expression = '=IIF(Fields!score.Value >= 80, Fields!label.Value & " pass", Fields!label.Value & " review")';
  const before = expressionPlanCacheStatistics();
  assert.equal(ev(expression, { score: 90, label: 'Alpha' }), 'Alpha pass');
  assert.equal(ev(expression, { score: 40, label: 'Beta' }), 'Beta review');
  const after = expressionPlanCacheStatistics();
  assert.ok(after.entries >= before.entries);
  assert.ok(after.misses > before.misses);
  assert.ok(after.hits > before.hits);
});

test('expression plan caching has a complete operational rollback path', () => {
  configureExpressionPlanCache(false);
  try {
    assert.equal(ev('=Fields!value.Value + 1', { value: 1 }), 2);
    assert.equal(ev('=Fields!value.Value + 1', { value: 5 }), 6);
    const statistics = expressionPlanCacheStatistics();
    assert.equal(statistics.enabled, false);
    assert.equal(statistics.entries, 0);
    assert.equal(statistics.hits, 0);
    assert.ok(statistics.misses > 0);
  } finally {
    configureExpressionPlanCache(true);
  }
});
