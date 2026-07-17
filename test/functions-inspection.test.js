import { test } from 'node:test';
import assert from 'node:assert/strict';
import { evaluateExpression } from '../src/rdl/expression.js';

test('IsNumeric', () => {
  assert.equal(evaluateExpression('=IsNumeric("12.5")', {}), true);
  assert.equal(evaluateExpression('=IsNumeric("abc")', {}), false);
  assert.equal(evaluateExpression('=IsNumeric("")', {}), false);
});

test('IsDate', () => {
  assert.equal(evaluateExpression('=IsDate("2026-07-17")', {}), true);
  assert.equal(evaluateExpression('=IsDate("nope")', {}), false);
});

test('InScope', () => {
  assert.equal(evaluateExpression('=InScope("CatGroup")', { scopes: { CatGroup: [] } }), true);
  assert.equal(evaluateExpression('=InScope("Other")', { scopes: { CatGroup: [] } }), false);
});

test('Level', () => {
  assert.equal(evaluateExpression('=Level()', { scopes: { A: [], B: [] } }), 1);
  assert.equal(evaluateExpression('=Level()', {}), 0);
});

test('case-insensitivity', () => {
  assert.equal(evaluateExpression('=ISNUMERIC("42")', {}), true);
});

test('IsArray', () => {
  assert.equal(evaluateExpression('=IsArray(Fields!x.Value)', { fields: { x: [1, 2] } }), true);
  assert.equal(evaluateExpression('=IsArray(Fields!x.Value)', { fields: { x: 42 } }), false);
  assert.equal(evaluateExpression('=IsArray("abc")', {}), false);
});

test('IsError', () => {
  assert.equal(evaluateExpression('=IsError(1)', {}), false);
  assert.equal(evaluateExpression('=IsError("abc")', {}), false);
  assert.equal(evaluateExpression('=IsError(1/0)', {}), false);
});

test('RGB', () => {
  assert.equal(evaluateExpression('=RGB(255,0,0)', {}), '#FF0000');
  assert.equal(evaluateExpression('=RGB(0,255,0)', {}), '#00FF00');
  assert.equal(evaluateExpression('=RGB(0,0,255)', {}), '#0000FF');
  // clamps out-of-range and rounds
  assert.equal(evaluateExpression('=RGB(300,-10,15.6)', {}), '#FF0010');
});
