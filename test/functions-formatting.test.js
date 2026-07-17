import { test } from 'node:test';
import assert from 'node:assert/strict';
import { evaluateExpression } from '../src/rdl/expression.js';

test('FormatNumber applies thousands separators and decimals', () => {
  assert.equal(evaluateExpression('=FormatNumber(1234.5, 2)', {}), '1,234.50');
  assert.equal(evaluateExpression('=FormatNumber(1234.5)', {}), '1,234.50');
  assert.equal(evaluateExpression('=FormatNumber(1000, 0)', {}), '1,000');
  assert.equal(evaluateExpression('=FormatNumber("abc")', {}), '');
  assert.equal(evaluateExpression('=FormatNumber("")', {}), '');
});

test('FormatCurrency uses USD symbol and 2 default decimals', () => {
  assert.equal(evaluateExpression('=FormatCurrency(1234.5)', {}), '$1,234.50');
  assert.equal(evaluateExpression('=FormatCurrency(1234.5, 0)', {}), '$1,235');
  assert.equal(evaluateExpression('=FormatCurrency("nope")', {}), '');
});

test('FormatPercent multiplies ratio by 100 and appends %', () => {
  assert.equal(evaluateExpression('=FormatPercent(0.125, 1)', {}), '12.5%');
  assert.equal(evaluateExpression('=FormatPercent(0.5)', {}), '50.00%');
  assert.equal(evaluateExpression('=FormatPercent("bad")', {}), '');
});

test('FormatDateTime honors the VB DateFormat enum', () => {
  assert.equal(evaluateExpression('=FormatDateTime("2026-07-17", 1)', {}), 'Friday, July 17, 2026');
  assert.equal(evaluateExpression('=FormatDateTime("2026-07-17", 2)', {}), '7/17/2026');
  assert.equal(evaluateExpression('=FormatDateTime("2026-07-17T12:30:00Z", 3)', {}), '12:30:00 PM');
  assert.equal(evaluateExpression('=FormatDateTime("2026-07-17T12:30:00Z", 4)', {}), '12:30 PM');
  assert.equal(evaluateExpression('=FormatDateTime("2026-07-17", 0)', {}), '7/17/2026');
  assert.equal(evaluateExpression('=FormatDateTime("2026-07-17T12:30:00Z")', {}), '7/17/2026, 12:30:00 PM');
  assert.equal(evaluateExpression('=FormatDateTime("2026-07-17", 0)', {}), '7/17/2026');
  assert.equal(evaluateExpression('=FormatDateTime("not-a-date", 2)', {}), '');
});

test('FormatDateTime accepts the enum name string', () => {
  assert.equal(evaluateExpression('=FormatDateTime("2026-07-17", "ShortDate")', {}), '7/17/2026');
  assert.equal(evaluateExpression('=FormatDateTime("2026-07-17", "longdate")', {}), 'Friday, July 17, 2026');
});

test('formatting function names are case-insensitive', () => {
  assert.equal(evaluateExpression('=formatnumber(1234.5, 2)', {}), '1,234.50');
  assert.equal(evaluateExpression('=FORMATCURRENCY(1234.5)', {}), '$1,234.50');
});
