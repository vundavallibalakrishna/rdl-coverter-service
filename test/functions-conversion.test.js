// End-to-end tests for conversion functions via the expression evaluator.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { evaluateExpression } from '../src/rdl/expression.js';

test('CInt uses banker\'s rounding', () => {
  assert.equal(evaluateExpression('=CInt(2.5)', {}), 2);
  assert.equal(evaluateExpression('=CInt(3.5)', {}), 4);
  assert.equal(evaluateExpression('=CInt(2.4)', {}), 2);
  assert.equal(evaluateExpression('=CInt(-2.5)', {}), -2);
  assert.equal(evaluateExpression('=CInt("abc")', {}), null);
});

test('CDbl and CDec convert to float', () => {
  assert.equal(evaluateExpression('=CDbl("3.14")', {}), 3.14);
  assert.equal(evaluateExpression('=CDec("2.5")', {}), 2.5);
  assert.equal(evaluateExpression('=CDbl("")', {}), null);
});

test('CStr converts to string', () => {
  assert.equal(evaluateExpression('=CStr(123)', {}), '123');
  assert.equal(evaluateExpression('=CStr(Nothing)', {}), '');
  assert.equal(evaluateExpression('=CStr(True)', {}), 'True');
});

test('CBool follows VB truthiness', () => {
  assert.equal(evaluateExpression('=CBool(0)', {}), false);
  assert.equal(evaluateExpression('=CBool(5)', {}), true);
  assert.equal(evaluateExpression('=CBool("True")', {}), true);
  assert.equal(evaluateExpression('=CBool("false")', {}), false);
  assert.equal(evaluateExpression('=CBool("-1")', {}), true);
  assert.equal(evaluateExpression('=CBool("0")', {}), false);
});

test('Val parses leading numeric portion', () => {
  assert.equal(evaluateExpression('=Val("12.5abc")', {}), 12.5);
  assert.equal(evaluateExpression('=Val("abc")', {}), 0);
  assert.equal(evaluateExpression('=Val("3+4")', {}), 3);
  assert.equal(evaluateExpression('=Val("  12.5abc")', {}), 12.5);
});

test('CLng uses banker\'s rounding', () => {
  assert.equal(evaluateExpression('=CLng(2.5)', {}), 2);
  assert.equal(evaluateExpression('=CLng(3.5)', {}), 4);
  assert.equal(evaluateExpression('=CLng("10")', {}), 10);
  assert.equal(evaluateExpression('=CLng("")', {}), null);
});

test('CSng converts to float', () => {
  assert.equal(evaluateExpression('=CSng("3.14")', {}), 3.14);
  assert.equal(evaluateExpression('=CSng(2)', {}), 2);
  assert.equal(evaluateExpression('=CSng("abc")', {}), null);
});

test('CByte clamps to 0..255', () => {
  assert.equal(evaluateExpression('=CByte(300)', {}), null);
  assert.equal(evaluateExpression('=CByte(-1)', {}), null);
  assert.equal(evaluateExpression('=CByte(2.5)', {}), 2);
  assert.equal(evaluateExpression('=CByte(255)', {}), 255);
  assert.equal(evaluateExpression('=CByte("")', {}), null);
});

test('Hex returns uppercase hexadecimal string', () => {
  assert.equal(evaluateExpression('=Hex(255)', {}), 'FF');
  assert.equal(evaluateExpression('=Hex(0)', {}), '0');
  assert.equal(evaluateExpression('=Hex("abc")', {}), '');
});

test('Oct returns octal string', () => {
  assert.equal(evaluateExpression('=Oct(8)', {}), '10');
  assert.equal(evaluateExpression('=Oct(0)', {}), '0');
  assert.equal(evaluateExpression('=Oct("")', {}), '');
});

test('function names are case-insensitive', () => {
  assert.equal(evaluateExpression('=cint(3.5)', {}), 4);
  assert.equal(evaluateExpression('=VAL("7x")', {}), 7);
});
