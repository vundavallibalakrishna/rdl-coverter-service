// End-to-end tests for RDL math functions via the expression evaluator.
import test from 'node:test';
import assert from 'node:assert/strict';
import { evaluateExpression } from '../src/rdl/expression.js';

const ev = (expr) => evaluateExpression(expr, {});

test('Round with digits', () => {
  assert.equal(ev('=Round(3.14159, 2)'), 3.14);
  assert.equal(ev('=Round(2.345, 2)'), 2.34);
});

test('Round uses bankers rounding (half to even)', () => {
  assert.equal(ev('=Round(2.5)'), 2);
  assert.equal(ev('=Round(3.5)'), 4);
  assert.equal(ev('=Round(0.5)'), 0);
  assert.equal(ev('=Round(1.5)'), 2);
  assert.equal(ev('=Round(-2.5)'), -2);
});

test('Ceiling and Floor', () => {
  assert.equal(ev('=Ceiling(2.1)'), 3);
  assert.equal(ev('=Floor(2.9)'), 2);
  assert.equal(ev('=Floor(-2.1)'), -3);
});

test('Abs', () => {
  assert.equal(ev('=Abs(-7.5)'), 7.5);
  assert.equal(ev('=Abs(4)'), 4);
});

test('Int vs Fix on negatives', () => {
  assert.equal(ev('=Int(-2.5)'), -3); // floor toward -inf
  assert.equal(ev('=Fix(-2.5)'), -2); // truncate toward zero
  assert.equal(ev('=Int(2.9)'), 2);
  assert.equal(ev('=Fix(2.9)'), 2);
});

test('Sqrt (negative -> null)', () => {
  assert.equal(ev('=Sqrt(9)'), 3);
  assert.equal(ev('=Sqrt(-4)'), null);
});

test('Sign', () => {
  assert.equal(ev('=Sign(-12)'), -1);
  assert.equal(ev('=Sign(0)'), 0);
  assert.equal(ev('=Sign(3)'), 1);
});

test('Power', () => {
  assert.equal(ev('=Power(2, 10)'), 1024);
  assert.equal(ev('=Power(9, 0.5)'), 3);
});

test('blank / non-numeric input returns null', () => {
  assert.equal(ev('=Round("abc")'), null);
  assert.equal(ev('=Abs("")'), null);
});

test('function names are case-insensitive', () => {
  assert.equal(ev('=round(3.5)'), 4);
  assert.equal(ev('=SQRT(16)'), 4);
});

test('Exp', () => {
  assert.equal(ev('=Exp(0)'), 1);
  assert.ok(Math.abs(ev('=Exp(1)') - Math.E) < 1e-9);
});

test('Log natural and with base', () => {
  assert.equal(ev('=Log(1)'), 0);
  assert.ok(Math.abs(ev('=Log(2.718281828459045)') - 1) < 1e-9);
  assert.ok(Math.abs(ev('=Log(8, 2)') - 3) < 1e-9);
  assert.ok(Math.abs(ev('=Log(1000, 10)') - 3) < 1e-9);
});

test('Log domain errors return null', () => {
  assert.equal(ev('=Log(0)'), null);
  assert.equal(ev('=Log(-5)'), null);
  assert.equal(ev('=Log(8, 1)'), null); // invalid base
  assert.equal(ev('=Log(8, -2)'), null);
});

test('Log10', () => {
  assert.equal(ev('=Log10(1000)'), 3);
  assert.equal(ev('=Log10(1)'), 0);
  assert.equal(ev('=Log10(0)'), null); // domain
  assert.equal(ev('=Log10(-10)'), null);
});

test('Truncate toward zero', () => {
  assert.equal(ev('=Truncate(-2.7)'), -2);
  assert.equal(ev('=Truncate(2.7)'), 2);
  assert.equal(Math.abs(ev('=Truncate(-0.9)')), 0); // toward zero (may be -0)
});

test('trig functions in radians', () => {
  assert.ok(Math.abs(ev('=Sin(0)')) < 1e-9);
  assert.ok(Math.abs(ev('=Cos(0)') - 1) < 1e-9);
  assert.ok(Math.abs(ev('=Tan(0)')) < 1e-9);
  assert.ok(Math.abs(ev('=Atan(0)')) < 1e-9);
  assert.ok(Math.abs(ev('=Sin(1.5707963267948966)') - 1) < 1e-9); // sin(pi/2)
});

test('new math functions return null on non-numeric input', () => {
  assert.equal(ev('=Exp("abc")'), null);
  assert.equal(ev('=Sin("")'), null);
  assert.equal(ev('=Truncate("x")'), null);
});
