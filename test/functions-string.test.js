import test from 'node:test';
import assert from 'node:assert/strict';
import { evaluateExpression } from '../src/rdl/expression.js';

const evalExpr = (expr) => evaluateExpression(expr, {});

test('Left', () => {
  assert.equal(evalExpr('=Left("Hello", 2)'), 'He');
  assert.equal(evalExpr('=Left("Hi", -1)'), '');
  assert.equal(evalExpr('=Left("Hi", 10)'), 'Hi');
  assert.equal(evalExpr('=left("Hi", 1)'), 'H'); // case-insensitive
});

test('Right', () => {
  assert.equal(evalExpr('=Right("Hello", 2)'), 'lo');
  assert.equal(evalExpr('=Right("Hi", -1)'), '');
  assert.equal(evalExpr('=Right("Hi", 10)'), 'Hi');
});

test('Mid', () => {
  assert.equal(evalExpr('=Mid("Hello", 2, 3)'), 'ell');
  assert.equal(evalExpr('=Mid("Hello", 2)'), 'ello');
  assert.equal(evalExpr('=Mid("Hello", 0, 2)'), 'He'); // start clamps to 1
  assert.equal(evalExpr('=Mid("Hello", 99)'), ''); // out of range
});

test('Len', () => {
  assert.equal(evalExpr('=Len("Hello")'), 5);
  assert.equal(evalExpr('=Len(Nothing)'), 0);
});

test('Trim / LTrim / RTrim', () => {
  assert.equal(evalExpr('=Trim("  hi  ")'), 'hi');
  assert.equal(evalExpr('=LTrim("  hi  ")'), 'hi  ');
  assert.equal(evalExpr('=RTrim("  hi  ")'), '  hi');
  assert.equal(evalExpr('=Trim(Nothing)'), '');
});

test('UCase / LCase', () => {
  assert.equal(evalExpr('=UCase("Hello")'), 'HELLO');
  assert.equal(evalExpr('=LCase("Hello")'), 'hello');
  assert.equal(evalExpr('=ucase("abc")'), 'ABC'); // case-insensitive name
});

test('Replace', () => {
  assert.equal(evalExpr('=Replace("a.b.c", ".", "-")'), 'a-b-c');
  assert.equal(evalExpr('=Replace("abc", "", "-")'), 'abc'); // empty find -> unchanged
});

test('InStr', () => {
  assert.equal(evalExpr('=InStr("Hello", "ll")'), 3); // 1-based
  assert.equal(evalExpr('=InStr("Hello", "z")'), 0); // not found
  assert.equal(evalExpr('=InStr(4, "abcabc", "a")'), 4); // start form
});

test('StrReverse', () => {
  assert.equal(evalExpr('=StrReverse("abc")'), 'cba');
  assert.equal(evalExpr('=StrReverse(Nothing)'), '');
});

test('Split (via Join)', () => {
  assert.equal(evalExpr('=Join(Split("a,b,c", ","), "-")'), 'a-b-c');
  assert.equal(evalExpr('=Join(Split("a b c"), "-")'), 'a-b-c'); // default space delimiter
});

test('InStrRev', () => {
  assert.equal(evalExpr('=InStrRev("a-b-c", "-")'), 4); // last occurrence, 1-based
  assert.equal(evalExpr('=InStrRev("a-b-c", "-", 3)'), 2); // at/before start
  assert.equal(evalExpr('=InStrRev("Hello", "z")'), 0); // not found
});

test('Space', () => {
  assert.equal(evalExpr('=Space(3)'), '   ');
  assert.equal(evalExpr('=Space(0)'), '');
  assert.equal(evalExpr('=Space(-1)'), ''); // negative -> ''
});

test('StrComp', () => {
  assert.equal(evalExpr('=StrComp("a", "b")'), -1);
  assert.equal(evalExpr('=StrComp("b", "a")'), 1);
  assert.equal(evalExpr('=StrComp("a", "a")'), 0);
  assert.equal(evalExpr('=StrComp("A", "a", 1)'), 0); // case-insensitive mode
});

test('StartsWith', () => {
  assert.equal(evalExpr('=StartsWith("Hello", "He")'), true);
  assert.equal(evalExpr('=StartsWith("Hello", "lo")'), false);
  assert.equal(evalExpr('=StartsWith("Hello", "")'), true); // empty prefix
});

test('EndsWith', () => {
  assert.equal(evalExpr('=EndsWith("Hello", "lo")'), true);
  assert.equal(evalExpr('=EndsWith("Hello", "He")'), false);
  assert.equal(evalExpr('=EndsWith(Nothing, "")'), true); // blank input, empty suffix
});

test('Contains', () => {
  assert.equal(evalExpr('=Contains("Hello", "ell")'), true);
  assert.equal(evalExpr('=Contains("Hello", "z")'), false);
  assert.equal(evalExpr('=Contains("Hello", "")'), true); // empty substr -> true
});
