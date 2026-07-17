import assert from 'node:assert/strict';
import test from 'node:test';
import { formatNet } from '../src/rdl/format.js';

// Expectations are pinned to the actual output produced on this machine's ICU
// build, formatting all dates in UTC.

test('null and undefined values render as empty string', () => {
  assert.equal(formatNet(null, 'N2'), '');
  assert.equal(formatNet(undefined, 'C2'), '');
  assert.equal(formatNet(null, ''), '');
});

test('blank or non-string format falls back to String(value)', () => {
  assert.equal(formatNet(42, ''), '42');
  assert.equal(formatNet(42, '   '), '42');
  assert.equal(formatNet('abc', undefined), 'abc');
});

test('standard numeric specifiers', () => {
  assert.equal(formatNet(1234.5, 'N2'), '1,234.50');
  assert.equal(formatNet(1234, 'N0'), '1,234');
  assert.equal(formatNet(1234.5, 'C2'), '$1,234.50');
  assert.equal(formatNet(1234.5, 'C'), '$1,234.50');
  assert.equal(formatNet(0.125, 'P1'), '12.5%');
  assert.equal(formatNet(0.5, 'P0'), '50%');
  assert.equal(formatNet(3.14159, 'F2'), '3.14');
  assert.equal(formatNet(42, 'D5'), '00042');
  assert.equal(formatNet(-42, 'D5'), '-00042');
  assert.equal(formatNet(7, 'D'), '7');
});

test('custom numeric formats', () => {
  assert.equal(formatNet(1234.567, '#,##0.00'), '1,234.57');
  assert.equal(formatNet(0.5, '0.0%'), '50.0%');
  assert.equal(formatNet(5, '000'), '005');
  assert.equal(formatNet(1234.5, '#,##0'), '1,235');
  assert.equal(formatNet(0.25, '#0.##'), '0.25');
});

test('section syntax: positive / negative / zero', () => {
  const fmt = '$#,##0.00;($#,##0.00);"-"';
  assert.equal(formatNet(1234.5, fmt), '$1,234.50');
  assert.equal(formatNet(-1234.5, fmt), '($1,234.50)');
  assert.equal(formatNet(0, fmt), '-');
});

test('single-section negative values get an automatic minus sign', () => {
  assert.equal(formatNet(-1234.5, '#,##0.00'), '-1,234.50');
});

test('date custom formats in UTC', () => {
  assert.equal(
    formatNet(new Date('2026-07-17T00:00:00Z'), 'MMMM d, yyyy'),
    'July 17, 2026',
  );
  assert.equal(
    formatNet(new Date('2026-07-17T00:00:00Z'), 'dd/MM/yyyy'),
    '17/07/2026',
  );
  assert.equal(formatNet(new Date('2026-07-17T13:05:00Z'), 'HH:mm'), '13:05');
  assert.equal(formatNet(new Date('2026-07-17T13:05:00Z'), 'hh:mm tt'), '01:05 PM');
  assert.equal(formatNet(new Date('2026-07-17T00:00:00Z'), 'MMMM'), 'July');
});

test('date standard specifiers in UTC', () => {
  const d = new Date('2026-07-17T13:05:00Z');
  assert.equal(formatNet(d, 'd'), '7/17/26');
  assert.equal(formatNet(d, 'D'), 'July 17, 2026');
  assert.equal(formatNet(d, 't'), '1:05 PM');
  assert.equal(formatNet(d, 'T'), '1:05:00 PM');
  assert.equal(formatNet(d, 'g'), '7/17/26 1:05 PM');
});

test('date-parseable strings are formatted as dates when the format is a date format', () => {
  assert.equal(formatNet('2026-07-17T00:00:00Z', 'MMMM d, yyyy'), 'July 17, 2026');
});

test('numeric strings are formatted as numbers', () => {
  assert.equal(formatNet('1234.5', 'N2'), '1,234.50');
});

test('booleans coerce to 0 / 1 for numeric formats', () => {
  assert.equal(formatNet(true, 'N0'), '1');
  assert.equal(formatNet(false, 'N0'), '0');
});

test('unmatched value / format combinations fall back to the raw string', () => {
  assert.equal(formatNet('hello', 'N2'), 'hello');
});

test('never throws on odd input', () => {
  assert.doesNotThrow(() => formatNet({}, 'N2'));
  assert.doesNotThrow(() => formatNet(Number.NaN, 'N2'));
  assert.doesNotThrow(() => formatNet(Infinity, '#,##0.00'));
});
