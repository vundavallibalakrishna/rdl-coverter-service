// End-to-end tests for RDL date/time expression functions via the expression evaluator.
import test from 'node:test';
import assert from 'node:assert/strict';
import { evaluateExpression } from '../src/rdl/expression.js';

test('Year/Month/Day/Hour/Minute/Second extract UTC parts', () => {
  assert.equal(evaluateExpression('=Year("2026-07-17")', {}), 2026);
  assert.equal(evaluateExpression('=Month("2026-07-17")', {}), 7);
  assert.equal(evaluateExpression('=Day("2026-07-17")', {}), 17);
  assert.equal(evaluateExpression('=Hour("2026-07-17T13:45:30Z")', {}), 13);
  assert.equal(evaluateExpression('=Minute("2026-07-17T13:45:30Z")', {}), 45);
  assert.equal(evaluateExpression('=Second("2026-07-17T13:45:30Z")', {}), 30);
});

test('invalid input returns null, not a throw', () => {
  assert.equal(evaluateExpression('=Year("not-a-date")', {}), null);
  assert.equal(evaluateExpression('=Month("")', {}), null);
});

test('Weekday defaults to Sunday=1', () => {
  // 2026-07-19 is a Sunday, 2026-07-17 is a Friday.
  assert.equal(evaluateExpression('=Weekday("2026-07-19")', {}), 1);
  assert.equal(evaluateExpression('=Weekday("2026-07-17")', {}), 6);
  // firstDayOfWeek = 2 (Monday) -> Sunday becomes 7.
  assert.equal(evaluateExpression('=Weekday("2026-07-19", 2)', {}), 7);
});

test('MonthName and WeekdayName produce English names', () => {
  assert.equal(evaluateExpression('=MonthName(7)', {}), 'July');
  assert.equal(evaluateExpression('=MonthName(7, True)', {}), 'Jul');
  assert.equal(evaluateExpression('=MonthName(13)', {}), '');
  assert.equal(evaluateExpression('=WeekdayName(1)', {}), 'Sunday');
  assert.equal(evaluateExpression('=WeekdayName(7, True)', {}), 'Sat');
});

test('DateSerial constructs UTC midnight and rolls over', () => {
  assert.equal(evaluateExpression('=Year(DateSerial(2026, 7, 17))', {}), 2026);
  assert.equal(evaluateExpression('=Month(DateSerial(2026, 7, 17))', {}), 7);
  assert.equal(evaluateExpression('=Hour(DateSerial(2026, 7, 17))', {}), 0);
  // Month 13 rolls into the next year -> January 2027.
  assert.equal(evaluateExpression('=Year(DateSerial(2026, 13, 1))', {}), 2027);
  assert.equal(evaluateExpression('=Month(DateSerial(2026, 13, 1))', {}), 1);
});

test('DateValue keeps date part; TimeValue keeps time part', () => {
  assert.equal(evaluateExpression('=Hour(DateValue("2026-07-17T13:45:30Z"))', {}), 0);
  assert.equal(evaluateExpression('=Day(DateValue("2026-07-17T13:45:30Z"))', {}), 17);
  assert.equal(evaluateExpression('=Hour(TimeValue("2026-07-17T13:45:30Z"))', {}), 13);
  assert.equal(evaluateExpression('=Minute(TimeValue("2026-07-17T13:45:30Z"))', {}), 45);
  assert.equal(evaluateExpression('=Second(TimeValue("2026-07-17T13:45:30Z"))', {}), 30);
});

test('Now/Today use ExecutionTime global deterministically', () => {
  const context = { globals: { ExecutionTime: new Date('2030-01-01T00:00:00Z') } };
  assert.equal(evaluateExpression('=Year(Now())', context), 2030);
  assert.equal(evaluateExpression('=Year(Today())', context), 2030);
  assert.equal(evaluateExpression('=Hour(Today())', context), 0);
});

test('function names are case-insensitive', () => {
  assert.equal(evaluateExpression('=year("2026-07-17")', {}), 2026);
  assert.equal(evaluateExpression('=MONTHNAME(7)', {}), 'July');
});
