import assert from 'node:assert/strict';
import test from 'node:test';
import { cellString, excelNumberFormat } from '../src/render/excelFormat.js';

test('standard .NET numeric specifiers translate to Excel format codes', () => {
  assert.equal(excelNumberFormat('C2'), '$#,##0.00');
  assert.equal(excelNumberFormat('C'), '$#,##0.00');
  assert.equal(excelNumberFormat('N0'), '#,##0');
  assert.equal(excelNumberFormat('N2'), '#,##0.00');
  assert.equal(excelNumberFormat('F2'), '0.00');
  assert.equal(excelNumberFormat('P1'), '0.0%');
  assert.equal(excelNumberFormat('D'), '0');
});

test('Excel-compatible custom numeric patterns pass through', () => {
  assert.equal(excelNumberFormat('#,##0.00'), '#,##0.00');
  assert.equal(excelNumberFormat('0.0%'), '0.0%');
  assert.equal(excelNumberFormat('$#,##0'), '$#,##0');
});

test('unrecognized, empty, expression, and date formats return null (write as text/General)', () => {
  assert.equal(excelNumberFormat(null), null);
  assert.equal(excelNumberFormat(''), null);
  assert.equal(excelNumberFormat('=Fields!X.Value'), null);
  assert.equal(excelNumberFormat('yyyy-MM-dd'), null);
  assert.equal(excelNumberFormat('General'), null);
});

test('cellString does NOT prefix formula-like values (XLSX typed strings are already safe)', () => {
  // Unlike CSV, an XLSX string cell is typed and never evaluated, so the CSV-era apostrophe guard would
  // only corrupt legitimate data. These must round-trip verbatim.
  assert.equal(cellString('=1+1'), '=1+1');
  assert.equal(cellString('-N/A'), '-N/A');
  assert.equal(cellString('@owner'), '@owner');
  assert.equal(cellString('+27 21 555'), '+27 21 555');
  assert.equal(cellString(null), '');
  assert.equal(cellString(42), '42');
});
