import assert from 'node:assert/strict';
import test from 'node:test';
import { normalizeDisplayText } from '../src/rdl/text.js';

const TAB = String.fromCharCode(9);
const NUL = String.fromCharCode(0);
const VERTICAL_TAB = String.fromCharCode(11);
const CR = String.fromCharCode(13);
const LF = String.fromCharCode(10);

test('normalizeDisplayText expands tabs so bullet text does not render as a missing-glyph box', () => {
  assert.equal(normalizeDisplayText(`•${TAB}Scope of work`), '• Scope of work');
});

test('normalizeDisplayText strips non-glyph C0 control characters but keeps line breaks', () => {
  assert.equal(normalizeDisplayText(`a${NUL}b${VERTICAL_TAB}c`), 'abc');
  assert.equal(normalizeDisplayText(`line1${LF}line2`), 'line1\nline2');
  assert.equal(normalizeDisplayText(`a${CR}${LF}b`), 'a\nb');
});

test('normalizeDisplayText passes non-string values through untouched', () => {
  assert.equal(normalizeDisplayText(42), 42);
  assert.equal(normalizeDisplayText(null), null);
  assert.equal(normalizeDisplayText(undefined), undefined);
});
