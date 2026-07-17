import assert from 'node:assert/strict';
import test from 'node:test';
import { computeCellPlacements } from '../src/render/tableGrid.js';

test('places interlocking horizontal and vertical spans on the intended columns', () => {
  const rows = [
    { cells: Array.from({ length: 4 }, () => ({ colSpan: 2, rowSpan: 1 })) },
    {
      cells: [
        { colSpan: 1, rowSpan: 2 },
        { colSpan: 1, rowSpan: 2 },
        { colSpan: 1, rowSpan: 2 },
        { colSpan: 1, rowSpan: 2 },
        { colSpan: 2, rowSpan: 1 },
        { colSpan: 1, rowSpan: 2 },
        { colSpan: 1, rowSpan: 2 },
      ],
    },
    { cells: [{ colSpan: 1, rowSpan: 1 }, { colSpan: 1, rowSpan: 1 }] },
  ];

  assert.deepEqual(computeCellPlacements(rows, 8), [
    [0, 2, 4, 6],
    [0, 1, 2, 3, 4, 6, 7],
    [4, 5],
  ]);
});
