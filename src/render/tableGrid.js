export function computeCellPlacements(rows, columnCount) {
  const occupied = Array.from({ length: columnCount }, () => 0);
  return rows.map((row) => {
    const placements = [];
    let columnIndex = 0;
    for (const cell of row.cells || []) {
      while (columnIndex < columnCount && occupied[columnIndex] > 0) columnIndex += 1;
      const colSpan = Math.max(1, cell.colSpan || 1);
      placements.push(Math.min(columnIndex, Math.max(0, columnCount - 1)));
      if ((cell.rowSpan || 1) > 1) {
        for (let offset = 0; offset < colSpan && columnIndex + offset < columnCount; offset += 1) {
          occupied[columnIndex + offset] = Math.max(occupied[columnIndex + offset], cell.rowSpan || 1);
        }
      }
      columnIndex += colSpan;
    }
    for (let index = 0; index < occupied.length; index += 1) occupied[index] = Math.max(0, occupied[index] - 1);
    return placements;
  });
}
