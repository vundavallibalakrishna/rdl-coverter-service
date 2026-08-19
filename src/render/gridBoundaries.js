// Shared grid-line quantization for the renderers that rebuild an RDL page as a native grid: the
// page-locked Word canvas table and the Excel worksheet. Both turn every report-item edge into a grid
// boundary, and both then have to draw each item's borders on the cells between those boundaries.
//
// RDL authors routinely give a container and the item it holds edges that differ by a fraction of a
// point, because an inch-valued container width and an inch-valued child width rarely land on the same
// quarter point. A fixed-layout renderer strokes both edges and the two strokes overlap into one visible
// rule. A grid renderer cannot: it has to materialize the gap as a real band, and a band narrower than
// the certified geometry tolerance is not drawable at that width. Word and Excel widen it to fit the two
// cell borders, so a single canonical rule becomes a visible double line and the perpendicular rule loses
// its corner join.
//
// Collapsing those boundaries into one grid line is safe because the certification contract already
// treats geometry inside 0.5 point as identical. Clustering stops as soon as a cluster's span would reach
// the tolerance, so no edge ever moves by a half point or more. A span whose two ends both anchor the
// same report item is never collapsed: that item would otherwise lose its only row or column.
export const GRID_BOUNDARY_TOLERANCE_POINTS = 0.5;

function spanKey(from, to) {
  return `${from}|${to}`;
}

export function protectedSpanSet(spans) {
  const keys = new Set();
  for (const [from, to] of spans || []) {
    if (Number.isFinite(from) && Number.isFinite(to) && to > from) keys.add(spanKey(from, to));
  }
  return keys;
}

/**
 * Collapses grid boundaries that lie closer together than the certified geometry tolerance.
 *
 * @param {Iterable<number>} values every edge coordinate the grid must be able to address
 * @param {object} [options]
 * @param {number} [options.tolerance] maximum distance between boundaries treated as one grid line
 * @param {Iterable<[number, number]>} [options.protectedSpans] item extents that must keep both ends
 * @returns {{ boundaries: number[], indexOf: (value: number) => number }}
 */
export function buildGridBoundaries(values, options = {}) {
  const tolerance = Number(options.tolerance ?? GRID_BOUNDARY_TOLERANCE_POINTS);
  const protectedSpans = protectedSpanSet(options.protectedSpans);
  const counts = new Map();
  for (const value of values) {
    if (!Number.isFinite(value)) continue;
    counts.set(value, (counts.get(value) || 0) + 1);
  }
  const distinct = [...counts.keys()].sort((left, right) => left - right);
  const clusters = [];
  for (const value of distinct) {
    const current = clusters[clusters.length - 1];
    const joins = current
      && value - current[0] < tolerance
      && !current.some((member) => protectedSpans.has(spanKey(member, value)));
    if (joins) current.push(value);
    else clusters.push([value]);
  }
  // The canvas origin and the canvas extent are structural: the grid has to start and end exactly where
  // the page does, or the rebuilt table no longer spans it. Every other cluster keeps the coordinate the
  // most edges already sit on, which leaves the largest number of items exactly where the PDF put them.
  const boundaries = clusters.map((cluster, index) => {
    if (index === 0) return cluster[0];
    if (index === clusters.length - 1) return cluster[cluster.length - 1];
    return cluster.reduce((best, value) => (counts.get(value) > counts.get(best) ? value : best), cluster[0]);
  });
  if (boundaries.length === 1 && distinct.length > 1) {
    boundaries.splice(0, 1, distinct[0], distinct[distinct.length - 1]);
  }
  const aliases = new Map();
  clusters.forEach((cluster, index) => {
    const resolved = Math.min(index, boundaries.length - 1);
    for (const value of cluster) aliases.set(value, resolved);
  });
  const indexOf = (value) => {
    const exact = aliases.get(value);
    if (exact !== undefined) return exact;
    // Header and footer sub-grids address the same boundary list with coordinates that were not part of
    // its input set. Resolve those to the nearest grid line inside the same tolerance.
    let best = -1;
    let bestDistance = Infinity;
    for (let index = 0; index < boundaries.length; index += 1) {
      const distance = Math.abs(boundaries[index] - value);
      if (distance < bestDistance) {
        bestDistance = distance;
        best = index;
      }
    }
    return bestDistance < tolerance ? best : -1;
  };
  return { boundaries, indexOf };
}
