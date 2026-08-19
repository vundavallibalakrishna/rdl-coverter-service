const TRACE_VERSION = 1;

function roundPoint(value) {
  return Math.round(Number(value || 0) * 1000) / 1000;
}

export function createLayoutTrace(model, request) {
  return {
    version: TRACE_VERSION,
    reportName: request.outputFileName || model.name,
    definitionSha256: model.identity?.definitionSha256 || null,
    coordinatePrecisionPt: 0.25,
    page: {
      width: roundPoint(model.page.width),
      height: roundPoint(model.page.height),
      marginTop: roundPoint(model.page.marginTop),
      marginRight: roundPoint(model.page.marginRight),
      marginBottom: roundPoint(model.page.marginBottom),
      marginLeft: roundPoint(model.page.marginLeft),
      headerHeight: roundPoint(model.page.header?.height || 0),
      footerHeight: roundPoint(model.page.footer?.height || 0),
    },
    pages: [],
  };
}

export function attachLayoutTrace(doc, trace) {
  Object.defineProperty(doc, '_rdlLayoutTrace', {
    configurable: true,
    enumerable: false,
    writable: true,
    value: trace || null,
  });
  doc._rdlLayoutTracePage = -1;
}

export function beginLayoutTracePage(doc, geometry) {
  if (!doc._rdlLayoutTrace) return;
  const page = {
    number: doc._rdlLayoutTrace.pages.length + 1,
    width: roundPoint(geometry.width),
    height: roundPoint(geometry.height),
    bodyTop: roundPoint(geometry.bodyTop),
    bodyBottom: roundPoint(geometry.bodyBottom),
    regions: {
      header: {
        x: roundPoint(geometry.marginLeft),
        y: roundPoint(geometry.marginTop),
        width: roundPoint(geometry.width - geometry.marginLeft - geometry.marginRight),
        height: roundPoint(geometry.headerHeight),
      },
      body: {
        x: roundPoint(geometry.marginLeft),
        y: roundPoint(geometry.bodyTop),
        width: roundPoint(geometry.width - geometry.marginLeft - geometry.marginRight),
        height: roundPoint(geometry.bodyBottom - geometry.bodyTop),
      },
      footer: {
        x: roundPoint(geometry.marginLeft),
        y: roundPoint(geometry.bodyBottom),
        width: roundPoint(geometry.width - geometry.marginLeft - geometry.marginRight),
        height: roundPoint(geometry.footerHeight),
      },
    },
    items: [],
  };
  doc._rdlLayoutTrace.pages.push(page);
  doc._rdlLayoutTracePage = doc._rdlLayoutTrace.pages.length - 1;
}

export function selectLayoutTracePage(doc, index) {
  if (!doc._rdlLayoutTrace) return;
  doc._rdlLayoutTracePage = index;
}

export function recordLayoutItem(doc, item) {
  if (!doc._rdlLayoutTrace || doc._rdlLayoutTracePage < 0) return;
  const page = doc._rdlLayoutTrace.pages[doc._rdlLayoutTracePage];
  if (!page) return;
  const itemTop = Number(item.y || 0);
  const itemBottom = itemTop + Number(item.height || 0);
  const region = item.region || (
    itemBottom <= page.regions.header.y + page.regions.header.height + 0.125
      ? 'header'
      : itemTop >= page.regions.footer.y - 0.125
        ? 'footer'
        : 'body'
  );
  const record = {
    ...item,
    region,
    x: roundPoint(item.x),
    y: roundPoint(item.y),
    width: roundPoint(item.width),
    height: roundPoint(item.height),
  };
  page.items.push(record);
  // Returned so callers that only know a flow container's final height after laying out its children
  // (e.g. a Rectangle whose Tablix grows) can correct the recorded geometry in place.
  return record;
}

function uniquePoints(values) {
  const sorted = values.map(roundPoint).sort((left, right) => left - right);
  return sorted.filter((value, index) => index === 0 || Math.abs(value - sorted[index - 1]) > 0.125);
}

function cellsTouch(left, right) {
  const horizontalGap = Math.max(
    left.x - (right.x + right.width),
    right.x - (left.x + left.width),
    0,
  );
  const verticalGap = Math.max(
    left.y - (right.y + right.height),
    right.y - (left.y + left.height),
    0,
  );
  const xOverlap = Math.min(left.x + left.width, right.x + right.width) - Math.max(left.x, right.x);
  const yOverlap = Math.min(left.y + left.height, right.y + right.height) - Math.max(left.y, right.y);
  return (horizontalGap <= 0.125 && yOverlap >= -0.125)
    || (verticalGap <= 0.125 && xOverlap >= -0.125);
}

function connectedCellGroups(cells) {
  const remaining = new Set(cells);
  const groups = [];
  while (remaining.size > 0) {
    const first = remaining.values().next().value;
    remaining.delete(first);
    const group = [first];
    for (let cursor = 0; cursor < group.length; cursor += 1) {
      for (const candidate of [...remaining]) {
        if (cellsTouch(group[cursor], candidate)) {
          remaining.delete(candidate);
          group.push(candidate);
        }
      }
    }
    groups.push(group);
  }
  return groups;
}

function fragmentFor(page, tablixName, cells, fragmentIndex) {
  const left = Math.min(...cells.map((cell) => cell.x));
  const top = Math.min(...cells.map((cell) => cell.y));
  const right = Math.max(...cells.map((cell) => cell.x + cell.width));
  const bottom = Math.max(...cells.map((cell) => cell.y + cell.height));
  const columnEdges = uniquePoints(cells.flatMap((cell) => [cell.x, cell.x + cell.width]));
  const rowEdges = uniquePoints(cells.flatMap((cell) => [cell.y, cell.y + cell.height]));
  const borderLines = page.items.filter((item) => (
    item.kind === 'line'
    && item.traceRole === 'resolvedTablixFragmentBorder'
    && item.tablixName === tablixName
    && item.x + item.width >= left - 0.125
    && item.x <= right + 0.125
    && item.y + item.height >= top - 0.125
    && item.y <= bottom + 0.125
  ));
  return {
    id: `${tablixName || 'tablix'}:${page.number}:${fragmentIndex + 1}`,
    tablixName,
    x: roundPoint(left),
    y: roundPoint(top),
    width: roundPoint(right - left),
    height: roundPoint(bottom - top),
    columnEdges,
    columnWidths: columnEdges.slice(1).map((edge, index) => roundPoint(edge - columnEdges[index])),
    rowEdges,
    rowHeights: rowEdges.slice(1).map((edge, index) => roundPoint(edge - rowEdges[index])),
    repeatedHeaderRows: [...new Set(cells.filter((cell) => cell.repeatedHeader).map((cell) => cell.rowIndex))],
    continuation: cells.some((cell) => cell.continuation),
    cells: cells.map((cell) => ({
      itemIndex: page.items.indexOf(cell),
      itemName: cell.itemName,
      rowIndex: cell.rowIndex,
      columnIndex: cell.columnIndex,
      rowSpan: cell.rowSpan || 1,
      colSpan: cell.colSpan || 1,
      repeatedHeader: Boolean(cell.repeatedHeader),
      continuation: Boolean(cell.continuation),
      x: cell.x,
      y: cell.y,
      width: cell.width,
      height: cell.height,
      text: cell.text,
      borders: cell.borders,
      backgroundColor: cell.backgroundColor,
    })),
    resolvedFragmentBorders: borderLines.map((line) => ({
      side: line.fragmentSide,
      x: line.x,
      y: line.y,
      width: line.width,
      height: line.height,
      line: line.line,
    })),
  };
}

function tablixFragments(page) {
  const byName = new Map();
  for (const cell of page.items.filter((item) => item.kind === 'tablixCell')) {
    const key = cell.tablixName || '';
    if (!byName.has(key)) byName.set(key, []);
    byName.get(key).push(cell);
  }
  return [...byName.entries()].flatMap(([tablixName, cells]) => (
    connectedCellGroups(cells).map((group, index) => fragmentFor(page, tablixName, group, index))
  ));
}

export function finalizeLayoutTrace(trace) {
  if (!trace) return null;
  return {
    ...trace,
    pages: trace.pages.map((page) => ({
      ...page,
      tablixFragments: tablixFragments(page),
    })),
    pageCount: trace.pages.length,
  };
}

export function validateLayoutTrace(trace, expectedPageCount = null) {
  if (!trace || trace.version !== TRACE_VERSION || !Array.isArray(trace.pages)) {
    throw new Error('Canonical PDF layout trace has an unsupported or incomplete schema');
  }
  if (trace.pageCount !== trace.pages.length
    || (expectedPageCount !== null && trace.pageCount !== expectedPageCount)) {
    throw new Error('Canonical PDF layout trace page count does not match the rendered PDF');
  }
  for (const [pageIndex, page] of trace.pages.entries()) {
    if (page.number !== pageIndex + 1
      || !Number.isFinite(page.width)
      || !Number.isFinite(page.height)
      || page.width <= 0
      || page.height <= 0
      || !Array.isArray(page.items)
      || !Array.isArray(page.tablixFragments)) {
      throw new Error(`Canonical PDF layout trace page ${pageIndex + 1} is invalid`);
    }
    for (const item of page.items) {
      if (![item.x, item.y, item.width, item.height].every(Number.isFinite)
        || item.width < 0
        || item.height < 0
        || !['header', 'body', 'footer'].includes(item.region)) {
        throw new Error(`Canonical PDF layout trace item on page ${pageIndex + 1} is invalid`);
      }
      for (const line of item.lines || []) {
        if (![line.x, line.y, line.baseline, line.height, line.contentHeight].every(Number.isFinite)) {
          throw new Error(`Canonical PDF text line on page ${pageIndex + 1} is invalid`);
        }
      }
    }
    for (const fragment of page.tablixFragments) {
      if (!Array.isArray(fragment.columnWidths)
        || !Array.isArray(fragment.rowHeights)
        || !Array.isArray(fragment.cells)
        || fragment.columnWidths.some((width) => !Number.isFinite(width) || width <= 0)
        || fragment.rowHeights.some((height) => !Number.isFinite(height) || height <= 0)) {
        throw new Error(`Canonical PDF tablix fragment on page ${pageIndex + 1} is invalid`);
      }
    }
  }
  return trace;
}
