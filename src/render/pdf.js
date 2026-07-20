import PDFDocument from 'pdfkit';
import { PDFDocument as PdfLibDocument } from 'pdf-lib';
import { color, isHidden, normalizeDatasets, styleColor, styleSize, styleValue, styledTextForItem, tablixRows, textForItem, cellText, cellTextbox } from './common.js';
import { pdfFont } from './fonts.js';
import { computeCellPlacements } from './tableGrid.js';
import { cellGeometryPt, resolveGridColumns } from './tableLayout.js';
import { materializeChart } from './chartData.js';
import { drawChart } from './chart.js';

// Minimum border stroke width (points), set from config at the start of each render. A floor lets thin
// 1pt hairlines render at a crisp, uniform weight instead of rounding unevenly across screen zoom levels.
let borderWidthFloor = 0;

function collectDocument(doc) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    doc.on('data', (chunk) => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);
  });
}

function applyFont(doc, config, style, context = {}) {
  const bold = /bold|600|700|800|900/i.test(String(styleValue(style.fontWeight, context, 'Normal')));
  const italic = /italic/i.test(String(styleValue(style.fontStyle, context, 'Normal')));
  doc.font(pdfFont(config, styleValue(style.fontFamily, context, 'Arial'), bold, italic)).fontSize(styleSize(style.fontSize, context, 10) || 10).fillColor(styleColor(style.color, context));
}

function drawBorderEdge(doc, x, y, width, height, side, border, context = {}) {
  if (!border) return;
  const borderStyle = String(styleValue(border.style, context, 'None'));
  const borderColor = styleColor(border.color, context, null);
  const borderWidth = styleSize(border.width, context, 1);
  // A conditional width of 0 (e.g. =IIF(rn=1,"1pt","0pt")) means the border is intentionally absent.
  if (/^none$/i.test(borderStyle) || !borderColor || borderWidth <= 0) return;
  const segments = {
    top: [x, y, x + width, y], right: [x + width, y, x + width, y + height],
    bottom: [x, y + height, x + width, y + height], left: [x, y, x, y + height],
  };
  const [x1, y1, x2, y2] = segments[side];
  doc.save().lineWidth(Math.max(borderWidth, borderWidthFloor)).strokeColor(borderColor);
  if (/dash/i.test(borderStyle)) doc.dash(Math.max(2, borderWidth * 3));
  else if (/dot/i.test(borderStyle)) doc.dash(Math.max(1, borderWidth), { space: Math.max(1, borderWidth * 2) });
  // Solid edges use a projecting (square) cap so each per-side segment overlaps its neighbours by half the
  // line width, closing the hairline notches/nicks that a butt cap leaves at cell corners, T-junctions and
  // fragment seams (visible only at high zoom). Dashed/dotted keep the butt cap so gaps stay open.
  else doc.lineCap('square').lineJoin('miter');
  doc.moveTo(x1, y1).lineTo(x2, y2).stroke().restore();
}

function drawBorder(doc, x, y, width, height, style, context = {}) {
  const sides = style?.borders || (style?.border ? { top: style.border, right: style.border, bottom: style.border, left: style.border } : null);
  if (!sides) return;
  for (const [side, border] of Object.entries(sides)) drawBorderEdge(doc, x, y, width, height, side, border, context);
}

// Height of `text` as doc.text() will actually draw it, for vertical alignment. A source line with no
// spaces cannot be broken, so doc.text() always draws it on ONE line — but heightOfString applies a
// conservative wrap margin and can report an extra line for a long token (e.g. a 9-digit number in a
// narrow numeric cell), which pushed vertically-centred values up out of line with their neighbours.
// Lines that do contain spaces wrap normally and are measured as before.
function renderedTextHeight(doc, text, width) {
  const lines = String(text ?? '').split('\n');
  // heightOfString ignores a single trailing empty line; mirror that so existing layouts are unchanged.
  if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop();
  const lineHeight = doc.currentLineHeight(true);
  return lines.reduce((total, line) => total
    + (line.includes(' ') ? doc.heightOfString(line, { width, lineGap: 0 }) : lineHeight), 0);
}

function styledSegmentsForText(item, context, requestedText) {
  const paragraphs = styledTextForItem(item, context);
  if (!paragraphs) return null;
  const segments = [];
  let fullText = '';
  for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
    for (const run of paragraph.runs) {
      const text = String(run.text ?? '');
      segments.push({ text, style: run.style || item.style, paragraphStyle: paragraph.style || item.style, paragraphIndex });
      fullText += text;
    }
    if (paragraphIndex < paragraphs.length - 1) {
      segments.push({
        text: '\n',
        style: paragraph.style || item.style,
        paragraphStyle: paragraph.style || item.style,
        paragraphIndex,
      });
      fullText += '\n';
    }
  }

  const target = requestedText === undefined ? fullText : String(requestedText ?? '');
  if (target.length === 0) return { text: target, segments: [] };
  let start = 0;
  if (target !== fullText) {
    if (fullText.startsWith(target)) start = 0;
    else if (fullText.endsWith(target)) start = fullText.length - target.length;
    else start = fullText.indexOf(target);
    // A materialized cell can combine several conditional textboxes. If its override is not a contiguous
    // slice of this textbox, retain the existing plain-text path rather than assigning misleading styles.
    if (start < 0) return null;
  }
  const end = start + target.length;
  let offset = 0;
  const sliced = [];
  for (const segment of segments) {
    const segmentStart = offset;
    const segmentEnd = offset + segment.text.length;
    offset = segmentEnd;
    const from = Math.max(start, segmentStart);
    const to = Math.min(end, segmentEnd);
    if (to <= from) continue;
    sliced.push({ ...segment, text: segment.text.slice(from - segmentStart, to - segmentStart) });
  }
  return { text: target, segments: sliced };
}

function lineHeightForStyle(doc, config, style, context) {
  applyFont(doc, config, style, context);
  return doc.currentLineHeight(true);
}

function layoutStyledText(doc, config, item, context, text, width) {
  const source = styledSegmentsForText(item, context, text);
  if (!source) return null;
  if (source.text.length === 0) return { lines: [], height: 0 };
  const lines = [];
  let line = null;
  const startLine = (style, paragraphStyle) => ({
    runs: [],
    width: 0,
    height: lineHeightForStyle(doc, config, style || item.style, context),
    paragraphStyle: paragraphStyle || item.style,
    paragraphEnd: false,
  });
  const finishLine = (paragraphEnd = false) => {
    if (!line) line = startLine(item.style, item.style);
    line.paragraphEnd = paragraphEnd;
    lines.push(line);
    line = null;
  };
  const addToken = (token, segment) => {
    if (!line) line = startLine(segment.style, segment.paragraphStyle);
    const whitespace = /^\s+$/.test(token);
    applyFont(doc, config, segment.style, context);
    const tokenWidth = doc.widthOfString(token);
    const tokenHeight = doc.currentLineHeight(true);
    if (line.runs.length > 0 && line.width + tokenWidth > width) {
      finishLine(false);
      line = startLine(segment.style, segment.paragraphStyle);
      // Wrapping consumes the separating whitespace, matching PDFKit's normal word-wrap behaviour.
      if (whitespace) return;
    }
    line.runs.push({ text: token, style: segment.style, width: tokenWidth });
    line.width += tokenWidth;
    line.height = Math.max(line.height, tokenHeight);
  };

  for (const segment of source.segments) {
    const parts = segment.text.split(/(\n)/);
    for (const part of parts) {
      if (part === '') continue;
      if (part === '\n') {
        if (!line) line = startLine(segment.style, segment.paragraphStyle);
        finishLine(true);
        continue;
      }
      const tokens = part.match(/[^\S\n]+|[^\s\n]+/g) || [];
      for (const token of tokens) addToken(token, segment);
    }
  }
  // PDFKit ignores one trailing empty line. Preserve empty lines between paragraphs, but not a terminal one.
  if (line?.runs.length || lines.length === 0) finishLine(true);
  return { lines, height: lines.reduce((sum, current) => sum + current.height, 0) };
}

function drawStyledText(doc, config, item, context, layout, x, y, width) {
  let lineY = y;
  for (const line of layout.lines) {
    const alignment = String(styleValue(line.paragraphStyle?.textAlign ?? item.style?.textAlign, context, 'left')).toLowerCase();
    let lineX = x;
    if (alignment === 'center') lineX += Math.max(0, (width - line.width) / 2);
    else if (alignment === 'right') lineX += Math.max(0, width - line.width);
    const spaces = line.runs.reduce((count, run) => count + ((run.text.match(/ /g) || []).length), 0);
    const justifyExtra = alignment === 'justify' && !line.paragraphEnd && spaces > 0
      ? Math.max(0, width - line.width) / spaces
      : 0;
    for (const run of line.runs) {
      applyFont(doc, config, run.style, context);
      doc.text(run.text, lineX, lineY, {
        lineBreak: false,
        underline: /underline/i.test(String(styleValue(run.style?.textDecoration, context, 'None'))),
        strike: /line.?through/i.test(String(styleValue(run.style?.textDecoration, context, 'None'))),
      });
      lineX += run.width + justifyExtra * ((run.text.match(/ /g) || []).length);
    }
    lineY += line.height;
  }
}

function drawTextbox(doc, config, item, x, y, context, override = {}) {
  if (isHidden(item.hidden, context)) return;
  const style = item.style;
  const width = override.width ?? item.width;
  const height = override.height ?? item.height;
  const backgroundColor = styleColor(style.backgroundColor, context, null);
  if (backgroundColor) doc.save().fillColor(backgroundColor).rect(x, y, width, height).fill().restore();
  // Tablix cells resolve their borders against neighbouring cells (SSRS shared-edge model) and draw
  // them separately, so they ask drawTextbox to skip its own per-cell border.
  if (!override.skipBorder) drawBorder(doc, x, y, width, height, style, context);
  const text = override.text ?? textForItem(item, context);
  // Padding is an RdlSize ExpressionType — resolve per row via styleSize (a literal passes straight through).
  const padTop = styleSize(style.paddingTop, context, 2);
  const padRight = styleSize(style.paddingRight, context, 2);
  const padBottom = styleSize(style.paddingBottom, context, 2);
  const paddingLeft = styleSize(style.paddingLeft, context, 2) + (override.padLeft || 0);
  const innerWidth = Math.max(1, width - paddingLeft - padRight);
  const innerHeight = Math.max(1, height - padTop - padBottom);
  const styledLayout = layoutStyledText(doc, config, item, context, text, innerWidth);
  if (!styledLayout) applyFont(doc, config, style, context);
  const measuredHeight = styledLayout?.height ?? renderedTextHeight(doc, text, innerWidth);
  // VerticalAlign can be an expression; resolve it before matching, otherwise the regex tests the raw
  // expression source (e.g. `=IIF(c,"Middle","Top")` always contains "Middle").
  const verticalAlign = String(styleValue(style.verticalAlign, context, 'top')).toLowerCase();
  let textY = y + padTop;
  if (/middle|center/.test(verticalAlign)) textY = y + Math.max(padTop, (height - measuredHeight) / 2);
  if (/bottom/.test(verticalAlign)) textY = y + Math.max(padTop, height - measuredHeight - padBottom);
  // Clip text to the cell box so it can never bleed into an adjacent cell or, when the cell is
  // clamped at the page/footer boundary, into the reserved footer band. Background and borders are
  // drawn above (unclipped) so their edges stay crisp.
  doc.save().rect(x, y, width, height).clip();
  if (styledLayout) drawStyledText(doc, config, item, context, styledLayout, x + paddingLeft, textY, innerWidth);
  else {
    doc.text(text, x + paddingLeft, textY, {
      width: innerWidth,
      height: innerHeight,
      align: String(styleValue(style.textAlign, context, 'left')).toLowerCase(),
      lineBreak: true,
      ellipsis: !item.canGrow,
      underline: /underline/i.test(String(styleValue(style.textDecoration, context, 'None'))),
      strike: /line.?through/i.test(String(styleValue(style.textDecoration, context, 'None'))),
    });
  }
  doc.restore();
}

function drawImage(doc, model, item, x, y, context = {}) {
  if (item.source !== 'Embedded') return;
  // Image Value and Sizing can be expressions (e.g. Value=`=Fields!Logo.Value`); resolve before use, or
  // the raw expression string misses the embeddedImages map and the image is silently dropped.
  const image = model.embeddedImages[styleValue(item.value, context, item.value)];
  if (!image?.data) return;
  const data = Buffer.from(image.data.replace(/\s+/g, ''), 'base64');
  // Honour the RDL Image Sizing. FitProportional (the RDL default) scales to fit the box while keeping
  // aspect; Fit stretches to fill the box exactly (SSRS behaviour — the box, not the source, wins);
  // Clip draws at native size clipped to the box; AutoSize draws at native size.
  const sizing = String(styleValue(item.sizing, context, 'FitProportional') || 'FitProportional');
  if (/^Fit$/i.test(sizing)) {
    doc.image(data, x, y, { width: item.width, height: item.height });
  } else if (/^Clip$/i.test(sizing)) {
    doc.save().rect(x, y, item.width, item.height).clip();
    doc.image(data, x, y);
    doc.restore();
  } else if (/^AutoSize$/i.test(sizing)) {
    doc.image(data, x, y);
  } else {
    doc.image(data, x, y, { fit: [item.width, item.height], align: 'center', valign: 'center' });
  }
}

function drawSimpleItem(doc, config, model, item, x, y, context) {
  if (isHidden(item.hidden, context)) return;
  if (item.type === 'Textbox') drawTextbox(doc, config, item, x, y, context);
  else if (item.type === 'Chart') {
    const data = materializeChart(item, context.datasets || {}, context.parameters || {}, context.globals || {});
    drawChart(doc, config, item, data, x, y, item.width, item.height, context);
  } else if (item.type === 'Image') drawImage(doc, model, item, x, y, context);
  else if (item.type === 'Line') doc.save().lineWidth(styleSize(item.style?.border?.width, context, 1) || 1).strokeColor(styleColor(item.style?.border?.color, context, '#000000')).moveTo(x, y).lineTo(x + item.width, y + item.height).stroke().restore();
  else if (item.type === 'Rectangle') {
    const backgroundColor = styleColor(item.style.backgroundColor, context, null);
    if (backgroundColor) doc.save().fillColor(backgroundColor).rect(x, y, item.width, item.height).fill().restore();
    drawBorder(doc, x, y, item.width, item.height, item.style, context);
    for (const child of [...item.items].sort((left, right) => left.zIndex - right.zIndex || left.top - right.top || left.left - right.left)) {
      drawSimpleItem(doc, config, model, child, x + child.left, y + child.top, context);
    }
  }
}

function textHeight(doc, config, textbox, context, text, width) {
  if (!textbox || !text) return 0;
  const innerWidth = Math.max(1, width - styleSize(textbox.style.paddingLeft, context, 2) - styleSize(textbox.style.paddingRight, context, 2));
  const styledLayout = layoutStyledText(doc, config, textbox, context, text, innerWidth);
  if (styledLayout) return styledLayout.height;
  applyFont(doc, config, textbox.style, context);
  return doc.heightOfString(text, { width: innerWidth, lineGap: 0 });
}

function splitTextForHeight(doc, config, textbox, context, text, width, height) {
  const value = String(text || '');
  if (!value || !textbox) return { head: value, tail: '' };
  const available = Math.max(1, height - styleSize(textbox.style.paddingTop, context, 2) - styleSize(textbox.style.paddingBottom, context, 2));
  if (textHeight(doc, config, textbox, context, value, width) <= available) return { head: value, tail: '' };
  let low = 1;
  let high = value.length;
  let best = 0;
  while (low <= high) {
    const middle = Math.floor((low + high) / 2);
    if (textHeight(doc, config, textbox, context, value.slice(0, middle), width) <= available) {
      best = middle;
      low = middle + 1;
    } else high = middle - 1;
  }
  if (best === 0) best = 1;
  const newline = value.lastIndexOf('\n', best - 1);
  const space = value.lastIndexOf(' ', best - 1);
  const boundary = Math.max(newline, space);
  const splitAt = boundary > 0 ? boundary : best;
  return {
    head: value.slice(0, splitAt).trimEnd(),
    tail: value.slice(splitAt + (boundary >= 0 ? 1 : 0)).trimStart(),
  };
}

function renderTablix({ doc, config, model, item, request, startX, startY, pageBottom, addPage, globals }) {
  const { rows, columns } = tablixRows(item, request, globals, model);
  const datasets = normalizeDatasets(model, request);
  // A matrix expands to a data-dependent column grid wider than the design width; use its natural
  // total so columns are not scaled down. Static-column tablixes keep item's declared width scaling.
  const layoutItem = item.hasColumnGroups ? { ...item, columns, width: columns.reduce((sum, width) => sum + width, 0) } : item;
  const { columnsPt: columnWidths, totalPt: totalWidth } = resolveGridColumns(layoutItem);
  const headers = rows.filter((row) => row.isHeader);
  const placements = computeCellPlacements(rows, columnWidths.length);
  const rowIndexes = new Map(rows.map((row, index) => [row, index]));
  const outerContext = { parameters: request.parameters || {}, globals, dataset: datasets[item.datasetName] || [], datasets, fields: {} };

  // Grid occupancy map: which cell (and its owning row) covers each grid position, so a cell can find
  // the neighbour on each side. Populated from the placements plus col/row spans.
  const gridOwners = rows.map(() => new Array(columnWidths.length).fill(null));
  rows.forEach((row, rowIndex) => {
    row.cells.forEach((cell, index) => {
      const start = placements[rowIndex][index];
      for (let r = 0; r < (cell.rowSpan || 1) && rowIndex + r < rows.length; r += 1) {
        for (let c = 0; c < (cell.colSpan || 1) && start + c < columnWidths.length; c += 1) {
          gridOwners[rowIndex + r][start + c] = { cell, rowIndex };
        }
      }
    });
  });
  const contextForRow = (rowIndex) => ({ fields: rows[rowIndex]?.fields || {}, parameters: request.parameters || {}, globals, dataset: datasets[item.datasetName] || [], datasets });
  // A side's border for a cell, evaluated in that cell's row context. Returns null when the side is
  // absent or resolves to None so the caller can fall back to the neighbouring cell's opposite side.
  const resolveSide = (owner, side) => {
    if (!owner) return null;
    const cellStyle = owner.cell.containerWrapped ? item.style : (cellTextbox(owner.cell)?.style || item.style);
    const border = cellStyle?.borders?.[side];
    if (!border) return null;
    const context = contextForRow(owner.rowIndex);
    if (/^none$/i.test(String(styleValue(border.style, context, 'None')))) return null;
    return { border, context };
  };
  // Shared-edge resolution: each edge is drawn from this cell's border, or the adjacent cell's border
  // on the shared edge when this cell declares None. Blank cells therefore inherit their neighbours'
  // grid lines instead of leaving gaps.
  const resolveEdges = (owner, start, span) => {
    const { rowIndex } = owner;
    const rowSpan = owner.cell.rowSpan || 1;
    const above = rowIndex > 0 ? gridOwners[rowIndex - 1][start] : null;
    const below = rowIndex + rowSpan < rows.length ? gridOwners[rowIndex + rowSpan][start] : null;
    const left = start > 0 ? gridOwners[rowIndex][start - 1] : null;
    const right = start + span < columnWidths.length ? gridOwners[rowIndex][start + span] : null;
    return {
      top: resolveSide(owner, 'top') || resolveSide(above, 'bottom'),
      bottom: resolveSide(owner, 'bottom') || resolveSide(below, 'top'),
      left: resolveSide(owner, 'left') || resolveSide(left, 'right'),
      right: resolveSide(owner, 'right') || resolveSide(right, 'left'),
    };
  };
  // A grid line is shared by many cells (adjacent cells, and a merged cell's full-length edge vs its
  // neighbours' per-row/per-column edges). Stroking it once per owner paints the same anti-aliased line
  // several times, so it looks ~1px heavier — or, worse, a merged cell's clean full-length edge gets built
  // from choppy per-row pieces. Instead, COLLECT every edge, then at each page boundary MERGE the collinear
  // pieces of each identical line into maximal runs and stroke each run exactly once — the way SSRS draws.
  const posKey = (v) => Math.round(v * 4); // 0.25pt precision so coincident edges match
  let pendingEdges = [];
  const collectEdge = (x, y, width, height, side, border, context) => {
    if (!border) return;
    const vertical = side === 'left' || side === 'right';
    const pos = vertical ? (side === 'right' ? x + width : x) : (side === 'bottom' ? y + height : y);
    const [a, b] = vertical ? [y, y + height] : [x, x + width];
    const sig = `${styleValue(border.style, context, 'None')}|${styleColor(border.color, context, null)}|${styleSize(border.width, context, 1)}`;
    pendingEdges.push({ orient: vertical ? 'V' : 'H', pos, a, b, border, context, sig });
  };
  const flushEdges = () => {
    const groups = new Map();
    for (const edge of pendingEdges) {
      const key = `${edge.orient}|${posKey(edge.pos)}|${edge.sig}`;
      if (!groups.has(key)) groups.set(key, { edge, intervals: [] });
      groups.get(key).intervals.push([edge.a, edge.b]);
    }
    for (const { edge, intervals } of groups.values()) {
      intervals.sort((p, q) => p[0] - q[0]);
      let [runStart, runEnd] = intervals[0];
      const runs = [];
      for (let i = 1; i < intervals.length; i += 1) {
        const [s, e] = intervals[i];
        if (s <= runEnd + 0.75) runEnd = Math.max(runEnd, e); // collinear & touching → extend the run
        else { runs.push([runStart, runEnd]); [runStart, runEnd] = [s, e]; }
      }
      runs.push([runStart, runEnd]);
      for (const [s, e] of runs) {
        if (edge.orient === 'V') drawBorderEdge(doc, edge.pos, s, 0, e - s, 'left', edge.border, edge.context);
        else drawBorderEdge(doc, s, edge.pos, e - s, 0, 'top', edge.border, edge.context);
      }
    }
    pendingEdges = [];
  };
  const drawEdges = (x, y, width, height, edges) => {
    for (const side of ['top', 'right', 'bottom', 'left']) {
      if (edges[side]) collectEdge(x, y, width, height, side, edges[side].border, edges[side].context);
    }
  };
  let y = startY;
  let fragmentStartY = startY;
  let firstFragment = true;
  let addedHeight = 0;
  // Open merged (row-span) cells. A merged cell's fill/border/value are drawn to follow the ACTUAL extent
  // of its spanned rows — not a precomputed height — because a spanned row can grow when it splits across a
  // page. Each span stays open from its first row until its last spanned row is fully drawn, closing a
  // segment (and re-drawing the value, SSRS-style) at every page break in between.
  let openSpans = [];

  const closeOuterBorderFragment = (endY) => {
    const fragmentHeight = Math.max(0, endY - fragmentStartY);
    if (firstFragment) collectEdge(startX, fragmentStartY, totalWidth, fragmentHeight, 'top', item.style?.borders?.top, outerContext);
    collectEdge(startX, fragmentStartY, totalWidth, fragmentHeight, 'left', item.style?.borders?.left, outerContext);
    collectEdge(startX, fragmentStartY, totalWidth, fragmentHeight, 'right', item.style?.borders?.right, outerContext);
    collectEdge(startX, fragmentStartY, totalWidth, fragmentHeight, 'bottom', item.style?.borders?.bottom, outerContext);
    firstFragment = false;
    flushEdges(); // draw this page fragment's borders as merged, single strokes
  };

  const layoutsForRow = (row, texts = row.cells.map((cell) => cellText(cell))) => {
    const rowIndex = rowIndexes.get(row);
    return row.cells.map((cell, index) => {
      const textbox = cellTextbox(cell);
      const span = cell.colSpan || 1;
      const columnIndex = placements[rowIndex][index];
      const { widthPt: width } = cellGeometryPt(columnWidths, columnIndex, span);
      const context = { fields: row.fields, parameters: request.parameters || {}, globals, dataset: datasets[item.datasetName] || [], datasets };
      return { cell, textbox, width, columnIndex, context, text: texts[index] || '' };
    });
  };

  const measureRow = (row, texts) => layoutsForRow(row, texts).reduce((height, layout) => Math.max(
    height,
    textHeight(doc, config, layout.textbox, layout.context, layout.text, layout.width)
      + styleSize(layout.textbox?.style.paddingTop, layout.context, 2) + styleSize(layout.textbox?.style.paddingBottom, layout.context, 2),
  ), row.height);
  const measuredHeights = rows.map((row) => measureRow(row));

  // Draw one segment of an open merged cell: its fill + value + borders from segStartY down to endY,
  // clamped to the reserved body area so it never bleeds into the footer band. When the value is taller
  // than the segment, the overflow is recorded in `pendingTail` so it continues on the next page instead
  // of being clipped; a value that fits leaves no tail and therefore repeats at the top of each page.
  const drawSpanSegment = (span, endY) => {
    const segmentHeight = Math.min(endY - span.segStartY, Math.max(0, pageBottom - span.segStartY));
    if (segmentHeight <= 0.5) return;
    span.pendingTail = null;
    if (span.textbox && !span.cell.hidden) {
      const { head, tail } = splitTextForHeight(doc, config, span.textbox, span.context, span.text, span.width, segmentHeight);
      if (tail && tail.length > 0) span.pendingTail = tail;
      drawTextbox(doc, config, span.textbox, span.x, span.segStartY, span.context, { width: span.width, height: segmentHeight, text: head, skipBorder: true });
    }
    drawEdges(span.x, span.segStartY, span.width, segmentHeight, span.edges);
  };

  // Close every span whose last spanned row has now been fully drawn, ending it at the current y.
  const closeSpansEndingAt = (rowIndex) => {
    for (const span of openSpans) if (span.endRowIndex <= rowIndex) drawSpanSegment(span, y);
    openSpans = openSpans.filter((span) => span.endRowIndex > rowIndex);
  };

  const drawRowContent = (row, height, texts = row.cells.map((cell) => cellText(cell)), rowComplete = true) => {
    const rowIndex = rowIndexes.get(row);
    for (const [index, cell] of row.cells.entries()) {
      const span = cell.colSpan || 1;
      const columnIndex = placements[rowIndex][index];
      const { xOffsetPt, widthPt: width } = cellGeometryPt(columnWidths, columnIndex, span);
      const x = startX + xOffsetPt;
      const textbox = cellTextbox(cell);
      const cellContext = { fields: row.fields, parameters: request.parameters || {}, globals, dataset: datasets[item.datasetName] || [], datasets };
      const edges = resolveEdges({ cell, rowIndex }, columnIndex, span);
      // Merged (row-span) cells are drawn lazily via drawSpanSegment so they track the real extent of their
      // spanned rows across page splits; single-row cells draw here directly.
      if ((cell.rowSpan || 1) > 1) {
        openSpans.push({ x, width, textbox, cell, text: texts[index] || '', context: cellContext, edges, segStartY: y, endRowIndex: rowIndex + (cell.rowSpan || 1) - 1 });
        continue;
      }
      const renderedHeight = Math.min(height, Math.max(0, pageBottom - y));
      if (renderedHeight <= 0) continue;
      // Recursive (parent/child) groups render expanded with the first cell indented by depth.
      const padLeft = index === 0 && row.indentLevel ? row.indentLevel * 12 : 0;
      if (textbox && !cell.hidden) drawTextbox(doc, config, textbox, x, y, cellContext, { width, height: renderedHeight, text: texts[index] || '', skipBorder: true, padLeft });
      drawEdges(x, y, width, renderedHeight, edges);
    }
    y += height;
    addedHeight += height;
    // Only close spans on a fully-rendered row; a split row's head must keep its spans open for the tail.
    if (rowComplete) closeSpansEndingAt(rowIndex);
  };

  const startContinuationPage = () => {
    // End each open span at this page's content bottom, break, repeat the headers, then re-open the spans
    // just below the repeated headers so their value redraws at the top of the new page.
    for (const span of openSpans) drawSpanSegment(span, y);
    closeOuterBorderFragment(y);
    addPage();
    y = addPage.bodyTop;
    fragmentStartY = y;
    for (const header of headers) drawRowContent(header, measureRow(header));
    // Continue overflowing values from where they were clipped; repeat values that fully fit.
    for (const span of openSpans) {
      if (span.pendingTail) { span.text = span.pendingTail; span.pendingTail = null; }
      span.segStartY = y;
    }
  };

  const drawRow = (row) => {
    // A row-group page break starts the group's first row on a fresh page (unless we're already at the
    // top of a page). Reuses the continuation-page machinery so repeated headers redraw.
    if (row.pageBreakBefore && y > addPage.bodyTop) startContinuationPage();
    let remainingTexts = row.cells.map((cell) => cellText(cell));
    let measured = measureRow(row, remainingTexts);
    const repeatedHeaderHeight = headers.reduce((sum, header) => sum + measureRow(header), 0);
    const freshPageCapacity = pageBottom - addPage.bodyTop - repeatedHeaderHeight;
    const rowIndex = rowIndexes.get(row);
    const protectedHeight = row.cells.reduce((maximum, cell) => Math.max(
      maximum,
      measuredHeights.slice(rowIndex, rowIndex + Math.max(1, cell.rowSpan || 1)).reduce((sum, value) => sum + value, 0),
    ), measured);
    if (y + protectedHeight > pageBottom && protectedHeight <= freshPageCapacity) startContinuationPage();
    measured = measureRow(row, remainingTexts);
    if (y + measured <= pageBottom) {
      drawRowContent(row, measured, remainingTexts);
      return;
    }

    while (remainingTexts.some(Boolean)) {
      if (pageBottom - y < Math.max(12, row.height)) startContinuationPage();
      const availableHeight = pageBottom - y;
      const layouts = layoutsForRow(row, remainingTexts);
      // Row-span (merged) cells are continued only by activeSpans/redrawActiveSpans, which re-draws the whole
      // value at the top of each continuation page (SSRS merged-cell behaviour). They must never also produce a
      // split tail here, or the tail and the redrawn value overlap on the continuation page.
      const parts = layouts.map((layout) => ((layout.cell.rowSpan || 1) > 1
        ? { head: layout.text, tail: '' }
        : splitTextForHeight(doc, config, layout.textbox, layout.context, layout.text, layout.width, availableHeight)));
      const hasContinuation = parts.some((part) => part.tail.length > 0);
      const heads = parts.map((part) => part.head);
      const segmentHeight = hasContinuation ? availableHeight : Math.min(availableHeight, measureRow(row, heads));
      drawRowContent(row, segmentHeight, heads, !hasContinuation);
      remainingTexts = parts.map((part) => part.tail);
      if (hasContinuation) startContinuationPage();
    }
  };

  for (const row of rows) drawRow(row);
  for (const span of openSpans) drawSpanSegment(span, y); // flush any spans still open at the tablix end
  openSpans = [];
  closeOuterBorderFragment(y);
  return { height: Math.max(item.height, addedHeight), endY: y };
}

export async function renderPdf(model, request, config) {
  borderWidthFloor = config?.borderWidthFloorPt || 0;
  const doc = new PDFDocument({ autoFirstPage: false, bufferPages: true, compress: true, info: { Title: request.outputFileName || model.name, Producer: 'RDL Converter Service' } });
  const completion = collectDocument(doc);
  const page = model.page;
  const headerHeight = page.header?.height || 0;
  const footerHeight = page.footer?.height || 0;
  const bodyTop = page.marginTop + headerHeight;
  const pageBottom = page.height - page.marginBottom - footerHeight;
  const globals = { PageNumber: 0, TotalPages: 1, ReportName: request.outputFileName || model.name, ExecutionTime: new Date(), variables: model.variables || {} };
  const addPage = () => {
    doc.addPage({ size: [page.width, page.height], margins: { top: 0, right: 0, bottom: 0, left: 0 } });
    globals.PageNumber += 1;
  };
  addPage.bodyTop = bodyTop;
  addPage();

  const datasets = normalizeDatasets(model, request);
  const items = [...model.body.items].sort((left, right) => left.top - right.top || left.left - right.left || left.zIndex - right.zIndex);
  let cursorY = bodyTop;
  let previousDesignBottom = 0;
  let pageHasContent = false;
  let forcePageBreak = false;
  for (const item of items) {
    const context = { parameters: request.parameters || {}, globals, datasets, dataset: [], fields: {} };
    if (isHidden(item.hidden, context)) continue;
    const breakDisabled = item.pageBreak ? isHidden(item.pageBreak.disabled, context) : true;
    const breakLocation = breakDisabled ? 'None' : String(item.pageBreak?.location || 'None');
    if (forcePageBreak || (/^(Start|StartAndEnd)$/i.test(breakLocation) && pageHasContent)) {
      addPage();
      cursorY = bodyTop;
      pageHasContent = false;
      forcePageBreak = false;
    }
    const gap = pageHasContent ? Math.max(0, item.top - previousDesignBottom) : 0;
    let y = cursorY + gap;
    if (y >= pageBottom || (item.type !== 'Tablix' && y + item.height > pageBottom && pageHasContent)) {
      addPage();
      y = bodyTop;
      pageHasContent = false;
    }
    const x = page.marginLeft + item.left;
    if (item.type === 'Tablix') {
      const rendered = renderTablix({ doc, config, model, item, request, startX: x, startY: y, pageBottom, addPage, globals });
      cursorY = rendered.endY;
    } else {
      drawSimpleItem(doc, config, model, item, x, y, context);
      cursorY = y + item.height;
    }
    pageHasContent = true;
    previousDesignBottom = Math.max(previousDesignBottom, item.top + item.height);
    if (/^(End|StartAndEnd)$/i.test(breakLocation)) {
      forcePageBreak = true;
    }
  }

  const range = doc.bufferedPageRange();
  globals.TotalPages = range.count;
  for (let index = 0; index < range.count; index += 1) {
    doc.switchToPage(range.start + index);
    globals.PageNumber = index + 1;
    const context = { parameters: request.parameters || {}, globals, datasets, dataset: [], fields: {} };
    if (page.header && (index > 0 || page.header.printOnFirstPage) && (index < range.count - 1 || page.header.printOnLastPage)) {
      for (const item of [...page.header.items].sort((left, right) => left.zIndex - right.zIndex || left.top - right.top || left.left - right.left)) drawSimpleItem(doc, config, model, item, page.marginLeft + item.left, page.marginTop + item.top, context);
    }
    if (page.footer && (index > 0 || page.footer.printOnFirstPage) && (index < range.count - 1 || page.footer.printOnLastPage)) {
      const footerTop = page.height - page.marginBottom - page.footer.height;
      for (const item of [...page.footer.items].sort((left, right) => left.zIndex - right.zIndex || left.top - right.top || left.left - right.left)) drawSimpleItem(doc, config, model, item, page.marginLeft + item.left, footerTop + item.top, context);
    }
  }
  doc.end();
  const buffer = await completion;
  const parsed = await PdfLibDocument.load(buffer);
  return { buffer, pageCount: parsed.getPageCount(), mimeType: 'application/pdf', extension: 'pdf' };
}
