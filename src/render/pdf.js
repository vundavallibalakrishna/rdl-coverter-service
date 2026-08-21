import PDFDocument from 'pdfkit';
import { PDFDocument as PdfLibDocument } from 'pdf-lib';
import { ServiceError } from '../errors.js';
import { CONTINUATION_MARKERS, cellBorderStyle, cellText, cellTextbox, color, continuationMarkersEnabled, enforcedBottomBorder, isHidden, matchingChangedGroupOwnerRowBoundary, materializedCellContext, materializedCellVisualSignature, normalizeDatasets, shouldEnforceTablixBottom, styleColor, styleSize, styleText, styleValue, styledSegmentsForText, styledTextForItem, tablixRows, textForItem } from './common.js';
import { fontVerticalMetrics, pdfFont } from './fonts.js';
import { computeCellPlacements } from './tableGrid.js';
import { cellGeometryPt, resolveGridColumns } from './tableLayout.js';
import { materializeChart } from './chartData.js';
import { drawChart } from './chart.js';
import { resolveReportCulture } from '../rdl/expression.js';
import {
  attachLayoutTrace,
  beginLayoutTracePage,
  createLayoutTrace,
  finalizeLayoutTrace,
  recordLayoutItem,
  selectLayoutTracePage,
} from './layoutTrace.js';

// Minimum border stroke width (points), set from config at the start of each render. A floor lets thin
// 1pt hairlines render at a crisp, uniform weight instead of rounding unevenly across screen zoom levels.
let borderWidthFloor = 0;
const COINCIDENT_EDGE_TOLERANCE_PT = 0.25;
// Minimum number of text lines a growable textbox fragment must be able to carry before the block is
// allowed to start on the current page. A break that strands fewer lines than this reads as a
// typesetting error rather than as flow, so the block moves to the next page instead. This is the
// standard typographic orphan minimum and matches Word's default widow/orphan control.
const MINIMUM_FLOWED_TEXT_LINES = 2;

// Two free-form canvas peers share a horizontal lane when their declared boxes overlap on the x axis.
// Growth only displaces peers in the same lane, so items designed side by side stay side by side.
function canvasLanesOverlap(left, right) {
  const leftStart = left.left || 0;
  const leftEnd = leftStart + (left.width || 0);
  const rightStart = right.left || 0;
  const rightEnd = rightStart + (right.width || 0);
  return leftStart < rightEnd - COINCIDENT_EDGE_TOLERANCE_PT
    && rightStart < leftEnd - COINCIDENT_EDGE_TOLERANCE_PT;
}

// PageBreak is a property of EVERY RDL report item, not only of a direct <Body> child: a rectangle nested
// inside another rectangle, or a tablix inside one, carries the same property with the same meaning. RDL
// makes BreakLocation a plain enum while Disabled is expression-capable, so only Disabled resolves through
// the style helpers.
function activeBreakLocation(item, context) {
  if (!item?.pageBreak || isHidden(item.pageBreak.disabled, context)) return 'None';
  return String(item.pageBreak.location || 'None');
}

const breaksBeforeItem = (location) => /^(Start|StartAndEnd)$/i.test(location);
const breaksAfterItem = (location) => /^(End|StartAndEnd)$/i.test(location);

// Structural (context-free) question: can this subtree move the page cursor by itself? A declared break
// makes the item a flow participant rather than a fixed-coordinate one, exactly like a tablix or a growable
// textbox. Disabled is per-row, so this deliberately over-approximates: a disabled break costs the item its
// fixed-coordinate fast path, never a wrong page.
function declaresPageBreak(item) {
  if (item?.pageBreak && !/^None$/i.test(String(item.pageBreak.location || 'None'))) return true;
  return (item?.items || []).some(declaresPageBreak);
}

// `isolate` marks an item that must not share a layout band with its neighbours. A break-carrying item is
// its own flow unit: banding it with a coordinate peer would render the peer before the break is applied.
function containerLayoutBands(items, isolate = () => false) {
  const ordered = [...items].sort((left, right) => (
    (left.top || 0) - (right.top || 0)
    || (left.left || 0) - (right.left || 0)
    || (left.zIndex || 0) - (right.zIndex || 0)
  ));
  const bands = [];
  for (const item of ordered) {
    const top = item.top || 0;
    const bottom = top + (item.height || 0);
    const band = bands[bands.length - 1];
    const isolated = Boolean(isolate(item));
    const coincidentTop = band
      && Math.abs(top - band.top) <= COINCIDENT_EDGE_TOLERANCE_PT;
    const overlapsBand = band
      && top < band.designBottom - COINCIDENT_EDGE_TOLERANCE_PT;
    if (!band || isolated || band.isolated || (!coincidentTop && !overlapsBand)) {
      bands.push({ top, designBottom: bottom, items: [item], isolated });
      continue;
    }
    band.designBottom = Math.max(band.designBottom, bottom);
    band.items.push(item);
  }
  return bands;
}

function collectDocument(doc) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    doc.on('data', (chunk) => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);
  });
}

function applyFont(doc, config, style, context = {}, text = '') {
  const bold = /bold|600|700|800|900/i.test(String(styleValue(style.fontWeight, context, 'Normal')));
  const italic = /italic/i.test(String(styleValue(style.fontStyle, context, 'Normal')));
  doc.font(pdfFont(config, styleText(style.fontFamily, context, 'Arial'), bold, italic, text)).fontSize(styleSize(style.fontSize, context, 10) || 10).fillColor(styleColor(style.color, context));
}

function resolvedTraceFont(config, style, context, text = '') {
  const bold = /bold|600|700|800|900/i.test(String(styleValue(style?.fontWeight, context, 'Normal')));
  const italic = /italic/i.test(String(styleValue(style?.fontStyle, context, 'Normal')));
  const family = String(styleText(style?.fontFamily, context, 'Arial'));
  const file = pdfFont(config, family, bold, italic, text);
  const size = styleSize(style?.fontSize, context, 10) || 10;
  return {
    family,
    file,
    size,
    bold,
    italic,
    underline: /underline/i.test(String(styleValue(style?.textDecoration, context, 'None'))),
    strike: /line.?through/i.test(String(styleValue(style?.textDecoration, context, 'None'))),
    color: styleColor(style?.color, context, '#000000'),
    metrics: fontVerticalMetrics(file, size),
  };
}

function resolvedTraceBorder(border, context = {}) {
  if (!border) return null;
  const style = String(styleValue(border.style, context, 'None'));
  const width = styleSize(border.width, context, 1);
  const borderColor = styleColor(border.color, context, '#000000');
  if (/^none$/i.test(style) || !borderColor || width <= 0) return null;
  return { style, width, color: borderColor };
}

function resolvedTraceBorders(style, context = {}, explicitEdges = null) {
  if (explicitEdges) {
    return Object.fromEntries(['top', 'right', 'bottom', 'left'].map((side) => {
      const edge = explicitEdges[side];
      return [side, edge ? resolvedTraceBorder(edge.border, edge.context || context) : null];
    }));
  }
  const sides = style?.borders || (style?.border
    ? { top: style.border, right: style.border, bottom: style.border, left: style.border }
    : {});
  return Object.fromEntries(['top', 'right', 'bottom', 'left'].map((side) => (
    [side, resolvedTraceBorder(sides?.[side], context)]
  )));
}

function traceTextbox(doc, config, item, x, y, context, details) {
  // Direct PDF output has no attached trace. Do not resolve trace-only fonts, vertical metrics, lines, or
  // borders merely to pass them to recordLayoutItem(), which would discard them. DOCX_EDITABLE attaches
  // the canonical trace before drawing and therefore retains the complete existing path.
  if (details.trace === false || !doc._rdlLayoutTrace) return;
  const layout = details.styledLayout;
  // Resolved lazily: it is consumed only by the no-styled-layout branch below, and asking for one font that
  // suits the entire textbox is a strictly harder question than the per-run resolution the styled path uses
  // (a single string mixing scripts, or Latin with a pictograph, has no single covering face). Computing it
  // eagerly imposed that harder question on every traced textbox and threw away the answer.
  let resolvedFallbackFont = null;
  const fallbackFont = () => {
    if (!resolvedFallbackFont) resolvedFallbackFont = resolvedTraceFont(config, item.style || {}, context, details.text);
    return resolvedFallbackFont;
  };
  let lineOffset = details.localTextY || 0;
  const innerWidth = Math.max(1, details.width - details.padding.left - details.padding.right);
  const lines = layout?.lines?.map((line) => {
    const alignment = String(
      styleValue(line.paragraphStyle?.textAlign ?? item.style?.textAlign, context, 'left'),
    ).toLowerCase();
    const alignedOffset = alignment === 'center'
      ? Math.max(0, (innerWidth - line.width) / 2)
      : alignment === 'right'
        ? Math.max(0, innerWidth - line.width)
        : 0;
    const spaces = line.runs.reduce((count, run) => count + ((run.text.match(/ /g) || []).length), 0);
    const justifyExtra = alignment === 'justify' && !line.paragraphEnd && spaces > 0
      ? Math.max(0, innerWidth - line.width) / spaces
      : 0;
    const absoluteTextTop = y + details.padding.top + lineOffset + line.before;
    let runX = x + details.padding.left + alignedOffset;
    const runs = line.runs.map((run) => {
      const font = resolvedTraceFont(config, run.style || item.style || {}, context, run.text);
      const tracedRun = {
        text: run.text,
        x: runX,
        y: absoluteTextTop,
        baseline: font.metrics ? absoluteTextTop + font.metrics.ascender : null,
        width: run.width,
        font,
      };
      runX += run.width + justifyExtra * ((run.text.match(/ /g) || []).length);
      return tracedRun;
    });
    const traced = {
      width: line.width,
      height: line.height,
      contentHeight: line.contentHeight,
      before: line.before,
      after: line.after,
      top: details.padding.top + lineOffset,
      textTop: details.padding.top + lineOffset + line.before,
      x: x + details.padding.left + alignedOffset,
      y: absoluteTextTop,
      baseline: runs.length > 0 ? Math.max(...runs.map((run) => run.baseline || absoluteTextTop)) : absoluteTextTop,
      paragraphEnd: line.paragraphEnd,
      wrapped: !line.paragraphEnd,
      alignment,
      runs,
    };
    lineOffset += line.height;
    return traced;
  }) || String(details.text ?? '').split('\n').map((text, index) => {
    const font = fallbackFont();
    const contentHeight = font.size * 1.2;
    const textTop = details.padding.top + (details.localTextY || 0) + index * contentHeight;
    const absoluteTextTop = y + textTop;
    const runX = x + details.padding.left;
    const baseline = font.metrics
      ? absoluteTextTop + font.metrics.ascender
      : absoluteTextTop;
    return {
      width: null,
      height: contentHeight,
      contentHeight,
      before: 0,
      after: 0,
      top: textTop,
      textTop,
      x: runX,
      y: absoluteTextTop,
      baseline,
      paragraphEnd: true,
      wrapped: false,
      alignment: String(styleValue(item.style?.textAlign, context, 'left')).toLowerCase(),
      runs: [{ text, x: runX, y: absoluteTextTop, baseline, width: null, font }],
    };
  });
  recordLayoutItem(doc, {
    kind: details.traceMeta?.kind || 'textbox',
    itemName: item.name || null,
    zIndex: item.zIndex || 0,
    x,
    y,
    width: details.width,
    height: details.height,
    text: String(details.text ?? ''),
    lines,
    writingMode: details.writingMode,
    verticalAlign: String(styleValue(item.style?.verticalAlign, context, 'top')).toLowerCase(),
    padding: details.padding,
    textOffsetY: details.localTextY || 0,
    backgroundColor: styleColor(item.style?.backgroundColor, context, null),
    borders: resolvedTraceBorders(item.style, context, details.traceEdges),
    ...details.traceMeta,
  });
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
  // Double: two parallel strands with a gap between them (line + gap + line, each ~1/3 of the nominal width —
  // the CSS/OOXML double-border model). PDFKit has no double line style, so it is drawn explicitly. This
  // keeps the PDF in parity with the editable DOCX (which maps Double -> a real double rule) and with SSRS,
  // which renders Double as two rules rather than one thick line. Solid/dashed/dotted are unaffected.
  if (/double/i.test(borderStyle)) {
    const strand = Math.max(0.25, borderWidth / 3);
    const vertical = x1 === x2;
    const [ox, oy] = vertical ? [strand, 0] : [0, strand]; // offset perpendicular to the edge
    doc.save().lineWidth(strand).strokeColor(borderColor).lineCap('square').lineJoin('miter');
    doc.moveTo(x1 - ox, y1 - oy).lineTo(x2 - ox, y2 - oy).stroke();
    doc.moveTo(x1 + ox, y1 + oy).lineTo(x2 + ox, y2 + oy).stroke();
    doc.restore();
    return;
  }
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

function lineHeightForStyle(doc, config, style, context) {
  applyFont(doc, config, style, context);
  return styleSize(style?.lineHeight, context, 0) || doc.currentLineHeight(true);
}

const graphemeSegmenter = new Intl.Segmenter(undefined, { granularity: 'grapheme' });

function graphemes(text) {
  return Array.from(graphemeSegmenter.segment(text), (part) => part.segment);
}

// Returns the largest non-empty grapheme prefix that fits `width`. Exact PDF font metrics are used rather
// than character counts so the same rule works for proportional fonts, bold/italic variants, combining
// marks, emoji sequences and expression-selected families. If one grapheme is intrinsically wider than the
// cell, return it alone: no legal line break can make that glyph narrower, and advancing guarantees progress.
function fittingGraphemeEnd(doc, parts, start, width) {
  let low = start + 1;
  let high = parts.length;
  let best = start;
  while (low <= high) {
    const middle = Math.floor((low + high) / 2);
    const candidate = parts.slice(start, middle).join('');
    if (doc.widthOfString(candidate) <= width) {
      best = middle;
      low = middle + 1;
    } else {
      high = middle - 1;
    }
  }
  return best > start ? best : start + 1;
}

function layoutStyledText(doc, config, item, context, text, width) {
  const source = styledSegmentsForText(item, context, text);
  if (!source) return null;
  if (source.text.length === 0) return { lines: [], height: 0 };
  const lines = [];
  const startedParagraphs = new Set();
  let line = null;
  const startLine = (style, paragraphStyle, paragraphIndex = 0) => {
    const effectiveParagraphStyle = paragraphStyle || item.style;
    const firstLine = !startedParagraphs.has(paragraphIndex);
    startedParagraphs.add(paragraphIndex);
    const before = firstLine ? styleSize(effectiveParagraphStyle?.spaceBefore, context, 0) : 0;
    const contentHeight = lineHeightForStyle(doc, config, style || item.style, context);
    return {
      runs: [],
      width: 0,
      contentHeight,
      before,
      after: 0,
      height: before + contentHeight,
      paragraphStyle: effectiveParagraphStyle,
      paragraphIndex,
      paragraphEnd: false,
    };
  };
  const finishLine = (paragraphEnd = false) => {
    if (!line) line = startLine(item.style, item.style);
    line.paragraphEnd = paragraphEnd;
    line.after = paragraphEnd ? styleSize(line.paragraphStyle?.spaceAfter, context, 0) : 0;
    line.height = line.before + line.contentHeight + line.after;
    lines.push(line);
    line = null;
  };
  const addToken = (token, segment) => {
    if (!line) line = startLine(segment.style, segment.paragraphStyle, segment.paragraphIndex);
    const whitespace = /^\s+$/.test(token);
    applyFont(doc, config, segment.style, context, token);
    const tokenWidth = doc.widthOfString(token);
    const tokenHeight = styleSize(segment.style?.lineHeight, context, 0) || doc.currentLineHeight(true);
    if (line.runs.length > 0 && line.width + tokenWidth > width) {
      finishLine(false);
      line = startLine(segment.style, segment.paragraphStyle, segment.paragraphIndex);
      // Wrapping consumes the separating whitespace, matching PDFKit's normal word-wrap behaviour.
      if (whitespace) return;
    }
    if (whitespace || tokenWidth <= width) {
      line.runs.push({ text: token, style: segment.style, width: tokenWidth });
      line.width += tokenWidth;
      line.contentHeight = Math.max(line.contentHeight, tokenHeight);
      line.height = line.before + line.contentHeight;
      return;
    }

    // A token wider than an otherwise-empty line must wrap within the word. Treating it as indivisible
    // draws it past the inner cell width; the mandatory cell clip then truncates the suffix. Split only on
    // Unicode grapheme boundaries so no surrogate pair, combining mark or joined emoji is corrupted.
    const parts = graphemes(token);
    let start = 0;
    while (start < parts.length) {
      const end = fittingGraphemeEnd(doc, parts, start, width);
      const fragment = parts.slice(start, end).join('');
      const fragmentWidth = doc.widthOfString(fragment);
      line.runs.push({ text: fragment, style: segment.style, width: fragmentWidth });
      line.width += fragmentWidth;
      line.contentHeight = Math.max(line.contentHeight, tokenHeight);
      line.height = line.before + line.contentHeight;
      start = end;
      if (start < parts.length) {
        finishLine(false);
        line = startLine(segment.style, segment.paragraphStyle, segment.paragraphIndex);
      }
    }
  };

  for (const segment of source.segments) {
    const parts = segment.text.split(/(\n)/);
    for (const part of parts) {
      if (part === '') continue;
      if (part === '\n') {
        if (!line) line = startLine(segment.style, segment.paragraphStyle, segment.paragraphIndex);
        finishLine(segment.paragraphBreak === true);
        continue;
      }
      const tokens = part.match(/[^\S\n]+|[^\s\n]+/g) || [];
      for (const token of tokens) addToken(token, segment);
    }
  }
  // PDFKit ignores one trailing empty line. Preserve empty lines between paragraphs, but not a terminal one.
  if (line?.runs.length || lines.length === 0) finishLine(true);
  else if (lines.length > 0 && !lines.at(-1).paragraphEnd) {
    // A trailing TextRun newline does not create another visible line. Close its RDL paragraph once so
    // SpaceAfter is retained exactly once instead of either being lost or charged per embedded newline.
    const last = lines.at(-1);
    last.paragraphEnd = true;
    last.after = styleSize(last.paragraphStyle?.spaceAfter, context, 0);
    last.height = last.before + last.contentHeight + last.after;
  }
  return { lines, height: lines.reduce((sum, current) => sum + current.height, 0) };
}

function drawStyledText(doc, config, item, context, layout, x, y, width) {
  let lineY = y;
  for (const line of layout.lines) {
    const textY = lineY + line.before;
    const alignment = String(styleValue(line.paragraphStyle?.textAlign ?? item.style?.textAlign, context, 'left')).toLowerCase();
    let lineX = x;
    if (alignment === 'center') lineX += Math.max(0, (width - line.width) / 2);
    else if (alignment === 'right') lineX += Math.max(0, width - line.width);
    const spaces = line.runs.reduce((count, run) => count + ((run.text.match(/ /g) || []).length), 0);
    const justifyExtra = alignment === 'justify' && !line.paragraphEnd && spaces > 0
      ? Math.max(0, width - line.width) / spaces
      : 0;
    for (const run of line.runs) {
      applyFont(doc, config, run.style, context, run.text);
      const decoration = String(styleValue(run.style?.textDecoration, context, 'None'));
      const underline = /underline/i.test(decoration);
      const strike = /line.?through/i.test(decoration);
      // PDFKit 0.17 computes a NaN decoration endpoint when underline/strike is combined with
      // `lineBreak:false` (its fragment has no `textWidth`). Styled layout already measured the exact run
      // width, so draw the decoration explicitly using that grounded geometry instead of asking PDFKit to
      // re-measure. This also keeps the rule aligned with mixed-font runs.
      doc.text(run.text, lineX, textY, { lineBreak: false });
      if ((underline || strike) && run.width > 0) {
        const fontSize = styleSize(run.style?.fontSize, context, 10) || 10;
        const lineWidth = fontSize < 10 ? 0.5 : Math.floor(fontSize / 10);
        const decorationY = underline ? textY + doc.currentLineHeight() - lineWidth : textY + doc.currentLineHeight() / 2;
        doc.save()
          .strokeColor(styleColor(run.style?.color, context, '#000000'))
          .lineWidth(lineWidth)
          .moveTo(lineX, decorationY)
          .lineTo(lineX + run.width, decorationY)
          .stroke()
          .restore();
      }
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
  const writingMode = String(styleValue(style.writingMode, context, 'Default') || 'Default').toLowerCase();
  const rotated = writingMode === 'rotate270' || writingMode === 'vertical';
  const layoutWidth = rotated ? innerHeight : innerWidth;
  const layoutHeight = rotated ? innerWidth : innerHeight;
  const styledLayout = layoutStyledText(doc, config, item, context, text, layoutWidth);
  if (!styledLayout) applyFont(doc, config, style, context, text);
  const measuredHeight = styledLayout?.height ?? renderedTextHeight(doc, text, layoutWidth);
  // VerticalAlign can be an expression; resolve it before matching, otherwise the regex tests the raw
  // expression source (e.g. `=IIF(c,"Middle","Top")` always contains "Middle").
  const verticalAlign = String(styleValue(style.verticalAlign, context, 'top')).toLowerCase();
  let localTextY = 0;
  if (/middle|center/.test(verticalAlign)) localTextY = Math.max(0, (layoutHeight - measuredHeight) / 2);
  if (/bottom/.test(verticalAlign)) localTextY = Math.max(0, layoutHeight - measuredHeight);
  traceTextbox(doc, config, item, x, y, context, {
    trace: override.trace,
    traceMeta: override.traceMeta,
    traceEdges: override.traceEdges,
    width,
    height,
    text,
    styledLayout,
    writingMode,
    padding: {
      top: padTop,
      right: padRight,
      bottom: padBottom,
      left: paddingLeft,
    },
    localTextY,
  });
  // Clip text to the cell box so it can never bleed into an adjacent cell or, when the cell is
  // clamped at the page/footer boundary, into the reserved footer band. Background and borders are
  // drawn above (unclipped) so their edges stay crisp.
  doc.save().rect(x, y, width, height).clip();
  if (writingMode === 'rotate270') {
    // Local horizontal text is mapped into the physical content box bottom-to-top:
    //   physicalX = localY + left, physicalY = -localX + bottom.
    // The textbox background and borders remain in the original coordinate system.
    doc.transform(0, -1, 1, 0, x + paddingLeft, y + padTop + innerHeight);
  } else if (writingMode === 'vertical') {
    // SSRS Vertical writes Latin glyphs top-to-bottom. Map the horizontal local line clockwise into
    // the physical box; East-Asian glyph shaping remains delegated to the selected font/PDFKit.
    doc.transform(0, 1, -1, 0, x + paddingLeft + innerWidth, y + padTop);
  } else {
    doc.translate(x + paddingLeft, y + padTop);
  }
  if (styledLayout) drawStyledText(doc, config, item, context, styledLayout, 0, localTextY, layoutWidth);
  else {
    doc.text(text, 0, localTextY, {
      width: layoutWidth,
      height: layoutHeight,
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
  recordLayoutItem(doc, {
    kind: 'image',
    itemName: item.name || null,
    embeddedImage: styleValue(item.value, context, item.value),
    sizing: String(styleValue(item.sizing, context, 'FitProportional') || 'FitProportional'),
    zIndex: item.zIndex || 0,
    x,
    y,
    width: item.width,
    height: item.height,
  });
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
    recordLayoutItem(doc, {
      kind: 'chart',
      itemName: item.name || null,
      zIndex: item.zIndex || 0,
      x,
      y,
      width: item.width,
      height: item.height,
    });
    drawChart(doc, config, item, data, x, y, item.width, item.height, context);
  } else if (item.type === 'Image') drawImage(doc, model, item, x, y, context);
  else if (item.type === 'Line') {
    const lineWidth = styleSize(item.style?.border?.width, context, 1) || 1;
    const lineColor = styleColor(item.style?.border?.color, context, '#000000');
    recordLayoutItem(doc, {
      kind: 'line',
      itemName: item.name || null,
      zIndex: item.zIndex || 0,
      x,
      y,
      width: item.width,
      height: item.height,
      // drawSimpleItem strokes standalone RDL Line primitives as solid PDF paths. Record the resolved
      // stroke that was actually painted so page-locked Word can materialize the same visible border.
      line: { style: 'Solid', width: lineWidth, color: lineColor },
    });
    doc.save().lineWidth(lineWidth).strokeColor(lineColor).moveTo(x, y).lineTo(x + item.width, y + item.height).stroke().restore();
  }
  else if (item.type === 'Rectangle') {
    const backgroundColor = styleColor(item.style.backgroundColor, context, null);
    recordLayoutItem(doc, {
      kind: 'rectangle',
      itemName: item.name || null,
      zIndex: item.zIndex || 0,
      x,
      y,
      width: item.width,
      height: item.height,
      backgroundColor,
      borders: resolvedTraceBorders(item.style, context),
    });
    if (backgroundColor) doc.save().fillColor(backgroundColor).rect(x, y, item.width, item.height).fill().restore();
    drawBorder(doc, x, y, item.width, item.height, item.style, context);
    for (const child of [...item.items].sort((left, right) => left.zIndex - right.zIndex || left.top - right.top || left.left - right.left)) {
      drawSimpleItem(doc, config, model, child, x + child.left, y + child.top, context);
    }
  }
}

// Shared by the native XLSX report renderer so wrapped row heights derive from the same font metrics and
// rich-text run boundaries as PDF rather than a character-count approximation.
export function measureTextboxHeight(doc, config, textbox, context, text, width) {
  if (!textbox || !text) return 0;
  const writingMode = String(styleValue(textbox.style?.writingMode, context, 'Default') || 'Default').toLowerCase();
  // A rotated textbox consumes the row's declared vertical extent; wrapping grows across the physical
  // width, not down the page. Returning its cross-axis line height prevents Excel/DOCX row measurement
  // from expanding a vertical header to the unrotated string height.
  if (writingMode === 'rotate270' || writingMode === 'vertical') {
    return styleSize(textbox.style?.fontSize, context, 10) || 10;
  }
  const innerWidth = Math.max(1, width - styleSize(textbox.style.paddingLeft, context, 2) - styleSize(textbox.style.paddingRight, context, 2));
  const styledLayout = layoutStyledText(doc, config, textbox, context, text, innerWidth);
  if (styledLayout) return styledLayout.height;
  applyFont(doc, config, textbox.style, context, text);
  return doc.heightOfString(text, { width: innerWidth, lineGap: 0 });
}

function splitTextForHeight(doc, config, textbox, context, text, width, height) {
  const value = String(text || '');
  if (!value || !textbox) return { head: value, tail: '' };
  const available = Math.max(1, height - styleSize(textbox.style.paddingTop, context, 2) - styleSize(textbox.style.paddingBottom, context, 2));
  if (measureTextboxHeight(doc, config, textbox, context, value, width) <= available) return { head: value, tail: '' };
  let low = 1;
  let high = value.length;
  let best = 0;
  while (low <= high) {
    const middle = Math.floor((low + high) / 2);
    if (measureTextboxHeight(doc, config, textbox, context, value.slice(0, middle), width) <= available) {
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

function renderTablix({ doc, config, model, item, request, startX, startY, pageBottom, addPage, globals, statistics }) {
  const tablixStartedAt = performance.now();
  const { rows, columns } = tablixRows(item, request, globals, model);
  const materializedAt = performance.now();
  statistics.tablixCount += 1;
  statistics.tablixRowCount += rows.length;
  statistics.tablixCellCount += rows.reduce((sum, row) => sum + row.cells.length, 0);
  statistics.tablixMaterializationMs += materializedAt - tablixStartedAt;
  const enforceBottomClosure = shouldEnforceTablixBottom(rows, item);
  const datasets = normalizeDatasets(model, request);
  // A matrix expands to a data-dependent column grid wider than the design width; use its natural
  // total so columns are not scaled down. Static-column tablixes keep item's declared width scaling.
  const layoutItem = item.hasColumnGroups ? { ...item, columns, width: columns.reduce((sum, width) => sum + width, 0) } : item;
  const { columnsPt: columnWidths, totalPt: totalWidth } = resolveGridColumns(layoutItem);
  const headers = rows.filter((row) => row.isHeader);
  // Initial placement concerns only the leading contiguous header block. Other repeatable rows can be
  // meaningful to continuation replay in grouped tablixes, so preserve the renderer's established
  // `headers` set and use this narrower block only for the keep-with-first-body preflight below.
  const leadingHeaders = [];
  for (const row of rows) {
    if (!row.isHeader) break;
    leadingHeaders.push(row);
  }
  const placements = computeCellPlacements(rows, columnWidths.length);
  const rowIndexes = new Map(rows.map((row, index) => [row, index]));
  const outerContext = { parameters: request.parameters || {}, globals, dataset: datasets[item.datasetName] || [], datasets, fields: {} };
  const showContinuationMarkers = continuationMarkersEnabled(request);

  const markerDetails = (row) => {
    const textbox = row?.cells?.map((cell) => cellTextbox(cell)).find(Boolean);
    const style = textbox?.style || item.style || {};
    const context = row
      ? { ...outerContext, fields: row.fields || {} }
      : outerContext;
    const sourceSize = styleSize(style.fontSize, context, 10) || 10;
    const fontSize = Math.max(7, sourceSize * 0.8);
    return { style, context, fontSize, height: Math.max(10, fontSize * 1.5) };
  };

  const drawContinuationMarker = (label, row) => {
    const { style, context, fontSize, height } = markerDetails(row);
    const markerY = y;
    const markerStyle = {
      ...style,
      fontSize,
      fontStyle: 'Italic',
      fontWeight: 'Normal',
      color: '#000000',
      textAlign: 'Right',
      verticalAlign: 'Middle',
    };
    doc.save();
    doc.rect(startX, markerY, totalWidth, height).fill('#FFFFFF');
    const family = styleText(style.fontFamily, context, 'Arial');
    const fontFile = pdfFont(config, family, false, true, label);
    doc.font(fontFile)
      .fontSize(fontSize)
      .fillColor('#000000')
      .text(label, startX + 2, markerY + Math.max(0, (height - fontSize) / 2 - 1), {
        width: Math.max(1, totalWidth - 4), height, align: 'right', lineBreak: false,
      });
    // The marker already occupies this exact PDF region and advances `y` below. Trace the resolved native
    // text into that existing region so page-locked Word fills the canonical gap instead of inserting a
    // flow paragraph that could repaginate the table. Recording is deliberately side-effect-free: all
    // metrics come from the font state already used to draw the marker.
    const textWidth = doc.widthOfString(label);
    const contentHeight = doc.currentLineHeight(true);
    const textY = markerY + Math.max(0, (height - fontSize) / 2 - 1);
    const font = {
      ...resolvedTraceFont(config, markerStyle, context, label),
      family: String(family),
      file: fontFile,
    };
    const textX = startX + totalWidth - 2 - textWidth;
    const baseline = font.metrics ? textY + font.metrics.ascender : textY;
    recordLayoutItem(doc, {
      kind: 'textbox',
      itemName: null,
      tablixName: item.name || null,
      traceRole: 'continuationMarker',
      continuation: true,
      zIndex: item.zIndex || 0,
      x: startX,
      y: markerY,
      width: totalWidth,
      height,
      text: label,
      lines: [{
        width: textWidth,
        height: contentHeight,
        contentHeight,
        before: 0,
        after: 0,
        top: textY - markerY,
        textTop: textY - markerY,
        x: textX,
        y: textY,
        baseline,
        paragraphEnd: true,
        wrapped: false,
        alignment: 'right',
        runs: [{ text: label, x: textX, y: textY, baseline, width: textWidth, font }],
      }],
      writingMode: 'default',
      verticalAlign: 'middle',
      padding: { top: 0, right: 2, bottom: 0, left: 2 },
      textOffsetY: textY - markerY,
      backgroundColor: '#FFFFFF',
      borders: { top: null, right: null, bottom: null, left: null },
    });
    doc.restore();
    y += height;
    addedHeight += height;
  };

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
  const contextForCell = (rowIndex, cell = null) => materializedCellContext(cell, rows[rowIndex], {
    parameters: request.parameters || {},
    globals,
    dataset: datasets[item.datasetName] || [],
    datasets,
  });
  // A side's border for a cell, evaluated in that cell's row context. Returns null when the side is
  // absent or resolves to None so the caller can fall back to the neighbouring cell's opposite side.
  const resolveSide = (owner, side) => {
    if (!owner) return null;
    const cellStyle = cellBorderStyle(owner.cell, item);
    const border = cellStyle?.borders?.[side];
    if (!border) return null;
    const context = contextForCell(owner.rowIndex, owner.cell);
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
    const top = resolveSide(owner, 'top') || resolveSide(above, 'bottom') || matchingChangedGroupOwnerRowBoundary(
      owner,
      above,
      left,
      right,
      resolveSide,
      ({ border, context }) => `${styleValue(border.style, context, 'None')}|${styleColor(border.color, context, null)}|${styleSize(border.width, context, 1)}`,
      (candidate) => {
        const context = contextForCell(candidate.rowIndex, candidate.cell);
        const style = cellTextbox(candidate.cell)?.style || item.style;
        return materializedCellVisualSignature(candidate.cell, style, context);
      },
    );
    return {
      top,
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
  const collectEdge = (x, y, width, height, side, border, context, traceMeta = null) => {
    if (!border) return;
    statistics.borderEdgesCollected += 1;
    const vertical = side === 'left' || side === 'right';
    const pos = vertical ? (side === 'right' ? x + width : x) : (side === 'bottom' ? y + height : y);
    const [a, b] = vertical ? [y, y + height] : [x, x + width];
    const sig = `${styleValue(border.style, context, 'None')}|${styleColor(border.color, context, null)}|${styleSize(border.width, context, 1)}`;
    pendingEdges.push({
      orient: vertical ? 'V' : 'H', pos, a, b, border, context, sig, traceMeta,
    });
  };
  const flushEdges = () => {
    const groups = new Map();
    for (const edge of pendingEdges) {
      const key = `${edge.orient}|${posKey(edge.pos)}|${edge.sig}`;
      if (!groups.has(key)) groups.set(key, { edge, intervals: [], traceMeta: null });
      const group = groups.get(key);
      group.intervals.push([edge.a, edge.b]);
      if (edge.traceMeta) group.traceMeta = edge.traceMeta;
    }
    for (const { edge, intervals, traceMeta } of groups.values()) {
      intervals.sort((p, q) => p[0] - q[0]);
      let [runStart, runEnd] = intervals[0];
      const runs = [];
      for (let i = 1; i < intervals.length; i += 1) {
        const [s, e] = intervals[i];
        if (s <= runEnd + 0.75) runEnd = Math.max(runEnd, e); // collinear & touching → extend the run
        else { runs.push([runStart, runEnd]); [runStart, runEnd] = [s, e]; }
      }
      runs.push([runStart, runEnd]);
      statistics.borderRunsDrawn += runs.length;
      for (const [s, e] of runs) {
        const vertical = edge.orient === 'V';
        if (traceMeta && !/^none$/i.test(String(styleValue(edge.border.style, edge.context, 'None')))) {
          // Record the exact outer fragment stroke produced by the PDF shared-edge resolver. Cell-level
          // metadata cannot describe a synthesized fragment closure (notably the mandatory bottom edge
          // of a data tablix), while recording every internal stroke would incorrectly turn nested-grid
          // borders into free-standing lines. These records are trace-only and cannot change PDF output.
          recordLayoutItem(doc, {
            kind: 'line',
            itemName: item.name || null,
            tablixName: item.name || null,
            traceRole: 'resolvedTablixFragmentBorder',
            fragmentSide: traceMeta.side,
            zIndex: item.zIndex || 0,
            x: vertical ? edge.pos : s,
            y: vertical ? s : edge.pos,
            width: vertical ? 0 : e - s,
            height: vertical ? e - s : 0,
            line: {
              style: styleValue(edge.border.style, edge.context, 'None'),
              color: styleColor(edge.border.color, edge.context, '#000000'),
              width: styleSize(edge.border.width, edge.context, 1),
            },
          });
        }
        if (vertical) drawBorderEdge(doc, edge.pos, s, 0, e - s, 'left', edge.border, edge.context);
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
  const nestedLayout = (nested, availableWidth) => {
    const nestedParameters = nested.parameters || request.parameters || {};
    const nestedDatasets = nested.datasets || datasets;
    const nestedGlobals = nested.globals || globals;
    const naturalColumns = nested.columns || nested.item.columns || [];
    const naturalWidth = Math.max(1, naturalColumns.reduce((sum, value) => sum + value, 0));
    const usableWidth = Math.max(1, availableWidth - Math.max(0, nested.item.left || 0));
    const width = Math.min(naturalWidth, usableWidth);
    const scale = width / naturalWidth;
    const columns = naturalColumns.map((value) => value * scale);
    const placements = computeCellPlacements(nested.rows || [], columns.length);
    const heights = (nested.rows || []).map((row, rowIndex) => row.cells.reduce((maximum, cell, cellIndex) => {
      const textbox = cellTextbox(cell);
      const columnIndex = placements[rowIndex][cellIndex];
      const cellWidth = cellGeometryPt(columns, columnIndex, cell.colSpan || 1).widthPt;
      const context = materializedCellContext(cell, row, {
        parameters: nestedParameters,
        globals: nestedGlobals,
        dataset: nestedDatasets[nested.item.datasetName] || [],
        datasets: nestedDatasets,
      });
      const textHeight = textbox && !cell.hidden
        ? measureTextboxHeight(doc, config, textbox, context, cellText(cell), cellWidth)
          + styleSize(textbox.style?.paddingTop, context, 2) + styleSize(textbox.style?.paddingBottom, context, 2)
        : 0;
      const childHeight = Math.max(0, ...(cell.nestedTablixes || []).map((child) => {
        const layout = nestedLayout(child, cellWidth);
        return (child.item.top || 0) + layout.height;
      }));
      return Math.max(maximum, textHeight, childHeight);
    }, row.height || 0));
    return { columns, placements, heights, width, height: heights.reduce((sum, value) => sum + value, 0) };
  };
  const drawNestedTablix = (nested, parentX, parentY, availableWidth) => {
    const nestedParameters = nested.parameters || request.parameters || {};
    const nestedDatasets = nested.datasets || datasets;
    const nestedGlobals = nested.globals || globals;
    const layout = nestedLayout(nested, availableWidth);
    const start = {
      x: parentX + (nested.item.left || 0),
      y: parentY + (nested.item.top || 0),
    };
    const rowOffsets = [0];
    layout.heights.forEach((height) => rowOffsets.push(rowOffsets[rowOffsets.length - 1] + height));
    const columnOffsets = [0];
    layout.columns.forEach((width) => columnOffsets.push(columnOffsets[columnOffsets.length - 1] + width));
    for (const [rowIndex, row] of (nested.rows || []).entries()) {
      for (const [cellIndex, cell] of row.cells.entries()) {
        const columnIndex = layout.placements[rowIndex][cellIndex];
        const colSpan = Math.max(1, cell.colSpan || 1);
        const rowSpan = Math.max(1, cell.rowSpan || 1);
        const x = start.x + columnOffsets[columnIndex];
        const cellY = start.y + rowOffsets[rowIndex];
        const width = columnOffsets[Math.min(layout.columns.length, columnIndex + colSpan)] - columnOffsets[columnIndex];
        const height = rowOffsets[Math.min(layout.heights.length, rowIndex + rowSpan)] - rowOffsets[rowIndex];
        const textbox = cellTextbox(cell);
        const context = materializedCellContext(cell, row, {
          parameters: nestedParameters,
          globals: nestedGlobals,
          dataset: nestedDatasets[nested.item.datasetName] || [],
          datasets: nestedDatasets,
        });
        const style = textbox?.style || nested.item.style || {};
        const borderStyle = cellBorderStyle(cell, nested.item) || {};
        const background = styleColor(style.backgroundColor, context, null);
        if (background) doc.save().rect(x, cellY, width, height).fill(background).restore();
        if (textbox && !cell.hidden) {
          drawTextbox(doc, config, textbox, x, cellY, context, {
            width,
            height,
            text: cellText(cell),
            skipBorder: true,
            traceEdges: Object.fromEntries(
              ['top', 'right', 'bottom', 'left'].map((side) => {
                const border = borderStyle.borders?.[side];
                return [side, border && !/^none$/i.test(String(styleValue(border.style, context, 'None')))
                  ? { border, context }
                  : null];
              }),
            ),
            traceMeta: {
              kind: 'tablixCell',
              tablixName: nested.item.name || null,
              rowIndex,
              columnIndex,
              colSpan,
              rowSpan,
              repeatedHeader: Boolean(row.isHeader),
              nested: true,
            },
          });
        } else {
          recordLayoutItem(doc, {
            kind: 'tablixCell',
            itemName: null,
            tablixName: nested.item.name || null,
            rowIndex,
            columnIndex,
            colSpan,
            rowSpan,
            repeatedHeader: Boolean(row.isHeader),
            nested: true,
            zIndex: nested.item.zIndex || 0,
            x,
            y: cellY,
            width,
            height,
            text: '',
            lines: [],
            writingMode: 'default',
            verticalAlign: 'top',
            padding: { top: 0, right: 0, bottom: 0, left: 0 },
            backgroundColor: background,
            borders: resolvedTraceBorders(borderStyle, context),
          });
        }
        const borders = borderStyle.borders || {};
        drawEdges(x, cellY, width, height, Object.fromEntries(
          ['top', 'right', 'bottom', 'left'].map((side) => {
            const border = borders[side];
            return [side, border && !/^none$/i.test(String(styleValue(border.style, context, 'None')))
              ? { border, context }
              : null];
          }),
        ));
        for (const child of cell.nestedTablixes || []) drawNestedTablix(child, x, cellY, width);
      }
    }
    const outer = nested.item.style?.borders || {};
    const nestedOuterContext = {
      parameters: nestedParameters,
      globals: nestedGlobals,
      dataset: nestedDatasets[nested.item.datasetName] || [],
      datasets: nestedDatasets,
      fields: {},
    };
    drawEdges(start.x, start.y, layout.width, layout.height, Object.fromEntries(
      ['top', 'right', 'bottom', 'left'].map((side) => {
        const border = outer[side];
        return [side, border && !/^none$/i.test(String(styleValue(border.style, nestedOuterContext, 'None')))
          ? { border, context: nestedOuterContext }
          : null];
      }),
    ));
    return layout.height;
  };
  // Draw a nested data region that is taller than one printable page, continuing its rows on further
  // pages instead of painting a single block through the footer band and off the sheet. The region's own
  // leading header rows repeat above every fragment, exactly like a top-level tablix continuation.
  // `screenY` is the page coordinate of the region's first row; the returned value is the page coordinate
  // its last fragment ended at, on whatever page the walk finished on.
  const drawNestedTablixAcrossPages = (nested, parentX, screenY, availableWidth) => {
    const top = nested.item.top || 0;
    const allRows = nested.rows || [];
    const full = nestedLayout(nested, availableWidth);
    const heightByRow = new Map(allRows.map((row, index) => [row, full.heights[index]]));
    const leadingHeaders = [];
    for (const row of allRows) {
      if (!row.isHeader) break;
      leadingHeaders.push(row);
    }
    const dataRows = allRows.slice(leadingHeaders.length);
    const headerHeight = leadingHeaders.reduce((sum, row) => sum + (heightByRow.get(row) || 0), 0);
    let cursor = screenY;
    let offset = 0;
    try {
      while (true) {
        const available = pageBottom - cursor;
        // First estimate the slice from the whole-region row heights, then measure the slice itself. A
        // fragment is a tablix in its own right: dropping rows changes which grid column each remaining
        // cell occupies (a row-span from an earlier row no longer covers it), so its rows can measure
        // taller than they did in the full region. Shrink until the measured fragment really fits.
        const budget = Math.max(1, available - headerHeight);
        let end = offset;
        let used = 0;
        while (end < dataRows.length) {
          const next = heightByRow.get(dataRows[end]) || 0;
          if (end > offset && used + next > budget) break;
          used += next;
          end += 1;
          if (used >= budget) break;
        }
        if (end === offset) end = offset + 1;
        let fragmentHeight = 0;
        while (true) {
          nested.rows = [...leadingHeaders, ...dataRows.slice(offset, end)];
          fragmentHeight = nestedLayout(nested, availableWidth).height;
          if (fragmentHeight <= available + 0.5 || end <= offset + 1) break;
          end -= 1;
        }
        if (fragmentHeight > available + 0.5) {
          // One child row plus the repeated headers does not fit here. Retry on an empty page; a row that
          // does not fit even there cannot be represented by row-level continuation, so fail closed rather
          // than drawing it through the footer band.
          if (cursor > addPage.bodyTop + 0.5) {
            addPage();
            cursor = addPage.bodyTop;
            continue;
          }
          throw new ServiceError(
            'UNSUPPORTED_FEATURE',
            `One row of nested tablix ${nested.item.name || 'unnamed'} exceeds one printable page`,
          );
        }
        drawNestedTablix(nested, parentX, cursor - top, availableWidth);
        cursor += fragmentHeight;
        offset = end;
        if (offset >= dataRows.length) break;
        addPage();
        cursor = addPage.bodyTop;
      }
    } finally {
      nested.rows = allRows;
    }
    return cursor;
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
    if (firstFragment) collectEdge(
      startX, fragmentStartY, totalWidth, fragmentHeight, 'top',
      item.style?.borders?.top, outerContext, { side: 'top' },
    );
    collectEdge(
      startX, fragmentStartY, totalWidth, fragmentHeight, 'left',
      item.style?.borders?.left, outerContext, { side: 'left' },
    );
    collectEdge(
      startX, fragmentStartY, totalWidth, fragmentHeight, 'right',
      item.style?.borders?.right, outerContext, { side: 'right' },
    );
    // Data fragments always close, including overflow pages. Static layout tablixes instead honor the
    // declared bottom edge, so Border=None cannot create a decorative rule after headings or prose.
    collectEdge(
      startX,
      fragmentStartY,
      totalWidth,
      fragmentHeight,
      'bottom',
      enforceBottomClosure ? enforcedBottomBorder(item.style) : item.style?.borders?.bottom,
      outerContext,
      { side: 'bottom' },
    );
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
      const context = contextForCell(rowIndex, cell);
      return { cell, textbox, width, columnIndex, context, text: texts[index] || '' };
    });
  };

  // A canvas cell is an SSRS List / free-form layout — a Rectangle flattened into several positioned items
  // (multiple textboxes, or any line/chart/image) drawn item-by-item. Its height is the extent of those
  // items, not the tablix's large design row height (which would falsely exceed a page).
  // Gated to cells that carry a Line, Chart, or Image — content that previously could not render in a cell
  // at all (it was refused at analyze time). A cell with only textboxes keeps its existing single-textbox
  // path, so ordinary multi-textbox cells are unchanged and no existing report regresses.
  const isCanvasCell = (cell) => (cell.items || []).some((entry) => (
    entry.type === 'Line' || entry.type === 'Chart' || entry.type === 'Image'
  ));
  // RDL CellContents holds exactly ONE report item. A Textbox there fills the cell, so its declared
  // Top/Left/Width/Height are irrelevant and the single-textbox fast path is exact. A Rectangle there is a
  // CANVAS: the flattened children keep their own declared position and size inside it. Drawing one of those
  // children stretched across the whole cell puts it in the wrong place and makes it swallow its positioned
  // siblings — which the page-locked Word renderer must then refuse as overlapping regions. A textbox that
  // shares its cell with a nested data region is therefore drawn item-by-item, like a Line/Chart/Image
  // canvas. Row measurement is deliberately untouched: only free-form media cells take their height from
  // the content extent, so this row keeps its declared height and pagination is unchanged.
  const isFreeFormCell = (cell) => {
    const items = cell.items || [];
    if (isCanvasCell(cell)) return true;
    const textboxes = items.filter((entry) => entry.type === 'Textbox');
    // Every textbox must declare its own box on the canvas. RDL defaults an omitted Width/Height to zero,
    // so drawing an undeclared one positioned would silently erase its text; such a cell keeps the
    // fill-the-cell path and, if that genuinely overlaps, still fails closed rather than losing content.
    return textboxes.length > 0
      && textboxes.every((entry) => (entry.width || 0) > 0 && (entry.height || 0) > 0)
      && items.some((entry) => entry.type === 'Tablix' || entry.type === 'Subreport');
  };
  const canvasCellExtent = (cell, width) => Math.max(
    0,
    ...(cell.items || []).filter((entry) => entry.type !== 'Tablix' && entry.type !== 'Subreport')
      .map((entry) => (entry.top || 0) + (entry.height || 0)),
    ...(cell.nestedTablixes || []).map((nested) => (nested.item.top || 0) + nestedLayout(nested, width).height),
  );

  const rowMeasurementCache = new WeakMap();
  const measureRow = (row, texts = row.cells.map((cell) => cellText(cell))) => {
    statistics.rowMeasurementRequests += 1;
    // Nested child rows may be temporarily replaced with one page fragment during subreport pagination.
    // Their geometry is mutable, so they deliberately bypass this otherwise exact memoization path.
    const cacheable = config.pdfLayoutOptimizations !== false
      && !row.cells.some((cell) => (cell.nestedTablixes || []).length > 0);
    const pageNumber = globals.PageNumber;
    const textKey = cacheable ? JSON.stringify(texts) : null;
    let pageCache;
    if (cacheable) {
      let rowCache = rowMeasurementCache.get(row);
      if (!rowCache) {
        rowCache = new Map();
        rowMeasurementCache.set(row, rowCache);
      }
      pageCache = rowCache.get(pageNumber);
      if (!pageCache) {
        pageCache = new Map();
        rowCache.set(pageNumber, pageCache);
      }
      if (pageCache.has(textKey)) {
        statistics.rowMeasurementCacheHits += 1;
        return pageCache.get(textKey);
      }
    }
    statistics.rowMeasurementsComputed += 1;
    const layouts = layoutsForRow(row, texts);
    const canvasCells = row.cells.filter(isCanvasCell);
    let measured;
    if (canvasCells.length > 0) {
      // A List/canvas row is sized to its content extent, not the tablix's design row height.
      const widthOf = new Map(layouts.map((layout) => [layout.cell, layout.width]));
      measured = Math.max(0, ...canvasCells.map((cell) => canvasCellExtent(cell, widthOf.get(cell) || totalWidth)));
    } else {
      measured = layouts.reduce((height, layout) => Math.max(
        height,
        measureTextboxHeight(doc, config, layout.textbox, layout.context, layout.text, layout.width)
          + styleSize(layout.textbox?.style.paddingTop, layout.context, 2) + styleSize(layout.textbox?.style.paddingBottom, layout.context, 2),
        Math.max(0, ...(layout.cell.nestedTablixes || []).map((nested) => (
          (nested.item.top || 0) + nestedLayout(nested, layout.width).height
        ))),
      ), row.height);
    }
    if (cacheable) pageCache.set(textKey, measured);
    return measured;
  };
  const measurementStartedAt = performance.now();
  statistics.tablixSetupMs += measurementStartedAt - materializedAt;
  const measuredHeights = rows.map((row) => measureRow(row));
  statistics.tablixInitialMeasurementMs += performance.now() - measurementStartedAt;
  const headerMeasurementsByPage = new Map();
  const headerMeasurements = () => {
    const pageNumber = globals.PageNumber;
    if (config.pdfLayoutOptimizations !== false && headerMeasurementsByPage.has(pageNumber)) {
      statistics.headerMeasurementCacheHits += 1;
      return headerMeasurementsByPage.get(pageNumber);
    }
    statistics.headerMeasurementsComputed += 1;
    const heights = headers.map((header) => measureRow(header));
    const result = { heights, total: heights.reduce((sum, height) => sum + height, 0) };
    if (config.pdfLayoutOptimizations !== false) headerMeasurementsByPage.set(pageNumber, result);
    return result;
  };

  // Draw one segment of an open merged cell: its fill + value + borders from segStartY down to endY,
  // clamped to the reserved body area so it never bleeds into the footer band. When the value is taller
  // than the segment, the overflow is recorded in `pendingTail` so it continues on the next page instead
  // of being clipped; a value that fits leaves no tail and therefore repeats at the top of each page.
  const drawSpanSegment = (span, endY, finalSegment = false) => {
    const segmentHeight = Math.min(endY - span.segStartY, Math.max(0, pageBottom - span.segStartY));
    if (segmentHeight <= 0.5) return;
    span.pendingTail = null;
    if (span.textbox && !span.cell.hidden) {
      const { head, tail } = splitTextForHeight(doc, config, span.textbox, span.context, span.text, span.width, segmentHeight);
      if (tail && tail.length > 0) span.pendingTail = tail;
      drawTextbox(doc, config, span.textbox, span.x, span.segStartY, span.context, {
        width: span.width,
        height: segmentHeight,
        text: head,
        skipBorder: true,
        traceEdges: span.edges,
        traceMeta: {
          kind: 'tablixCell',
          tablixName: item.name || null,
          rowIndex: span.rowIndex,
          columnIndex: span.columnIndex,
          colSpan: span.colSpan,
          rowSpan: span.rowSpan,
          repeatedHeader: Boolean(span.sourceRow?.isHeader),
          continuation: span.segStartY !== span.firstSegmentY,
        },
      });
    } else {
      recordLayoutItem(doc, {
        kind: 'tablixCell',
        itemName: null,
        tablixName: item.name || null,
        rowIndex: span.rowIndex,
        columnIndex: span.columnIndex,
        colSpan: span.colSpan,
        rowSpan: span.rowSpan,
        repeatedHeader: Boolean(span.sourceRow?.isHeader),
        continuation: span.segStartY !== span.firstSegmentY,
        zIndex: item.zIndex || 0,
        x: span.x,
        y: span.segStartY,
        width: span.width,
        height: segmentHeight,
        text: '',
        lines: [],
        writingMode: 'default',
        verticalAlign: 'top',
        padding: { top: 0, right: 0, bottom: 0, left: 0 },
        backgroundColor: styleColor(item.style?.backgroundColor, span.context, null),
        borders: resolvedTraceBorders(item.style, span.context, span.edges),
      });
    }
    // A nested data region can live in a row-header cell whose owning group spans several physical parent
    // rows. Those cells are rendered lazily by this open-span path, so drawing nested regions only from
    // drawRowContent silently produced an empty merged cell. Render the child once in the first page
    // segment that can contain it. If no segment can contain it, fail closed rather than dropping or
    // clipping child rows; synchronized pagination of a row-spanned parent and oversized child is a
    // different construct from the supported oversized child in an ordinary cell.
    if (!span.nestedDrawn && span.nestedTablixes.length > 0) {
      const requiredHeight = Math.max(...span.nestedTablixes.map((nested) => (
        (nested.item.top || 0) + nestedLayout(nested, span.width).height
      )));
      if (requiredHeight <= segmentHeight + 0.5) {
        for (const nested of span.nestedTablixes) drawNestedTablix(nested, span.x, span.segStartY, span.width);
        span.nestedDrawn = true;
      } else if (finalSegment) {
        throw new ServiceError(
          'UNSUPPORTED_FEATURE',
          `Row-spanned nested tablix in ${item.name || 'unnamed'} exceeds its printable page segment`,
        );
      }
    }
    drawEdges(span.x, span.segStartY, span.width, segmentHeight, span.edges);
  };

  // Close every span whose last spanned row has now been fully drawn, ending it at the current y.
  const closeSpansEndingAt = (rowIndex) => {
    for (const span of openSpans) if (span.endRowIndex <= rowIndex) drawSpanSegment(span, y, true);
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
      const cellContext = contextForCell(rowIndex, cell);
      const edges = resolveEdges({ cell, rowIndex }, columnIndex, span);
      // Merged (row-span) cells are drawn lazily via drawSpanSegment so they track the real extent of their
      // spanned rows across page splits; single-row cells draw here directly.
      if ((cell.rowSpan || 1) > 1) {
        openSpans.push({
          x,
          width,
          textbox,
          cell,
          text: texts[index] || '',
          context: cellContext,
          edges,
          segStartY: y,
          endRowIndex: rowIndex + (cell.rowSpan || 1) - 1,
          sourceRow: row,
          rowIndex,
          columnIndex,
          colSpan: span,
          rowSpan: cell.rowSpan || 1,
          firstSegmentY: y,
          nestedTablixes: cell.nestedTablixes || [],
          nestedDrawn: false,
        });
        continue;
      }
      const renderedHeight = Math.min(height, Math.max(0, pageBottom - y));
      if (renderedHeight <= 0) continue;
      // Recursive (parent/child) groups render expanded with the first cell indented by depth.
      const padLeft = index === 0 && row.indentLevel ? row.indentLevel * 12 : 0;
      // A canvas cell (an SSRS List / free-form layout: a Rectangle flattened into several positioned items
      // — multiple textboxes, lines, charts, images) is drawn item-by-item at each item's position, rather
      // than as one textbox filling the cell. Its nested tablixes are drawn by the loop below, so those are
      // excluded here. A plain data cell (one textbox) keeps the fast path unchanged.
      const cellItems = cell.items || [];
      if (isFreeFormCell(cell) && !cell.hidden) {
        const cellStyle = textbox?.style || item.style || {};
        const cellBackground = styleColor(cellStyle.backgroundColor, cellContext, null);
        if (cellBackground) doc.save().fillColor(cellBackground).rect(x, y, width, renderedHeight).fill().restore();
        recordLayoutItem(doc, {
          kind: 'tablixCell',
          itemName: null,
          tablixName: item.name || null,
          rowIndex,
          columnIndex,
          colSpan: span,
          rowSpan: 1,
          repeatedHeader: Boolean(row.isHeader),
          zIndex: item.zIndex || 0,
          x,
          y,
          width,
          height: renderedHeight,
          text: '',
          lines: [],
          writingMode: 'default',
          verticalAlign: 'top',
          padding: { top: 0, right: 0, bottom: 0, left: 0 },
          backgroundColor: cellBackground,
          borders: resolvedTraceBorders(cellBorderStyle(cell, item) || {}, cellContext, edges),
        });
        // Charts/tablixes inside a List cell are scoped to the group instance, so the region dataset resolves
        // to the current group's rows rather than the whole dataset.
        const canvasContext = item.datasetName
          ? { ...cellContext, datasets: { ...cellContext.datasets, [item.datasetName]: cellContext.dataset } }
          : cellContext;
        for (const canvasItem of [...cellItems].sort((left, right) => (left.zIndex || 0) - (right.zIndex || 0) || (left.top || 0) - (right.top || 0) || (left.left || 0) - (right.left || 0))) {
          if (canvasItem.type === 'Tablix' || canvasItem.type === 'Subreport') continue;
          drawSimpleItem(doc, config, model, canvasItem, x + (canvasItem.left || 0), y + (canvasItem.top || 0), canvasContext);
        }
      } else if (textbox && !cell.hidden) {
        drawTextbox(doc, config, textbox, x, y, cellContext, {
          width,
          height: renderedHeight,
          text: texts[index] || '',
          skipBorder: true,
          padLeft,
          traceEdges: edges,
          traceMeta: {
            kind: 'tablixCell',
            tablixName: item.name || null,
            rowIndex,
            columnIndex,
            colSpan: span,
            rowSpan: 1,
            repeatedHeader: Boolean(row.isHeader),
          },
        });
      } else {
        const style = textbox?.style || item.style || {};
        const borderStyle = cellBorderStyle(cell, item) || {};
        recordLayoutItem(doc, {
          kind: 'tablixCell',
          itemName: null,
          tablixName: item.name || null,
          rowIndex,
          columnIndex,
          colSpan: span,
          rowSpan: 1,
          repeatedHeader: Boolean(row.isHeader),
          zIndex: item.zIndex || 0,
          x,
          y,
          width,
          height: renderedHeight,
          text: '',
          lines: [],
          writingMode: 'default',
          verticalAlign: 'top',
          padding: { top: 0, right: 0, bottom: 0, left: 0 },
          backgroundColor: styleColor(style.backgroundColor, cellContext, null),
          borders: resolvedTraceBorders(borderStyle, cellContext, edges),
        });
      }
      for (const nested of cell.nestedTablixes || []) drawNestedTablix(nested, x, y, width);
      drawEdges(x, y, width, renderedHeight, edges);
    }
    y += height;
    addedHeight += height;
    // Only close spans on a fully-rendered row; a split row's head must keep its spans open for the tail.
    if (rowComplete) closeSpansEndingAt(rowIndex);
  };

  const startContinuationPage = (continuedRow = null) => {
    // End each open span at this page's content bottom, break, repeat the headers, then re-open the spans
    // just below the repeated headers so their value redraws at the top of the new page.
    // An open row-span proves that the same logical group crosses this page boundary, even when no single
    // textbox needed line-level splitting. Label that boundary just like a split physical row.
    const logicalContinuationRow = continuedRow || openSpans[0]?.sourceRow || null;
    for (const span of openSpans) drawSpanSegment(span, y);
    // Close and flush the actual table fragment before breaking so the final data row retains its bottom
    // border. The continuation annotation belongs only to the next page, above its repeated table header.
    closeOuterBorderFragment(y);
    addPage();
    y = addPage.bodyTop;
    if (showContinuationMarkers && logicalContinuationRow) drawContinuationMarker(CONTINUATION_MARKERS.fromPrevious, logicalContinuationRow);
    fragmentStartY = y;
    const repeatedHeaders = headerMeasurements();
    headers.forEach((header, index) => drawRowContent(header, repeatedHeaders.heights[index]));
    // Continue overflowing values from where they were clipped; repeat values that fully fit.
    for (const span of openSpans) {
      if (span.pendingTail) { span.text = span.pendingTail; span.pendingTail = null; }
      span.segStartY = y;
    }
  };

  const startFreshTablePage = () => {
    // No row from this tablix has been drawn yet, so there is no table fragment to close and no header to
    // replay. This is an initial-placement page move, not a continuation. Leaving `firstFragment` intact
    // also ensures the outer top border is emitted on the page where the table actually begins.
    addPage();
    y = addPage.bodyTop;
    fragmentStartY = y;
  };

  // Draw a List / canvas row: a single cell holding a Rectangle-flattened free-form layout that can be
  // taller than a page. Items (textboxes, lines, charts, images, nested tablixes) are placed by their
  // canvas-relative top and reflowed across pages — when the next item would not fit the current page, a
  // page is added and that item begins the new page. Each item is atomic (SSRS keeps a report item
  // together); a single item taller than a full page is drawn where it starts rather than looping forever.
  const drawCanvasRow = (row) => {
    const cell = row.cells[0];
    const rowIndex = rowIndexes.get(row);
    const columnIndex = placements[rowIndex][0];
    const span = cell.colSpan || 1;
    const { xOffsetPt, widthPt: width } = cellGeometryPt(columnWidths, columnIndex, span);
    const x = startX + xOffsetPt;
    const cellContext = contextForCell(rowIndex, cell);
    const canvasContext = item.datasetName
      ? { ...cellContext, datasets: { ...cellContext.datasets, [item.datasetName]: cellContext.dataset } }
      : cellContext;
    const cellBackground = styleColor((cellTextbox(cell)?.style || item.style || {}).backgroundColor, cellContext, null);
    const drawables = [];
    for (const canvasItem of cell.items || []) {
      if (canvasItem.type === 'Tablix' || canvasItem.type === 'Subreport') continue;
      const declaredHeight = canvasItem.height || 0;
      let actualHeight = declaredHeight;
      let drawHeight;
      // A CanGrow textbox on a canvas is sized by its content, exactly as it is inside a tablix cell.
      // Measuring it here is what lets the displacement below push its later peers clear of the extra
      // lines; drawing it at the declared height instead let long values run under the next item.
      if (canvasItem.type === 'Textbox' && canvasItem.canGrow && !isHidden(canvasItem.hidden, canvasContext)) {
        const text = textForItem(canvasItem, canvasContext);
        const padding = styleSize(canvasItem.style?.paddingTop, canvasContext, 2)
          + styleSize(canvasItem.style?.paddingBottom, canvasContext, 2);
        actualHeight = Math.max(
          declaredHeight,
          measureTextboxHeight(doc, config, canvasItem, canvasContext, text, canvasItem.width) + padding,
        );
        drawHeight = actualHeight;
      }
      drawables.push({
        top: canvasItem.top || 0,
        left: canvasItem.left || 0,
        width: canvasItem.width || 0,
        declaredHeight,
        actualHeight,
        draw: (screenY) => (drawHeight === undefined
          ? drawSimpleItem(doc, config, model, canvasItem, x + (canvasItem.left || 0), screenY, canvasContext)
          : drawTextbox(doc, config, canvasItem, x + (canvasItem.left || 0), screenY, canvasContext, { height: drawHeight })),
      });
    }
    for (const nested of cell.nestedTablixes || []) {
      const top = nested.item.top || 0;
      const nestedGeometry = nestedLayout(nested, width);
      drawables.push({
        top,
        left: nested.item.left || 0,
        width: nestedGeometry.width,
        declaredHeight: nested.item.height || 0,
        actualHeight: nestedGeometry.height,
        // drawNestedTablix positions at parentY + item.top, so pass parentY = screenY - item.top.
        draw: (screenY) => drawNestedTablix(nested, x, screenY - top, width),
        // A child data region is a flow, not an atomic report item: when it is taller than a printable
        // page SSRS continues its rows on the next page. Expose that so the canvas loop can paginate it
        // instead of drawing one over-long block past the body boundary.
        paginate: (screenY) => drawNestedTablixAcrossPages(nested, x, screenY, width),
      });
    }
    drawables.sort((left, right) => left.top - right.top || (left.declaredHeight - right.declaredHeight));
    // SSRS vertical displacement: an item that renders taller than its design height pushes the items
    // BELOW it down by that growth, preserving their declared gap. Only a peer that shares its horizontal
    // lane is pushed — two items side by side in disjoint left/right lanes stay side by side instead of
    // being serialized. This is the same rule the section and container layouts apply.
    for (const [index, drawable] of drawables.entries()) {
      drawable.resolvedTop = drawable.top;
      for (let previousIndex = 0; previousIndex < index; previousIndex += 1) {
        const previous = drawables[previousIndex];
        const previousDesignBottom = previous.top + previous.declaredHeight;
        if (previousDesignBottom > drawable.top + COINCIDENT_EDGE_TOLERANCE_PT) continue;
        if (!canvasLanesOverlap(previous, drawable)) continue;
        const declaredGap = Math.max(0, drawable.top - previousDesignBottom);
        drawable.resolvedTop = Math.max(
          drawable.resolvedTop,
          previous.resolvedTop + previous.actualHeight + declaredGap,
        );
      }
    }
    let windowStart = 0;
    let pageTop = y;
    let windowBottom = windowStart + (pageBottom - pageTop);
    let cursor = y;
    for (const drawable of drawables) {
      const displacedTop = drawable.resolvedTop;
      const itemBottom = displacedTop + drawable.actualHeight;
      const startsBelowWindow = displacedTop >= windowBottom - 0.5;
      const overflowsWindow = itemBottom > windowBottom + 0.5;
      const canFitFreshPage = drawable.actualHeight <= (pageBottom - addPage.bodyTop) + 0.5;
      if ((startsBelowWindow || overflowsWindow) && canFitFreshPage && cursor > addPage.bodyTop + 0.5) {
        addPage();
        windowStart = displacedTop;
        pageTop = addPage.bodyTop;
        windowBottom = windowStart + (pageBottom - pageTop);
      }
      const screenY = pageTop + (displacedTop - windowStart);
      if (!canFitFreshPage && drawable.paginate) {
        // The region is taller than an empty page, so it continues across pages by rows. Re-anchor the
        // canvas window on the page and offset the last fragment ended at, so the following canvas items
        // keep their design gaps relative to the region's real bottom instead of its off-page one.
        const endY = drawable.paginate(screenY);
        windowStart = displacedTop + drawable.actualHeight;
        pageTop = endY;
        windowBottom = windowStart + (pageBottom - pageTop);
        cursor = endY;
        continue;
      }
      if (cellBackground) doc.save().fillColor(cellBackground).rect(x, screenY, width, Math.max(0, drawable.actualHeight)).fill().restore();
      drawable.draw(screenY);
      cursor = Math.max(cursor, screenY + drawable.actualHeight);
    }
    y = cursor;
  };

  const drawRow = (row) => {
    // A row-group page break starts the group's first row on a fresh page (unless we're already at the
    // top of a page). Reuses the continuation-page machinery so repeated headers redraw.
    if (row.pageBreakBefore && y > addPage.bodyTop) startContinuationPage();
    // A List / canvas row (one cell of free-form positioned items) reflows across pages on its own path.
    if (row.cells.length === 1 && isCanvasCell(row.cells[0])) {
      drawCanvasRow(row);
      return;
    }
    let remainingTexts = row.cells.map((cell) => cellText(cell));
    let measured = measureRow(row, remainingTexts);
    const repeatedHeaderHeight = headerMeasurements().total;
    const freshPageCapacity = pageBottom - addPage.bodyTop - repeatedHeaderHeight;
    // KeepTogether is best-effort: when a physical data row does not fit in the current remainder, move it
    // before splitting any cell text. If an oversized row is already at the fresh-page content boundary, it
    // must split there; attempting another break would create an endless blank-page loop. Repeating headers
    // use their existing pagination path and must not recursively request a continuation page themselves.
    const atFreshContentStart = y <= addPage.bodyTop + repeatedHeaderHeight + 0.5;
    const hasNestedTablix = row.cells.some((cell) => (cell.nestedTablixes || []).length > 0);
    // A nested data region is a complete child grid. Move the containing row intact when it fits on a fresh
    // page; splitting it as if it were plain textbox characters would lose child rows and shared borders.
    // If the complete child grid is itself taller than a page, fail closed until recursive child pagination
    // can preserve both grids rather than emitting a misleading partial table.
    if (hasNestedTablix && y + measured > pageBottom) {
      if (!atFreshContentStart) {
        startContinuationPage();
        measured = measureRow(row, remainingTexts);
      }
      if (y + measured > pageBottom) {
        const nestedOwners = row.cells.flatMap((cell, cellIndex) => (cell.nestedTablixes || [])
          .map((nested) => ({ cell, cellIndex, nested })));
        const oversized = nestedOwners.filter(({ nested, cellIndex }) => (
          (nested.item.top || 0) + nestedLayout(
            nested,
            layoutsForRow(row)[cellIndex]?.width || totalWidth,
          ).height > freshPageCapacity
        ));
        // A bundled subreport is an independent child data region and may legitimately exceed one parent
        // page. Split its physical child rows across parent page fragments. Multiple independently tall
        // children in one parent row require synchronized grids and remain fail-closed.
        if (oversized.length !== 1 || !oversized[0].nested.subreport) {
          throw new ServiceError('UNSUPPORTED_FEATURE', `Nested tablix in ${item.name || 'unnamed'} exceeds one printable page`);
        }
        const owner = oversized[0];
        const ownerLayout = layoutsForRow(row)[owner.cellIndex];
        const nested = owner.nested;
        const originalRows = nested.rows;
        const leadingHeaders = [];
        for (const nestedRow of originalRows) {
          if (!nestedRow.isHeader) break;
          leadingHeaders.push(nestedRow);
        }
        const dataRows = originalRows.slice(leadingHeaders.length);
        const fullLayout = nestedLayout(nested, ownerLayout.width);
        const heightByRow = new Map(originalRows.map((nestedRow, index) => [nestedRow, fullLayout.heights[index]]));
        const headerHeight = leadingHeaders.reduce((sum, nestedRow) => sum + (heightByRow.get(nestedRow) || 0), 0);
        let offset = 0;
        let firstChildFragment = true;
        try {
          while (offset < dataRows.length || (dataRows.length === 0 && firstChildFragment)) {
            const available = pageBottom - y;
            const childBudget = Math.max(1, available - (nested.item.top || 0) - headerHeight);
            let end = offset;
            let used = 0;
            while (end < dataRows.length) {
              const next = heightByRow.get(dataRows[end]) || dataRows[end].height || 0;
              if (end > offset && used + next > childBudget) break;
              used += next;
              end += 1;
              if (used >= childBudget) break;
            }
            if (end === offset && offset < dataRows.length) end += 1;
            nested.rows = [...leadingHeaders, ...dataRows.slice(offset, end)];
            const fragmentTexts = firstChildFragment ? remainingTexts : remainingTexts.map(() => '');
            const fragmentHeight = measureRow(row, fragmentTexts);
            if (fragmentHeight > available + 0.5) {
              throw new ServiceError(
                'UNSUPPORTED_FEATURE',
                `One nested subreport row in ${item.name || 'unnamed'} exceeds one printable page`,
              );
            }
            const complete = end >= dataRows.length;
            drawRowContent(row, fragmentHeight, fragmentTexts, complete);
            offset = end;
            firstChildFragment = false;
            if (!complete) startContinuationPage(row);
            if (dataRows.length === 0) break;
          }
        } finally {
          nested.rows = originalRows;
        }
        return;
      }
      drawRowContent(row, measured, remainingTexts);
      return;
    }
    if (!row.isHeader && row.keepTogether && y + measured > pageBottom && !atFreshContentStart) {
      startContinuationPage();
      measured = measureRow(row, remainingTexts);
    }
    // A vertical merge is not inherently KeepTogether. When its member allows splitting, activeSpans closes
    // the merge at the body boundary and redraws/reopens it after the repeated header on the next page. This
    // matches SSRS and lets the current page consume every physical row that fits. Reserving the complete
    // span is appropriate only when the owning tablix member explicitly declares KeepTogether.
    if (row.keepSpanTogether) {
      const rowIndex = rowIndexes.get(row);
      let protectedHeight;
      if (config.pdfLayoutOptimizations === false) {
        protectedHeight = row.cells.reduce((maximum, cell) => Math.max(
          maximum,
          measuredHeights.slice(rowIndex, rowIndex + Math.max(1, cell.rowSpan || 1)).reduce((sum, value) => sum + value, 0),
        ), measured);
        statistics.rowSpanHeightCalculations += row.cells.length;
      } else {
        const maximumRowSpan = row.cells.reduce((maximum, cell) => Math.max(maximum, Math.max(1, cell.rowSpan || 1)), 1);
        // Row heights are non-negative. Therefore the longest declared span always has the largest protected
        // height. Sum that range once in the same top-to-bottom order as the previous per-cell reductions.
        let spannedHeight = 0;
        const spanEnd = Math.min(measuredHeights.length, rowIndex + maximumRowSpan);
        for (let spanIndex = rowIndex; spanIndex < spanEnd; spanIndex += 1) spannedHeight += measuredHeights[spanIndex];
        statistics.rowSpanHeightCalculations += 1;
        protectedHeight = Math.max(measured, spannedHeight);
      }
      if (y + protectedHeight > pageBottom && protectedHeight <= freshPageCapacity) startContinuationPage();
    }
    measured = measureRow(row, remainingTexts);
    if (y + measured <= pageBottom) {
      drawRowContent(row, measured, remainingTexts);
      return;
    }

    // The row does not fit and the split loop below only advances rows that have splittable text. A row
    // whose cells carry no continuation-able text (e.g. the trailing row of a row-span group, whose columns
    // are covered by the merged header) would otherwise fall through the loop undrawn: its height is never
    // consumed and its open row-span stays open, so the header's residual later closes against the *next*
    // group's cursor and paints on top of it (two same-named cells at one origin — unrepresentable in Word).
    // Move the whole row to a fresh page so it draws in order and its span closes against its own row.
    if (!remainingTexts.some(Boolean) && !atFreshContentStart && measured <= freshPageCapacity) {
      startContinuationPage();
      drawRowContent(row, measureRow(row, remainingTexts), remainingTexts);
      return;
    }

    while (remainingTexts.some(Boolean)) {
      if (pageBottom - y < Math.max(12, row.height)) startContinuationPage();
      const availableHeight = pageBottom - y;
      const layouts = layoutsForRow(row, remainingTexts);
      // Row-span (merged) cells are continued only by activeSpans/redrawActiveSpans, which re-draws the whole
      // value at the top of each continuation page (SSRS merged-cell behaviour). They must never also produce a
      // split tail here, or the tail and the redrawn value overlap on the continuation page.
      const parts = layouts.map((layout) => {
        if ((layout.cell.rowSpan || 1) > 1) return { head: layout.text, tail: '' };
        statistics.textSplitRequests += 1;
        return splitTextForHeight(doc, config, layout.textbox, layout.context, layout.text, layout.width, availableHeight);
      });
      const hasContinuation = parts.some((part) => part.tail.length > 0);
      const heads = parts.map((part) => part.head);
      const segmentHeight = hasContinuation ? availableHeight : Math.min(availableHeight, measureRow(row, heads));
      drawRowContent(row, segmentHeight, heads, !hasContinuation);
      remainingTexts = parts.map((part) => part.tail);
      if (hasContinuation) {
        startContinuationPage(row);
      }
    }
  };

  const drawingStartedAt = performance.now();
  if (leadingHeaders.length > 0 && y > addPage.bodyTop + 0.5) {
    const headerHeight = leadingHeaders.reduce((total, row) => total + measureRow(row), 0);
    const firstBodyRow = rows[leadingHeaders.length] || null;
    const firstBodyHeight = firstBodyRow ? measureRow(firstBodyRow) : 0;
    const freshPageHeight = pageBottom - addPage.bodyTop;
    // SSRS keeps a leading header with the first body row when the complete unit can fit on a fresh page.
    // If the first body row is intrinsically oversized, keep only the header block intact and allow the
    // existing row-splitting path to handle that body row after it starts.
    const preferredInitialHeight = headerHeight + firstBodyHeight;
    const requiredInitialHeight = firstBodyRow && preferredInitialHeight <= freshPageHeight + 0.5
      ? preferredInitialHeight
      : headerHeight;
    if (y + requiredInitialHeight > pageBottom + 0.5 && requiredInitialHeight <= freshPageHeight + 0.5) {
      startFreshTablePage();
    }
  }
  for (const row of rows) drawRow(row);
  for (const span of openSpans) drawSpanSegment(span, y); // flush any spans still open at the tablix end
  openSpans = [];
  closeOuterBorderFragment(y);
  statistics.tablixDrawingMs += performance.now() - drawingStartedAt;
  return { height: Math.max(item.height, addedHeight), endY: y };
}

export async function renderPdf(model, request, config, options = {}) {
  // Timing callbacks are observation-only. They receive bounded counts/sizes and must never inspect or
  // mutate PDFKit state, report content, coordinates, or the canonical trace.
  const reportTelemetry = (phase, metrics = {}) => {
    try { options.telemetry?.(phase, metrics); } catch { /* Telemetry cannot affect canonical PDF output. */ }
  };
  borderWidthFloor = config?.borderWidthFloorPt || 0;
  const doc = new PDFDocument({ autoFirstPage: false, bufferPages: true, compress: true, info: { Title: request.outputFileName || model.name, Producer: 'RDL Converter Service' } });
  const layoutTrace = options.captureLayoutTrace ? createLayoutTrace(model, request) : null;
  attachLayoutTrace(doc, layoutTrace);
  const completion = collectDocument(doc);
  const page = model.page;
  const headerHeight = page.header?.height || 0;
  const footerHeight = page.footer?.height || 0;
  const bodyTop = page.marginTop + headerHeight;
  const pageBottom = page.height - page.marginBottom - footerHeight;
  const globals = { PageNumber: 0, TotalPages: 1, ReportName: request.outputFileName || model.name, ExecutionTime: new Date(), variables: model.variables || {}, culture: resolveReportCulture(model, { parameters: request.parameters || {} }) };
  const addPage = () => {
    doc.addPage({ size: [page.width, page.height], margins: { top: 0, right: 0, bottom: 0, left: 0 } });
    globals.PageNumber += 1;
    beginLayoutTracePage(doc, {
      width: page.width,
      height: page.height,
      marginTop: page.marginTop,
      marginRight: page.marginRight,
      marginBottom: page.marginBottom,
      marginLeft: page.marginLeft,
      headerHeight,
      footerHeight,
      bodyTop,
      bodyBottom: pageBottom,
    });
  };
  addPage.bodyTop = bodyTop;
  addPage();
  const switchBufferedPage = (pageNumber) => {
    const range = doc.bufferedPageRange();
    const pageIndex = Number(pageNumber) - 1;
    if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= range.count) {
      throw new ServiceError('INTERNAL_ERROR', 'Buffered PDF page navigation became inconsistent', 500);
    }
    doc.switchToPage(range.start + pageIndex);
    selectLayoutTracePage(doc, pageIndex);
    globals.PageNumber = pageIndex + 1;
  };

  const datasets = normalizeDatasets(model, request);
  const statistics = {
    optimizationsEnabled: config.pdfLayoutOptimizations !== false,
    tablixCount: 0,
    tablixRowCount: 0,
    tablixCellCount: 0,
    tablixMaterializationMs: 0,
    tablixSetupMs: 0,
    tablixInitialMeasurementMs: 0,
    tablixDrawingMs: 0,
    rowMeasurementRequests: 0,
    rowMeasurementsComputed: 0,
    rowMeasurementCacheHits: 0,
    headerMeasurementsComputed: 0,
    headerMeasurementCacheHits: 0,
    rowSpanHeightCalculations: 0,
    textSplitRequests: 0,
    borderEdgesCollected: 0,
    borderRunsDrawn: 0,
  };
  reportTelemetry('pdf.initialized', {
    bodyItemCount: model.body?.items?.length || 0,
    normalizedDatasetCount: Object.keys(datasets || {}).length,
    captureLayoutTrace: Boolean(layoutTrace),
  });
  const renderBodyItem = (item, x, y, context, pagination = {}) => {
    const pageAdvance = pagination.addPage || addPage;
    if (isHidden(item.hidden, context)) return { endY: y };
    if (item.type === 'Tablix') {
      return renderTablix({
        doc,
        config,
        model,
        item,
        request,
        startX: x,
        startY: y,
        pageBottom,
        addPage: pageAdvance,
        globals,
        statistics,
      });
    }
    if (item.type === 'Textbox' && item.canGrow) {
      let remaining = textForItem(item, context);
      let currentY = y;
      let firstSegment = true;
      const verticalPadding = styleSize(item.style?.paddingTop, context, 2)
        + styleSize(item.style?.paddingBottom, context, 2);
      // A growable textbox whose *text* outgrows the page remainder is a flowing block, not an atomic unit
      // like a tablix row: it can be most of a page tall. Deferring such a block whenever it merely fits on
      // an empty page discards the entire remainder, so a block one line too tall leaves a blank band its
      // own height. It instead fills the remainder and continues on the next page, which is what
      // splitTextForHeight below already implements, unless the remainder is below the orphan minimum.
      //
      // The RDL KeepTogether flag is deliberately not a veto here. It is defined as best-effort ("keep on
      // one page if possible"), and designers emit it on essentially every textbox, so honouring it
      // absolutely would stop any long textbox in any report from ever crossing a page boundary. Atomic
      // keep-together remains enforced where the unit really is indivisible: tablix rows.
      const orphanHeight = verticalPadding
        + MINIMUM_FLOWED_TEXT_LINES * lineHeightForStyle(doc, config, item.style, context);
      while (true) {
        const measured = measureTextboxHeight(doc, config, item, context, remaining, item.width)
          + verticalPadding;
        const desiredHeight = Math.max(firstSegment ? item.height : 0, measured);
        const available = pageBottom - currentY;
        if (desiredHeight <= available + 0.5) {
          drawTextbox(doc, config, item, x, currentY, context, {
            width: item.width,
            height: desiredHeight,
            text: remaining,
          });
          return { height: currentY + desiredHeight - y, endY: currentY + desiredHeight };
        }
        // Only the declared height overflows: the text itself still fits. The declared height is a minimum
        // reservation rather than content, so there is nothing to flow — splitting here would silently
        // squash the box into the remainder. Move the reserved block to a fresh page as a unit instead.
        if (measured <= available + 0.5 && currentY > bodyTop + 0.5) {
          pageAdvance();
          currentY = bodyTop;
          firstSegment = false;
          continue;
        }
        if (available < orphanHeight && currentY > bodyTop + 0.5) {
          pageAdvance();
          currentY = bodyTop;
          firstSegment = false;
          continue;
        }
        const split = splitTextForHeight(
          doc,
          config,
          item,
          context,
          remaining,
          item.width,
          available,
        );
        if (!split.head && split.tail) {
          pageAdvance();
          currentY = bodyTop;
          firstSegment = false;
          continue;
        }
        drawTextbox(doc, config, item, x, currentY, context, {
          width: item.width,
          height: available,
          text: split.head,
        });
        if (!split.tail) return { height: currentY + available - y, endY: currentY + available };
        remaining = split.tail;
        pageAdvance();
        currentY = bodyTop;
        firstSegment = false;
      }
    }
    if (item.type === 'Rectangle') {
      const backgroundColor = styleColor(item.style.backgroundColor, context, null);
      const rectangleTrace = recordLayoutItem(doc, {
        kind: 'rectangle',
        itemName: item.name || null,
        zIndex: item.zIndex || 0,
        x,
        y,
        width: item.width,
        height: item.height,
        backgroundColor,
        borders: resolvedTraceBorders(item.style, context),
      });
      if (backgroundColor) doc.save().fillColor(backgroundColor).rect(x, y, item.width, item.height).fill().restore();
      // The border is deferred to the end of this branch, after the child bands are laid out, so its
      // bottom edge tracks the container's *rendered* height. A flow child such as a Tablix with a CanGrow
      // cell can grow past the rectangle's declared height; drawing the border here at item.height would
      // strand it above the grown content and paint a second, higher bottom line beneath the coincident
      // tablix edge (the "double line" defect). A bordered/filled rectangle is guarded against page
      // fragmentation below, so the grown extent stays on one page and endY - y is a valid single-page
      // height.
      const visibleChildren = (item.items || []).filter((child) => !isHidden(child.hidden, context));
      const bands = containerLayoutBands(
        visibleChildren,
        (child) => !/^None$/i.test(activeBreakLocation(child, context)),
      );
      // Fragmenting the container is only safe when nothing paints its extent: a fill or border would be
      // drawn once, at the design height, on the page the container started on. The same guard already
      // covers content-driven fragmentation below; a declared break reaches it through this helper.
      const refuseFragmentationWithVisibleExtent = () => {
        const hasVisibleBorder = Object.values(resolvedTraceBorders(item.style, context)).some(Boolean);
        if (!backgroundColor && !hasVisibleBorder) return;
        throw new ServiceError(
          'UNSUPPORTED_FEATURE',
          'A page-spanning rectangle with a visible fill or border cannot be safely fragmented',
          422,
          { item: item.name || null },
        );
      };
      let endY = y;
      let previousDesignBottom = 0;
      let hasRenderedBand = false;
      // Carried across bands, and out of this container, exactly like the body-level flow: an End break is
      // owned by the item that declares it but is spent on whatever comes next, which may be a later band,
      // a following sibling, or a sibling of an ancestor.
      let pendingBreak = false;
      for (const band of bands) {
        const bandBreak = band.isolated ? activeBreakLocation(band.items[0], context) : 'None';
        const gap = hasRenderedBand ? Math.max(0, band.top - previousDesignBottom) : band.top;
        let bandY = endY + gap;
        const bandHeight = band.designBottom - band.top;
        const fixedBand = band.items.every(isFixedCoordinateItem);
        const atPageTop = bandY <= bodyTop + COINCIDENT_EDGE_TOLERANCE_PT;
        if (pendingBreak || (breaksBeforeItem(bandBreak) && !atPageTop)) {
          refuseFragmentationWithVisibleExtent();
          pageAdvance();
          bandY = bodyTop;
          pendingBreak = false;
        } else if (bandY >= pageBottom || (
          fixedBand
          && bandY + bandHeight > pageBottom
          // Same rule as the body-level flow: a band taller than an empty page gains nothing from being
          // deferred, so it must not consume the remainder of the page it is standing on.
          && bandHeight <= (pageBottom - bodyTop) + COINCIDENT_EDGE_TOLERANCE_PT
          && endY > bodyTop + COINCIDENT_EDGE_TOLERANCE_PT
        )) {
          pageAdvance();
          bandY = bodyTop;
        }
        let bandEndY = bandY;
        // A child container can end with a break of its own; the break belongs to the enclosing flow, so it
        // surfaces here and is spent on the next band or handed further out.
        let bandPropagatedBreak = false;
        const orderedBandItems = [...band.items].sort((left, right) => (
          (left.zIndex || 0) - (right.zIndex || 0)
          || (left.top || 0) - (right.top || 0)
          || (left.left || 0) - (right.left || 0)
        ));
        if (orderedBandItems.length > 1 && !fixedBand) {
          const startPage = globals.PageNumber;
          const state = { lastPage: startPage };
          const childEnds = [];
          for (const child of orderedBandItems) {
            switchBufferedPage(startPage);
            const synchronizedAdvance = () => {
              const nextPage = globals.PageNumber + 1;
              if (nextPage <= state.lastPage) {
                switchBufferedPage(nextPage);
              } else {
                pageAdvance();
                state.lastPage = Math.max(state.lastPage, globals.PageNumber);
              }
            };
            synchronizedAdvance.bodyTop = pageAdvance.bodyTop ?? bodyTop;
            const childY = bandY + (child.top || 0) - band.top;
            const rendered = renderBodyItem(
              child,
              x + (child.left || 0),
              childY,
              context,
              { addPage: synchronizedAdvance },
            );
            state.lastPage = Math.max(state.lastPage, globals.PageNumber);
            bandPropagatedBreak = bandPropagatedBreak || Boolean(rendered.pageBreakAfter);
            childEnds.push({
              page: globals.PageNumber,
              endY: rendered.endY ?? childY,
            });
          }
          if (state.lastPage > startPage) refuseFragmentationWithVisibleExtent();
          switchBufferedPage(state.lastPage);
          bandEndY = Math.max(
            bandY,
            ...childEnds
              .filter((entry) => entry.page === state.lastPage)
              .map((entry) => entry.endY),
          );
        } else {
          for (const child of orderedBandItems) {
            const childY = bandY + (child.top || 0) - band.top;
            const pageNumberBeforeChild = globals.PageNumber;
            const rendered = renderBodyItem(
              child,
              x + (child.left || 0),
              childY,
              context,
              { addPage: pageAdvance },
            );
            const renderedEndY = rendered.endY ?? childY;
            bandPropagatedBreak = bandPropagatedBreak || Boolean(rendered.pageBreakAfter);
            // Y coordinates are page-local. Once a single-child band advances to another page, its
            // final-page endpoint replaces the prior-page band coordinate; comparing the two with
            // Math.max would retain a stale Y and can push otherwise fitting later bands forward.
            bandEndY = globals.PageNumber === pageNumberBeforeChild
              ? Math.max(bandEndY, renderedEndY)
              : renderedEndY;
          }
        }
        endY = bandEndY;
        previousDesignBottom = band.designBottom;
        hasRenderedBand = true;
        if (breaksAfterItem(bandBreak) || bandPropagatedBreak) pendingBreak = true;
      }
      if (hasRenderedBand) {
        endY += Math.max(0, item.height - previousDesignBottom);
      }
      const renderedHeight = Math.max(item.height, endY - y);
      // Correct the recorded geometry and draw the border/box at the rendered height so PDF, and the
      // trace-driven editable DOCX built from it, keep the container's bottom edge coincident with grown
      // flow content instead of doubling it at the declared height.
      if (rectangleTrace) rectangleTrace.height = Math.round(renderedHeight * 1000) / 1000;
      drawBorder(doc, x, y, item.width, renderedHeight, item.style, context);
      return {
        height: renderedHeight,
        endY: hasRenderedBand ? endY : y + item.height,
        pageBreakAfter: pendingBreak,
      };
    }
    drawSimpleItem(doc, config, model, item, x, y, context);
    return { height: item.height, endY: y + item.height };
  };
  const items = [...model.body.items].sort((left, right) => left.top - right.top || left.left - right.left || left.zIndex - right.zIndex);
  // RDL body items are coordinate-positioned. Items with the same Top are peers in one horizontal
  // design band, not consecutive flow blocks. Keep pagination-capable items on the established flow
  // path: rendering two independently paginating objects beside each other needs synchronized page
  // fragments, while fixed items can safely share their declared page-relative Y coordinate.
  const isFixedCoordinateItem = (item) => {
    if (item.type === 'Tablix' || (item.type === 'Textbox' && item.canGrow)) return false;
    // A declared break makes the item move the page cursor, so it is a flow participant however simple its
    // own content is. Without this a break nested in a textbox-only rectangle keeps the fixed-coordinate
    // fast path and its page advance is measured against a stale band origin.
    if (declaresPageBreak(item)) return false;
    if (item.type === 'Rectangle') return (item.items || []).every(isFixedCoordinateItem);
    return true;
  };
  const sameDesignTop = (left, right) => Math.abs((left.top || 0) - (right.top || 0)) <= 0.25;
  const horizontallyDisjoint = (left, right) => (
    (left.left || 0) + (left.width || 0) <= (right.left || 0) + COINCIDENT_EDGE_TOLERANCE_PT
    || (right.left || 0) + (right.width || 0) <= (left.left || 0) + COINCIDENT_EDGE_TOLERANCE_PT
  );
  // A subtree that declares a break anywhere is its own flow unit at body level too: it must not be banded
  // with a coordinate peer, because the peer would be drawn before the nested break moved the page.
  const carriesPageBreak = (item, context) => (
    !/^None$/i.test(activeBreakLocation(item, context)) || declaresPageBreak(item)
  );
  let cursorY = bodyTop;
  let previousDesignBottom = 0;
  let pageHasContent = false;
  let forcePageBreak = false;
  for (let itemIndex = 0; itemIndex < items.length;) {
    const item = items[itemIndex];
    const context = { parameters: request.parameters || {}, globals, datasets, dataset: [], fields: {} };
    if (isHidden(item.hidden, context)) {
      itemIndex += 1;
      continue;
    }
    const breakLocation = activeBreakLocation(item, context);
    let band = [item];
    if (!carriesPageBreak(item, context)) {
      let nextIndex = itemIndex + 1;
      let designBottom = (item.top || 0) + (item.height || 0);
      while (nextIndex < items.length) {
        const candidate = items[nextIndex];
        if (isHidden(candidate.hidden, context)) {
          nextIndex += 1;
          continue;
        }
        if (carriesPageBreak(candidate, context)) break;
        const coincidentTop = sameDesignTop(item, candidate);
        const overlapsVertically = (candidate.top || 0) < designBottom - COINCIDENT_EDGE_TOLERANCE_PT;
        // Different horizontal lanes are coordinate peers while their declared vertical intervals
        // overlap. Items in the same lane remain sequential so a growing tablix still displaces its
        // following label/legend. Exact-top items retain the established z-layer behaviour.
        const independentLane = band.every((peer) => horizontallyDisjoint(peer, candidate));
        if (!coincidentTop && !(overlapsVertically && independentLane)) break;
        band.push(candidate);
        designBottom = Math.max(designBottom, (candidate.top || 0) + (candidate.height || 0));
        nextIndex += 1;
      }
      itemIndex = nextIndex;
    } else {
      itemIndex += 1;
    }
    if (forcePageBreak || (breaksBeforeItem(breakLocation) && pageHasContent)) {
      addPage();
      cursorY = bodyTop;
      pageHasContent = false;
      forcePageBreak = false;
    }
    // Body coordinates are relative to the body's printable origin. Preserve the first visible band's
    // declared Top on the initial report page so a designed spacer between the page header and body is not
    // collapsed. A page break at Start deliberately establishes a new page-local origin, and continuation
    // or overflow pages also resume at bodyTop, so neither consumes the report's absolute leading offset.
    const initialBodyGap = globals.PageNumber === 1
      && !pageHasContent
      && !breaksBeforeItem(breakLocation)
      ? Math.max(0, item.top || 0)
      : 0;
    const gap = pageHasContent
      ? Math.max(0, item.top - previousDesignBottom)
      : initialBodyGap;
    let y = cursorY + gap;
    const bandBottomOffset = Math.max(...band.map((candidate) => (
      candidate.top - item.top + candidate.height
    )));
    // Deferring a band to a fresh page is only worth the remainder it throws away when the band actually
    // fits on a fresh page. A band taller than the whole printable body cannot be rescued by moving it: it
    // will overflow wherever it starts, so deferring costs the entire remainder of this page and leaves a
    // blank band its own height. Such a band starts where it stands and fragments across pages, exactly as
    // the growable-textbox flow above already reasons about its own oversized blocks.
    const freshPageCapacity = pageBottom - bodyTop;
    if (y >= pageBottom || (
      band.every((candidate) => candidate.type !== 'Tablix')
      && y + bandBottomOffset > pageBottom
      && bandBottomOffset <= freshPageCapacity + COINCIDENT_EDGE_TOLERANCE_PT
      && pageHasContent
    )) {
      addPage();
      y = bodyTop;
      pageHasContent = false;
    }
    let bandEndY = y;
    // A break declared on the last child of a container is owned by the enclosing flow: the container
    // reports it here and the next body band spends it.
    let propagatedBreak = false;
    const fixedBand = band.every(isFixedCoordinateItem);
    if (band.length > 1 && !fixedBand) {
      const startPage = globals.PageNumber;
      const state = { lastPage: startPage };
      const childEnds = [];
      const orderedBandItems = [...band].sort((left, right) => (
        (left.zIndex || 0) - (right.zIndex || 0)
        || (left.top || 0) - (right.top || 0)
        || (left.left || 0) - (right.left || 0)
      ));
      for (const candidate of orderedBandItems) {
        switchBufferedPage(startPage);
        const synchronizedAdvance = () => {
          const nextPage = globals.PageNumber + 1;
          if (nextPage <= state.lastPage) switchBufferedPage(nextPage);
          else {
            addPage();
            state.lastPage = Math.max(state.lastPage, globals.PageNumber);
          }
        };
        synchronizedAdvance.bodyTop = bodyTop;
        const candidateY = y + candidate.top - item.top;
        const x = page.marginLeft + candidate.left;
        const rendered = renderBodyItem(candidate, x, candidateY, context, {
          addPage: synchronizedAdvance,
        });
        state.lastPage = Math.max(state.lastPage, globals.PageNumber);
        propagatedBreak = propagatedBreak || Boolean(rendered.pageBreakAfter);
        childEnds.push({
          page: globals.PageNumber,
          endY: rendered.endY ?? candidateY,
        });
      }
      switchBufferedPage(state.lastPage);
      bandEndY = Math.max(
        state.lastPage === startPage ? y : bodyTop,
        ...childEnds
          .filter((entry) => entry.page === state.lastPage)
          .map((entry) => entry.endY),
      );
    } else {
      for (const candidate of band) {
        const candidateY = y + candidate.top - item.top;
        const x = page.marginLeft + candidate.left;
        const rendered = renderBodyItem(candidate, x, candidateY, context);
        propagatedBreak = propagatedBreak || Boolean(rendered.pageBreakAfter);
        bandEndY = Math.max(bandEndY, rendered.endY);
      }
    }
    previousDesignBottom = Math.max(
      previousDesignBottom,
      ...band.map((candidate) => candidate.top + candidate.height),
    );
    cursorY = bandEndY;
    pageHasContent = true;
    if (breaksAfterItem(breakLocation) || propagatedBreak) {
      forcePageBreak = true;
    }
  }

  reportTelemetry('pdf.body-layout-completed', {
    pageCount: doc.bufferedPageRange().count,
    bodyItemCount: items.length,
    ...statistics,
  });

  const range = doc.bufferedPageRange();
  globals.TotalPages = range.count;
  for (let index = 0; index < range.count; index += 1) {
    doc.switchToPage(range.start + index);
    selectLayoutTracePage(doc, index);
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
  reportTelemetry('pdf.page-bands-completed', {
    pageCount: range.count,
    headerItemCount: page.header?.items?.length || 0,
    footerItemCount: page.footer?.items?.length || 0,
  });
  doc.end();
  const buffer = await completion;
  reportTelemetry('pdf.serialized', { pageCount: range.count, outputBytes: buffer.length });
  const parsed = await PdfLibDocument.load(buffer);
  const parsedPageCount = parsed.getPageCount();
  reportTelemetry('pdf.validated', { pageCount: parsedPageCount, outputBytes: buffer.length });
  return {
    buffer,
    pageCount: parsedPageCount,
    mimeType: 'application/pdf',
    extension: 'pdf',
    ...(layoutTrace ? { layoutTrace: finalizeLayoutTrace(layoutTrace) } : {}),
  };
}
