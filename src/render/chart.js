import { evaluateExpression } from '../rdl/expression.js';
import { pdfFont } from './fonts.js';
import { color as resolveColor, isHidden, styleColor, styleSize, styleText, styleValue } from './common.js';

const AXIS_COLOR = '#d9d9d9';
const TICK_LABEL_COLOR = '#595959';
const LABEL_COLOR = '#000000';
// SSRS shape-chart callout geometry: the radial stub and the horizontal elbow are each this fraction of
// the shape radius, and the label starts this far past the elbow.
const PIE_CALLOUT_SEGMENT_RATIO = 0.2;
const PIE_CALLOUT_TEXT_GAP_PT = 3.5;
// A bisector pointing straight up or down has a cosine of ±1e-16; treat that as pointing right.
const COINCIDENT_ANGLE_TOLERANCE = 1e-9;

// Rounds an axis maximum up to a readable value and picks a tick interval, matching how SSRS lays out
// a numeric value axis (0-based, "nice" 1/2/5·10ⁿ steps).
function niceScale(maxValue) {
  if (!(maxValue > 0)) return { max: 1, interval: 1 };
  const rough = maxValue / 5;
  const magnitude = 10 ** Math.floor(Math.log10(rough));
  const normalized = rough / magnitude;
  const step = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  const interval = step * magnitude;
  return { max: Math.ceil(maxValue / interval) * interval, interval };
}

// A tick value is an accumulated multiple of the interval, so binary floating point leaves noise
// (3 * 0.2 === 0.6000000000000001). Rendered verbatim in the narrow axis-label gutter that long string
// wraps to several lines. Rounding to the interval's own decimal precision yields the clean label SSRS
// shows ("0.6"), and String(Number(...)) drops any trailing zero ("1.0" -> "1").
function formatTick(value, interval) {
  const decimals = Math.min(10, Math.max(0, -Math.floor(Math.log10(interval || 1))));
  return String(Number(value.toFixed(decimals)));
}

function setFont(doc, config, { size = 8, bold = false, italic = false, family = 'Arial' } = {}) {
  doc.font(pdfFont(config, family, bold, italic)).fontSize(size);
}

// Every string a chart draws is absolutely positioned, never flowed. Without an explicit `height` PDFKit
// treats a string that reaches the bottom of the page as overflowing body copy: its line wrapper calls
// `continueOnNewPage()`, so that label AND everything drawn after it land on a page nobody asked for. In
// the chart-image document — one page exactly the size of the chart — that silently dropped whole slices
// from the Word/Excel picture while the same chart drawn on a tall report page looked fine. Passing a
// height bounds the wrapper (`LineWrapper` refuses to paginate once `height != null`) so chart text clips
// at the chart edge, which is what SSRS does, instead of paginating.
const CHART_TEXT_BLOCK_LINES = 4;

function fillText(doc, text, x, y, options = {}) {
  const height = options.height ?? Math.max(1, doc.currentLineHeight(true) * CHART_TEXT_BLOCK_LINES);
  doc.save().fillColor(options.color || LABEL_COLOR)
    .text(String(text), x, y, { lineBreak: false, ...options, height })
    .restore();
}

function styleFont(doc, config, style, context, fallbackSize = 8) {
  const weight = String(styleValue(style?.fontWeight, context, 'Normal'));
  const fontStyle = String(styleValue(style?.fontStyle, context, 'Normal'));
  const appearance = {
    family: String(styleText(style?.fontFamily, context, 'Arial')),
    size: styleSize(style?.fontSize, context, fallbackSize) || fallbackSize,
    bold: /bold|[6-9]00/i.test(weight),
    italic: /italic/i.test(fontStyle),
    color: styleColor(style?.color, context, LABEL_COLOR),
  };
  setFont(doc, config, appearance);
  return appearance;
}

function chartPaint(value, context, fallback = null) {
  const resolved = styleColor(value, context, fallback);
  if (!resolved) return null;
  const argb = /^#([0-9a-f]{8})$/i.exec(resolved);
  if (!argb) return { color: resolved, opacity: 1 };
  const opacity = Number.parseInt(argb[1].slice(0, 2), 16) / 255;
  if (opacity <= 0) return null;
  return { color: `#${argb[1].slice(2)}`, opacity };
}

function strokeChartEdge(doc, x1, y1, x2, y2, border, context) {
  if (!border) return;
  const style = String(styleValue(border.style, context, 'None'));
  const paint = chartPaint(border.color, context, '#000000');
  const width = styleSize(border.width, context, 1);
  if (/^none$/i.test(style) || !paint || width <= 0) return;
  doc.save().strokeColor(paint.color).strokeOpacity(paint.opacity).lineWidth(Math.max(0.25, width));
  if (/dash/i.test(style)) doc.dash(Math.max(2, width * 3));
  else if (/dot/i.test(style)) doc.dash(Math.max(1, width), { space: Math.max(1, width * 2) });
  else doc.lineCap('square').lineJoin('miter');
  if (/double/i.test(style)) {
    const strand = Math.max(0.25, width / 3);
    const vertical = x1 === x2;
    const offsetX = vertical ? strand : 0;
    const offsetY = vertical ? 0 : strand;
    doc.lineWidth(strand);
    doc.moveTo(x1 - offsetX, y1 - offsetY).lineTo(x2 - offsetX, y2 - offsetY).stroke();
    doc.moveTo(x1 + offsetX, y1 + offsetY).lineTo(x2 + offsetX, y2 + offsetY).stroke();
  } else {
    doc.moveTo(x1, y1).lineTo(x2, y2).stroke();
  }
  doc.restore();
}

function drawStyledBox(doc, x, y, width, height, style, context, { fill = true, border = true } = {}) {
  if (!(width > 0 && height > 0)) return;
  if (fill) {
    const paint = chartPaint(style?.backgroundColor, context, null);
    if (paint) doc.save().fillColor(paint.color).fillOpacity(paint.opacity).rect(x, y, width, height).fill().restore();
  }
  if (!border) return;
  const borders = style?.borders || {};
  strokeChartEdge(doc, x, y, x + width, y, borders.top, context);
  strokeChartEdge(doc, x + width, y, x + width, y + height, borders.right, context);
  strokeChartEdge(doc, x, y + height, x + width, y + height, borders.bottom, context);
  strokeChartEdge(doc, x, y, x, y + height, borders.left, context);
}

function axisFont(doc, config, axis, context, labels, slot, fallback = 8) {
  const style = axis?.style || {};
  const configured = styleSize(style.fontSize, context, fallback) || fallback;
  const disabled = String(styleValue(axis?.labelsAutoFitDisabled, context, 'false')).toLowerCase() === 'true';
  let size = configured;
  const appearance = styleFont(doc, config, style, context, configured);
  // SSRS auto-fit can shrink/offset/rotate labels. This renderer implements the deterministic shrink part;
  // when auto-fit is disabled the declared font size is retained exactly.
  if (!disabled && slot > 0 && labels?.length) {
    const widest = Math.max(0, ...labels.map((label) => doc.widthOfString(String(label ?? ''))));
    if (widest > slot) size = Math.max(6, configured * slot / widest);
  }
  setFont(doc, config, { ...appearance, size });
  return { ...appearance, color: styleColor(style.color, context, TICK_LABEL_COLOR), size };
}

function categoryLabelRotation(doc, axis, context, labels, slot) {
  if (!(slot > 0) || !labels?.length) return 0;
  const widest = Math.max(0, ...labels.map((label) => doc.widthOfString(String(label ?? ''))));
  if (widest <= slot) return 0;
  const declared = String(styleValue(axis?.allowLabelRotation, context, 'None')).toLowerCase();
  if (declared === 'rotate90') return 90;
  if (declared === 'rotate45') return 45;
  if (declared === 'rotate30') return 30;
  if (declared === 'rotate15') return 15;
  return 0;
}

function categoryLabelBand(doc, config, axis, context, labels, slot, fallback = 7) {
  const appearance = axisFont(doc, config, axis, context, labels, slot, fallback);
  const rotation = categoryLabelRotation(doc, axis, context, labels, slot);
  const lineHeight = doc.currentLineHeight(true);
  const tallest = Math.max(0, ...labels.map((label) => {
    const width = doc.widthOfString(String(label ?? ''));
    if (!rotation) return doc.heightOfString(String(label ?? ''), { width: Math.max(1, slot) });
    const radians = rotation * Math.PI / 180;
    return Math.abs(width * Math.sin(radians)) + Math.abs(lineHeight * Math.cos(radians));
  }));
  return { appearance, rotation, height: tallest };
}

function drawBottomCategoryLabel(doc, label, slotX, slot, y, appearance, rotation) {
  if (!rotation) {
    fillText(doc, label, slotX, y, { width: slot, align: 'center', color: appearance.color });
    return;
  }
  const centerX = slotX + slot / 2;
  const width = doc.widthOfString(String(label ?? ''));
  doc.save().fillColor(appearance.color || TICK_LABEL_COLOR)
    .rotate(-rotation, { origin: [centerX, y] })
    .text(String(label ?? ''), centerX - width / 2, y, { lineBreak: false, height: doc.currentLineHeight(true) })
    .restore();
}

function legendLayout(doc, config, legend, entries, maxWidth, maxHeight, context, orientation) {
  const appearance = styleFont(doc, config, legend.style || {}, context, 8);
  const padding = 4;
  // SSRS' default series-key is a horizontal rectangle, not a text-sized square. Keep its height tied
  // to the legend font while reserving the wider symbol in both horizontal and vertical layouts.
  const swatchHeight = Math.max(6, appearance.size * 0.75);
  const swatchWidth = Math.max(12, appearance.size * 1.5);
  const gap = Math.max(3, appearance.size * 0.5);
  const spacing = Math.max(8, appearance.size * 1.5);
  const lineHeight = Math.max(swatchHeight, appearance.size * 1.25) + 4;
  const chips = entries.map((entry) => ({
    ...entry,
    width: swatchWidth + gap + doc.widthOfString(String(entry.label)) + spacing,
  }));
  if (orientation === 'vertical') {
    const rowsPerColumn = Math.max(1, Math.floor(Math.max(lineHeight, maxHeight - padding * 2) / lineHeight));
    const columns = [];
    for (let index = 0; index < chips.length; index += rowsPerColumn) {
      const values = chips.slice(index, index + rowsPerColumn);
      columns.push({ chips: values, width: Math.max(0, ...values.map((chip) => chip.width)) });
    }
    return {
      appearance, chips, columns, orientation, padding, swatchWidth, swatchHeight, gap, spacing, lineHeight,
      width: Math.min(maxWidth, padding * 2 + columns.reduce((sum, column) => sum + column.width, 0)),
      height: Math.min(maxHeight, padding * 2 + Math.min(rowsPerColumn, chips.length) * lineHeight),
    };
  }
  const rows = [[]];
  const usableWidth = Math.max(1, maxWidth - padding * 2);
  let rowWidth = 0;
  for (const chip of chips) {
    if (rowWidth + chip.width > usableWidth && rows[rows.length - 1].length) { rows.push([]); rowWidth = 0; }
    rows[rows.length - 1].push(chip);
    rowWidth += chip.width;
  }
  return {
    appearance, chips, rows, orientation, padding, swatchWidth, swatchHeight, gap, spacing, lineHeight,
    width: Math.min(maxWidth, padding * 2 + Math.max(0, ...rows.map((row) => row.reduce((sum, chip) => sum + chip.width, 0) - spacing))),
    height: Math.min(maxHeight, padding * 2 + rows.length * lineHeight),
  };
}

function drawLegend(doc, config, layout, legend, x, y, context, alignment = 'center') {
  if (!layout.chips.length) return;
  drawStyledBox(doc, x, y, layout.width, layout.height, legend.style, context);
  setFont(doc, config, layout.appearance);
  const textColor = layout.appearance.color || TICK_LABEL_COLOR;
  if (layout.orientation === 'vertical') {
    let columnX = x + layout.padding;
    for (const column of layout.columns) {
      column.chips.forEach((chip, rowIndex) => {
        const rowY = y + layout.padding + rowIndex * layout.lineHeight;
        const swatchY = rowY + (layout.lineHeight - layout.swatchHeight) / 2;
        doc.save().fillColor(resolveColor(chip.color, '#808080')).rect(columnX, swatchY, layout.swatchWidth, layout.swatchHeight).fill().restore();
        fillText(doc, chip.label, columnX + layout.swatchWidth + layout.gap, rowY + 1, { color: textColor });
      });
      columnX += column.width;
    }
    return;
  }
  layout.rows.forEach((row, rowIndex) => {
    const rowWidth = row.reduce((sum, chip) => sum + chip.width, 0) - layout.spacing;
    const offset = alignment === 'left' ? 0 : alignment === 'right' ? layout.width - layout.padding * 2 - rowWidth : (layout.width - layout.padding * 2 - rowWidth) / 2;
    let cursor = x + layout.padding + Math.max(0, offset);
    const rowY = y + layout.padding + rowIndex * layout.lineHeight;
    for (const chip of row) {
      const swatchY = rowY + (layout.lineHeight - layout.swatchHeight) / 2;
      doc.save().fillColor(resolveColor(chip.color, '#808080')).rect(cursor, swatchY, layout.swatchWidth, layout.swatchHeight).fill().restore();
      fillText(doc, chip.label, cursor + layout.swatchWidth + layout.gap, rowY + 1, { color: textColor });
      cursor += chip.width;
    }
  });
}

function titleHeight(chart, context) {
  const style = chart.title?.style || {};
  return Math.max(14, (styleSize(style.fontSize, context, 9) || 9) * 1.35
    + styleSize(style.paddingTop, context, 2) + styleSize(style.paddingBottom, context, 2));
}

function titleWidth(doc, config, chart, context, maximum) {
  const caption = String(evaluateExpression(chart.title.caption, context) ?? '').trim();
  const style = chart.title.style || {};
  styleFont(doc, config, style, context, 9);
  const paddingLeft = styleSize(style.paddingLeft, context, 2);
  const paddingRight = styleSize(style.paddingRight, context, 2);
  const widestLine = Math.max(0, ...caption.split(/\r?\n/).map((line) => doc.widthOfString(line)));
  return Math.min(maximum, Math.max(1, widestLine + paddingLeft + paddingRight));
}

function anchoredStart(start, end, size, alignment) {
  if (/^(?:left|top)$/i.test(alignment)) return start;
  if (/^(?:right|bottom)$/i.test(alignment)) return end - size;
  return start + (end - start - size) / 2;
}

function drawTitle(doc, config, chart, x, y, width, height, context) {
  const caption = String(evaluateExpression(chart.title.caption, context) ?? '').trim();
  const style = chart.title.style || {};
  drawStyledBox(doc, x, y, width, height, style, context);
  const appearance = styleFont(doc, config, style, context, 9);
  const align = String(styleValue(style.textAlign, context, 'Center')).toLowerCase();
  const paddingLeft = styleSize(style.paddingLeft, context, 2);
  const paddingRight = styleSize(style.paddingRight, context, 2);
  const paddingTop = styleSize(style.paddingTop, context, 2);
  fillText(doc, caption, x + paddingLeft, y + paddingTop, {
    width: Math.max(1, width - paddingLeft - paddingRight),
    align: ['left', 'right', 'center'].includes(align) ? align : 'center',
    color: appearance.color,
  });
}

function drawValueGrid(doc, config, scale, plot, orientation, axis, context) {
  const axisAppearance = axisFont(doc, config, axis, context, [], 0, 8);
  const ticks = Math.round(scale.max / scale.interval);
  for (let tick = 0; tick <= ticks; tick += 1) {
    const value = tick * scale.interval;
    if (orientation === 'horizontal') {
      const gx = plot.x + (value / scale.max) * plot.width;
      doc.save().lineWidth(0.5).strokeColor(AXIS_COLOR).moveTo(gx, plot.y).lineTo(gx, plot.y + plot.height).stroke().restore();
      fillText(doc, formatTick(value, scale.interval), gx - 10, plot.y + plot.height + 4, { width: 20, align: 'center', color: axisAppearance.color });
    } else {
      const gy = plot.y + plot.height - (value / scale.max) * plot.height;
      doc.save().lineWidth(0.5).strokeColor(AXIS_COLOR).moveTo(plot.x, gy).lineTo(plot.x + plot.width, gy).stroke().restore();
      fillText(doc, formatTick(value, scale.interval), plot.x - 34, gy - 4, { width: 30, align: 'right', color: axisAppearance.color });
    }
  }
}

// Per-category stack total, and the value-axis maximum for a stacked/percent chart.
function stackTotal(data, categoryIndex) {
  return data.series.reduce((sum, series) => sum + Math.max(0, series.points[categoryIndex]?.y ?? 0), 0);
}
function scaleFor(data, stacked) {
  if (stacked === 'percent') return niceScale(100);
  if (stacked === 'stacked') return niceScale(Math.max(1, ...data.categories.map((_, index) => stackTotal(data, index))));
  return niceScale(data.maxY);
}

function drawBarChart(doc, config, chart, data, plot, stacked, context) {
  const scale = scaleFor(data, stacked);
  drawValueGrid(doc, config, scale, plot, 'horizontal', chart.valueAxis, context);
  const slot = plot.height / (data.categories.length || 1);
  const configuredWidth = Number(styleValue(chart.seriesDefs?.[0]?.customProperties?.PointWidth, context, 0.8));
  const barHeight = slot * Math.min(1, Math.max(0.05, Number.isFinite(configuredWidth) ? configuredWidth : 0.8));
  const categoryAppearance = axisFont(doc, config, chart.categoryAxis, context, data.categories.map((entry) => entry.label), 152, 8);
  data.categories.forEach((category, index) => {
    // First category at the bottom (SSRS horizontal-bar order).
    const centerY = plot.y + plot.height - (index + 0.5) * slot;
    fillText(doc, category.label ?? '', plot.x - 158, centerY - 4, { width: 152, align: 'right', color: categoryAppearance.color });
    if (stacked === 'none') {
      const point = data.series[0]?.points[index];
      if (!point || point.y === null || point.y <= 0) return;
      const barWidth = (point.y / scale.max) * plot.width;
      doc.save().fillColor(resolveColor(point.color, '#808080')).rect(plot.x, centerY - barHeight / 2, barWidth, barHeight).fill().restore();
      if (point.label) fillText(doc, point.label, plot.x + barWidth + 3, centerY - 4, { color: LABEL_COLOR });
      return;
    }
    const total = stackTotal(data, index) || 1;
    let cursorX = plot.x;
    for (const series of data.series) {
      const point = series.points[index];
      const raw = Math.max(0, point?.y ?? 0);
      if (raw <= 0) continue;
      const width = ((stacked === 'percent' ? (raw / total) * 100 : raw) / scale.max) * plot.width;
      doc.save().fillColor(resolveColor(series.color, '#808080')).rect(cursorX, centerY - barHeight / 2, width, barHeight).fill().restore();
      // SSRS renders an enabled data-point label inside its own stacked segment.  The series loop used
      // to discard the materialized label here, so stacked bars lost labels while columns, lines and pies
      // retained theirs.  Suppress only labels that cannot fit in the declared segment rather than letting
      // a label overlap the adjacent series.
      if (point?.label) {
        const label = String(point.label);
        const labelWidth = doc.widthOfString(label);
        const lineHeight = doc.currentLineHeight(true);
        if (width >= labelWidth + 4 && barHeight >= lineHeight + 2) {
          fillText(doc, label, cursorX, centerY - lineHeight / 2, {
            width,
            height: lineHeight,
            align: 'center',
            color: LABEL_COLOR,
          });
        }
      }
      cursorX += width;
    }
  });
}

function drawColumnChart(doc, config, chart, data, plot, stacked, context) {
  const scale = scaleFor(data, stacked);
  drawValueGrid(doc, config, scale, plot, 'vertical', chart.valueAxis, context);
  const categoryCount = data.categories.length || 1;
  const seriesCount = data.series.length || 1;
  const slot = plot.width / categoryCount;
  const categoryLabels = data.categories.map((entry) => entry.label);
  const categoryLayout = categoryLabelBand(doc, config, chart.categoryAxis, context, categoryLabels, slot, 7);
  const configuredWidth = Number(styleValue(chart.seriesDefs?.[0]?.customProperties?.PointWidth, context, 0.8));
  const pointWidth = Math.min(1, Math.max(0.05, Number.isFinite(configuredWidth) ? configuredWidth : 0.8));
  data.categories.forEach((category, categoryIndex) => {
    const slotX = plot.x + categoryIndex * slot;
    drawBottomCategoryLabel(doc, category.label ?? '', slotX, slot, plot.y + plot.height + 4, categoryLayout.appearance, categoryLayout.rotation);
    if (stacked === 'none') {
      const groupWidth = slot * pointWidth;
      const barWidth = groupWidth / seriesCount;
      data.series.forEach((series, seriesIndex) => {
        const point = series.points[categoryIndex];
        if (!point || point.y === null || point.y <= 0) return;
        const barHeight = (point.y / scale.max) * plot.height;
        const barX = slotX + (slot - groupWidth) / 2 + seriesIndex * barWidth;
        const barY = plot.y + plot.height - barHeight;
        doc.save().fillColor(resolveColor(point.color, '#808080')).rect(barX, barY, barWidth * 0.9, barHeight).fill().restore();
        if (point.label) fillText(doc, point.label, barX - barWidth, barY - 9, { width: barWidth * 3, align: 'center', color: LABEL_COLOR });
      });
      return;
    }
    const barWidth = slot * pointWidth;
    const barX = slotX + (slot - barWidth) / 2;
    const total = stackTotal(data, categoryIndex) || 1;
    let cursorY = plot.y + plot.height;
    for (const series of data.series) {
      const raw = Math.max(0, series.points[categoryIndex]?.y ?? 0);
      if (raw <= 0) continue;
      const segmentHeight = ((stacked === 'percent' ? (raw / total) * 100 : raw) / scale.max) * plot.height;
      cursorY -= segmentHeight;
      doc.save().fillColor(resolveColor(series.color, '#808080')).rect(barX, cursorY, barWidth, segmentHeight).fill().restore();
      if (segmentHeight > 9) fillText(doc, String(raw), barX, cursorY + segmentHeight / 2 - 4, { width: barWidth, align: 'center', color: '#ffffff' });
    }
  });
}

// Pie (innerRatio 0) and doughnut (innerRatio > 0) share this: each slice is an annulus segment.
// The RDL exploded subtypes move every slice away from the common centre along its bisector. The ratio is
// a renderer-level semantic used for every exploded shape chart; the radius is reduced first so no slice,
// label, or callout can leave the declared plot rectangle.
function drawPieChart(doc, config, chart, data, plot, innerRatio, context) {
  // A zero-value data point is a real value, not an empty one: SSRS keeps its palette colour and legend
  // entry, draws a zero-width slice, and still prints its data label at that angle. Only a point with no
  // value at all (an empty category/series intersection) is dropped, which materializeChart already marks.
  const points = (data.series[0]?.points || []).filter((point) => Number.isFinite(point.y) && point.y >= 0);
  const total = points.reduce((sum, point) => sum + point.y, 0);
  if (total <= 0) return;
  const explosionRatio = chart.exploded && points.length > 1 ? 0.1 : 0;
  const outerLimit = Math.max(10, Math.min(plot.width, plot.height) / 2 - 8);
  // SSRS draws an outside shape-chart label as a callout — a radial stub off the slice edge, then a
  // horizontal elbow, with the label starting at the elbow. It shrinks the shape so the whole callout
  // (both segments plus the gap before the text) stays inside the plot rectangle; only the label text
  // itself may reach past it. Reserve that band here instead of letting the callout run over the edge.
  // WHERE a shape-chart label goes is decided by the series' PieLabelStyle custom property, not by how
  // much room the wedge has: SSRS defaults to Inside and keeps drawing at the label radius however narrow
  // the slice becomes (a 0% sliver still prints "0%" inside the pie, overlapping its neighbour if need
  // be). Only PieLabelStyle=Outside — or an explicitly Outside data-label Position — moves labels out to
  // callouts, and Disabled suppresses them. Auto-promoting a label that "does not fit" invented callouts
  // SSRS never draws.
  const pieLabelStyle = String(styleValue(chart.seriesDefs?.[0]?.customProperties?.PieLabelStyle, context, 'Inside'));
  const labelsDisabled = /^disabled$/i.test(pieLabelStyle);
  const labelsOutside = /^outside$/i.test(pieLabelStyle);
  const labelPlacement = (point) => {
    const appearance = styleFont(doc, config, point.labelStyle || {}, context, 8);
    const labelWidth = doc.widthOfString(String(point.label ?? '')) + 1;
    const lineHeight = doc.currentLineHeight(true);
    const outside = Boolean(point.label)
      && (labelsOutside || /outside/i.test(String(point.labelPosition || 'Auto')));
    return { outside, appearance, labelWidth, lineHeight };
  };
  const hasOutsideLabels = points.some((point) => labelPlacement(point).outside);
  const calloutBudget = hasOutsideLabels
    ? Math.max(10, (outerLimit - PIE_CALLOUT_TEXT_GAP_PT) / (1 + 2 * PIE_CALLOUT_SEGMENT_RATIO))
    : outerLimit;
  const radius = Math.max(10, calloutBudget / (1 + explosionRatio));
  const explosion = radius * explosionRatio;
  const inner = radius * innerRatio;
  const centerX = plot.x + plot.width / 2;
  const centerY = plot.y + plot.height / 2;
  const at = (angle, r, offsetX = 0, offsetY = 0) => [
    centerX + offsetX + Math.cos(angle) * r,
    centerY + offsetY + Math.sin(angle) * r,
  ];
  // A ChartBorderSkin is a decorative FRAME around the chart image (Emboss, Raised, FrameThin…), and it
  // is inert unless the RDL declares a BorderSkinType — which defaults to None. It is not a slice outline:
  // the BI designer emits <ChartBorderSkin><Style><BackgroundColor>Gray</BackgroundColor><Color>White…
  // into practically every chart, and SSRS still draws pie slices as bare adjacent fills, so painting the
  // skin colour around each slice added a white radial rule SSRS never draws. A slice outline would come
  // from the data point's own Style/Border, which these reports do not declare.
  // SSRS pie and doughnut charts use PieStartAngle=0 by default, which places the first value at
  // 3 o'clock (90 degrees clockwise from the top). A declared PieStartAngle rotates clockwise in
  // degrees; 270 therefore places the first value at 12 o'clock.
  const rawStartAngle = styleValue(chart.seriesDefs?.[0]?.customProperties?.PieStartAngle, context, 0);
  const startAngleDegrees = Number(rawStartAngle);
  let angle = (Number.isFinite(startAngleDegrees) ? startAngleDegrees : 0) * (Math.PI / 180);
  // Geometry first, then every slice, then every label. Labels belong to the shape as a whole: a label
  // that sits inside a hairline slice is drawn at the label radius and therefore lands on the NEXT
  // slice's fill, which would paint over it if slices and labels were interleaved in one pass.
  const slices = points.map((point) => {
    const sweep = (point.y / total) * Math.PI * 2;
    const middle = angle + sweep / 2;
    const slice = {
      point,
      sweep,
      middle,
      start: angle,
      offsetX: Math.cos(middle) * explosion,
      offsetY: Math.sin(middle) * explosion,
      placement: labelPlacement(point),
    };
    angle += sweep;
    return slice;
  });
  for (const { point, start, sweep, offsetX, offsetY } of slices) {
    const steps = Math.max(2, Math.ceil((sweep / (Math.PI * 2)) * 120));
    doc.save().fillColor(resolveColor(point.color, '#808080'));
    doc.moveTo(...at(start, radius, offsetX, offsetY));
    for (let step = 1; step <= steps; step += 1) doc.lineTo(...at(start + (sweep * step) / steps, radius, offsetX, offsetY));
    for (let step = steps; step >= 0; step -= 1) doc.lineTo(...at(start + (sweep * step) / steps, inner, offsetX, offsetY));
    doc.fill();
    doc.restore();
  }
  if (labelsDisabled) return;
  for (const { point, middle, offsetX, offsetY, placement } of slices) {
    if (!point.label) continue;
    // The label box is measured from the text itself. A fixed-width box wraps a value that is merely
    // a few points wider than the guess ("1 (50.0%)" became two lines), which SSRS never does.
    const { labelWidth, lineHeight, outside, appearance } = placement;
    // Measurement happened in the geometry pass, so the point's own data-label font has to be re-applied
    // before the text is drawn: fillText paints with whatever face is current on the document.
    setFont(doc, config, appearance);
    const labelHalfHeight = lineHeight / 2;
    if (outside) {
      const segment = radius * PIE_CALLOUT_SEGMENT_RATIO;
      const [stubX, stubY] = at(middle, radius, offsetX, offsetY);
      const [elbowX, elbowY] = at(middle, radius + segment, offsetX, offsetY);
      // The elbow follows the horizontal direction of the bisector. A bisector that is vertical has a
      // cosine of ±1e-16, so compare against a tolerance: a floating-point sign must not decide which
      // side of the shape the label sits on (SSRS puts a straight-up/straight-down callout on the right).
      const right = Math.cos(middle) >= -COINCIDENT_ANGLE_TOLERANCE;
      const elbowEndX = elbowX + (right ? segment : -segment);
      const calloutColor = styleColor(chart.seriesDefs?.[0]?.customProperties?.PieLineColor, context, '#000000');
      doc.save().lineWidth(0.75).strokeColor(calloutColor)
        .moveTo(stubX, stubY).lineTo(elbowX, elbowY).lineTo(elbowEndX, elbowY)
        .stroke().restore();
      const textX = right
        ? elbowEndX + PIE_CALLOUT_TEXT_GAP_PT
        : elbowEndX - PIE_CALLOUT_TEXT_GAP_PT - labelWidth;
      fillText(doc, point.label, textX, elbowY - labelHalfHeight, {
        width: labelWidth,
        height: lineHeight,
        align: right ? 'left' : 'right',
        color: appearance.color,
      });
    } else {
      const [labelX, labelY] = at(middle, inner + (radius - inner) * 0.55, offsetX, offsetY);
      fillText(doc, point.label, labelX - labelWidth / 2, labelY - labelHalfHeight, {
        width: labelWidth,
        height: lineHeight,
        align: 'center',
        color: appearance.color,
      });
    }
  }
}

// X positions and scaled Y for each series point, shared by the line and area charts.
function seriesLines(data, scale, plot) {
  const slot = plot.width / (data.categories.length || 1);
  return data.series.map((series) => ({
    color: series.color || series.points.find((point) => point?.color)?.color,
    points: series.points.map((point, index) => (point && point.y !== null ? {
      x: plot.x + (index + 0.5) * slot,
      y: plot.y + plot.height - (point.y / scale.max) * plot.height,
      label: point.label,
    } : null)),
  }));
}

function drawCategoryLabels(doc, config, chart, data, plot, context) {
  const slot = plot.width / (data.categories.length || 1);
  const labels = data.categories.map((entry) => entry.label);
  const layout = categoryLabelBand(doc, config, chart.categoryAxis, context, labels, slot, 7);
  data.categories.forEach((category, index) => {
    drawBottomCategoryLabel(doc, category.label ?? '', plot.x + index * slot, slot, plot.y + plot.height + 4, layout.appearance, layout.rotation);
  });
}

function drawLineChart(doc, config, chart, data, plot, context) {
  const scale = niceScale(data.maxY);
  drawValueGrid(doc, config, scale, plot, 'vertical', chart.valueAxis, context);
  drawCategoryLabels(doc, config, chart, data, plot, context);
  setFont(doc, config, { size: 7 });
  for (const series of seriesLines(data, scale, plot)) {
    const points = series.points.filter(Boolean);
    if (!points.length) continue;
    const color = resolveColor(series.color, '#4472c4');
    doc.save().lineWidth(1.5).strokeColor(color);
    points.forEach((point, index) => (index === 0 ? doc.moveTo(point.x, point.y) : doc.lineTo(point.x, point.y)));
    doc.stroke().restore();
    for (const point of points) {
      doc.save().fillColor(color).circle(point.x, point.y, 2).fill().restore();
      if (point.label) fillText(doc, point.label, point.x - 14, point.y - 12, { width: 28, align: 'center', color: LABEL_COLOR });
    }
  }
}

function drawAreaChart(doc, config, chart, data, plot, stacked, context) {
  const scale = scaleFor(data, stacked);
  drawValueGrid(doc, config, scale, plot, 'vertical', chart.valueAxis, context);
  drawCategoryLabels(doc, config, chart, data, plot, context);
  setFont(doc, config, { size: 7 });
  const baseline = plot.y + plot.height;
  if (stacked === 'none') {
    for (const series of seriesLines(data, scale, plot)) {
      const points = series.points.filter(Boolean);
      if (!points.length) continue;
      const color = resolveColor(series.color, '#4472c4');
      doc.save().fillOpacity(0.35).fillColor(color).moveTo(points[0].x, baseline);
      for (const point of points) doc.lineTo(point.x, point.y);
      doc.lineTo(points[points.length - 1].x, baseline).fill().restore();
      doc.save().lineWidth(1.5).strokeColor(color);
      points.forEach((point, index) => (index === 0 ? doc.moveTo(point.x, point.y) : doc.lineTo(point.x, point.y)));
      doc.stroke().restore();
      for (const point of points) {
        if (point.label) fillText(doc, point.label, point.x - 14, point.y - 12, { width: 28, align: 'center', color: LABEL_COLOR });
      }
    }
    return;
  }
  // Stacked/percent: each series is a band drawn between the running lower and upper cumulative totals.
  const slot = plot.width / (data.categories.length || 1);
  const xAt = (index) => plot.x + (index + 0.5) * slot;
  const yFor = (value) => plot.y + plot.height - (value / scale.max) * plot.height;
  const cumulative = data.categories.map(() => 0);
  for (const series of data.series) {
    const lower = [];
    const upper = [];
    data.categories.forEach((category, index) => {
      const raw = Math.max(0, series.points[index]?.y ?? 0);
      const value = stacked === 'percent' ? (raw / (stackTotal(data, index) || 1)) * 100 : raw;
      lower.push(cumulative[index]);
      cumulative[index] += value;
      upper.push(cumulative[index]);
    });
    doc.save().fillOpacity(0.75).fillColor(resolveColor(series.color, '#4472c4'));
    doc.moveTo(xAt(0), yFor(upper[0]));
    data.categories.forEach((category, index) => doc.lineTo(xAt(index), yFor(upper[index])));
    for (let index = data.categories.length - 1; index >= 0; index -= 1) doc.lineTo(xAt(index), yFor(lower[index]));
    doc.fill().restore();
  }
}

function drawScatterChart(doc, config, chart, data, plot, context) {
  const scale = niceScale(data.maxY);
  drawValueGrid(doc, config, scale, plot, 'vertical', chart.valueAxis, context);
  drawCategoryLabels(doc, config, chart, data, plot, context);
  setFont(doc, config, { size: 7 });
  for (const series of seriesLines(data, scale, plot)) {
    const color = resolveColor(series.color, '#4472c4');
    for (const point of series.points.filter(Boolean)) {
      doc.save().fillColor(color).circle(point.x, point.y, 3).fill().restore();
      if (point.label) fillText(doc, point.label, point.x - 14, point.y - 12, { width: 28, align: 'center', color: LABEL_COLOR });
    }
  }
}

// Draws a chart into the current PDFKit document at (x,y) within width x height. Supports the
// bar/column/pie/doughnut/line/area/scatter types the parser accepts (bar/column/area honour the
// stacked & percent-stacked subtypes); everything else is fail-closed before we get here.
export function drawChart(doc, config, chart, data, x, y, width, height, outerContext) {
  // Chart-level expressions — the title, axis titles, legend title, and expression-backed styles — are
  // evaluated in the CHART's data scope, not in the scope of whatever positioned the chart. SSRS resolves
  // a bare Fields! reference there against the chart's own dataset, so those expressions see the first row
  // the chart was materialized from. Without this every field-backed caption rendered empty, because a
  // body, canvas, or nested-region context carries no row of the chart's dataset.
  const chartRows = chart.datasetName ? outerContext?.datasets?.[chart.datasetName] : null;
  const context = Array.isArray(chartRows) && chartRows.length > 0
    ? { ...outerContext, dataset: chartRows, fields: chartRows[0] }
    : outerContext;
  drawStyledBox(doc, x, y, width, height, chart.style, context, { fill: true, border: false });
  doc.save().rect(x, y, width, height).clip();
  const content = {
    left: x + 8,
    top: y + 6,
    right: x + width - 8,
    bottom: y + height - 6,
  };

  const titleVisible = chart.title && !isHidden(chart.title.hidden, context);
  if (titleVisible) {
    const measuredHeight = titleHeight(chart, context);
    const titlePosition = String(styleValue(chart.title.position, context, 'TopCenter'))
      .replace(/\s+/g, '')
      .toLowerCase();
    const measuredWidth = titleWidth(doc, config, chart, context, content.right - content.left);
    if (titlePosition.startsWith('bottom')) {
      const alignment = titlePosition.slice('bottom'.length) || 'center';
      const titleX = anchoredStart(content.left, content.right, measuredWidth, alignment);
      drawTitle(doc, config, chart, titleX, content.bottom - measuredHeight, measuredWidth, measuredHeight, context);
      content.bottom -= measuredHeight + 6;
    } else if (titlePosition.startsWith('left') || titlePosition.startsWith('right')) {
      const side = titlePosition.startsWith('left') ? 'left' : 'right';
      const alignment = titlePosition.slice(side.length) || 'center';
      const titleY = anchoredStart(content.top, content.bottom, measuredHeight, alignment);
      const titleX = side === 'left' ? content.left : content.right - measuredWidth;
      drawTitle(doc, config, chart, titleX, titleY, measuredWidth, measuredHeight, context);
      if (side === 'left') content.left += measuredWidth + 6;
      else content.right -= measuredWidth + 6;
    } else {
      const alignment = titlePosition.startsWith('top')
        ? titlePosition.slice('top'.length) || 'center'
        : 'center';
      const titleX = anchoredStart(content.left, content.right, measuredWidth, alignment);
      drawTitle(doc, config, chart, titleX, content.top, measuredWidth, measuredHeight, context);
      content.top += measuredHeight + 6;
    }
  }

  const legendVisible = chart.legend?.visible && !isHidden(chart.legend.hidden, context) && data.legend.length;
  if (legendVisible) {
    const position = String(styleValue(chart.legend.position, context, 'RightTop')).replace(/\s+/g, '').toLowerCase();
    const layoutSetting = String(styleValue(chart.legend.layout, context, 'AutoTable')).toLowerCase();
    const side = position.startsWith('left') ? 'left'
      : position.startsWith('right') ? 'right'
        : position.startsWith('top') ? 'top' : 'bottom';
    const orientation = /^row$/i.test(layoutSetting) ? 'horizontal'
      : /^column$/i.test(layoutSetting) ? 'vertical'
        : (side === 'left' || side === 'right' ? 'vertical' : 'horizontal');
    const contentWidth = Math.max(1, content.right - content.left);
    const contentHeight = Math.max(1, content.bottom - content.top);
    const maxLegendWidth = side === 'left' || side === 'right' ? Math.min(180, contentWidth * 0.42) : contentWidth;
    const maxLegendHeight = side === 'top' || side === 'bottom' ? Math.min(90, contentHeight * 0.35) : contentHeight;
    const layout = legendLayout(doc, config, chart.legend, data.legend, maxLegendWidth, maxLegendHeight, context, orientation);
    let legendX = content.left;
    let legendY = content.top;
    let horizontalAlignment = 'center';
    if (side === 'left' || side === 'right') {
      legendX = side === 'left' ? content.left : content.right - layout.width;
      const verticalAlignment = position.endsWith('bottom') ? 'bottom' : position.endsWith('center') ? 'center' : 'top';
      legendY = verticalAlignment === 'bottom' ? content.bottom - layout.height
        : verticalAlignment === 'center' ? content.top + (contentHeight - layout.height) / 2
          : content.top;
      if (side === 'left') content.left += layout.width + 8;
      else content.right -= layout.width + 8;
    } else {
      horizontalAlignment = position.endsWith('left') ? 'left' : position.endsWith('right') ? 'right' : 'center';
      legendX = horizontalAlignment === 'left' ? content.left
        : horizontalAlignment === 'right' ? content.right - layout.width
          : content.left + (contentWidth - layout.width) / 2;
      legendY = side === 'top' ? content.top : content.bottom - layout.height;
      if (side === 'top') content.top += layout.height + 8;
      else content.bottom -= layout.height + 8;
    }
    drawLegend(doc, config, layout, chart.legend, legendX, legendY, context, horizontalAlignment);
  }

  if (!data.hasData) {
    setFont(doc, config, { size: 10, bold: true });
    fillText(doc, chart.noDataMessage || 'No Data Available', content.left, content.top + (content.bottom - content.top) / 2 - 6, {
      width: content.right - content.left,
      align: 'center',
      color: TICK_LABEL_COLOR,
    });
    doc.restore();
    drawStyledBox(doc, x, y, width, height, chart.style, context, { fill: false, border: true });
    return;
  }

  // Plot area, leaving gutters for axis/category labels.
  const circular = chart.chartType === 'pie' || chart.chartType === 'doughnut';
  const leftGutter = chart.chartType === 'bar' ? Math.min(165, (content.right - content.left) * 0.4) : circular ? 0 : 40;
  // Reserve the true rendered height of the bottom category labels so a long label that wraps to a second
  // line inside its narrow column slot is not clipped by the chart's outer clip rectangle. Bar charts put
  // their category labels on the left gutter, so their bottom band only carries short numeric value ticks.
  // The plot width does not depend on the bottom gutter, so it can be estimated here to size the slot the
  // same way the series renderers do, then measured with the auto-fitted axis font.
  let bottomGutter = circular ? 0 : 16;
  const bottomCategoryLabels = new Set(['column', 'line', 'area', 'scatter']);
  if (!circular && bottomCategoryLabels.has(chart.chartType) && data.categories?.length) {
    const estimatedPlotWidth = Math.max(1, content.right - content.left - leftGutter - 24);
    const slot = estimatedPlotWidth / data.categories.length;
    const labels = data.categories.map((entry) => String(entry.label ?? ''));
    const labelBand = categoryLabelBand(doc, config, chart.categoryAxis, context, labels, slot, 7);
    bottomGutter = Math.max(16, Math.ceil(labelBand.height) + 6);
  }
  const plot = {
    x: content.left + leftGutter,
    y: content.top + 6,
    width: content.right - content.left - leftGutter - (circular ? 0 : 24),
    height: content.bottom - content.top - 6 - bottomGutter,
  };
  if (plot.width > 10 && plot.height > 10) {
    drawStyledBox(doc, plot.x, plot.y, plot.width, plot.height, chart.chartArea?.style, context, { fill: true, border: false });
    const stacked = chart.stacked || 'none';
    if (chart.chartType === 'bar') drawBarChart(doc, config, chart, data, plot, stacked, context);
    else if (chart.chartType === 'column') drawColumnChart(doc, config, chart, data, plot, stacked, context);
    else if (chart.chartType === 'pie') drawPieChart(doc, config, chart, data, plot, 0, context);
    else if (chart.chartType === 'doughnut') drawPieChart(doc, config, chart, data, plot, 0.55, context);
    else if (chart.chartType === 'line') drawLineChart(doc, config, chart, data, plot, context);
    else if (chart.chartType === 'area') drawAreaChart(doc, config, chart, data, plot, stacked, context);
    else if (chart.chartType === 'scatter') drawScatterChart(doc, config, chart, data, plot, context);
    drawStyledBox(doc, plot.x, plot.y, plot.width, plot.height, chart.chartArea?.style, context, { fill: false, border: true });
  }
  doc.restore();
  drawStyledBox(doc, x, y, width, height, chart.style, context, { fill: false, border: true });
}
