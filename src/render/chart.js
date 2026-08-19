import { evaluateExpression } from '../rdl/expression.js';
import { pdfFont } from './fonts.js';
import { color as resolveColor, isHidden, styleColor, styleSize, styleValue } from './common.js';

const AXIS_COLOR = '#d9d9d9';
const TICK_LABEL_COLOR = '#595959';
const LABEL_COLOR = '#000000';

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

function fillText(doc, text, x, y, options = {}) {
  doc.save().fillColor(options.color || LABEL_COLOR).text(String(text), x, y, { lineBreak: false, ...options }).restore();
}

function styleFont(doc, config, style, context, fallbackSize = 8) {
  const weight = String(styleValue(style?.fontWeight, context, 'Normal'));
  const fontStyle = String(styleValue(style?.fontStyle, context, 'Normal'));
  const appearance = {
    family: String(styleValue(style?.fontFamily, context, 'Arial')),
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

function legendLayout(doc, config, legend, entries, maxWidth, maxHeight, context, orientation) {
  const appearance = styleFont(doc, config, legend.style || {}, context, 8);
  const padding = 4;
  const swatch = Math.max(8, appearance.size);
  const gap = Math.max(3, appearance.size * 0.5);
  const spacing = Math.max(8, appearance.size * 1.5);
  const lineHeight = Math.max(swatch, appearance.size * 1.25) + 4;
  const chips = entries.map((entry) => ({
    ...entry,
    width: swatch + gap + doc.widthOfString(String(entry.label)) + spacing,
  }));
  if (orientation === 'vertical') {
    const rowsPerColumn = Math.max(1, Math.floor(Math.max(lineHeight, maxHeight - padding * 2) / lineHeight));
    const columns = [];
    for (let index = 0; index < chips.length; index += rowsPerColumn) {
      const values = chips.slice(index, index + rowsPerColumn);
      columns.push({ chips: values, width: Math.max(0, ...values.map((chip) => chip.width)) });
    }
    return {
      appearance, chips, columns, orientation, padding, swatch, gap, spacing, lineHeight,
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
    appearance, chips, rows, orientation, padding, swatch, gap, spacing, lineHeight,
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
        doc.save().fillColor(resolveColor(chip.color, '#808080')).rect(columnX, rowY, layout.swatch, layout.swatch).fill().restore();
        fillText(doc, chip.label, columnX + layout.swatch + layout.gap, rowY + 1, { color: textColor });
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
      doc.save().fillColor(resolveColor(chip.color, '#808080')).rect(cursor, rowY, layout.swatch, layout.swatch).fill().restore();
      fillText(doc, chip.label, cursor + layout.swatch + layout.gap, rowY + 1, { color: textColor });
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
      const raw = Math.max(0, series.points[index]?.y ?? 0);
      if (raw <= 0) continue;
      const width = ((stacked === 'percent' ? (raw / total) * 100 : raw) / scale.max) * plot.width;
      doc.save().fillColor(resolveColor(series.color, '#808080')).rect(cursorX, centerY - barHeight / 2, width, barHeight).fill().restore();
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
  const categoryAppearance = axisFont(doc, config, chart.categoryAxis, context, data.categories.map((entry) => entry.label), slot, 7);
  const configuredWidth = Number(styleValue(chart.seriesDefs?.[0]?.customProperties?.PointWidth, context, 0.8));
  const pointWidth = Math.min(1, Math.max(0.05, Number.isFinite(configuredWidth) ? configuredWidth : 0.8));
  data.categories.forEach((category, categoryIndex) => {
    const slotX = plot.x + categoryIndex * slot;
    fillText(doc, category.label ?? '', slotX, plot.y + plot.height + 4, { width: slot, align: 'center', color: categoryAppearance.color });
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
  const points = (data.series[0]?.points || []).filter((point) => point.y && point.y > 0);
  const total = points.reduce((sum, point) => sum + point.y, 0);
  if (total <= 0) return;
  const explosionRatio = chart.exploded && points.length > 1 ? 0.1 : 0;
  const outerLimit = Math.max(10, Math.min(plot.width, plot.height) / 2 - 8);
  const radius = Math.max(10, outerLimit / (1 + explosionRatio));
  const explosion = radius * explosionRatio;
  const inner = radius * innerRatio;
  const centerX = plot.x + plot.width / 2;
  const centerY = plot.y + plot.height / 2;
  const at = (angle, r, offsetX = 0, offsetY = 0) => [
    centerX + offsetX + Math.cos(angle) * r,
    centerY + offsetY + Math.sin(angle) * r,
  ];
  setFont(doc, config, { size: 8 });
  // SSRS pie and doughnut charts use PieStartAngle=0 by default, which places the first value at
  // 3 o'clock (90 degrees clockwise from the top). A declared PieStartAngle rotates clockwise in
  // degrees; 270 therefore places the first value at 12 o'clock.
  const rawStartAngle = styleValue(chart.seriesDefs?.[0]?.customProperties?.PieStartAngle, context, 0);
  const startAngleDegrees = Number(rawStartAngle);
  let angle = (Number.isFinite(startAngleDegrees) ? startAngleDegrees : 0) * (Math.PI / 180);
  for (const point of points) {
    const sweep = (point.y / total) * Math.PI * 2;
    const middle = angle + sweep / 2;
    const offsetX = Math.cos(middle) * explosion;
    const offsetY = Math.sin(middle) * explosion;
    const steps = Math.max(2, Math.ceil((sweep / (Math.PI * 2)) * 120));
    doc.save().fillColor(resolveColor(point.color, '#808080'));
    doc.moveTo(...at(angle, radius, offsetX, offsetY));
    for (let step = 1; step <= steps; step += 1) doc.lineTo(...at(angle + (sweep * step) / steps, radius, offsetX, offsetY));
    for (let step = steps; step >= 0; step -= 1) doc.lineTo(...at(angle + (sweep * step) / steps, inner, offsetX, offsetY));
    doc.fill().restore();
    if (point.label) {
      if (/outside/i.test(point.labelPosition || '')) {
        const [lineStartX, lineStartY] = at(middle, radius * 0.92, offsetX, offsetY);
        const [lineEndX, lineEndY] = at(middle, radius + 9, offsetX, offsetY);
        const calloutColor = styleColor(chart.seriesDefs?.[0]?.customProperties?.PieLineColor, context, '#000000');
        doc.save().lineWidth(0.75).strokeColor(calloutColor).moveTo(lineStartX, lineStartY).lineTo(lineEndX, lineEndY).stroke().restore();
        const right = Math.cos(middle) >= 0;
        fillText(doc, point.label, right ? lineEndX + 2 : lineEndX - 42, lineEndY - 4, { width: 40, align: right ? 'left' : 'right' });
      } else {
        const [labelX, labelY] = at(middle, inner + (radius - inner) * 0.55, offsetX, offsetY);
        fillText(doc, point.label, labelX - 12, labelY - 4, { width: 24, align: 'center' });
      }
    }
    angle += sweep;
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
  const appearance = axisFont(doc, config, chart.categoryAxis, context, data.categories.map((entry) => entry.label), slot, 7);
  data.categories.forEach((category, index) => {
    fillText(doc, category.label ?? '', plot.x + index * slot, plot.y + plot.height + 4, { width: slot, align: 'center', color: appearance.color });
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
export function drawChart(doc, config, chart, data, x, y, width, height, context) {
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
    axisFont(doc, config, chart.categoryAxis, context, labels, slot, 7);
    const labelHeight = Math.max(0, ...labels.map((label) => doc.heightOfString(label, { width: Math.max(1, slot) })));
    bottomGutter = Math.max(16, Math.ceil(labelHeight) + 6);
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
