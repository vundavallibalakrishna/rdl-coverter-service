import { evaluateExpression } from '../rdl/expression.js';
import { pdfFont } from './fonts.js';
import { color as resolveColor, styleColor, styleSize, styleValue } from './common.js';

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

function setFont(doc, config, { size = 8, bold = false } = {}) {
  doc.font(pdfFont(config, 'Arial', bold, false)).fontSize(size);
}

function fillText(doc, text, x, y, options = {}) {
  doc.save().fillColor(options.color || LABEL_COLOR).text(String(text), x, y, { lineBreak: false, ...options }).restore();
}

function axisFont(doc, config, axis, context, labels, slot, fallback = 8) {
  const style = axis?.style || {};
  const configured = styleSize(style.fontSize, context, fallback) || fallback;
  const disabled = String(styleValue(axis?.labelsAutoFitDisabled, context, 'false')).toLowerCase() === 'true';
  let size = configured;
  setFont(doc, config, { size, bold: /bold|[6-9]00/i.test(String(styleValue(style.fontWeight, context, 'Normal'))) });
  // SSRS auto-fit can shrink/offset/rotate labels. This renderer implements the deterministic shrink part;
  // when auto-fit is disabled the declared font size is retained exactly.
  if (!disabled && slot > 0 && labels?.length) {
    const widest = Math.max(0, ...labels.map((label) => doc.widthOfString(String(label ?? ''))));
    if (widest > slot) size = Math.max(6, configured * slot / widest);
  }
  setFont(doc, config, { size, bold: /bold|[6-9]00/i.test(String(styleValue(style.fontWeight, context, 'Normal'))) });
  return { color: styleColor(style.color, context, TICK_LABEL_COLOR), size };
}

// Legend: colour swatch + label chips wrapped across the width and centred. Returns the height used.
function drawLegend(doc, config, entries, x, y, width) {
  if (!entries.length) return 0;
  setFont(doc, config, { size: 8 });
  const swatch = 10;
  const gap = 4;
  const spacing = 14;
  const lineHeight = 16;
  const chips = entries.map((entry) => ({ ...entry, chipWidth: swatch + gap + doc.widthOfString(String(entry.label)) + spacing }));
  const rows = [[]];
  let rowWidth = 0;
  for (const chip of chips) {
    if (rowWidth + chip.chipWidth > width && rows[rows.length - 1].length) { rows.push([]); rowWidth = 0; }
    rows[rows.length - 1].push(chip);
    rowWidth += chip.chipWidth;
  }
  rows.forEach((row, rowIndex) => {
    const total = row.reduce((sum, chip) => sum + chip.chipWidth, 0) - spacing;
    let cursor = x + Math.max(0, (width - total) / 2);
    const rowY = y + rowIndex * lineHeight;
    for (const chip of row) {
      doc.save().fillColor(resolveColor(chip.color, '#808080')).rect(cursor, rowY, swatch, swatch).fill().restore();
      fillText(doc, chip.label, cursor + swatch + gap, rowY + 1, { color: TICK_LABEL_COLOR });
      cursor += chip.chipWidth;
    }
  });
  return rows.length * lineHeight;
}

function drawTitle(doc, config, chart, data, x, y, width, context) {
  const caption = String(evaluateExpression(chart.title.caption, context) ?? '').trim();
  const style = chart.title.style || {};
  const height = 18;
  const background = styleColor(style.backgroundColor, context, null);
  if (background) doc.save().fillColor(background).rect(x, y, width, height).fill().restore();
  setFont(doc, config, { size: 9, bold: true });
  fillText(doc, caption, x, y + 4, { width, align: 'center', color: styleColor(style.color, context, '#000000') });
  return height;
}

function drawValueGrid(doc, config, scale, plot, orientation, axis, context) {
  const axisAppearance = axisFont(doc, config, axis, context, [], 0, 8);
  const ticks = Math.round(scale.max / scale.interval);
  for (let tick = 0; tick <= ticks; tick += 1) {
    const value = tick * scale.interval;
    if (orientation === 'horizontal') {
      const gx = plot.x + (value / scale.max) * plot.width;
      doc.save().lineWidth(0.5).strokeColor(AXIS_COLOR).moveTo(gx, plot.y).lineTo(gx, plot.y + plot.height).stroke().restore();
      fillText(doc, value, gx - 10, plot.y + plot.height + 4, { width: 20, align: 'center', color: axisAppearance.color });
    } else {
      const gy = plot.y + plot.height - (value / scale.max) * plot.height;
      doc.save().lineWidth(0.5).strokeColor(AXIS_COLOR).moveTo(plot.x, gy).lineTo(plot.x + plot.width, gy).stroke().restore();
      fillText(doc, value, plot.x - 34, gy - 4, { width: 30, align: 'right', color: axisAppearance.color });
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
function drawPieChart(doc, config, chart, data, plot, innerRatio, context) {
  const points = (data.series[0]?.points || []).filter((point) => point.y && point.y > 0);
  const total = points.reduce((sum, point) => sum + point.y, 0);
  if (total <= 0) return;
  const radius = Math.max(10, Math.min(plot.width, plot.height) / 2 - 8);
  const inner = radius * innerRatio;
  const centerX = plot.x + plot.width / 2;
  const centerY = plot.y + plot.height / 2;
  const at = (angle, r) => [centerX + Math.cos(angle) * r, centerY + Math.sin(angle) * r];
  setFont(doc, config, { size: 8 });
  let angle = -Math.PI / 2; // start at 12 o'clock
  for (const point of points) {
    const sweep = (point.y / total) * Math.PI * 2;
    const steps = Math.max(2, Math.ceil((sweep / (Math.PI * 2)) * 120));
    doc.save().fillColor(resolveColor(point.color, '#808080'));
    doc.moveTo(...at(angle, radius));
    for (let step = 1; step <= steps; step += 1) doc.lineTo(...at(angle + (sweep * step) / steps, radius));
    for (let step = steps; step >= 0; step -= 1) doc.lineTo(...at(angle + (sweep * step) / steps, inner));
    doc.fill().restore();
    if (point.label) {
      const middle = angle + sweep / 2;
      if (/outside/i.test(point.labelPosition || '')) {
        const [lineStartX, lineStartY] = at(middle, radius * 0.92);
        const [lineEndX, lineEndY] = at(middle, radius + 9);
        const calloutColor = styleColor(chart.seriesDefs?.[0]?.customProperties?.PieLineColor, context, '#000000');
        doc.save().lineWidth(0.75).strokeColor(calloutColor).moveTo(lineStartX, lineStartY).lineTo(lineEndX, lineEndY).stroke().restore();
        const right = Math.cos(middle) >= 0;
        fillText(doc, point.label, right ? lineEndX + 2 : lineEndX - 42, lineEndY - 4, { width: 40, align: right ? 'left' : 'right' });
      } else {
        const [labelX, labelY] = at(middle, inner + (radius - inner) * 0.55);
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
  doc.save().rect(x, y, width, height).clip();
  let top = y + 6;
  let bottom = y + height - 6;
  if (chart.title) top += drawTitle(doc, config, chart, data, x + 4, top, width - 8, context) + 6;
  if (!data.hasData) {
    setFont(doc, config, { size: 10, bold: true });
    fillText(doc, chart.noDataMessage || 'No Data Available', x, y + height / 2 - 6, { width, align: 'center', color: TICK_LABEL_COLOR });
    doc.restore();
    return;
  }
  if (chart.legend?.visible && data.legend.length) {
    const band = measureLegendHeight(doc, config, data.legend, width - 16);
    drawLegend(doc, config, data.legend, x + 8, bottom - band, width - 16);
    bottom -= band + 8;
  }
  // Plot area, leaving gutters for axis/category labels.
  const circular = chart.chartType === 'pie' || chart.chartType === 'doughnut';
  const leftGutter = chart.chartType === 'bar' ? 165 : 40;
  const bottomGutter = circular ? 0 : 16;
  const plot = { x: x + leftGutter, y: top + 6, width: width - leftGutter - 24, height: bottom - top - 6 - bottomGutter };
  if (plot.width > 10 && plot.height > 10) {
    const stacked = chart.stacked || 'none';
    if (chart.chartType === 'bar') drawBarChart(doc, config, chart, data, plot, stacked, context);
    else if (chart.chartType === 'column') drawColumnChart(doc, config, chart, data, plot, stacked, context);
    else if (chart.chartType === 'pie') drawPieChart(doc, config, chart, data, plot, 0, context);
    else if (chart.chartType === 'doughnut') drawPieChart(doc, config, chart, data, plot, 0.55, context);
    else if (chart.chartType === 'line') drawLineChart(doc, config, chart, data, plot, context);
    else if (chart.chartType === 'area') drawAreaChart(doc, config, chart, data, plot, stacked, context);
    else if (chart.chartType === 'scatter') drawScatterChart(doc, config, chart, data, plot, context);
  }
  doc.restore();
}

// Pre-measures wrapped legend height so the plot area can reserve exactly the right band.
function measureLegendHeight(doc, config, entries, width) {
  if (!entries.length) return 0;
  setFont(doc, config, { size: 8 });
  const swatch = 10; const gap = 4; const spacing = 14; const lineHeight = 16;
  let rows = 1; let rowWidth = 0;
  for (const entry of entries) {
    const chipWidth = swatch + gap + doc.widthOfString(String(entry.label)) + spacing;
    if (rowWidth + chipWidth > width && rowWidth > 0) { rows += 1; rowWidth = 0; }
    rowWidth += chipWidth;
  }
  return rows * lineHeight;
}
