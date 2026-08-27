import { evaluateExpression } from '../rdl/expression.js';
import { filterMatches } from '../rdl/validation.js';
import { color as resolveNamedColor } from './common.js';

// Legacy deterministic fallback retained for reports that explicitly request the RDL "Default" palette.
const FALLBACK_PALETTE = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#264478', '#9e480e'];

// Reporting Services' modern default palette. Shape charts consume one colour per data point; other
// charts consume one colour per series. Keeping this named palette separate from the fallback makes
// Palette=Pacific deterministic and lets an explicit point/series colour continue to take precedence.
const PACIFIC_PALETTE = [
  '#01b8aa', '#374649', '#fd625e', '#f2c80f', '#5f6b6d',
  '#8ad4eb', '#fe9666', '#a66999', '#3599b8', '#dfbfbf',
];

function evalIn(expression, row, base) {
  return evaluateExpression(expression, { ...base, fields: row || {} });
}

function labelFor(member, row, base) {
  if (!member?.label) return null;
  const value = evalIn(member.label, row, base);
  return value === null || value === undefined ? '' : String(value);
}

function compareBySort(sorts, rowA, rowB, base) {
  for (const sort of sorts) {
    const left = evalIn(sort.value, rowA, base);
    const right = evalIn(sort.value, rowB, base);
    const compared = typeof left === 'string' || typeof right === 'string'
      ? String(left ?? '').localeCompare(String(right ?? ''))
      : Number(left ?? 0) - Number(right ?? 0);
    if (compared !== 0) return /^desc/i.test(sort.direction) ? -compared : compared;
  }
  return 0;
}

// Groups rows by a chart member's group expressions, labelled and sorted per the member. A member
// without a group (a single static series/category) yields one group covering all rows.
function groupsFor(member, rows, base) {
  if (!member?.group?.expressions?.length) {
    return [{ key: 'all', label: labelFor(member, rows[0], base), rows }];
  }
  const map = new Map();
  for (const row of rows) {
    const key = JSON.stringify(member.group.expressions.map((expression) => evalIn(expression, row, base)));
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(row);
  }
  const groups = [...map.values()].map((groupRows) => ({ label: labelFor(member, groupRows[0], base), rows: groupRows }));
  if (member.sortExpressions?.length) groups.sort((left, right) => compareBySort(member.sortExpressions, left.rows[0], right.rows[0], base));
  return groups;
}

function chartPalette(chart, base) {
  const requested = String(evaluateExpression(chart.palette || 'Default', base) ?? 'Default').trim();
  if (/^custom$/i.test(requested)) {
    const colors = (chart.customPaletteColors || [])
      .map((entry) => evaluateExpression(entry, base))
      .filter((entry) => entry !== null && entry !== undefined && entry !== '')
      .map((entry) => resolveNamedColor(String(entry), '#ffffff'));
    // MS-RDL specifies white when Palette=Custom has no custom colour collection.
    return colors.length ? colors : ['#ffffff'];
  }
  if (/^pacific(?:light|semitransparent)?$/i.test(requested)) return PACIFIC_PALETTE;
  return FALLBACK_PALETTE;
}

function pointColor(expression, row, base, index, palette) {
  const evaluated = expression ? evalIn(expression, row, base) : null;
  const fallback = palette[index % palette.length];
  if (evaluated === null || evaluated === undefined || evaluated === '') return fallback;
  return resolveNamedColor(String(evaluated), fallback);
}

function pointLabel(dataLabel, y, row, base) {
  if (!dataLabel?.visible) return null;
  if (dataLabel.useValueAsLabel) return y === null ? '' : String(y);
  if (dataLabel.expression) {
    // SSRS's chart keyword binds the label to the resolved Y value, not a literal string.
    if (/^#VALY$/i.test(String(dataLabel.expression).trim())) return y === null ? '' : String(y);
    const value = evalIn(dataLabel.expression, row, base);
    return value === null || value === undefined ? '' : String(value);
  }
  return null;
}

// Turns a parsed chart plus its (already field-normalized) dataset rows into a render-ready model:
// categories, one entry per series, and a legend. Y values are evaluated per (series, category) cell
// with the cell's rows as the aggregate scope, mirroring SSRS.
export function materializeChart(chart, datasetsByName, parameters = {}, globals = {}) {
  const rows = datasetsByName[chart.datasetName] || [];
  const base = { parameters, globals, datasets: datasetsByName, dataset: rows, fields: {} };
  const seriesDefs = chart.seriesDefs.length ? chart.seriesDefs : [{}];
  const palette = chartPalette(chart, base);

  const categories = groupsFor(chart.category, rows, base);
  const grouped = Boolean(chart.series?.group);
  const seriesRows = grouped
    ? rows.filter((row) => (chart.series.filters || []).every((filter) => filterMatches(filter, row, parameters, globals, rows, datasetsByName)))
    : rows;
  const seriesGroups = grouped ? groupsFor(chart.series, seriesRows, base) : [];

  // A shape chart (pie/doughnut) consumes one palette colour and one legend entry PER DATA POINT, so an
  // empty point changes what every later point looks like. Other chart types key the legend/palette off
  // the series, where an empty cell only leaves a gap in that series' own line of points.
  const shapeChart = chart.chartType === 'pie' || chart.chartType === 'doughnut';

  // A static series hierarchy has one chart series for every ChartSeries definition. A dynamic
  // hierarchy instead expands the declared definition across its data-driven member instances.
  const entries = grouped
    ? seriesGroups.map((group, index) => ({ group, definition: seriesDefs[index % seriesDefs.length], index }))
    : seriesDefs.map((definition, index) => ({
      group: { label: chart.staticSeriesLabels?.[index] ?? definition.name, rows: seriesRows }, definition, index,
    }));
  const series = entries.map(({ group, definition: seriesDef, index: seriesIndex }) => {
    const groupSet = new Set(group.rows);
    // For a grouped chart the colour is a property of the SERIES (e.g. Incident Type), so evaluate it
    // once from a row of this series. Evaluating per-cell would mis-colour empty cells, whose fallback
    // row belongs to a different series.
    const seriesColor = (!shapeChart || grouped)
      ? pointColor(seriesDef.color, group.rows[0] || rows[0] || {}, base, seriesIndex, palette)
      : null;
    let palettePosition = 0;
    const points = categories.map((category, categoryIndex) => {
      const cellRows = category.rows.filter((row) => groupSet.has(row));
      // A missing dynamic series/category intersection is an empty point. Falling back to the
      // category's first row leaks another series' value into this cell (every stacked segment
      // becomes identical), which is unlike SSRS and corrupts palettes, labels, and totals.
      const first = cellRows[0] || {};
      const context = { parameters, globals, datasets: datasetsByName, dataset: cellRows, fields: first };
      const raw = evaluateExpression(seriesDef.y, context);
      const numeric = raw === null || raw === undefined || raw === '' ? null : Number(raw);
      const y = Number.isFinite(numeric) ? numeric : null;
      // SSRS calls a point whose value is Nothing/null an EMPTY POINT. On a shape chart it draws no
      // slice, takes no legend entry, and consumes no palette colour, so the points after it keep the
      // colours they would have had if the empty point had never been in the category list. A zero is a
      // real value, not an empty point: it keeps its colour and legend entry and draws a zero-width slice.
      const empty = y === null;
      const paletteIndex = shapeChart ? palettePosition : categoryIndex;
      const color = shapeChart && !grouped
        ? pointColor(seriesDef.color, first, base, paletteIndex, palette)
        : seriesColor;
      if (!(shapeChart && empty)) palettePosition += 1;
      return {
        y,
        empty,
        color,
        label: pointLabel(seriesDef.dataLabel, y, first, base),
        labelPosition: seriesDef.dataLabel?.position || 'Auto',
        labelStyle: seriesDef.dataLabel?.style || {},
      };
    });
    return { label: group.label ?? seriesDef.name, color: seriesColor, points };
  });

  const legend = grouped || !shapeChart
    ? series.map((entry, index) => ({ label: entry.label, color: entry.color || palette[index % palette.length] }))
    : categories.reduce((entries, category, index) => {
      const point = series[0]?.points[index];
      if (shapeChart && point?.empty) return entries;
      entries.push({ label: category.label, color: point?.color || palette[index % palette.length] });
      return entries;
    }, []);

  const maxY = Math.max(0, ...series.flatMap((entry) => entry.points.map((point) => point.y || 0)));
  return {
    chartType: chart.chartType,
    categories: categories.map((category) => ({ label: category.label })),
    series,
    legend,
    maxY,
    hasData: series.some((entry) => entry.points.some((point) => point.y !== null && point.y !== 0)),
  };
}
