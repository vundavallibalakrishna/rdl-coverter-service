import { evaluateExpression } from '../rdl/expression.js';
import { filterMatches } from '../rdl/validation.js';
import { color as resolveNamedColor } from './common.js';

// Deterministic fallback palette for chart points whose colour expression evaluates to Nothing.
const FALLBACK_PALETTE = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#264478', '#9e480e'];

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

function pointColor(expression, row, base, index) {
  const evaluated = expression ? evalIn(expression, row, base) : null;
  if (evaluated === null || evaluated === undefined || evaluated === '') return FALLBACK_PALETTE[index % FALLBACK_PALETTE.length];
  return resolveNamedColor(String(evaluated), FALLBACK_PALETTE[index % FALLBACK_PALETTE.length]);
}

function pointLabel(dataLabel, y, row, base) {
  if (!dataLabel?.visible) return null;
  if (dataLabel.useValueAsLabel) return y === null ? '' : String(y);
  if (dataLabel.expression) {
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
  const seriesDef = chart.seriesDefs[0] || {};

  const categories = groupsFor(chart.category, rows, base);
  const grouped = Boolean(chart.series?.group);
  const seriesRows = grouped
    ? rows.filter((row) => (chart.series.filters || []).every((filter) => filterMatches(filter, row, parameters, globals, rows, datasetsByName)))
    : rows;
  const seriesGroups = grouped ? groupsFor(chart.series, seriesRows, base) : [{ label: seriesDef.name, rows: seriesRows }];

  const series = seriesGroups.map((group, seriesIndex) => {
    const groupSet = new Set(group.rows);
    // For a grouped chart the colour is a property of the SERIES (e.g. Incident Type), so evaluate it
    // once from a row of this series. Evaluating per-cell would mis-colour empty cells, whose fallback
    // row belongs to a different series.
    const seriesColor = grouped ? pointColor(seriesDef.color, group.rows[0] || rows[0] || {}, base, seriesIndex) : null;
    const points = categories.map((category, categoryIndex) => {
      const cellRows = category.rows.filter((row) => groupSet.has(row));
      const first = cellRows[0] || category.rows[0] || rows[0] || {};
      const context = { parameters, globals, datasets: datasetsByName, dataset: cellRows, fields: first };
      const raw = evaluateExpression(seriesDef.y, context);
      const numeric = raw === null || raw === undefined || raw === '' ? null : Number(raw);
      const y = Number.isFinite(numeric) ? numeric : null;
      const color = grouped ? seriesColor : pointColor(seriesDef.color, first, base, categoryIndex);
      return {
        y,
        color,
        label: pointLabel(seriesDef.dataLabel, y, first, base),
        labelPosition: seriesDef.dataLabel?.position || 'Auto',
      };
    });
    return { label: group.label ?? seriesDef.name, color: seriesColor, points };
  });

  const legend = grouped
    ? series.map((entry) => ({ label: entry.label, color: entry.color || FALLBACK_PALETTE[0] }))
    : categories.map((category, index) => ({ label: category.label, color: series[0]?.points[index]?.color || FALLBACK_PALETTE[index % FALLBACK_PALETTE.length] }));

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
