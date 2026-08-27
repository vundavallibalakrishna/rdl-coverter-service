import ExcelJS from 'exceljs';
import ExcelRange from 'exceljs/lib/doc/range.js';
import PDFDocument from 'pdfkit';
import fs from 'node:fs/promises';
import path from 'node:path';
import { ServiceError } from '../errors.js';
import { resolveExcelLayoutMode } from '../excelLayoutMode.js';
import { evaluateExpression, resolveReportCulture } from '../rdl/expression.js';
import { parseDateValue } from '../rdl/dateValue.js';
import { resolveSubreportInvocation } from '../rdl/validation.js';
import { cellBorderStyle, cellText, cellTextbox, color, enforcedBottomBorder, isDateLikeValue, isFreeFormCell, isHidden, matchingChangedGroupOwnerRowBoundary, materializedCellContext, materializedCellVisualSignature, normalizeDatasets, shouldEnforceTablixBottom, styledTextForItem, styleColor, styleSize, styleText, styleValue, tablixRows, textForItem } from './common.js';
import { computeCellPlacements } from './tableGrid.js';
import { resolveGridColumns } from './tableLayout.js';
import { materializeChart } from './chartData.js';
import { renderChartPng } from './chartImage.js';
import { measureTextboxHeight } from './pdf.js';
import { buildGridBoundaries } from './gridBoundaries.js';
import { DEFAULT_EXCEL_DATETIME_FORMAT, excelDateFormat, excelNumberFormat, cellString } from './excelFormat.js';
const SHEET_NAME_FORBIDDEN = /[\\/?*[\]:]/g;
const DEFAULT_ROW_POINTS = 15;
const COINCIDENT_EDGE_TOLERANCE_PT = 0.25;
const EXCEL_LAYOUT_DPI = 96;
const EXCEL_MAX_DIGIT_WIDTH_PX = 7;
// Cell content that is a picture rather than a cell value. A tablix cell whose CellContents Rectangle was
// flattened is a free-form canvas and can carry these alongside its textbox.

const EXCEL_TEXT_CLEARANCE_PT = 2;
// OOXML stores column widths in units derived from the workbook's maximum digit width. Excel then rounds
// those widths again for the active viewer/font metrics. Reserving one standard 7-pixel digit cell during
// measurement covers that character-unit conversion without changing the RDL column proportions.
const EXCEL_MAX_DIGIT_WIDTH_PT = EXCEL_MAX_DIGIT_WIDTH_PX * (72 / EXCEL_LAYOUT_DPI);
// Merge overlap checks are spatially bucketed by column and row interval. ExcelJS otherwise compares every
// new merge with every existing merge, which makes large grouped REPORT workbooks quadratic in the number
// of native merges. Sixty-four rows keeps each lookup bounded without depending on any report geometry.
const MERGE_ROW_BUCKET_SIZE = 64;

function sheetName(name, fallback) {
  const cleaned = String(name || fallback).replace(SHEET_NAME_FORBIDDEN, ' ').trim().slice(0, 31);
  return cleaned || fallback;
}

function hex(value, fallback = null) {
  if (!value) return fallback;
  const resolved = color(String(value));
  if (!resolved || resolved === 'transparent') return fallback;
  return resolved.replace('#', '').toUpperCase().padStart(6, '0').slice(0, 6);
}

function imageExtension(buffer) {
  if (buffer.length > 8 && buffer[0] === 0x89 && buffer.toString('ascii', 1, 4) === 'PNG') return 'png';
  if (buffer.length > 3 && buffer[0] === 0xFF && buffer[1] === 0xD8) return 'jpeg';
  if (buffer.length > 6 && buffer.toString('ascii', 0, 3) === 'GIF') return 'gif';
  return 'png';
}

function excelBorderSide(border, context) {
  const style = String(styleValue(border?.style, context, 'None'));
  if (!style || /^none$/i.test(style)) return undefined;
  const weight = /double/i.test(style) ? 'double' : /dash/i.test(style) ? 'dashed' : /dot/i.test(style) ? 'dotted' : 'thin';
  return { style: weight, color: { argb: `FF${hex(styleValue(border?.color, context, '#000000'), '000000')}` } };
}

// A textbox whose whole content is a single expression, so its raw typed value can be recovered for a live
// Excel number/date. SSRS commonly wraps one value between empty paragraphs (["", "=Fields!X.Value", ""]),
// so we count NON-EMPTY runs and accept exactly one; cells with several real runs, markup, or aggregated
// items fall back to display text.
function singleExpression(textbox) {
  if (Array.isArray(textbox?.paragraphs)) {
    const runs = textbox.paragraphs.flat().filter((run) => {
      const value = run?.value ?? run;
      return typeof value === 'string' ? value.trim() !== '' : value !== null && value !== undefined;
    });
    if (runs.length !== 1) return null;
    const value = runs[0]?.value ?? runs[0];
    return typeof value === 'string' ? value.trim() : null;
  }
  return typeof textbox?.value === 'string' && textbox.value.trim() !== '' ? textbox.value.trim() : null;
}

// Recovers the value behind a pure display-formatting call so Excel still gets a live number + format code.
// SSRS authors routinely write =Format(Fields!X.Value, "N2") or =FormatNumber(...) only for presentation;
// the underlying number is what a caller pivots on. Returns { inner, format } or null.
function unwrapFormatCall(expression) {
  const body = String(expression).replace(/^=/, '').trim();
  const twoArg = /^Format\s*\((.+),\s*"([^"]*)"\s*\)$/is.exec(body);
  if (twoArg) return { inner: `=${twoArg[1].trim()}`, format: twoArg[2] };
  const named = /^Format(Number|Currency|Percent)\s*\((.+?)(?:,\s*(\d+))?\s*\)$/is.exec(body);
  if (named) {
    const digits = named[3] ?? '2';
    const kind = named[1].toLowerCase();
    const code = kind === 'currency' ? `C${digits}` : kind === 'percent' ? `P${digits}` : `N${digits}`;
    return { inner: `=${named[2].trim()}`, format: code };
  }
  return null;
}

// The Excel value + optional number format for a materialized cell. Numeric and date fields are written as
// their native typed value so they stay computable; everything else is injection-safe display text.
function excelCellValue(cell, context) {
  const display = cellText(cell);
  // `values` is the authoritative materialized display. In particular, HideDuplicates deliberately turns
  // a repeated expression result into an empty string. Re-evaluating a numeric/date expression after that
  // suppression resurrects the duplicate in Excel even though PDF and Word correctly keep it blank.
  if (display === '') return { value: '' };
  const textbox = cellTextbox(cell);
  const expression = singleExpression(textbox);
  if (expression === null) return { value: cellString(display) };
  const unwrapped = unwrapFormatCall(expression);
  const valueExpression = unwrapped ? unwrapped.inner : expression;
  let raw;
  try {
    raw = valueExpression.startsWith('=') ? evaluateExpression(valueExpression, context) : valueExpression;
  } catch {
    return { value: cellString(display) };
  }
  const format = unwrapped ? unwrapped.format : styleValue(textbox?.style?.format, context, null);
  if (typeof raw === 'number' && Number.isFinite(raw)) {
    return { value: raw, numFmt: excelNumberFormat(format) || undefined };
  }
  // A DateTime value (a Date, or an ISO-8601 string from parameter/field coercion) becomes a live typed
  // Excel date so the workbook carries a real date, matching the PDF/DOCX text treatment of the same value.
  const asDate = raw instanceof Date ? raw : (isDateLikeValue(raw) ? parseDateValue(raw) : null);
  if (asDate && !Number.isNaN(asDate.getTime())) {
    // No format resolves to general date/time (date and time) to match SSRS and the PDF/DOCX text default.
    if (!format) return { value: asDate, numFmt: DEFAULT_EXCEL_DATETIME_FORMAT };
    // An explicit format is translated to a live Excel number format so the cell stays a typed date and
    // still displays what the author asked for; an untranslatable format writes the exact formatted string.
    const excelFmt = excelDateFormat(format);
    if (excelFmt) return { value: asDate, numFmt: excelFmt };
    return { value: cellString(display) };
  }
  return { value: cellString(display) };
}

// Longest single line of a cell's display text — the width the column needs before wrapping.
function displayWidth(cell) {
  return Math.max(0, ...String(cellText(cell) ?? '').split('\n').map((line) => line.length), 0);
}

// Writes one tablix as a block of styled rows starting at `startRow` (1-based). Accumulates the widest
// content per grid column into `columnMaxChars` (single-column cells only, so merged headers do not inflate
// widths) and returns the number of rows consumed.
function writeTablix(worksheet, model, item, request, globals, startRow, columnMaxChars) {
  const { rows, columns } = tablixRows(item, request, globals, model);
  const placements = computeCellPlacements(rows, columns.length);
  const datasets = normalizeDatasets(model, request);
  const merges = [];
  for (let r = 0; r < rows.length; r += 1) {
    const row = rows[r];
    const excelRowNumber = startRow + r;
    worksheet.getRow(excelRowNumber).height = row.height || DEFAULT_ROW_POINTS;
    for (let ci = 0; ci < row.cells.length; ci += 1) {
      const cell = row.cells[ci];
      if (cell.hidden) continue;
      const columnIndex = placements[r][ci];
      if (columnIndex === undefined || columnIndex < 0) continue;
      const textbox = cellTextbox(cell);
      // Rectangle-only cell wrappers affect the cell/grid perimeter, not the text they contain. Keep the
      // visible textbox as the content-style authority while retaining the tablix as the border authority.
      // Conflating the two drops paragraph/run formatting (font, size, colour and alignment) from wrapped
      // symbols and labels even though PDF correctly renders the inner textbox.
      const style = textbox?.style || item.style;
      const borderStyle = cellBorderStyle(cell, item);
      const context = { fields: row.fields, parameters: request.parameters || {}, globals, dataset: datasets[item.datasetName] || [], datasets };
      const target = worksheet.getCell(excelRowNumber, columnIndex + 1);
      const { value, numFmt } = excelCellValue(cell, context);
      target.value = value;
      if (numFmt) target.numFmt = numFmt;
      if ((cell.colSpan || 1) === 1) columnMaxChars[columnIndex] = Math.max(columnMaxChars[columnIndex] || 0, displayWidth(cell));
      const fill = hex(styleColor(style.backgroundColor, context, null));
      if (fill) target.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${fill}` } };
      target.font = {
        name: String(styleText(style.fontFamily, context, 'Arial')),
        size: styleSize(style.fontSize, context, 10) || 10,
        bold: /bold|[6-9]00/i.test(String(styleValue(style.fontWeight, context, 'Normal'))),
        italic: /italic/i.test(String(styleValue(style.fontStyle, context, 'Normal'))),
        color: { argb: `FF${hex(styleColor(style.color, context, '#000000'), '000000')}` },
      };
      // VerticalAlign/TextAlign can be expressions; resolve before matching so the regex does not test the
      // raw expression source (`=IIF(x,"Right","Left")` contains "Right").
      const vAlign = String(styleValue(style.verticalAlign, context, '') || '');
      const hAlign = String(styleValue(style.textAlign, context, '') || '');
      target.alignment = {
        vertical: /bottom/i.test(vAlign) ? 'bottom' : /middle|center/i.test(vAlign) ? 'middle' : 'top',
        horizontal: /center/i.test(hAlign) ? 'center' : /right/i.test(hAlign) ? 'right' : /justify/i.test(hAlign) ? 'justify' : 'left',
        wrapText: true,
      };
      const borders = borderStyle?.borders || {};
      target.border = {
        top: excelBorderSide(borders.top, context),
        bottom: excelBorderSide(borders.bottom, context),
        left: excelBorderSide(borders.left, context),
        right: excelBorderSide(borders.right, context),
      };
      const colSpan = cell.colSpan || 1;
      const rowSpan = cell.rowSpan || 1;
      if (colSpan > 1 || rowSpan > 1) {
        merges.push([excelRowNumber, columnIndex + 1, excelRowNumber + rowSpan - 1, columnIndex + colSpan]);
      }
    }
  }
  // A merge that overlaps an earlier one (malformed span) is skipped rather than throwing — the underlying
  // cells still hold their values, so no data is lost.
  for (const [top, left, bottom, right] of merges) {
    try {
      worksheet.mergeCells(top, left, bottom, right);
    } catch {
      // overlapping merge; leave the cells unmerged
    }
  }
  return { rowsConsumed: rows.length };
}

function autofitColumns(worksheet, columnMaxChars) {
  columnMaxChars.forEach((chars, index) => {
    if (!chars) return;
    worksheet.getColumn(index + 1).width = Math.min(60, Math.max(6, chars + 1.5));
  });
}

function embedImage(workbook, worksheet, buffer, startRow, widthPoints, heightPoints) {
  const imageId = workbook.addImage({ buffer, extension: imageExtension(buffer) });
  const widthPx = Math.round((widthPoints / 72) * 96);
  const heightPx = Math.round((heightPoints / 72) * 96);
  worksheet.addImage(imageId, {
    tl: { col: 0, row: startRow - 1 },
    ext: { width: widthPx, height: heightPx },
    editAs: 'oneCell',
  });
  return Math.max(1, Math.ceil(heightPoints / DEFAULT_ROW_POINTS));
}

async function writeItem(workbook, worksheet, model, item, request, globals, context, config, tempDir, cursor, columnMaxChars, chartIndex) {
  if (isHidden(item.hidden, context)) return { rows: 0, chartIndex };
  if (item.type === 'Tablix') {
    const { rowsConsumed } = writeTablix(worksheet, model, item, request, globals, cursor, columnMaxChars);
    return { rows: rowsConsumed, chartIndex };
  }
  if (item.type === 'Textbox') {
    const text = cellString(textForItem(item, context));
    const target = worksheet.getCell(cursor, 1);
    target.value = text;
    columnMaxChars[0] = Math.max(columnMaxChars[0] || 0, ...text.split('\n').map((line) => line.length));
    target.font = {
      name: String(styleText(item.style?.fontFamily, context, 'Arial')),
      size: styleSize(item.style?.fontSize, context, 10) || 10,
      bold: /bold|[6-9]00/i.test(String(styleValue(item.style?.fontWeight, context, 'Normal'))),
      color: { argb: `FF${hex(styleColor(item.style?.color, context, '#000000'), '000000')}` },
    };
    return { rows: 1, chartIndex };
  }
  if (item.type === 'Image') {
    // Image Value can be an expression (=Fields!Logo.Value); resolve before the embeddedImages lookup.
    const image = model.embeddedImages?.[styleValue(item.value, context, item.value)];
    if (!image?.data) return { rows: 0, chartIndex };
    const buffer = Buffer.from(image.data.replace(/\s+/g, ''), 'base64');
    return { rows: embedImage(workbook, worksheet, buffer, cursor, item.width || 100, item.height || 40), chartIndex };
  }
  if (item.type === 'Chart') {
    const data = materializeChart(item, context.datasets, context.parameters, context.globals);
    const png = await renderChartPng(item, data, config, tempDir, context, chartIndex);
    if (!png) return { rows: 0, chartIndex: chartIndex + 1 };
    return { rows: embedImage(workbook, worksheet, png.data, cursor, item.width || 400, item.height || 250), chartIndex: chartIndex + 1 };
  }
  if (item.type === 'Rectangle') {
    let rows = 0;
    let nextChart = chartIndex;
    for (const child of [...(item.items || [])].sort((a, b) => (a.top || 0) - (b.top || 0) || (a.left || 0) - (b.left || 0))) {
      const result = await writeItem(workbook, worksheet, model, child, request, globals, context, config, tempDir, cursor + rows, columnMaxChars, nextChart);
      rows += result.rows;
      nextChart = result.chartIndex;
    }
    return { rows, chartIndex: nextChart };
  }
  // Line and any other body construct have no meaningful cell-grid representation; skip without failing.
  return { rows: 0, chartIndex };
}

function newSheet(workbook, model, name) {
  return workbook.addWorksheet(name, {
    views: [{ state: 'normal' }],
    pageSetup: { orientation: model.page.width > model.page.height ? 'landscape' : 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
}

// Writes a list of items down one worksheet (stacked blocks, one spacer row between) and autofits its
// columns. The chart counter is shared across the workbook so per-chart temp filenames never collide.
async function fillSheet(workbook, worksheet, model, items, request, globals, context, config, tempDir, chartCounter) {
  const columnMaxChars = [];
  let cursor = 1;
  for (const item of items) {
    const result = await writeItem(workbook, worksheet, model, item, request, globals, context, config, tempDir, cursor, columnMaxChars, chartCounter.value);
    cursor += result.rows;
    chartCounter.value = result.chartIndex;
    if (result.rows > 0) cursor += 1; // one blank spacer row between blocks
  }
  autofitColumns(worksheet, columnMaxChars);
  return Math.max(0, cursor - 1);
}

async function renderDataExcel(model, request, config, tempDir) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'RDL Converter Service';
  workbook.title = request.outputFileName || model.name || 'Report';
  const globals = { PageNumber: 1, TotalPages: 1, ReportName: request.outputFileName || model.name, ExecutionTime: new Date(), variables: model.variables || {}, culture: resolveReportCulture(model, { parameters: request.parameters || {} }) };
  const context = { parameters: request.parameters || {}, globals, fields: {}, dataset: [], datasets: normalizeDatasets(model, request) };
  const chartCounter = { value: 0 };

  // Page-header band (logo, report title) leads so the export is recognizable; the page footer is
  // page-number/date chrome that has no place in a continuous sheet and is intentionally omitted.
  const bandItems = [...(model.page.header?.items || [])].sort((a, b) => (a.top || 0) - (b.top || 0) || (a.left || 0) - (b.left || 0));
  const bodyItems = [...(model.body.items || [])].sort((a, b) => (a.top || 0) - (b.top || 0) || (a.left || 0) - (b.left || 0) || (a.zIndex || 0) - (b.zIndex || 0));

  let rowCount = 0;
  // Opt-in: one worksheet per tablix. Each table then owns its own column widths, so tables with different
  // column counts no longer have to share a single compromised grid. Non-tablix content (title band, charts,
  // free-form textboxes) collects onto a leading "Overview" sheet.
  if (request.excel?.sheetPerTablix === true) {
    const tablixes = bodyItems.filter((item) => item.type === 'Tablix' && !isHidden(item.hidden, context));
    const others = [...bandItems, ...bodyItems.filter((item) => item.type !== 'Tablix')];
    if (others.length) {
      rowCount += await fillSheet(workbook, newSheet(workbook, model, 'Overview'), model, others, request, globals, context, config, tempDir, chartCounter);
    }
    for (let index = 0; index < tablixes.length; index += 1) {
      rowCount += await fillSheet(workbook, newSheet(workbook, model, `Table ${index + 1}`), model, [tablixes[index]], request, globals, context, config, tempDir, chartCounter);
    }
    if (workbook.worksheets.length === 0) newSheet(workbook, model, 'Report'); // never emit an empty workbook
  } else {
    const worksheet = newSheet(workbook, model, sheetName(request.outputFileName || model.name, 'Report'));
    rowCount = await fillSheet(workbook, worksheet, model, [...bandItems, ...bodyItems], request, globals, context, config, tempDir, chartCounter);
  }

  const buffer = Buffer.from(await workbook.xlsx.writeBuffer());
  return {
    buffer,
    pageCount: null, // a spreadsheet is continuous; Excel decides print pagination itself
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    extension: 'xlsx',
    layoutMode: 'spreadsheet',
    sheetCount: workbook.worksheets.length,
    rowCount,
  };
}

const POINT_PRECISION = 4; // quarter-point geometry, matching the PDF shared-edge coordinate key
const MAX_EXCEL_ROW_HEIGHT = 409;
// Excel reserves a fixed horizontal inset inside every cell — the same 5 device pixels excelWidthFromPoints
// adds back when converting a point width into a column width — so the width available to wrapped text is
// narrower than the column the item was given.
const EXCEL_CELL_TEXT_INSET_PT = 5 * (72 / 96);
// How much narrower a box is re-flowed at before its content counts as sitting on the wrap point. PDF
// measurement uses the font file's fractional glyph advances; Excel lays the same string out with
// integer-pixel advances at display resolution, and those roundings accumulate over a long line. This only
// decides whether to reserve one line of slack in a merged cell that cannot overflow — it never shrinks
// anything — so it is set well clear of the disagreement rather than tuned to one string.
const EXCEL_WRAP_SAFETY_FRACTION = 0.1;

function point(value) {
  return Math.round(Number(value || 0) * POINT_PRECISION) / POINT_PRECISION;
}

function visibleBorderWidth(border, context = {}) {
  const style = String(styleValue(border?.style, context, 'None'));
  return /^none$/i.test(style) ? 0 : Math.max(0, styleSize(border?.width, context, 1));
}

// RDL designers commonly make adjacent free-form boxes overlap by exactly their shared border width so
// raster/fixed renderers draw one continuous rule. Excel cannot represent two overlapping merged ranges.
// Record only sibling edges whose vertical spans overlap and whose horizontal difference is no greater
// than the visible shared-edge stroke. Those two physical coordinates then become one native grid edge.
// Larger/intentional overlaps remain distinct and still fail closed in mergeSafe.
//
// Every vertically-overlapping sibling pair is examined, not only adjacent entries in the left-sorted
// order. A sibling list interleaves items from different vertical bands — a report whose whole body is
// hoisted into one list routinely puts an unrelated data region's left edge between two boxes that
// genuinely share a rule — so the box that actually shares an edge is usually not the immediate
// predecessor, and its shared stroke went unrecorded until the merge itself failed closed.
function collectEquivalentXEdgePairs(items, pairs, context, parentLeft = 0, parentTop = 0) {
  const positioned = (items || []).map((item) => {
    // Match collectXBoundaries/renderFreeformItem exactly: round the positioned origin first, then add
    // dimensions. Rounding the combined expression instead creates a different quarter-point edge.
    const left = point(parentLeft + (item.left || 0));
    const top = point(parentTop + (item.top || 0));
    return {
      item,
      left,
      right: point(left + (item.width || 0)),
      top,
      bottom: point(top + (item.height || 0)),
      rightBorder: visibleBorderWidth(item.style?.borders?.right, context),
      leftBorder: visibleBorderWidth(item.style?.borders?.left, context),
    };
  }).sort((left, right) => left.left - right.left || left.top - right.top);

  // No pair can be equivalent once the gap exceeds the widest stroke any sibling declares, so the forward
  // scan stops there instead of comparing every item with every other one.
  const reach = Math.max(
    1 / POINT_PRECISION,
    ...positioned.map((entry) => Math.max(entry.rightBorder, entry.leftBorder)),
  );
  for (let index = 0; index < positioned.length; index += 1) {
    const previous = positioned[index];
    for (let other = index + 1; other < positioned.length; other += 1) {
      const current = positioned[other];
      if (current.left - previous.right > reach) break;
      const verticallyOverlaps = previous.top < current.bottom && current.top < previous.bottom;
      if (!verticallyOverlaps) continue;
      const tolerance = Math.max(1 / POINT_PRECISION, previous.rightBorder, current.leftBorder);
      if (Math.abs(previous.right - current.left) <= tolerance) {
        pairs.push({ from: previous.right, to: current.left, tolerance });
      }
    }
  }

  for (const entry of positioned) {
    collectEquivalentXEdgePairs(entry.item.items, pairs, context, entry.left, entry.top);
  }
}

// Resolves the recorded pairs into one canonical coordinate per equivalence cluster. Clustering rather
// than writing each pair straight into the map keeps the result independent of the order the pairs were
// discovered in: three boxes meeting at one rule contribute two pairs sharing a coordinate, and
// last-write-wins would otherwise let the second pair strand the first one on a stale canonical edge.
function resolveEquivalentXEdges(pairs) {
  const parent = new Map();
  const extent = new Map();
  const find = (value) => {
    let root = value;
    while (parent.get(root) !== root) root = parent.get(root);
    let cursor = value;
    while (parent.get(cursor) !== root) {
      const next = parent.get(cursor);
      parent.set(cursor, root);
      cursor = next;
    }
    return root;
  };
  const add = (value) => {
    if (parent.has(value)) return;
    parent.set(value, value);
    extent.set(value, [value, value]);
  };
  for (const { from, to, tolerance } of pairs) {
    add(from);
    add(to);
    const fromRoot = find(from);
    const toRoot = find(to);
    if (fromRoot === toRoot) continue;
    const [fromLow, fromHigh] = extent.get(fromRoot);
    const [toLow, toHigh] = extent.get(toRoot);
    const low = Math.min(fromLow, toLow);
    const high = Math.max(fromHigh, toHigh);
    // Chaining is legitimate — three boxes really can share one rule — but a cluster may never grow wider
    // than the stroke that justified joining it, or an edge would move further than the border it hides.
    if (high - low > tolerance) continue;
    const root = Math.min(fromRoot, toRoot);
    const merged = Math.max(fromRoot, toRoot);
    parent.set(merged, root);
    extent.set(root, [low, high]);
    extent.delete(merged);
  }
  const clusters = new Map();
  for (const value of parent.keys()) {
    const root = find(value);
    if (!clusters.has(root)) clusters.set(root, []);
    clusters.get(root).push(value);
  }
  const aliases = new Map();
  for (const members of clusters.values()) {
    if (members.length < 2) continue;
    const canonical = point((Math.min(...members) + Math.max(...members)) / 2);
    for (const value of members) aliases.set(value, canonical);
  }
  return aliases;
}

function visiblePageBreak(item, context) {
  if (!item.pageBreak || isHidden(item.pageBreak.disabled, context)) return 'None';
  return String(item.pageBreak.location || 'None');
}

function declaresNestedPageBreak(item, context) {
  return (item.items || []).some((child) => (
    !/^None$/i.test(visiblePageBreak(child, context))
    || declaresNestedPageBreak(child, context)
  ));
}

function paintsOwnExtent(style, context) {
  if (styleColor(style?.backgroundColor, context, null)) return true;
  return ['top', 'right', 'bottom', 'left'].some((side) => {
    const border = style?.borders?.[side] || style?.border;
    return border && !/^none$/i.test(String(styleValue(border.style, context, 'None')));
  });
}

// The report an expanded item belongs to. A Subreport outside a data region is invoked once, in the scope
// of the item that declares it, and SSRS lays the child report's body out inline at that position — so the
// child's items keep the CHILD model, request, globals, and context for every dataset, expression, image,
// and column measurement, while sitting in the parent's worksheet flow.
function scopeOf(item, fallback) {
  return item?.subreportScope || fallback;
}

function withSubreportScope(item, scope) {
  return {
    ...item,
    subreportScope: item.subreportScope || scope,
    ...(item.items ? { items: item.items.map((child) => withSubreportScope(child, scope)) } : {}),
  };
}

function containsSubreport(item) {
  if (item.type === 'Subreport') return true;
  return (item.items || []).some((child) => containsSubreport(child));
}

// Replaces every Subreport that is NOT inside a data region with the child report's own body items, at the
// invoking item's coordinates and carrying the child scope. A Subreport inside a tablix cell is deliberately
// untouched: it is invoked once per emitted row and is materialized with that row's scope instead.
function expandStandaloneSubreports(items, context) {
  const expanded = [];
  for (const item of items || []) {
    if (item.type === 'Subreport') {
      const invocation = resolveSubreportInvocation(item, context);
      if (!invocation) continue;
      const childRequest = { parameters: invocation.parameters, datasets: invocation.instance.datasets };
      const childContext = {
        parameters: invocation.parameters,
        globals: invocation.globals,
        datasets: normalizeDatasets(invocation.model, childRequest),
        dataset: [],
        fields: {},
      };
      const scope = {
        model: invocation.model,
        request: childRequest,
        globals: invocation.globals,
        context: childContext,
      };
      for (const child of expandStandaloneSubreports(invocation.model.body.items || [], childContext)) {
        expanded.push(withSubreportScope({
          ...child,
          left: point((item.left || 0) + (child.left || 0)),
          top: point((item.top || 0) + (child.top || 0)),
        }, scope));
      }
      continue;
    }
    if (containsSubreport(item)) {
      // A container keeps its own box; only the Subreport inside it is replaced. Once its children include
      // the child report's data regions, the container-flattening pass below hoists them as usual.
      expanded.push({ ...item, items: expandStandaloneSubreports(item.items || [], context) });
      continue;
    }
    expanded.push(item);
  }
  return expanded;
}

// PageBreak is a property of every RDL report item, not only of a direct Body child. A rectangle is a
// coordinate container rather than a page unit, so a break declared on one of its children splits the report
// flow *inside* it. REPORT worksheets are partitioned at breaks, so such a container is expanded into its
// children at absolute body coordinates and the partition falls between them. Containers without a nested
// break keep the existing container path untouched.
//
// A page-spanning container that paints its own fill or border cannot be expanded without losing that
// paint. That is the same construct the PDF renderer refuses to fragment, so it fails closed with the same
// code rather than producing a silently different worksheet.
//
// A container holding a tablix is expanded for a different reason. A Rectangle is a coordinate container,
// not a flow unit: its children keep their declared positions, so two children in disjoint horizontal lanes
// stay side by side. Only the coordinate scheduler honours that, and it needs a flat item list - a
// container holding a tablix otherwise falls back to the flow path, which appends each child below the
// previous one and collapses a two-column design into a single column. Expanding to absolute body
// coordinates keeps the declared layout. Nothing is lost when such a container paints: it keeps a childless
// copy of itself at the same coordinates, and the scheduler grows that copy over whatever its former
// children resolve to.
function expandBreakBearingContainers(items, context, offsetLeft = 0, offsetTop = 0) {
  const expanded = [];
  for (const item of items) {
    if (isHidden(item.hidden, context)) continue;
    // Copy only when a container above actually moved this item, so a report without nested breaks keeps
    // the identical item objects (and therefore the identical tablix layout cache keys) it had before.
    const shifted = offsetLeft || offsetTop
      ? { ...item, left: (item.left || 0) + offsetLeft, top: (item.top || 0) + offsetTop }
      : item;
    if (item.type !== 'Rectangle') {
      expanded.push(shifted);
      continue;
    }
    const fragments = declaresNestedPageBreak(item, context);
    const schedulesTablix = containsTablix(item);
    if (!fragments && !schedulesTablix) {
      expanded.push(shifted);
      continue;
    }
    const paints = paintsOwnExtent(item.style, context);
    if (fragments && paints) {
      throw new ServiceError(
        'UNSUPPORTED_FEATURE',
        'A page-spanning rectangle with a visible fill or border cannot be safely fragmented',
        422,
        { item: item.name || null },
      );
    }
    // The shell is emitted before its former children so they keep painting over it, exactly as the
    // container path drew them.
    if (paints) expanded.push({ ...shifted, items: [], containerPaintShell: true });
    expanded.push(...expandBreakBearingContainers(item.items || [], context, shifted.left, shifted.top));
  }
  return expanded;
}

export { resolveExcelLayoutMode };

function uniqueSheetName(workbook, requested, fallback) {
  const base = sheetName(requested, fallback);
  const existing = new Set(workbook.worksheets.map((sheet) => sheet.name.toLowerCase()));
  if (!existing.has(base.toLowerCase())) return base;
  for (let suffix = 2; suffix < 10_000; suffix += 1) {
    const marker = ` (${suffix})`;
    const candidate = `${base.slice(0, 31 - marker.length)}${marker}`;
    if (!existing.has(candidate.toLowerCase())) return candidate;
  }
  throw new ServiceError('RDL_INVALID', 'Excel worksheet names could not be made unique');
}

function partitionReportSections(model, context) {
  const items = expandBreakBearingContainers(expandStandaloneSubreports(model.body.items || [], context), context)
    .sort((left, right) => left.top - right.top || left.left - right.left || left.zIndex - right.zIndex);
  const sections = [];
  let current = [];
  const finish = () => {
    if (current.length) sections.push(current);
    current = [];
  };
  for (const item of items) {
    const location = visiblePageBreak(item, scopeOf(item, { context }).context);
    if (/^(Start|StartAndEnd)$/i.test(location)) finish();
    current.push(item);
    if (/^(End|StartAndEnd)$/i.test(location)) finish();
  }
  finish();
  return sections.length ? sections : [[]];
}

function firstVisibleText(items, model, request, globals) {
  const datasets = normalizeDatasets(model, request);
  const reportScope = {
    model,
    request,
    globals,
    context: { parameters: request.parameters || {}, globals, fields: {}, dataset: [], datasets },
  };
  const visit = (item, parentScope = reportScope) => {
    const scope = scopeOf(item, parentScope);
    const context = scope.context;
    if (item.type === 'Textbox') {
      try {
        const value = String(textForItem(item, context) || '').replace(/\s+/g, ' ').trim();
        if (value) return value;
      } catch { /* a field-scoped textbox is not a stable section title */ }
    }
    if (item.type === 'Tablix') {
      try {
        const { rows } = tablixRows(item, scope.request, scope.globals, scope.model);
        for (const row of rows.slice(0, 3)) {
          for (const cell of row.cells) {
            const value = cellText(cell).replace(/\s+/g, ' ').trim();
            if (value) return value;
          }
        }
      } catch { /* validation/rendering reports the real error later */ }
    }
    if (item.type === 'Chart' && item.title && !isHidden(item.title.hidden, context)) {
      try {
        const value = String(evaluateExpression(item.title.caption, context) || '').replace(/\s+/g, ' ').trim();
        if (value) return value;
      } catch { /* a field-scoped chart title is not a stable section title */ }
    }
    for (const child of [...(item.items || [])].sort((a, b) => a.top - b.top || a.left - b.left)) {
      const value = visit(child, scope);
      if (value) return value;
    }
    return '';
  };
  for (const item of items) {
    const value = visit(item);
    if (value) return value;
  }
  return '';
}

// Excel renders a stored column width `w` at exactly `w * maxDigitWidth` device pixels — it does NOT add
// the per-cell inset back. (Verified against Excel itself: stored 10 renders 70 px, stored 20 renders
// 140 px.) Subtracting the inset here therefore made every grid column ~5 px narrower than the RDL asked
// for, and an RDL column that the shared grid splits into N slices lost 5N px. Text merely clipped, but a
// date or number is never clipped by Excel — it becomes `#####`, so the value vanished entirely. The
// per-cell inset belongs to text measurement (EXCEL_CELL_TEXT_INSET_PT), not to the column geometry.
function excelWidthFromPoints(points) {
  const pixels = Math.max(1, points * (EXCEL_LAYOUT_DPI / 72));
  return Math.max(0.05, pixels / EXCEL_MAX_DIGIT_WIDTH_PX);
}

function tablixLayout(item, request, globals, model, cache) {
  if (cache.has(item)) return cache.get(item);
  const materialized = tablixRows(item, request, globals, model);
  const layoutItem = item.hasColumnGroups
    ? { ...item, columns: materialized.columns, width: materialized.columns.reduce((sum, width) => sum + width, 0) }
    : { ...item, columns: materialized.columns };
  const { columnsPt } = resolveGridColumns(layoutItem);
  const result = { rows: materialized.rows, columns: columnsPt };
  cache.set(item, result);
  return result;
}

// A worksheet band narrower than the certified geometry tolerance is not drawable at that width: Excel
// keeps both bounding cells and paints a border on each, so a container edge that sits a fraction of a
// point outside the tablix it holds becomes a visible double rule. Collapse those coordinates onto one
// grid line and record every collapsed value so the existing boundary lookups still resolve.
// See gridBoundaries.js for the shared reasoning.
function collapsedBoundaries(values, spans, existingAliases = new Map()) {
  const { boundaries, indexOf } = buildGridBoundaries(values, { protectedSpans: spans });
  const aliases = new Map();
  for (const [raw, target] of existingAliases) {
    const index = indexOf(target);
    if (index >= 0) aliases.set(raw, boundaries[index]);
  }
  for (const value of values) {
    if (aliases.has(value)) continue;
    const index = indexOf(value);
    if (index >= 0) aliases.set(value, boundaries[index]);
  }
  const resolved = [...boundaries];
  resolved.aliases = aliases;
  return resolved;
}

function collectXBoundaries(items, request, globals, model, tablixCache, target, spans, parentLeft = 0) {
  // A Chart or Image sitting on a tablix cell's flattened canvas is anchored as a floating picture at its
  // own left and width, so those edges have to be grid lines. Cell TEXT needs none: it is written into the
  // cell and inherits the cell's own columns.
  const collectCellCanvas = (cell, cellLeft) => {
    if (!isFreeFormCell(cell)) return;
    for (const child of cell.items || []) {
      if (child.type === 'Tablix' || child.type === 'Subreport') continue;
      const childLeft = point(cellLeft + (child.left || 0));
      const childRight = point(childLeft + (child.width || 0));
      target.add(childLeft);
      target.add(childRight);
      if (childRight > childLeft) spans.add(`${childLeft}|${childRight}`);
    }
  };
  const collectNested = (nested, cellLeft) => {
    const left = point(cellLeft + (nested.item.left || 0));
    target.add(left);
    let cursor = left;
    for (const width of nested.columns || nested.item.columns || []) {
      const previous = cursor;
      cursor = point(cursor + width);
      target.add(cursor);
      spans.add(`${previous}|${cursor}`);
    }
    const placements = computeCellPlacements(nested.rows || [], (nested.columns || []).length);
    for (const [rowIndex, row] of (nested.rows || []).entries()) {
      for (const [cellIndex, cell] of row.cells.entries()) {
        const columnIndex = placements[rowIndex][cellIndex];
        const nestedCellLeft = left + (nested.columns || []).slice(0, columnIndex).reduce((sum, width) => sum + width, 0);
        for (const child of cell.nestedTablixes || []) collectNested(child, nestedCellLeft);
        collectCellCanvas(cell, nestedCellLeft);
      }
    }
  };
  for (const item of items || []) {
    const scope = scopeOf(item, { model, request, globals });
    const left = point(parentLeft + (item.left || 0));
    const right = point(left + (item.width || 0));
    target.add(left);
    target.add(right);
    spans.add(`${left}|${right}`);
    if (item.type === 'Tablix') {
      const layout = tablixLayout(item, scope.request, scope.globals, scope.model, tablixCache);
      let cursor = left;
      const offsets = [0];
      for (const width of layout.columns) {
        const previous = cursor;
        cursor = point(cursor + width);
        target.add(cursor);
        spans.add(`${previous}|${cursor}`);
        offsets.push(cursor - left);
      }
      const placements = computeCellPlacements(layout.rows, layout.columns.length);
      for (const [rowIndex, row] of layout.rows.entries()) {
        for (const [cellIndex, cell] of row.cells.entries()) {
          const columnIndex = placements[rowIndex][cellIndex];
          const cellLeft = left + offsets[columnIndex];
          for (const nested of cell.nestedTablixes || []) collectNested(nested, cellLeft);
          collectCellCanvas(cell, cellLeft);
        }
      }
    }
    collectXBoundaries(item.items, scope.request, scope.globals, scope.model, tablixCache, target, spans, left);
  }
}

function reportGrid(model, section, request, globals, tablixCache) {
  const boundaries = new Set([0]);
  const spans = new Set();
  collectXBoundaries(model.page.header?.items || [], request, globals, model, tablixCache, boundaries, spans);
  collectXBoundaries(section, request, globals, model, tablixCache, boundaries, spans);
  const context = { parameters: request.parameters || {}, globals, fields: {} };
  // Header and body edges are resolved together: they address the same worksheet column grid, so a rule
  // shared across the band boundary has to land on one coordinate.
  const equivalentEdgePairs = [];
  collectEquivalentXEdgePairs(model.page.header?.items || [], equivalentEdgePairs, context);
  collectEquivalentXEdgePairs(section, equivalentEdgePairs, context);
  const aliases = resolveEquivalentXEdges(equivalentEdgePairs);
  const maximum = Math.max(...boundaries, point(model.page.width - model.page.marginLeft - model.page.marginRight));
  boundaries.add(maximum);
  const values = [...new Set([...boundaries]
    .map((value) => aliases.get(value) ?? value))]
    .filter((value) => value >= 0 && value <= maximum)
    .sort((a, b) => a - b);
  if (values.length < 2) values.push(point(maximum || 100));
  // Every collected span is one report item's or one tablix column's own extent, so both of its ends must
  // stay distinct grid lines or that item would lose the only column it can occupy.
  const protectedSpans = [...spans].map((key) => key.split('|').map(Number))
    .map(([from, to]) => [aliases.get(from) ?? from, aliases.get(to) ?? to]);
  return collapsedBoundaries(values, protectedSpans, aliases);
}

// boundaryIndex for a coordinate that is only expected to be on the grid, not required to be: returns null
// instead of failing the render when it is not.
function optionalBoundaryIndex(boundaries, value) {
  try {
    return boundaryIndex(boundaries, value);
  } catch {
    return null;
  }
}

function boundaryIndex(boundaries, value) {
  const raw = point(value);
  const expected = boundaries.aliases?.get(raw) ?? raw;
  let index = boundaries.findIndex((candidate) => candidate === expected);
  // Independently accumulated decimal RDL sizes can land one quarter-point apart at a containing band
  // edge. At the renderer's declared 0.25-point precision those coordinates are equivalent.
  if (index < 0) {
    const nearby = boundaries
      .map((candidate, candidateIndex) => ({ candidateIndex, distance: Math.abs(candidate - expected) }))
      .filter(({ distance }) => distance <= 1 / POINT_PRECISION)
      .sort((left, right) => left.distance - right.distance);
    if (nearby.length === 1 || (nearby.length > 1 && nearby[0].distance < nearby[1].distance)) {
      index = nearby[0].candidateIndex;
    }
  }
  if (index < 0) throw new ServiceError('RDL_INVALID', `Excel layout boundary is not on the section grid: ${expected}`);
  return index;
}

function gridRange(boundaries, left, width) {
  const start = boundaryIndex(boundaries, left);
  const end = boundaryIndex(boundaries, left + width);
  if (end <= start) throw new ServiceError('RDL_INVALID', 'Excel layout item has an empty grid range');
  return { startCol: start + 1, endCol: end };
}

function excelFont(style, context) {
  const decoration = String(styleValue(style?.textDecoration, context, 'None'));
  return {
    name: String(styleText(style?.fontFamily, context, 'Arial')),
    size: styleSize(style?.fontSize, context, 10) || 10,
    bold: /bold|[6-9]00/i.test(String(styleValue(style?.fontWeight, context, 'Normal'))),
    italic: /italic/i.test(String(styleValue(style?.fontStyle, context, 'Normal'))),
    underline: /underline/i.test(decoration),
    strike: /line.?through/i.test(decoration),
    color: { argb: `FF${hex(styleColor(style?.color, context, '#000000'), '000000')}` },
  };
}

function richTextValue(textbox, context, requestedText) {
  if (!textbox?.paragraphs) return null;
  const paragraphs = styledTextForItem(textbox, context);
  if (!paragraphs) return null;
  const richText = [];
  let fullText = '';
  paragraphs.forEach((paragraph, paragraphIndex) => {
    paragraph.runs.forEach((run) => {
      const text = String(run.text ?? '');
      fullText += text;
      if (text) richText.push({ text, font: excelFont(run.style || textbox.style, context) });
    });
    if (paragraphIndex < paragraphs.length - 1) {
      fullText += '\n';
      richText.push({ text: '\n', font: excelFont(paragraph.style || textbox.style, context) });
    }
  });
  if (fullText !== String(requestedText ?? '')) return null;
  // Keep ordinary one-run strings as normal Excel text cells. Rich text is only necessary when the RDL
  // actually has multiple run/paragraph boundaries; this preserves normal filtering and value typing.
  return richText.length > 1 ? { richText } : null;
}

function applyFillFontAlignment(cell, style, context) {
  const fill = hex(styleColor(style?.backgroundColor, context, null));
  if (fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${fill}` } };
  cell.font = excelFont(style, context);
  const vertical = String(styleValue(style?.verticalAlign, context, 'Top'));
  const horizontal = String(styleValue(style?.textAlign, context, 'Left'));
  const writingMode = String(styleValue(style?.writingMode, context, 'Default') || 'Default').toLowerCase();
  cell.alignment = {
    vertical: /bottom/i.test(vertical) ? 'bottom' : /middle|center/i.test(vertical) ? 'middle' : 'top',
    horizontal: /center/i.test(horizontal) ? 'center' : /right/i.test(horizontal) ? 'right' : /justify/i.test(horizontal) ? 'justify' : 'left',
    wrapText: true,
    textRotation: writingMode === 'rotate270' ? 90 : writingMode === 'vertical' ? -90 : 0,
  };
}

// A container's declared border ADDS an edge on its own perimeter; it never removes one, and its fill sits
// behind whatever its contents already painted. Applied after the contents, so the contents' own resolved
// edges survive wherever the two perimeters coincide.
function paintContainerExtent(worksheet, range, item, context) {
  const fill = hex(styleColor(item.style?.backgroundColor, context, null));
  const borders = item.style?.borders || {};
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startCol; column <= range.endCol; column += 1) {
      const target = worksheet.getCell(row, column);
      if (fill && !target.fill?.type && (target.value === null || target.value === undefined)) {
        target.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${fill}` } };
      }
      const containerEdge = (side, onEdge) => (
        (onEdge ? excelBorderSide(borders[side], context) : undefined) || target.border?.[side]
      );
      target.border = {
        ...(target.border || {}),
        top: containerEdge('top', row === range.startRow),
        bottom: containerEdge('bottom', row === range.endRow),
        left: containerEdge('left', column === range.startCol),
        right: containerEdge('right', column === range.endCol),
      };
    }
  }
}

function applyRegionStyle(worksheet, range, style, context, { includeBorders = true } = {}) {
  const borders = style?.borders || {};
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startCol; column <= range.endCol; column += 1) {
      const cell = worksheet.getCell(row, column);
      applyFillFontAlignment(cell, style, context);
      if (!includeBorders) continue;
      cell.border = {
        top: row === range.startRow ? excelBorderSide(borders.top, context) : undefined,
        bottom: row === range.endRow ? excelBorderSide(borders.bottom, context) : undefined,
        left: column === range.startCol ? excelBorderSide(borders.left, context) : undefined,
        right: column === range.endCol ? excelBorderSide(borders.right, context) : undefined,
      };
    }
  }
}

// Distributes one logical cell's RESOLVED edges around the perimeter of the physical Excel region it
// occupies. A cell hosting a nested data region cannot be merged (the child grid needs the individual
// cells to place its own rows), so writing all four sides onto the anchor cell drew the cell's bottom and
// right rules across its FIRST physical row — a horizontal line through the middle of a tall cell that
// SSRS closes only at its outer edge. Merged cells keep taking the border from their anchor, which is how
// Excel renders a merged range.
function applyRegionBorder(worksheet, range, border = {}) {
  const onPerimeter = {
    top: (row) => row === range.startRow,
    bottom: (row) => row === range.endRow,
    left: (row, column) => column === range.startCol,
    right: (row, column) => column === range.endCol,
  };
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startCol; column <= range.endCol; column += 1) {
      const cell = worksheet.getCell(row, column);
      // Add the enclosing edges; never clear what the child grid already drew in this cell.
      const merged = { ...(cell.border || {}) };
      for (const side of ['top', 'bottom', 'left', 'right']) {
        if (border[side] && onPerimeter[side](row, column)) merged[side] = border[side];
      }
      // Only mark the cell when an edge is actually set: an all-undefined border object still makes the
      // cell look styled, and the copy step then materializes the whole region as a merge per column.
      if (Object.values(merged).some(Boolean)) cell.border = merged;
    }
  }
}

function rangesOverlap(left, right) {
  return left.startRow <= right.endRow && right.startRow <= left.endRow
    && left.startCol <= right.endCol && right.startCol <= left.endCol;
}

function mergeBucket(row) {
  return Math.floor((row - 1) / MERGE_ROW_BUCKET_SIZE);
}

function createMergeIndex() {
  return { bucketsByColumn: new Map() };
}

function mergeIndexBuckets(index, range, create = false) {
  const buckets = [];
  const firstBucket = mergeBucket(range.startRow);
  const lastBucket = mergeBucket(range.endRow);
  for (let column = range.startCol; column <= range.endCol; column += 1) {
    let columnBuckets = index.bucketsByColumn.get(column);
    if (!columnBuckets && create) {
      columnBuckets = new Map();
      index.bucketsByColumn.set(column, columnBuckets);
    }
    if (!columnBuckets) continue;
    for (let bucket = firstBucket; bucket <= lastBucket; bucket += 1) {
      let candidates = columnBuckets.get(bucket);
      if (!candidates && create) {
        candidates = new Set();
        columnBuckets.set(bucket, candidates);
      }
      if (candidates) buckets.push(candidates);
    }
  }
  return buckets;
}

function findIndexedMergeOverlap(index, range) {
  for (const candidates of mergeIndexBuckets(index, range)) {
    for (const candidate of candidates) {
      if (rangesOverlap(candidate, range)) return candidate;
    }
  }
  return null;
}

function indexMerge(index, range) {
  for (const candidates of mergeIndexBuckets(index, range, true)) candidates.add(range);
}

// The range has already passed our overlap check. Reproduce ExcelJS's merge operation without calling its
// public mergeCells(), whose implementation redundantly scans the complete worksheet merge collection for
// every insertion. Keep this small adapter aligned with the pinned ExcelJS operation: followers reference
// the master and inherit its style, while the Range remains in _merges for normal OOXML serialization.
function applyPrevalidatedMerge(worksheet, range) {
  if (!worksheet?._merges || typeof worksheet.getCell !== 'function') {
    throw new ServiceError('RENDER_FAILED', 'Excel worksheet merge support is unavailable');
  }
  const dimensions = new ExcelRange(range.startRow, range.startCol, range.endRow, range.endCol);
  const master = worksheet.getCell(dimensions.top, dimensions.left);
  for (let row = dimensions.top; row <= dimensions.bottom; row += 1) {
    for (let column = dimensions.left; column <= dimensions.right; column += 1) {
      if (row !== dimensions.top || column !== dimensions.left) worksheet.getCell(row, column).merge(master);
    }
  }
  worksheet._merges[master.address] = dimensions;
}

function mergeSafe(worksheet, range, mergeIndex, owner = null) {
  if (range.startRow === range.endRow && range.startCol === range.endCol) return;
  const existing = findIndexedMergeOverlap(mergeIndex, range);
  if (existing) {
    throw new ServiceError(
      'RDL_INVALID',
      `RDL produced overlapping Excel merged-cell ranges${owner ? ` for ${owner}` : ''} (${existing.startRow},${existing.startCol}:${existing.endRow},${existing.endCol} and ${range.startRow},${range.startCol}:${range.endRow},${range.endCol})`,
    );
  }
  applyPrevalidatedMerge(worksheet, range);
  indexMerge(mergeIndex, range);
}

function splitTallRowIntervals(boundaries) {
  const output = [boundaries[0] || 0];
  for (let index = 0; index < boundaries.length - 1; index += 1) {
    const end = boundaries[index + 1];
    let cursor = output.at(-1);
    while (end - cursor > MAX_EXCEL_ROW_HEIGHT) {
      cursor = point(cursor + MAX_EXCEL_ROW_HEIGHT);
      output.push(cursor);
    }
    if (end > output.at(-1)) output.push(end);
  }
  return output;
}

function allocateHeightRows(worksheet, boundaries, startRow) {
  for (let index = 0; index < boundaries.length - 1; index += 1) {
    const height = boundaries[index + 1] - boundaries[index];
    if (height > MAX_EXCEL_ROW_HEIGHT) throw new ServiceError('UNSUPPORTED_FEATURE', 'An Excel layout row exceeds the 409-point row-height limit');
    worksheet.getRow(startRow + index).height = Math.max(2, height);
  }
  return boundaries.length - 1;
}

function collectYBoundaries(items, target, spans, parentTop = 0) {
  for (const item of items || []) {
    const top = point(parentTop + (item.top || 0));
    const bottom = point(top + (item.height || 0));
    target.add(top);
    target.add(bottom);
    spans.add(`${top}|${bottom}`);
    collectYBoundaries(item.items, target, spans, top);
  }
}

function freeformRows(worksheet, items, height, startRow) {
  const boundaries = new Set([0, point(height)]);
  const spans = new Set();
  collectYBoundaries(items, boundaries, spans);
  const collapsed = collapsedBoundaries(
    [...boundaries].filter((value) => value >= 0 && value <= point(height)).sort((a, b) => a - b),
    [...spans].map((key) => key.split('|').map(Number)),
  );
  // splitTallRowIntervals only inserts extra intermediate boundaries, so the collapse aliases stay valid.
  const values = splitTallRowIntervals(collapsed);
  values.aliases = collapsed.aliases;
  allocateHeightRows(worksheet, values, startRow);
  return values;
}

// A free-form CanGrow textbox is laid out from its declared RDL height alone, but the worksheet re-wraps
// its text with Excel's engine inside a narrower usable width. When Excel produces more lines than the
// declared box holds, the overflow is simply not displayed — the row does not grow, and the cell is merged
// so it cannot spill sideways either. Raise the row band to what the text needs. Returns whether it grew.
function growFreeformRows(worksheet, range, requiredHeight) {
  if (!(requiredHeight > 0)) return false;
  const rows = [];
  let allocated = 0;
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const target = worksheet.getRow(row);
    rows.push(target);
    allocated += target.height || 0;
  }
  let deficit = point(requiredHeight - allocated);
  if (deficit <= 1 / POINT_PRECISION) return false;
  // Excel caps a single row at 409 points, so the extra height is absorbed across the band the item already
  // occupies, last row first — that band is this item's own geometry, so nothing else is displaced. A block
  // needing more than the band can hold even at the cap keeps every point the cap allows: the rows were
  // sized from the declared RDL height, so this can only ever reveal more text than before, never less.
  for (let index = rows.length - 1; index >= 0 && deficit > 1 / POINT_PRECISION; index -= 1) {
    const current = rows[index].height || 0;
    const room = MAX_EXCEL_ROW_HEIGHT - current;
    if (room <= 0) continue;
    const applied = Math.min(room, deficit);
    rows[index].height = point(current + applied);
    deficit = point(deficit - applied);
  }
  return true;
}

// Height a free-form CanGrow textbox needs in a worksheet, as opposed to on a PDF page.
//
// excelCanGrowTextboxHeight already reserves one line for the "PDF breaks into N lines, Excel breaks into
// N+1" disagreement, but only once the content is already multi-line. A single line that fills its box is
// exposed to exactly the same disagreement and is the more damaging case, because the whole second line
// then falls outside a one-line row. Detect it by re-flowing against a slightly narrower box rather than by
// assuming anything about Excel's font metrics: if a small loss of width adds a line, the two engines can
// disagree here. This lives at the free-form call site so tablix cell geometry is untouched.
function excelFreeformTextMetrics(measureDoc, config, item, context, display, mergedWidth) {
  const usable = Math.max(1, mergedWidth - EXCEL_CELL_TEXT_INSET_PT);
  const required = excelCanGrowTextboxHeight(measureDoc, config, item, context, display, usable, item.style || {});
  if (!required) return { required: 0, multiLine: false };
  const single = measureTextboxHeight(measureDoc, config, item, context, 'M', usable);
  const content = measureTextboxHeight(measureDoc, config, item, context, display, usable);
  // Already multi-line: excelCanGrowTextboxHeight has applied its own reserve.
  if (content > single + (1 / POINT_PRECISION)) return { required, multiLine: true };
  const narrowed = measureTextboxHeight(
    measureDoc,
    config,
    item,
    context,
    display,
    Math.max(1, usable * (1 - EXCEL_WRAP_SAFETY_FRACTION)),
  );
  const nearWrap = narrowed > content + (1 / POINT_PRECISION);
  return { required: nearWrap ? required + single : required, multiLine: nearWrap };
}

function rowRange(boundaries, startRow, top, height) {
  const from = boundaryIndex(boundaries, top);
  const to = boundaryIndex(boundaries, top + height);
  return { startRow: startRow + from, endRow: startRow + to - 1 };
}

// Excel pictures float above cell borders.  When a picture reaches a bordered Rectangle's edge, the
// picture would therefore cover the rule even though that same rule remains visible in SSRS.  Keep the
// picture inside every coincident ancestor edge by the declared stroke width.  This is not Rectangle
// padding (which only applies to text); it is the XLSX representation of the parent canvas clip.
function fractionalGridPosition(boundaries, coordinate) {
  const value = point(coordinate);
  for (let index = 0; index < boundaries.length; index += 1) {
    if (Math.abs(boundaries[index] - value) <= COINCIDENT_EDGE_TOLERANCE_PT) return index;
    if (index + 1 >= boundaries.length || value > boundaries[index + 1]) continue;
    const span = boundaries[index + 1] - boundaries[index];
    return index + ((value - boundaries[index]) / Math.max(span, Number.EPSILON));
  }
  return boundaries.length - 1;
}

function pictureInsetsAtContainerEdges(left, top, width, height, containerFrames = []) {
  const right = point(left + width);
  const bottom = point(top + height);
  const result = { top: 0, bottom: 0, left: 0, right: 0 };
  for (const frame of containerFrames) {
    const borders = frame.item.style?.borders || {};
    if (Math.abs(left - frame.left) <= COINCIDENT_EDGE_TOLERANCE_PT) result.left = Math.max(result.left, visibleBorderWidth(borders.left, frame.context));
    if (Math.abs(right - frame.right) <= COINCIDENT_EDGE_TOLERANCE_PT) result.right = Math.max(result.right, visibleBorderWidth(borders.right, frame.context));
    if (Math.abs(top - frame.top) <= COINCIDENT_EDGE_TOLERANCE_PT) result.top = Math.max(result.top, visibleBorderWidth(borders.top, frame.context));
    if (Math.abs(bottom - frame.bottom) <= COINCIDENT_EDGE_TOLERANCE_PT) result.bottom = Math.max(result.bottom, visibleBorderWidth(borders.bottom, frame.context));
  }
  // Preserve a drawable area for an unusually narrow picture rather than inverting its anchors.
  const horizontal = Math.max(0, width - (1 / POINT_PRECISION));
  const vertical = Math.max(0, height - (1 / POINT_PRECISION));
  if (result.left + result.right > horizontal) {
    const scale = horizontal / Math.max(result.left + result.right, 1);
    result.left *= scale;
    result.right *= scale;
  }
  if (result.top + result.bottom > vertical) {
    const scale = vertical / Math.max(result.top + result.bottom, 1);
    result.top *= scale;
    result.bottom *= scale;
  }
  return result;
}

function hasVisibleExcelBorder(border) {
  return Boolean(border?.style && !/^none$/i.test(String(border.style)));
}

// Pictures float above worksheet cells.  The normalized-container chain can be flattened by a grouped
// tablix or bundled subreport, but the target worksheet already has the resolved native border in either
// case. Use that final border as the last authority before anchoring a picture, so a border is never
// hidden merely because its original Rectangle/tablix ancestor was normalized away.
function worksheetPictureBorderInsets(worksheet, startRow, xGrid, yGrid, left, top, width, height) {
  const result = { top: 0, bottom: 0, left: 0, right: 0 };
  const leftIndex = fractionalGridPosition(xGrid, left);
  const rightIndex = fractionalGridPosition(xGrid, point(left + width));
  const topIndex = fractionalGridPosition(yGrid, top);
  const bottomIndex = fractionalGridPosition(yGrid, point(top + height));
  const isBoundary = (value) => Math.abs(value - Math.round(value)) <= Number.EPSILON;
  if (!isBoundary(leftIndex) || !isBoundary(rightIndex) || !isBoundary(topIndex) || !isBoundary(bottomIndex)) return result;
  const firstColumn = Math.round(leftIndex) + 1;
  const afterLastColumn = Math.round(rightIndex) + 1;
  const firstRow = startRow + Math.round(topIndex);
  const afterLastRow = startRow + Math.round(bottomIndex);
  for (let row = firstRow; row < afterLastRow; row += 1) {
    const leftCell = worksheet.getCell(row, firstColumn);
    const rightCell = worksheet.getCell(row, Math.max(1, afterLastColumn - 1));
    const outsideLeft = firstColumn > 1 ? worksheet.getCell(row, firstColumn - 1) : null;
    const outsideRight = worksheet.getCell(row, afterLastColumn);
    if (hasVisibleExcelBorder(leftCell.border?.left) || hasVisibleExcelBorder(outsideLeft?.border?.right)) result.left = 1;
    if (hasVisibleExcelBorder(rightCell.border?.right) || hasVisibleExcelBorder(outsideRight?.border?.left)) result.right = 1;
  }
  for (let column = firstColumn; column < afterLastColumn; column += 1) {
    const topCell = worksheet.getCell(firstRow, column);
    const bottomCell = worksheet.getCell(Math.max(1, afterLastRow - 1), column);
    const outsideTop = firstRow > 1 ? worksheet.getCell(firstRow - 1, column) : null;
    const outsideBottom = worksheet.getCell(afterLastRow, column);
    if (hasVisibleExcelBorder(topCell.border?.top) || hasVisibleExcelBorder(outsideTop?.border?.bottom)) result.top = 1;
    if (hasVisibleExcelBorder(bottomCell.border?.bottom) || hasVisibleExcelBorder(outsideBottom?.border?.top)) result.bottom = 1;
  }
  return result;
}

function pictureAnchor(worksheet, xGrid, yGrid, startRow, left, top, width, height, containerFrames) {
  const insets = pictureInsetsAtContainerEdges(left, top, width, height, containerFrames);
  const worksheetInsets = worksheetPictureBorderInsets(worksheet, startRow, xGrid, yGrid, left, top, width, height);
  for (const side of ['top', 'right', 'bottom', 'left']) insets[side] = Math.max(insets[side], worksheetInsets[side]);
  return {
    tl: {
      col: fractionalGridPosition(xGrid, point(left + insets.left)),
      row: startRow - 1 + fractionalGridPosition(yGrid, point(top + insets.top)),
    },
    br: {
      col: fractionalGridPosition(xGrid, point(left + width - insets.right)),
      row: startRow - 1 + fractionalGridPosition(yGrid, point(top + height - insets.bottom)),
    },
    editAs: 'oneCell',
  };
}

function canvasCellFrame(owner, left, top, width, height) {
  return {
    // The owning tablix cell is also a positioned canvas container. Its resolved cell style—not only an
    // explicit Rectangle style—owns the visible perimeter that an Excel floating picture must preserve.
    item: { style: owner.style || {} },
    context: owner.context,
    left: point(left),
    top: point(top),
    right: point(left + width),
    bottom: point(top + height),
  };
}

function addEmbeddedImage(workbook, worksheet, model, item, anchor, context) {
  if (item.source !== 'Embedded') return;
  const image = model.embeddedImages?.[styleValue(item.value, context, item.value)];
  if (!image?.data) throw new ServiceError('UNSUPPORTED_FEATURE', `Embedded Excel image is unavailable: ${item.name || 'unnamed'}`);
  const buffer = Buffer.from(image.data.replace(/\s+/g, ''), 'base64');
  const id = workbook.addImage({ buffer, extension: imageExtension(buffer) });
  worksheet.addImage(id, anchor);
}

async function renderFreeformItem({
  workbook,
  worksheet,
  model,
  item,
  context,
  config,
  tempDir,
  chartCounter,
  xGrid,
  yGrid,
  startRow,
  merges,
  measureDoc,
  parentLeft = 0,
  parentTop = 0,
  rowsOverride = null,
  containerFrames = [],
}) {
  if (isHidden(item.hidden, context)) return;
  const left = point(parentLeft + (item.left || 0));
  const top = point(parentTop + (item.top || 0));
  const columns = gridRange(xGrid, left, item.width || 0);
  const rows = rowsOverride || rowRange(yGrid, startRow, top, item.height || 0);
  const range = { ...columns, ...rows };
  const anchor = pictureAnchor(worksheet, xGrid, yGrid, startRow, left, top, item.width || 0, item.height || 0, containerFrames);
  if (item.type === 'Chart') {
    const data = materializeChart(item, context.datasets, context.parameters, context.globals);
    const png = await renderChartPng(item, data, config, tempDir, context, chartCounter.value++);
    if (!png?.data) throw new ServiceError('RENDER_FAILED', 'Excel chart picture could not be rendered', 500);
    const id = workbook.addImage({ buffer: png.data, extension: 'png' });
    worksheet.addImage(id, anchor);
    return;
  }
  if (item.type === 'Image') {
    addEmbeddedImage(workbook, worksheet, model, item, anchor, context);
    return;
  }
  if (item.type === 'Line') {
    const border = item.style?.border || item.style?.borders?.top;
    if ((item.width || 0) >= (item.height || 0)) {
      for (let column = range.startCol; column <= range.endCol; column += 1) worksheet.getCell(range.startRow, column).border = { top: excelBorderSide(border, context) };
    } else {
      for (let row = range.startRow; row <= range.endRow; row += 1) worksheet.getCell(row, range.startCol).border = { left: excelBorderSide(border, context) };
    }
    return;
  }
  applyRegionStyle(worksheet, range, item.style || {}, context);
  if (item.type === 'Textbox') {
    mergeSafe(worksheet, range, merges, item.name);
    const text = cellString(textForItem(item, context));
    const target = worksheet.getCell(range.startRow, range.startCol);
    target.value = richTextValue(item, context, text) || text;
    applyFillFontAlignment(target, item.style || {}, context);
    // Measure against the width the worksheet actually gives the cell, not the declared RDL width: the
    // merged range is snapped to the section's shared column grid, so the two differ whenever another
    // item's edge falls inside this one, and Excel wraps to the range it has.
    const metrics = measureDoc
      ? excelFreeformTextMetrics(measureDoc, config, item, context, text, xGrid[range.endCol] - xGrid[range.startCol - 1])
      : { required: 0, multiLine: false };
    const grown = growFreeformRows(worksheet, range, metrics.required);
    // `multiLine` covers the callers that already sized this item's band to its grown content height, where
    // the growth is real but invisible to growFreeformRows.
    if (grown || metrics.multiLine) {
      // The band is taller than the RDL box, and the extra height exists only to hold lines Excel wrapped
      // past that box. RDL VerticalAlign describes the declared box, where a grown box is exactly its own
      // text and Middle and Top coincide; keeping Middle against the enlarged band instead splits the
      // surplus symmetrically, which reads as a blank gap above the text. Anchor to the top so any surplus
      // trails below. A cell that did not grow keeps the alignment the RDL declared.
      target.alignment = { ...(target.alignment || {}), vertical: 'top' };
    }
    return;
  }
  if (item.type === 'Rectangle') {
    const frame = {
      item,
      context,
      left,
      top,
      right: point(left + (item.width || 0)),
      bottom: point(top + (item.height || 0)),
    };
    for (const child of [...(item.items || [])].sort((a, b) => a.zIndex - b.zIndex || a.top - b.top || a.left - b.left)) {
      await renderFreeformItem({
        workbook,
        worksheet,
        model,
        item: child,
        context,
        config,
        tempDir,
        chartCounter,
        xGrid,
        yGrid,
        startRow,
        merges,
        measureDoc,
        parentLeft: left,
        parentTop: top,
        containerFrames: [...containerFrames, frame],
      });
    }
    // The cell perimeter is the only native XLSX representation of a Rectangle border.  Apply it after
    // its children so a child textbox/tablix cannot erase it; floating pictures are separately inset
    // above because Excel always paints them above cells.
    paintContainerExtent(worksheet, range, item, context);
  }
}

async function renderFreeformBand({
  workbook,
  worksheet,
  model,
  items,
  height,
  context,
  config,
  tempDir,
  chartCounter,
  xGrid,
  startRow,
  merges,
  measureDoc,
}) {
  const reportScope = { model, context };
  const resolved = items.map((item, sourceIndex) => {
    const layout = resolveExcelFreeformLayout(item, scopeOf(item, reportScope).context, config, measureDoc);
    return {
      item: layout.item,
      occupiedHeight: layout.occupiedHeight,
      sourceIndex,
      designTop: item.top || 0,
      designHeight: item.height || 0,
      resolvedTop: item.top || 0,
    };
  });
  const layoutOrder = [...resolved].sort((left, right) => (
    left.designTop - right.designTop
    || (left.item.left || 0) - (right.item.left || 0)
    || left.sourceIndex - right.sourceIndex
  ));
  for (const [index, item] of layoutOrder.entries()) {
    for (let previousIndex = 0; previousIndex < index; previousIndex += 1) {
      const previous = layoutOrder[previousIndex];
      const previousDesignBottom = point(previous.designTop + previous.designHeight);
      if (previousDesignBottom > item.designTop + COINCIDENT_EDGE_TOLERANCE_PT) continue;
      if (!horizontalDesignOverlap(previous.item, item.item)) continue;
      const originalGap = Math.max(0, item.designTop - previousDesignBottom);
      item.resolvedTop = Math.max(
        item.resolvedTop,
        point(previous.resolvedTop + previous.occupiedHeight + originalGap),
      );
    }
    item.item = { ...item.item, top: item.resolvedTop, height: item.occupiedHeight };
  }
  const resolvedItems = resolved
    .sort((left, right) => left.sourceIndex - right.sourceIndex)
    .map((entry) => entry.item);
  const occupiedHeight = Math.max(
    height,
    ...resolved.map((entry) => point(entry.resolvedTop + entry.occupiedHeight)),
  );
  const yGrid = freeformRows(worksheet, resolvedItems, occupiedHeight, startRow);
  for (const item of [...resolvedItems].sort((a, b) => a.zIndex - b.zIndex || a.top - b.top || a.left - b.left)) {
    await renderFreeformItem({
      workbook,
      worksheet,
      model: scopeOf(item, reportScope).model,
      item,
      context: scopeOf(item, reportScope).context,
      config,
      tempDir,
      chartCounter,
      xGrid,
      yGrid,
      startRow,
      merges,
      measureDoc,
    });
  }
  return yGrid.length - 1;
}

function cellStyle(item, cell, context) {
  const textbox = cellTextbox(cell);
  const style = { ...(item.style || {}), ...(textbox?.style || {}) };
  if (textbox?.style?.backgroundColor === null || textbox?.style?.backgroundColor === undefined) {
    style.backgroundColor = item.style?.backgroundColor;
  }
  return {
    textbox,
    style,
    borderStyle: cellBorderStyle(cell, item),
    context,
  };
}

function resolvedOwnerBorder(owner, side) {
  // borderStyle is the resolved authority and may be deliberately null when the cell's only content is
  // hidden. Falling back to the content style there would resurrect the tablix border this resolves away.
  const border = owner.borderStyle?.borders?.[side];
  if (!border || /^none$/i.test(String(styleValue(border.style, owner.context, 'None')))) return null;
  return excelBorderSide(border, owner.context);
}

function reportCellBorders(gridOwners, owner, itemStyle, enforceBottomClosure, rows, tablix) {
  const rowIndex = owner.rowIndex;
  const columnIndex = owner.start;
  const span = owner.cell.colSpan || 1;
  const endRow = Math.min(gridOwners.length - 1, rowIndex + Math.max(1, owner.cell.rowSpan || 1) - 1);
  const above = rowIndex > 0 ? gridOwners[rowIndex - 1][columnIndex] : null;
  const below = endRow + 1 < gridOwners.length ? gridOwners[endRow + 1][columnIndex] : null;
  const left = columnIndex > 0 ? gridOwners[rowIndex][columnIndex - 1] : null;
  const right = columnIndex + span < gridOwners[rowIndex].length ? gridOwners[rowIndex][columnIndex + span] : null;
  const bottom = resolvedOwnerBorder(owner, 'bottom')
    || (below && resolvedOwnerBorder(below, 'top'))
    || (enforceBottomClosure && endRow === gridOwners.length - 1
      ? excelBorderSide(enforcedBottomBorder(itemStyle, rows, tablix), owner.context)
      : undefined);
  const top = above === owner ? undefined : resolvedOwnerBorder(owner, 'top')
    || (above && resolvedOwnerBorder(above, 'bottom'))
    || matchingChangedGroupOwnerRowBoundary(
      owner,
      above,
      left,
      right,
      resolvedOwnerBorder,
      (border) => `${border.style || ''}|${border.color?.argb || ''}`,
      (candidate) => materializedCellVisualSignature(candidate.cell, candidate.style, candidate.context),
    )
    || undefined;
  return {
    top,
    bottom,
    left: resolvedOwnerBorder(owner, 'left') || (left && resolvedOwnerBorder(left, 'right')) || undefined,
    right: resolvedOwnerBorder(owner, 'right') || (right && resolvedOwnerBorder(right, 'left')) || undefined,
  };
}

function visibleBorderBetween(upper, lower) {
  return Boolean(resolvedOwnerBorder(upper, 'bottom') || resolvedOwnerBorder(lower, 'top'));
}

function duplicateCellDescriptor(cell) {
  const items = cell.items || [];
  const duplicateItems = cell.duplicateItems || [];
  const duplicateIndexes = duplicateItems
    .map((entry, index) => (entry ? index : -1))
    .filter((index) => index >= 0);
  if (duplicateIndexes.length !== 1 || items.length !== 1 || (cell.nestedTablixes || []).length > 0) return null;
  const index = duplicateIndexes[0];
  const entry = duplicateItems[index];
  return {
    ...entry,
    itemName: items[index]?.name || null,
  };
}

// HideDuplicates suppresses text; it does not blindly merge equal values. In a borderless run where the
// materializer has explicitly identified one visible owner followed by suppressed duplicates in the same
// group instance, a native vertical Excel merge is the closest editable representation of the canonical
// PDF region. Declared internal borders remain authoritative, and ordinary repeated values (without
// HideDuplicates metadata) are never coalesced.
function coalesceBorderlessDuplicateOwners(gridOwners, owners) {
  for (const owner of owners) {
    if (owner.excelMergedInto) continue;
    const descriptor = duplicateCellDescriptor(owner.cell);
    if (!descriptor || descriptor.suppressed) continue;
    const span = Math.max(1, owner.cell.colSpan || 1);
    let totalRowSpan = Math.max(1, owner.cell.rowSpan || 1);
    let previous = owner;
    while (owner.rowIndex + totalRowSpan < gridOwners.length) {
      const nextRow = owner.rowIndex + totalRowSpan;
      const candidate = gridOwners[nextRow][owner.start];
      if (!candidate || candidate === owner || candidate.rowIndex !== nextRow) break;
      if (candidate.start !== owner.start || Math.max(1, candidate.cell.colSpan || 1) !== span) break;
      const next = duplicateCellDescriptor(candidate.cell);
      if (!next || !next.suppressed
        || next.key !== descriptor.key
        || next.scope !== descriptor.scope
        || next.value !== descriptor.value
        || next.itemName !== descriptor.itemName
        || visibleBorderBetween(previous, candidate)) break;
      const candidateRowSpan = Math.max(1, candidate.cell.rowSpan || 1);
      candidate.excelMergedInto = owner;
      for (let rowOffset = 0; rowOffset < candidateRowSpan; rowOffset += 1) {
        for (let columnOffset = 0; columnOffset < span; columnOffset += 1) {
          gridOwners[nextRow + rowOffset][owner.start + columnOffset] = owner;
        }
      }
      totalRowSpan += candidateRowSpan;
      previous = candidate;
    }
    if (totalRowSpan > Math.max(1, owner.cell.rowSpan || 1)) {
      owner.cell = { ...owner.cell, rowSpan: totalRowSpan };
    }
  }
}

function excelCanGrowTextboxHeight(measureDoc, config, textbox, context, display, width, style) {
  if (!textbox?.canGrow || !display) return 0;
  // Excel cells do not implement RDL/PDF paragraph LineHeight, SpaceBefore or SpaceAfter. Reusing the PDF
  // physical height therefore creates very tall merged cells even though Excel renders the same wrapped
  // lines at its native font leading. Keep the exact font/run width measurement, but neutralize paragraph
  // spacing properties Excel cannot represent. Explicit newlines remain in `display` and are still counted.
  const excelTextbox = {
    ...textbox,
    style: { ...(textbox.style || {}), lineHeight: null },
    paragraphStyles: (textbox.paragraphStyles || []).map((paragraphStyle) => ({
      ...paragraphStyle,
      lineHeight: null,
      spaceBefore: 0,
      spaceAfter: 0,
    })),
    paragraphs: (textbox.paragraphs || []).map((paragraph) => paragraph.map((run) => ({
      ...run,
      style: { ...(run.style || {}), lineHeight: null },
    }))),
  };
  const excelSafeWidth = Math.max(1, width - EXCEL_MAX_DIGIT_WIDTH_PT);
  const excelSafeContentHeight = measureTextboxHeight(
    measureDoc,
    config,
    excelTextbox,
    context,
    display,
    excelSafeWidth,
  );
  const firstVisibleCharacter = Array.from(String(display)).find((character) => !/\s/u.test(character)) || 'M';
  const nativeSingleLineHeight = measureTextboxHeight(
    measureDoc,
    config,
    excelTextbox,
    context,
    firstVisibleCharacter,
    excelSafeWidth,
  );
  const wrapped = excelSafeContentHeight > nativeSingleLineHeight + (1 / POINT_PRECISION);
  const borderClearance = ['top', 'bottom'].reduce((sum, side) => {
    const border = style?.borders?.[side];
    const borderStyle = String(styleValue(border?.style, context, 'None'));
    if (!border || /^none$/i.test(borderStyle)) return sum;
    // Half of each border is painted inside the cell display area. Include it so the last glyph baseline
    // cannot visually collide with a horizontal edge at common Excel zoom levels.
    return sum + (Math.max(0, styleSize(border.width, context, 1)) / 2);
  }, 0);
  return excelSafeContentHeight
    + styleSize(style?.paddingTop, context, 2)
    + styleSize(style?.paddingBottom, context, 2)
    // Excel's glyph descent and character-width rounding are viewer/font dependent. A small once-per-cell
    // clearance prevents the last baseline touching a border without reserving an entire extra line.
    + (wrapped ? EXCEL_TEXT_CLEARANCE_PT : 0)
    + borderClearance;
}

// Free-form report content is frequently nested in one or more borderless rectangles. Excel does not
// auto-fit merged cells, so measuring only top-level textboxes leaves a nested CanGrow textbox at its
// small design-time height and clips its wrapped value. Resolve each container recursively: grow textboxes
// from the shared PDF metrics, displace only later items in the same horizontal lane, and retain the
// rectangle's declared trailing space. The resulting item tree remains coordinate based and editable.
function resolveExcelFreeformLayout(item, context, config, measureDoc) {
  const declaredHeight = Math.max(0, item.height || 0);
  if (item.type === 'Textbox') {
    const display = cellString(textForItem(item, context));
    const occupiedHeight = Math.max(
      2,
      declaredHeight || DEFAULT_ROW_POINTS,
      excelCanGrowTextboxHeight(measureDoc, config, item, context, display, item.width || 0, item.style || {}),
    );
    return { item: { ...item, height: occupiedHeight }, occupiedHeight };
  }
  if (item.type !== 'Rectangle' || !(item.items || []).length) {
    return { item, occupiedHeight: Math.max(2, declaredHeight || DEFAULT_ROW_POINTS) };
  }

  const children = item.items.map((child, sourceIndex) => {
    const resolved = resolveExcelFreeformLayout(child, context, config, measureDoc);
    return {
      ...resolved,
      sourceIndex,
      designTop: child.top || 0,
      designHeight: child.height || 0,
      resolvedTop: child.top || 0,
    };
  });
  const layoutOrder = [...children].sort((left, right) => (
    left.designTop - right.designTop
    || (left.item.left || 0) - (right.item.left || 0)
    || left.sourceIndex - right.sourceIndex
  ));
  for (const [index, child] of layoutOrder.entries()) {
    for (let previousIndex = 0; previousIndex < index; previousIndex += 1) {
      const previous = layoutOrder[previousIndex];
      const previousDesignBottom = point(previous.designTop + previous.designHeight);
      if (previousDesignBottom > child.designTop + COINCIDENT_EDGE_TOLERANCE_PT) continue;
      if (!horizontalDesignOverlap(previous.item, child.item)) continue;
      const originalGap = Math.max(0, child.designTop - previousDesignBottom);
      child.resolvedTop = Math.max(
        child.resolvedTop,
        point(previous.resolvedTop + previous.occupiedHeight + originalGap),
      );
    }
    child.item = { ...child.item, top: child.resolvedTop, height: child.occupiedHeight };
  }
  const designBottom = Math.max(0, ...children.map((child) => child.designTop + child.designHeight));
  const resolvedBottom = Math.max(0, ...children.map((child) => child.resolvedTop + child.occupiedHeight));
  const trailingSpace = Math.max(0, declaredHeight - designBottom);
  const occupiedHeight = Math.max(2, declaredHeight, point(resolvedBottom + trailingSpace));
  const resolvedBySource = [...children]
    .sort((left, right) => left.sourceIndex - right.sourceIndex)
    .map((child) => child.item);
  return {
    item: { ...item, items: resolvedBySource, height: occupiedHeight },
    occupiedHeight,
  };
}

function renderReportTablix({ worksheet, model, item, request, globals, config, xGrid, startRow, merges, tablixCache, measureDoc }) {
  const { rows, columns } = tablixLayout(item, request, globals, model, tablixCache);
  const placements = computeCellPlacements(rows, columns.length);
  const datasets = normalizeDatasets(model, request);
  const enforceBottomClosure = shouldEnforceTablixBottom(rows, item);
  const gridOwners = rows.map(() => new Array(columns.length).fill(null));
  const owners = [];
  // Pictures a cell carries on its canvas, collected while the grid is written and anchored later.
  const canvasPlacements = [];
  // Bundled subreports expose their frame and their rendered child regions as sibling normalized
  // records. Keep the frame geometry so a chart/image that belongs to that child report still respects
  // the invoking Subreport's border when it is anchored later as an Excel floating drawing.
  const subreportFrames = [];
  // The columns this region actually painted into its scratch worksheet. A tablix's own grid is bounded by
  // its declared column widths, but a nested region or a subreport frame is positioned at its own left and
  // width and can legitimately extend past that bound. The copy step reads back exactly the columns named
  // here, so anything painted outside them — most visibly a subreport frame's right edge and the outer
  // rule above and below a sibling region — would otherwise be written to the scratch sheet and silently
  // dropped. Kept separate from the declared table range, which still defines the freeze pane/autofilter.
  const painted = { startCol: Infinity, endCol: -Infinity };
  const notePainted = (range) => {
    if (!range) return;
    painted.startCol = Math.min(painted.startCol, range.startCol);
    painted.endCol = Math.max(painted.endCol, range.endCol);
  };
  const columnOffsets = [0];
  columns.forEach((width) => columnOffsets.push(point(columnOffsets[columnOffsets.length - 1] + width)));
  const addNestedYBoundaries = (nested, baseTop, target) => {
    let cursor = point(baseTop + (nested.item.top || 0));
    target.add(cursor);
    // A subreport frame carries no rows: its declared box is the only boundary it contributes.
    if (nested.subreportFrame) {
      target.add(point(cursor + (nested.item.height || 0)));
      return;
    }
    for (const row of nested.rows || []) {
      const rowTop = cursor;
      cursor = point(cursor + (row.height || DEFAULT_ROW_POINTS));
      target.add(cursor);
      for (const cell of row.cells || []) {
        for (const child of cell.nestedTablixes || []) addNestedYBoundaries(child, rowTop, target);
      }
    }
  };
  const rowProfiles = rows.map((row) => {
    const boundaries = new Set([0, point(row.height || DEFAULT_ROW_POINTS)]);
    for (const cell of row.cells || []) {
      for (const nested of cell.nestedTablixes || []) addNestedYBoundaries(nested, 0, boundaries);
    }
    const maximum = Math.max(...boundaries);
    boundaries.add(maximum);
    return [...boundaries].sort((a, b) => a - b);
  });

  rows.forEach((row, rowIndex) => {
    row.cells.forEach((cell, cellIndex) => {
      const start = placements[rowIndex][cellIndex];
      if (start === undefined || start < 0) return;
      const context = materializedCellContext(cell, row, {
        parameters: request.parameters || {}, globals, dataset: datasets[item.datasetName] || [], datasets,
      });
      const presentation = cellStyle(item, cell, context);
      const owner = { cell, rowIndex, start, ...presentation };
      owners.push(owner);
      for (let r = 0; r < Math.max(1, cell.rowSpan || 1) && rowIndex + r < rows.length; r += 1) {
        for (let c = 0; c < Math.max(1, cell.colSpan || 1) && start + c < columns.length; c += 1) {
          if (gridOwners[rowIndex + r][start + c]) throw new ServiceError('RDL_INVALID', 'RDL produced overlapping Excel tablix cells');
          gridOwners[rowIndex + r][start + c] = owner;
        }
      }
    });
  });
  coalesceBorderlessDuplicateOwners(gridOwners, owners);

  // Measure a nested tablix using the same content-aware rule as its parent. Excel never auto-fits merged
  // cells, so relying on wrapText alone clips growing text inside nested data regions. Cache by materialized
  // nested instance: the same RDL item may occur under several group scopes with different rows and values.
  const nestedMeasurements = new WeakMap();
  const measureNestedTablix = (nested) => {
    const cached = nestedMeasurements.get(nested);
    if (cached) return cached;
    const nestedRows = nested.rows || [];
    const nestedColumns = nested.columns || nested.item.columns || [];
    const nestedLayoutItem = nested.item.hasColumnGroups
      ? { ...nested.item, columns: nestedColumns, width: nestedColumns.reduce((sum, width) => sum + width, 0) }
      : { ...nested.item, columns: nestedColumns };
    const { columnsPt } = resolveGridColumns(nestedLayoutItem);
    const nestedPlacements = computeCellPlacements(nestedRows, columnsPt.length);
    const heights = nestedRows.map((row) => Math.max(2, row.height || DEFAULT_ROW_POINTS));
    // Store before descending so malformed recursive containment cannot recurse forever.
    const result = { heights, columnsPt };
    nestedMeasurements.set(nested, result);

    nestedRows.forEach((row, rowIndex) => {
      row.cells.forEach((cell, cellIndex) => {
        if (cell.hidden) return;
        const start = nestedPlacements[rowIndex][cellIndex];
        if (start === undefined || start < 0) return;
        const context = materializedCellContext(cell, row, {
          parameters: request.parameters || {}, globals, dataset: datasets[nested.item.datasetName] || [], datasets,
        });
        const { textbox, style } = cellStyle(nested.item, cell, context);
        const span = Math.max(1, cell.colSpan || 1);
        const rowSpan = Math.max(1, cell.rowSpan || 1);
        const width = columnsPt.slice(start, start + span).reduce((sum, value) => sum + value, 0);
        let required = 0;
        const display = cellText(cell);
        required = excelCanGrowTextboxHeight(
          measureDoc,
          config,
          textbox,
          context,
          display,
          width,
          style,
        );
        for (const child of cell.nestedTablixes || []) {
          const childMeasurement = measureNestedTablix(child);
          required = Math.max(required,
            (child.item.top || 0) + childMeasurement.heights.reduce((sum, height) => sum + height, 0));
        }
        if (required <= 0) return;
        const endRow = Math.min(nestedRows.length - 1, rowIndex + rowSpan - 1);
        const available = heights.slice(rowIndex, endRow + 1).reduce((sum, height) => sum + height, 0);
        if (required > available) heights[endRow] += required - available;
      });
    });
    return result;
  };

  // SSRS free-form displacement, applied inside a tablix cell. A cell whose CellContents Rectangle was
  // flattened is a canvas of positioned peers, and a child data region that renders taller than its
  // declared height pushes the later peers in its horizontal lane down by that growth. Without it two
  // design-time disjoint child regions can be scheduled onto the same worksheet rows, which surfaces as
  // overlapping merged ranges. Same lane rule as the section and rectangle layouts.
  const resolvedNestedTops = new WeakMap();
  const nestedTopsFor = (cell) => {
    const cached = resolvedNestedTops.get(cell);
    if (cached) return cached;
    const peers = [];
    (cell.items || []).forEach((entry, sourceIndex) => {
      // Tablix/Subreport entries are the design-time source of the materialized nested regions added
      // below; counting them twice would displace every later peer by a whole extra region.
      if (entry.type === 'Tablix' || entry.type === 'Subreport') return;
      const height = Math.max(0, entry.height || 0);
      peers.push({
        sourceIndex, item: entry, designTop: point(entry.top || 0), designHeight: height, occupiedHeight: height, resolvedTop: point(entry.top || 0), nested: null,
      });
    });
    (cell.nestedTablixes || []).forEach((nested, index) => {
      const measured = measureNestedTablix(nested).heights.reduce((sum, height) => sum + height, 0);
      const declared = Math.max(0, nested.item.height || 0);
      peers.push({
        sourceIndex: (cell.items || []).length + index,
        item: nested.item,
        designTop: point(nested.item.top || 0),
        designHeight: declared,
        occupiedHeight: Math.max(declared, measured),
        resolvedTop: point(nested.item.top || 0),
        nested,
      });
    });
    const layoutOrder = [...peers].sort((left, right) => left.designTop - right.designTop
      || (left.item.left || 0) - (right.item.left || 0)
      || left.sourceIndex - right.sourceIndex);
    for (const [index, peer] of layoutOrder.entries()) {
      for (let previousIndex = 0; previousIndex < index; previousIndex += 1) {
        const previous = layoutOrder[previousIndex];
        const previousDesignBottom = point(previous.designTop + previous.designHeight);
        if (previousDesignBottom > peer.designTop + COINCIDENT_EDGE_TOLERANCE_PT) continue;
        if (!horizontalDesignOverlap(previous.item, peer.item)) continue;
        const originalGap = Math.max(0, peer.designTop - previousDesignBottom);
        peer.resolvedTop = Math.max(peer.resolvedTop, point(previous.resolvedTop + previous.occupiedHeight + originalGap));
      }
    }
    const tops = new WeakMap();
    // A canvas peer is displaced by the same rule, so record its resolved top too: a heading or a chart
    // that follows a grown child region has to move down with it, exactly as the PDF places it.
    for (const peer of peers) tops.set(peer.nested || peer.item, peer.resolvedTop);
    resolvedNestedTops.set(cell, tops);
    return tops;
  };
  const nestedTop = (cell, nested) => nestedTopsFor(cell).get(nested) ?? point(nested.item.top || 0);

  // Excel does not auto-fit merged cells. Measure every logical owner once, then grow the final row of its
  // vertical span only when CanGrow is enabled and the declared combined row heights cannot contain the
  // text. Nested tablixes contribute their measured height too. This preserves RDL row proportions and
  // CanGrow=false semantics while keeping merged and nested group content fully readable.
  const measuredHeights = rows.map((row, index) => Math.max(
    2,
    row.height || DEFAULT_ROW_POINTS,
    rowProfiles[index].at(-1) || 0,
  ));
  for (const owner of owners) {
    const display = cellText(owner.cell);
    const span = Math.max(1, owner.cell.colSpan || 1);
    const rowSpan = Math.max(1, owner.cell.rowSpan || 1);
    const width = point(columnOffsets[owner.start + span] - columnOffsets[owner.start]);
    let required = 0;
    required = excelCanGrowTextboxHeight(
      measureDoc,
      config,
      owner.textbox,
      owner.context,
      display,
      width,
      owner.style,
    );
    for (const nested of owner.cell.nestedTablixes || []) {
      const nestedMeasurement = measureNestedTablix(nested);
      required = Math.max(required,
        nestedTop(owner.cell, nested) + nestedMeasurement.heights.reduce((sum, height) => sum + height, 0));
    }
    if (required <= 0) continue;
    const endRow = Math.min(rows.length - 1, owner.rowIndex + rowSpan - 1);
    const available = measuredHeights.slice(owner.rowIndex, endRow + 1).reduce((sum, height) => sum + height, 0);
    if (required > available) measuredHeights[endRow] += required - available;
  }

  // Replace declared child-row boundaries with their measured CanGrow boundaries before mapping them to
  // physical Excel rows. Otherwise a nested row may grow in measurement but still be rendered into its old
  // one-line interval.
  // The rendered extent of an invoking Subreport's box: the union of the regions that invocation
  // produced, floored at the placeholder and child-body geometry the frame was created with.
  const subreportFrameExtent = (nested) => nested.subreportFrame.reduce((box, sibling) => {
    const measurement = measureNestedTablix(sibling);
    return {
      width: Math.max(box.width, (sibling.item.left || 0) - (nested.item.left || 0)
        + measurement.columnsPt.reduce((sum, value) => sum + value, 0)),
      height: Math.max(box.height, (sibling.item.top || 0) - (nested.item.top || 0)
        + measurement.heights.reduce((sum, value) => sum + value, 0)),
    };
  }, { width: nested.item.width || 0, height: nested.item.height || 0 });
  const addMeasuredNestedYBoundaries = (nested, baseTop, target, ownTop) => {
    let cursor = point(baseTop + ownTop);
    target.add(cursor);
    if (nested.subreportFrame) {
      target.add(point(cursor + subreportFrameExtent(nested).height));
      return;
    }
    const measurement = measureNestedTablix(nested);
    (nested.rows || []).forEach((row, index) => {
      const rowTop = cursor;
      cursor = point(cursor + measurement.heights[index]);
      target.add(cursor);
      for (const cell of row.cells || []) {
        for (const child of cell.nestedTablixes || []) addMeasuredNestedYBoundaries(child, rowTop, target, nestedTop(cell, child));
      }
    });
  };
  rows.forEach((row, rowIndex) => {
    const boundaries = new Set([0, point(row.height || DEFAULT_ROW_POINTS)]);
    for (const cell of row.cells || []) {
      for (const nested of cell.nestedTablixes || []) addMeasuredNestedYBoundaries(nested, 0, boundaries, nestedTop(cell, nested));
    }
    rowProfiles[rowIndex] = [...boundaries].sort((left, right) => left - right);
  });

  // A logical parent row may contain a child grid with several physical rows. Split only that parent row
  // at the child-grid boundaries; ordinary surrounding cells are vertically merged across the resulting
  // native Excel rows, keeping the workbook editable without flattening the child tablix to text.
  rowProfiles.forEach((profile, rowIndex) => {
    const declared = profile.at(-1) || 0;
    if (measuredHeights[rowIndex] > declared) {
      const containsNestedGrid = rows[rowIndex].cells.some((cell) => (cell.nestedTablixes || []).length > 0);
      if (containsNestedGrid) {
        // The child's final edge can coincide with the parent's declared row edge. When another cell grows
        // the parent, replacing that edge erases a boundary required to place the child; preserve it and
        // append the grown parent edge. This is the native-grid equivalent of a short child table followed
        // by unused vertical space in its taller invoking cell.
        profile.push(point(measuredHeights[rowIndex]));
      } else {
        profile[profile.length - 1] = point(measuredHeights[rowIndex]);
      }
    }
    rowProfiles[rowIndex] = splitTallRowIntervals(profile);
  });
  // A canvas cell's children sit at their own displaced tops and can extend past the height the nested
  // regions alone imply. Fold their extents into the row profile: the row then reserves the full canvas,
  // and every child lands on an exact boundary instead of being snapped onto a neighbour's row — which
  // is what let one group instance's heading collide with the previous instance's trailing paragraphs.
  rows.forEach((row, rowIndex) => {
    for (const cell of row.cells || []) {
      if (!isFreeFormCell(cell)) continue;
      const displaced = nestedTopsFor(cell);
      const profile = new Set(rowProfiles[rowIndex]);
      for (const child of cell.items || []) {
        if (child.type === 'Tablix' || child.type === 'Subreport') continue;
        const childTop = point(displaced.get(child) ?? (child.top || 0));
        profile.add(childTop);
        profile.add(point(childTop + (child.height || 0)));
      }
      // The regions themselves are displaced, and they RENDER at their measured (grown) row heights, not
      // their declared ones. Record the boundaries the renderer will actually use, so a canvas child that
      // follows a grown region is measured against the same extent the displacement gave it.
      const addMeasuredBoundaries = (nested, baseTop) => {
        let cursor = point(baseTop);
        profile.add(cursor);
        const heights = measureNestedTablix(nested).heights;
        (nested.rows || []).forEach((nestedRow, nestedIndex) => {
          const rowTop = cursor;
          cursor = point(cursor + (heights[nestedIndex] ?? nestedRow.height ?? DEFAULT_ROW_POINTS));
          profile.add(cursor);
          for (const nestedCell of nestedRow.cells || []) {
            for (const grandchild of nestedCell.nestedTablixes || []) {
              addMeasuredBoundaries(grandchild, rowTop + (grandchild.item.top || 0));
            }
          }
        });
      };
      for (const nested of cell.nestedTablixes || []) {
        addMeasuredBoundaries(nested, point(displaced.get(nested) ?? (nested.item.top || 0)));
      }
      rowProfiles[rowIndex] = [...profile].sort((left, right) => left - right);
    }
  });

  const physicalStarts = [];
  let physicalCursor = startRow;
  rowProfiles.forEach((profile) => {
    physicalStarts.push(physicalCursor);
    physicalCursor += Math.max(1, profile.length - 1);
  });
  // The point offset of each logical row from the tablix top, for anchoring cell pictures.
  const rowTops = [];
  let rowTopCursor = 0;
  rowProfiles.forEach((profile) => {
    rowTops.push(point(rowTopCursor));
    rowTopCursor += profile.at(-1) || 0;
  });
  rowProfiles.forEach((profile, rowIndex) => {
    for (let index = 0; index < profile.length - 1; index += 1) {
      const height = profile[index + 1] - profile[index];
      if (height > MAX_EXCEL_ROW_HEIGHT) throw new ServiceError('UNSUPPORTED_FEATURE', 'Rendered Excel cell text exceeds the 409-point row-height limit');
      worksheet.getRow(physicalStarts[rowIndex] + index).height = Math.max(2, height);
    }
  });

  // The invoking Subreport's own box. It carries no rows: SSRS paints the placeholder's border around the
  // whole child report that was rendered, so the range is the union of the sibling regions this invocation
  // produced, floored at the placeholder and child-body geometry. applyRegionBorder only adds perimeter
  // edges, so the child's own grid lines inside the box are untouched.
  const renderSubreportFrame = (nested, parentCellLeft, parentRowIndex, baseTop, ownTop) => {
    const frameLeft = nested.item.left || 0;
    const extent = subreportFrameExtent(nested);
    const context = {
      parameters: nested.parameters || request.parameters || {},
      globals: nested.globals || globals,
      datasets: nested.datasets || datasets,
      dataset: [],
      fields: {},
    };
    const borders = Object.fromEntries(['top', 'bottom', 'left', 'right']
      .map((side) => [side, excelBorderSide(nested.item.style?.borders?.[side], context)]));
    if (!Object.values(borders).some(Boolean)) return;
    const left = point(parentCellLeft + frameLeft);
    const top = point(baseTop + ownTop);
    const profile = rowProfiles[parentRowIndex];
    const range = {
      ...gridRange(xGrid, left, extent.width),
      startRow: physicalStarts[parentRowIndex] + boundaryIndex(profile, top),
      endRow: physicalStarts[parentRowIndex] + Math.max(
        boundaryIndex(profile, top) + 1,
        boundaryIndex(profile, point(top + extent.height)),
      ) - 1,
    };
    notePainted(range);
    applyRegionBorder(worksheet, range, borders);
    subreportFrames.push({
      item: nested.item,
      context,
      left,
      top: point(rowTops[parentRowIndex] + top),
      right: point(left + extent.width),
      bottom: point(rowTops[parentRowIndex] + top + extent.height),
    });
  };
  const framesContainingPicture = (left, top, width, height) => {
    const right = point(left + width);
    const bottom = point(top + height);
    return subreportFrames.filter((frame) => (
      left >= frame.left - COINCIDENT_EDGE_TOLERANCE_PT
      && right <= frame.right + COINCIDENT_EDGE_TOLERANCE_PT
      && top >= frame.top - COINCIDENT_EDGE_TOLERANCE_PT
      && bottom <= frame.bottom + COINCIDENT_EDGE_TOLERANCE_PT
    ));
  };
  const renderNested = (nested, parentCellLeft, parentRowIndex, baseTop, ownTop) => {
    if (nested.subreportFrame) {
      renderSubreportFrame(nested, parentCellLeft, parentRowIndex, baseTop, ownTop);
      return;
    }
    const nestedParameters = nested.parameters || request.parameters || {};
    const nestedDatasets = nested.datasets || datasets;
    const nestedGlobals = nested.globals || globals;
    const nestedRows = nested.rows || [];
    const nestedColumns = nested.columns || nested.item.columns || [];
    const nestedPlacements = computeCellPlacements(nestedRows, nestedColumns.length);
    const nestedOffsets = [0];
    nestedColumns.forEach((width) => nestedOffsets.push(point(nestedOffsets.at(-1) + width)));
    const nestedOwners = nestedRows.map(() => new Array(nestedColumns.length).fill(null));
    const rowOffsets = [point(baseTop + ownTop)];
    const nestedHeights = measureNestedTablix(nested).heights;
    nestedRows.forEach((row, index) => rowOffsets.push(point(rowOffsets.at(-1) + nestedHeights[index])));
    for (const [rowIndex, row] of nestedRows.entries()) {
      for (const [cellIndex, cell] of row.cells.entries()) {
        const start = nestedPlacements[rowIndex][cellIndex];
        const context = materializedCellContext(cell, row, {
          parameters: nestedParameters,
          globals: nestedGlobals,
          dataset: nestedDatasets[nested.item.datasetName] || [],
          datasets: nestedDatasets,
        });
        const presentation = cellStyle(nested.item, cell, context);
        const owner = { cell, rowIndex, start, ...presentation };
        for (let r = 0; r < Math.max(1, cell.rowSpan || 1) && rowIndex + r < nestedRows.length; r += 1) {
          for (let c = 0; c < Math.max(1, cell.colSpan || 1) && start + c < nestedColumns.length; c += 1) {
            if (nestedOwners[rowIndex + r][start + c]) throw new ServiceError('RDL_INVALID', 'Nested tablix produced overlapping Excel cells');
            nestedOwners[rowIndex + r][start + c] = owner;
          }
        }
      }
    }
    for (const [rowIndex, row] of nestedRows.entries()) {
      for (const [cellIndex, cell] of row.cells.entries()) {
        const start = nestedPlacements[rowIndex][cellIndex];
        const owner = nestedOwners[rowIndex][start];
        const colSpan = Math.max(1, cell.colSpan || 1);
        const rowSpan = Math.max(1, cell.rowSpan || 1);
        const left = point(parentCellLeft + (nested.item.left || 0) + nestedOffsets[start]);
        const width = point(nestedOffsets[Math.min(nestedColumns.length, start + colSpan)] - nestedOffsets[start]);
        const top = rowOffsets[rowIndex];
        const bottom = rowOffsets[Math.min(nestedRows.length, rowIndex + rowSpan)];
        const profile = rowProfiles[parentRowIndex];
        const range = {
          ...gridRange(xGrid, left, width),
          startRow: physicalStarts[parentRowIndex] + boundaryIndex(profile, top),
          endRow: physicalStarts[parentRowIndex] + boundaryIndex(profile, bottom) - 1,
        };
        notePainted(range);
        const borders = reportCellBorders(nestedOwners, owner, nested.item.style, shouldEnforceTablixBottom(nestedRows, nested.item), nestedRows, nested.item);
        if (isFreeFormCell(cell)) {
          const frame = canvasCellFrame(owner, left, rowTops[parentRowIndex] + top, width, bottom - top);
          for (const child of cell.items || []) {
            // Tablix/Subreport entries are the design-time source of the materialized child regions, which
            // the recursive call below renders.
            if (child.type === 'Tablix' || child.type === 'Subreport') continue;
            if (isHidden(child.hidden, owner.context)) continue;
            canvasPlacements.push({
              item: child,
              // A chart on a region's canvas is scoped to that region's rows, exactly as the canonical PDF
              // pass scopes it, rather than to the whole report dataset.
              context: nested.item.datasetName
                ? {
                  ...owner.context,
                  datasets: { ...owner.context.datasets, [nested.item.datasetName]: owner.context.dataset },
                }
                : owner.context,
              // The report the picture's definition came from, so an embedded child-report image resolves
              // against that report rather than the one that invoked it.
              model: nested.model || null,
              left: point(left + (child.left || 0)),
              // The scratch row this picture's top coincides with. Excel cannot make a row thinner than
              // two points, so a row profile built from sub-point RDL edges renders taller than it was
              // declared and every later boundary shifts down with it. A picture resolved from its raw
              // point offset then lands above the very cells it belongs to and covers their top rule —
              // most visibly the enclosing subreport frame, whose own edge went through this same row
              // mapping. Anchor to the row instead, so both land on one grid line.
              localStartRow: (() => {
                const index = optionalBoundaryIndex(rowProfiles[parentRowIndex], point(top + (child.top || 0)));
                return index === null ? null : physicalStarts[parentRowIndex] + index;
              })(),
              // Canvas placements are resolved by the caller against the section row grid, so they must
              // carry an offset from the *tablix* top. Every nested coordinate here — `rowOffsets`, and
              // the `baseTop`/`ownTop` a recursive call passes on — is local to the parent logical row,
              // because the nested grid is measured against that row's own profile. Without the parent
              // row's own offset, two instances of the same group each report the same local top, and the
              // caller schedules both pictures onto one range: the second overpaints the first and the
              // earlier group instance renders with no picture at all.
              top: point(rowTops[parentRowIndex] + top + (child.top || 0)),
              containerFrames: [frame, ...framesContainingPicture(
                left + (child.left || 0),
                rowTops[parentRowIndex] + top + (child.top || 0),
                child.width || 0,
                child.height || 0,
              )],
            });
          }
          applyRegionBorder(worksheet, range, borders);
        } else {
          if (range.startRow !== range.endRow || range.startCol !== range.endCol) mergeSafe(worksheet, range, merges);
          const target = worksheet.getCell(range.startRow, range.startCol);
          const { value, numFmt } = excelCellValue(cell, owner.context);
          const display = cellText(cell);
          target.value = typeof value === 'string' ? (richTextValue(owner.textbox, owner.context, display) || value) : value;
          if (numFmt) target.numFmt = numFmt;
          applyFillFontAlignment(target, owner.style || {}, owner.context);
          target.border = borders;
        }
        const childCellLeft = left;
        for (const child of cell.nestedTablixes || []) renderNested(child, childCellLeft, parentRowIndex, top, nestedTop(cell, child));
      }
    }
  };

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const excelRow = physicalStarts[rowIndex];
    for (let columnIndex = 0; columnIndex < columns.length;) {
      const owner = gridOwners[rowIndex][columnIndex];
      if (!owner) { columnIndex += 1; continue; }
      if (columnIndex !== owner.start) { columnIndex += 1; continue; }
      const span = Math.max(1, owner.cell.colSpan || 1);
      if (rowIndex !== owner.rowIndex) { columnIndex += span; continue; }
      const rowSpan = Math.max(1, owner.cell.rowSpan || 1);
      const left = point((item.left || 0) + columnOffsets[columnIndex]);
      const width = point(columnOffsets[columnIndex + span] - columnOffsets[columnIndex]);
      const columnsRange = gridRange(xGrid, left, width);
      const endLogicalRow = Math.min(rows.length - 1, rowIndex + rowSpan - 1);
      const range = {
        startRow: excelRow,
        endRow: physicalStarts[endLogicalRow] + Math.max(1, rowProfiles[endLogicalRow].length - 1) - 1,
        ...columnsRange,
      };
      notePainted(range);
      // The section-wide coordinate grid also contains boundaries from unrelated report items (for example,
      // a logo edge can fall inside a tablix column). A single logical RDL cell may therefore cover several
      // physical Excel columns even when its ColSpan is 1. Merge the occupied Excel region, not merely RDL
      // spans, or the extra grid slices appear as blank gaps inside the table.
      const hasNested = (owner.cell.nestedTablixes || []).length > 0;
      // A free-form cell is a flattened CellContents Rectangle: its children keep their own declared
      // position and size. Joining them into one cell value — right for every ordinary cell — collapsed a
      // whole page of headings and prose into a single, usually invisible, merged cell. Place each child
      // at its own coordinates instead, exactly as the canonical PDF pass does.
      const freeForm = isFreeFormCell(owner.cell);
      let target = null;
      if (!freeForm) {
        if (!hasNested && (range.startRow !== range.endRow || range.startCol !== range.endCol)) mergeSafe(worksheet, range, merges);
        target = worksheet.getCell(excelRow, range.startCol);
        const { value, numFmt } = excelCellValue(owner.cell, owner.context);
        const display = cellText(owner.cell);
        target.value = typeof value === 'string' ? (richTextValue(owner.textbox, owner.context, display) || value) : value;
        if (numFmt) target.numFmt = numFmt;
        applyFillFontAlignment(target, owner.style || {}, owner.context);
      }
      if (freeForm) {
        const frameHeight = rows.slice(rowIndex, endLogicalRow + 1)
          .reduce((sum, _row, index) => sum + (rowProfiles[rowIndex + index].at(-1) || 0), 0);
        const frame = canvasCellFrame(owner, left, rowTops[rowIndex], width, frameHeight);
        for (const child of owner.cell.items || []) {
          // Tablix/Subreport entries are the design-time source of the materialized nested regions,
          // which the child-region path below already renders.
          if (child.type === 'Tablix' || child.type === 'Subreport') continue;
          if (isHidden(child.hidden, owner.context)) continue;
          canvasPlacements.push({
            item: child,
            // A chart on a canvas cell is scoped to that group instance: point the region dataset at the
            // cell's own rows, as the canonical PDF pass does, or it aggregates the whole report.
            context: item.datasetName
              ? { ...owner.context, datasets: { ...owner.context.datasets, [item.datasetName]: owner.context.dataset } }
              : owner.context,
            left: point(left + (child.left || 0)),
            // Points from the tablix top. A child that follows a grown nested region is displaced by the
            // same SSRS rule the regions use, so it moves down with it exactly as the PDF places it.
            top: point(rowTops[rowIndex] + (nestedTopsFor(owner.cell).get(child) ?? (child.top || 0))),
            // See the nested push below: the row this picture's top coincides with, so it resolves onto
            // the same grid line as the cells around it once Excel's minimum row height has widened them.
            localStartRow: (() => {
              const index = optionalBoundaryIndex(
                rowProfiles[rowIndex],
                point(nestedTopsFor(owner.cell).get(child) ?? (child.top || 0)),
              );
              return index === null ? null : physicalStarts[rowIndex] + index;
            })(),
            containerFrames: [frame, ...framesContainingPicture(
              left + (child.left || 0),
              rowTops[rowIndex] + (nestedTopsFor(owner.cell).get(child) ?? (child.top || 0)),
              child.width || 0,
              child.height || 0,
            )],
          });
        }
      }
      if (hasNested) {
        if (!freeForm) applyRegionStyle(worksheet, range, owner.style || {}, owner.context, { includeBorders: false });
        for (const nested of owner.cell.nestedTablixes || []) renderNested(nested, left, rowIndex, 0, nestedTop(owner.cell, nested));
        // The child grid writes its own cells inside this region, so the enclosing cell's box goes on last
        // and only on the perimeter: applied first it would be overwritten by the child's first row, and
        // applied to the anchor cell it would draw this cell's bottom rule across its top physical row.
        applyRegionBorder(worksheet, range, reportCellBorders(gridOwners, owner, item.style, enforceBottomClosure, rows, item));
      } else {
        // A free-form cell wrote no anchor cell of its own; its box goes on the region it occupies.
        if (target) target.border = reportCellBorders(gridOwners, owner, item.style, enforceBottomClosure, rows, item);
        else applyRegionBorder(worksheet, range, reportCellBorders(gridOwners, owner, item.style, enforceBottomClosure, rows, item));
      }
      columnIndex += span;
    }
  }

  const headerRows = rows.reduce((sum, row, index) => (
    row.isHeader ? sum + Math.max(1, rowProfiles[index].length - 1) : sum
  ), 0);
  const tableRange = gridRange(xGrid, point(item.left || 0), columnOffsets[columnOffsets.length - 1]);
  const fixedLogicalColumns = item.fixedRowHeaders
    ? (item.rowHeaderColumns?.length || 0)
    : Math.min(
      item.rowHeaderColumns?.length || 0,
      (item.rowMemberPaths || []).filter((path) => path.some((member) => member.fixedData)).length,
    );
  const fixedColumnsSplit = fixedLogicalColumns > 0
    ? gridRange(xGrid, point(item.left || 0), columnOffsets[fixedLogicalColumns]).endCol
    : 0;
  return {
    rowsConsumed: physicalCursor - startRow,
    canvasPlacements,
    startRow,
    endRow: physicalCursor - 1,
    headerRows,
    dynamic: rows.some((row) => !row.isHeader && !row.isStatic),
    fixedColumnsSplit,
    ...tableRange,
    paintedStartCol: Math.min(tableRange.startCol, painted.startCol),
    paintedEndCol: Math.max(tableRange.endCol, painted.endCol),
  };
}

function horizontalDesignOverlap(left, right) {
  const leftStart = point(left.left || 0);
  const leftEnd = point(leftStart + (left.width || 0));
  const rightStart = point(right.left || 0);
  const rightEnd = point(rightStart + (right.width || 0));
  return leftStart < rightEnd - COINCIDENT_EDGE_TOLERANCE_PT
    && rightStart < leftEnd - COINCIDENT_EDGE_TOLERANCE_PT;
}

function worksheetRowBoundaries(worksheet, count) {
  const boundaries = [0];
  for (let row = 1; row <= count; row += 1) {
    boundaries.push(point(boundaries.at(-1) + (worksheet.getRow(row).height || DEFAULT_ROW_POINTS)));
  }
  return boundaries;
}

function decodeColumn(letters) {
  let value = 0;
  for (const character of letters) value = value * 26 + character.charCodeAt(0) - 64;
  return value;
}

function decodeAddress(address) {
  const match = /^([A-Z]+)(\d+)$/.exec(address);
  if (!match) throw new ServiceError('RDL_INVALID', `Invalid Excel cell address: ${address}`);
  return { column: decodeColumn(match[1]), row: Number(match[2]) };
}

function decodeMerge(range) {
  const [start, end] = range.split(':').map(decodeAddress);
  return { startRow: start.row, startCol: start.column, endRow: end.row, endCol: end.column };
}

function cloneExcelValue(value) {
  if (value === null || value === undefined) return value;
  return structuredClone(value);
}

// Where a tablix cell's canvas picture starts, in the section's point space. A picture anchored to one of
// the region's own rows takes that row's rendered top, so it lands on the same section grid line as the
// cells it sits among: Excel's two-point minimum row height moves every boundary below a sub-point RDL
// edge, and the raw declared offset no longer names the same line the region's own grid was built on.
function canvasPlacementTop(plan, placed) {
  const localTop = placed.localStartRow !== null && placed.localStartRow !== undefined
    ? plan.localBoundaries?.[placed.localStartRow - 1]
    : null;
  return point(plan.resolvedTop + (localTop ?? placed.top));
}

function copyPlannedTablix({ worksheet, plan, yGrid, startRow, merges }) {
  const offset = plan.resolvedTop;
  const local = plan.localBoundaries;
  const mergeByMaster = new Map((plan.worksheet.model.merges || []).map((range) => {
    const decoded = decodeMerge(range);
    return [`${decoded.startRow}:${decoded.startCol}`, decoded];
  }));
  const mapStart = (localRow) => startRow + boundaryIndex(yGrid, point(offset + local[localRow - 1]));
  const mapEnd = (localRow) => startRow + boundaryIndex(yGrid, point(offset + local[localRow])) - 1;

  // Read back every column the region painted, not only its declared table range: a nested region or a
  // subreport frame legitimately extends past the parent grid's own columns, and those cells carry the
  // frame's right edge and the rules that close it above and below a sibling region.
  const firstPaintedColumn = plan.region.paintedStartCol ?? plan.region.startCol;
  const lastPaintedColumn = plan.region.paintedEndCol ?? plan.region.endCol;
  for (let row = 1; row <= plan.region.rowsConsumed; row += 1) {
    for (let column = firstPaintedColumn; column <= lastPaintedColumn; column += 1) {
      const source = plan.worksheet.getCell(row, column);
      if (source.isMerged && source.master.address !== source.address) continue;
      const sourceMerge = mergeByMaster.get(`${row}:${column}`);
      const sourceRange = sourceMerge || { startRow: row, endRow: row, startCol: column, endCol: column };
      const hasContent = source.value !== null && source.value !== undefined && source.value !== '';
      const hasStyle = Object.keys(source.style || {}).length > 0;
      if (!hasContent && !hasStyle && !sourceMerge) continue;
      const range = {
        startRow: mapStart(sourceRange.startRow),
        endRow: mapEnd(sourceRange.endRow),
        startCol: sourceRange.startCol,
        endCol: sourceRange.endCol,
      };
      const target = worksheet.getCell(range.startRow, range.startCol);
      target.value = cloneExcelValue(source.value);
      target.style = structuredClone(source.style || {});
      if (source.dataValidation && Object.keys(source.dataValidation).length) {
        target.dataValidation = structuredClone(source.dataValidation);
      }
      if (source.note) target.note = structuredClone(source.note);
      // ExcelJS snapshots the master's style into merged followers when the merge is created. Styling the
      // master after merging leaves those follower cells without perimeter border records, which produces
      // gaps across wide header merges and missing bottoms on tall group merges. Copy the resolved style
      // first, then create the target merge so every edge is materialized in the worksheet XML.
      if (range.startRow !== range.endRow || range.startCol !== range.endCol) {
        mergeSafe(worksheet, range, merges, plan.item.name);
      }
    }
  }

  const mappedStart = mapStart(1);
  const mappedEnd = mapEnd(plan.region.rowsConsumed);
  const mappedHeaderEnd = plan.region.headerRows > 0 ? mapEnd(plan.region.headerRows) : mappedStart - 1;
  return {
    ...plan.region,
    startRow: mappedStart,
    endRow: mappedEnd,
    rowsConsumed: mappedEnd - mappedStart + 1,
    headerRows: Math.max(0, mappedHeaderEnd - mappedStart + 1),
    // Cell pictures carry point offsets from the tablix top; the caller resolves them against the section
    // row grid, so they float over the table instead of occupying its rows.
    canvasPlacements: plan.region.canvasPlacements || [],
  };
}

async function renderCoordinateScheduledSection({
  workbook,
  worksheet,
  model,
  section,
  request,
  globals,
  context,
  config,
  tempDir,
  chartCounter,
  xGrid,
  startRow,
  merges,
  tablixCache,
  measureDoc,
  preserveLeadingBodyGap,
}) {
  // Report body coordinates remain absolute even after an explicit page break starts a new logical Excel
  // section. Each worksheet has its own local origin, matching the legacy section renderer and SSRS's
  // page-local placement semantics; carrying the global body Top into the new sheet creates a giant blank
  // row containing the height of every earlier section.
  const sectionOriginTop = preserveLeadingBodyGap || !section.length
    ? 0
    : Math.min(...section.map((item) => point(item.top || 0)));
  const reportScope = { model, request, globals, context };
  const plans = [];
  for (const item of section) {
    const scope = scopeOf(item, reportScope);
    const designTop = point((item.top || 0) - sectionOriginTop);
    if (item.type !== 'Tablix') {
      const resolved = resolveExcelFreeformLayout(item, scope.context, config, measureDoc);
      plans.push({
        item: resolved.item,
        designTop,
        designHeight: item.height || 0,
        occupiedHeight: resolved.occupiedHeight,
        resolvedTop: designTop,
      });
      continue;
    }
    const planningWorkbook = new ExcelJS.Workbook();
    const planningSheet = planningWorkbook.addWorksheet('Tablix');
    const region = renderReportTablix({
      worksheet: planningSheet,
      model: scope.model,
      item,
      request: scope.request,
      globals: scope.globals,
      config,
      xGrid,
      startRow: 1,
      merges: createMergeIndex(),
      tablixCache,
      measureDoc,
    });
    const localBoundaries = worksheetRowBoundaries(planningSheet, region.rowsConsumed);
    plans.push({
      item,
      designTop,
      designHeight: item.height || 0,
      worksheet: planningSheet,
      region,
      localBoundaries,
      occupiedHeight: Math.max(item.height || 0, localBoundaries.at(-1)),
      resolvedTop: designTop,
    });
  }

  // RDL peers keep their declared coordinates. Growth only pushes an item down when a design-time peer
  // ended above it in the same horizontal lane, preserving the original minimum vertical gap. Items in
  // disjoint left/right lanes therefore remain side by side instead of being serialized by item type.
  const layoutOrder = plans
    .map((plan, sourceIndex) => ({ plan, sourceIndex }))
    .sort((left, right) => left.plan.designTop - right.plan.designTop
      || (left.plan.item.left || 0) - (right.plan.item.left || 0)
      || left.sourceIndex - right.sourceIndex)
    .map(({ plan }) => plan);
  for (const [index, plan] of layoutOrder.entries()) {
    const designTop = plan.designTop;
    for (let previousIndex = 0; previousIndex < index; previousIndex += 1) {
      const previous = layoutOrder[previousIndex];
      const previousDesignBottom = point(previous.designTop + previous.designHeight);
      if (previousDesignBottom > designTop + COINCIDENT_EDGE_TOLERANCE_PT) continue;
      if (!horizontalDesignOverlap(previous.item, plan.item)) continue;
      const originalGap = Math.max(0, designTop - previousDesignBottom);
      plan.resolvedTop = Math.max(plan.resolvedTop, point(previous.resolvedTop + previous.occupiedHeight + originalGap));
    }
  }

  // A container expanded into this list left behind a childless shell carrying its fill and border. The
  // shell was declared around content that has since been measured and may have grown, so stretch it back
  // over everything that used to be inside it. Containment is decided on the declared boxes the RDL gave
  // them, so this can never capture an item the container did not hold.
  for (const shell of plans.filter((plan) => plan.item.containerPaintShell)) {
    const declaredLeft = point(shell.item.left || 0);
    const declaredRight = point(declaredLeft + (shell.item.width || 0));
    for (const member of plans) {
      if (member === shell) continue;
      const memberLeft = point(member.item.left || 0);
      if (memberLeft < declaredLeft - COINCIDENT_EDGE_TOLERANCE_PT) continue;
      if (point(memberLeft + (member.item.width || 0)) > declaredRight + COINCIDENT_EDGE_TOLERANCE_PT) continue;
      if (member.designTop < shell.designTop - COINCIDENT_EDGE_TOLERANCE_PT) continue;
      if (member.designTop > point(shell.designTop + shell.designHeight) + COINCIDENT_EDGE_TOLERANCE_PT) continue;
      shell.occupiedHeight = Math.max(
        shell.occupiedHeight,
        point(member.resolvedTop + member.occupiedHeight - shell.resolvedTop),
      );
    }
  }

  const boundaries = new Set([0]);
  const spans = new Set();
  for (const plan of plans) {
    boundaries.add(plan.resolvedTop);
    boundaries.add(point(plan.resolvedTop + plan.occupiedHeight));
    spans.add(`${plan.resolvedTop}|${point(plan.resolvedTop + plan.occupiedHeight)}`);
    if (plan.item.type === 'Tablix') {
      let previous = plan.resolvedTop;
      for (const boundary of plan.localBoundaries) {
        const value = point(plan.resolvedTop + boundary);
        boundaries.add(value);
        if (value > previous) spans.add(`${previous}|${value}`);
        previous = value;
      }
      // A picture on a tablix cell's canvas floats at its own top and height, so those two coordinates
      // have to exist on the section grid for the anchor to resolve. They add boundaries only.
      for (const placed of plan.region?.canvasPlacements || []) {
        const from = canvasPlacementTop(plan, placed);
        const to = point(from + (placed.item.height || 0));
        boundaries.add(from);
        boundaries.add(to);
        if (to > from) spans.add(`${from}|${to}`);
      }
    } else {
      collectYBoundaries([{ ...plan.item, top: plan.resolvedTop }], boundaries, spans);
    }
  }
  const maximum = Math.max(...boundaries);
  boundaries.add(maximum);
  const collapsedY = collapsedBoundaries(
    [...boundaries].filter((value) => value >= 0 && value <= maximum).sort((a, b) => a - b),
    [...spans].map((key) => key.split('|').map(Number)),
  );
  const yGrid = splitTallRowIntervals(collapsedY);
  yGrid.aliases = collapsedY.aliases;
  allocateHeightRows(worksheet, yGrid, startRow);

  const detailRegions = [];
  for (const plan of plans) {
    // An expanded container's paint shell is applied last: its former children write their own resolved
    // edges into the same cells, and a container edge must add to those, never replace them.
    if (plan.item.containerPaintShell) continue;
    if (plan.item.type === 'Tablix') {
      const region = copyPlannedTablix({ worksheet, plan, yGrid, startRow, merges });
      // A free-form cell's children are placed at their own coordinates, through the same path a
      // section-level free-form item takes, so a canvas textbox, chart, image or line behaves in the
      // worksheet exactly as it would sitting directly on the body.
      for (const placed of region.canvasPlacements || []) {
        // The point model and the worksheet row grid can disagree by one interval where a grown child
        // region ends and the next canvas item begins. SSRS resolves exactly that case by displacement —
        // a later item moves below what grew — so apply the same rule against the rows already taken.
        const placedTop = canvasPlacementTop(plan, placed);
        const placedColumns = gridRange(xGrid, placed.left, placed.item.width || 0);
        let placedRows = rowRange(yGrid, startRow, placedTop, placed.item.height || 0);
        // A picture is anchored *over* the worksheet grid: it floats above the cells its range covers
        // and occupies none of them, exactly as the canonical PDF paints a chart or image on top of the
        // region beneath it. The merged cells of the tablix row the picture sits on are therefore not a
        // collision, and displacing the anchor past them pushes the picture off its own region — or off
        // the end of the sheet once several rows of that region are merged. Only an item that writes
        // into cells has to move below content that grew.
        const floats = placed.item.type === 'Chart' || placed.item.type === 'Image';
        for (let guard = 0; !floats && guard < 64; guard += 1) {
          const clash = findIndexedMergeOverlap(merges, { ...placedColumns, ...placedRows });
          if (!clash) break;
          const height = placedRows.endRow - placedRows.startRow;
          placedRows = { startRow: clash.endRow + 1, endRow: clash.endRow + 1 + height };
        }
        await renderFreeformItem({
          workbook,
          worksheet,
          model: placed.model || scopeOf(plan.item, reportScope).model,
          item: { ...placed.item, left: placed.left, top: placedTop },
          context: placed.context,
          config,
          tempDir,
          chartCounter,
          xGrid,
          yGrid,
          startRow,
          merges,
          measureDoc,
          parentLeft: 0,
          parentTop: 0,
          rowsOverride: placedRows,
          // Tablix canvas placements are stored relative to the tablix. Convert their parent-cell frame
          // into the section's resolved coordinates alongside the floating child item.
          containerFrames: (placed.containerFrames || []).map((frame) => ({
            ...frame,
            top: point(plan.resolvedTop + frame.top),
            bottom: point(plan.resolvedTop + frame.bottom),
          })),
        });
      }
      if (region.dynamic) detailRegions.push(region);
      continue;
    }
    await renderFreeformItem({
      workbook,
      worksheet,
      model: scopeOf(plan.item, reportScope).model,
      item: {
        ...plan.item,
        top: plan.resolvedTop,
        height: plan.occupiedHeight,
      },
      context: scopeOf(plan.item, reportScope).context,
      config,
      tempDir,
      chartCounter,
      xGrid,
      yGrid,
      startRow,
      merges,
      measureDoc,
    });
  }
  for (const shell of plans.filter((plan) => plan.item.containerPaintShell)) {
    paintContainerExtent(
      worksheet,
      {
        ...gridRange(xGrid, point(shell.item.left || 0), shell.item.width || 0),
        ...rowRange(yGrid, startRow, shell.resolvedTop, shell.occupiedHeight),
      },
      shell.item,
      context,
    );
  }
  return { rowsConsumed: Math.max(1, yGrid.length - 1), detailRegions };
}

function addGapRows(worksheet, startRow, points) {
  let remaining = Math.max(0, points);
  let rows = 0;
  while (remaining > 0.5) {
    const height = Math.min(MAX_EXCEL_ROW_HEIGHT, remaining);
    worksheet.getRow(startRow + rows).height = Math.max(2, height);
    remaining -= height;
    rows += 1;
  }
  return rows;
}

function footerText(model, request, globals, reportWidth) {
  const slots = { left: [], center: [], right: [] };
  const datasets = normalizeDatasets(model, request);
  const footerGlobals = { ...globals, PageNumber: '&P', TotalPages: '&N' };
  const context = { parameters: request.parameters || {}, globals: footerGlobals, fields: {}, dataset: [], datasets };
  const visit = (item, parentLeft = 0) => {
    if (item.type === 'Textbox' && !isHidden(item.hidden, context)) {
      const center = parentLeft + (item.left || 0) + (item.width || 0) / 2;
      const slot = center < reportWidth / 3 ? 'left' : center > reportWidth * 2 / 3 ? 'right' : 'center';
      const value = String(textForItem(item, context) || '').trim();
      if (value) slots[slot].push(value);
    }
    for (const child of item.items || []) visit(child, parentLeft + (item.left || 0));
  };
  for (const item of model.page.footer?.items || []) visit(item);
  return `&L${slots.left.join(' | ')}&C${slots.center.join(' | ')}&R${slots.right.join(' | ')}`;
}

function configureReportSheet(worksheet, model, request, globals, usedRows, usedColumns, detailRegions, headerBandRows, reportWidth) {
  const candidate = detailRegions.length === 1 ? detailRegions[0] : null;
  const ySplit = candidate?.headerRows > 0 ? candidate.startRow + candidate.headerRows - 1 : 0;
  const xSplit = candidate?.fixedColumnsSplit || 0;
  if (xSplit > 0 || ySplit > 0) {
    const topLeftCell = worksheet.getCell(ySplit + 1, xSplit + 1).address;
    worksheet.views = [{ state: 'frozen', xSplit: xSplit || undefined, ySplit: ySplit || undefined, topLeftCell, showGridLines: false }];
    if (ySplit > 0) worksheet.pageSetup.printTitlesRow = `${candidate.startRow}:${ySplit}`;
  } else {
    worksheet.views = [{ state: 'normal', showGridLines: false }];
    if (headerBandRows > 0) worksheet.pageSetup.printTitlesRow = `1:${headerBandRows}`;
  }
  if (candidate?.headerRows > 0) {
    worksheet.autoFilter = {
      from: { row: ySplit, column: candidate.startCol },
      to: { row: candidate.endRow, column: candidate.endCol },
    };
  }
  worksheet.pageSetup.orientation = model.page.width > model.page.height ? 'landscape' : 'portrait';
  worksheet.pageSetup.paperSize = 9; // ISO A4
  worksheet.pageSetup.fitToPage = true;
  worksheet.pageSetup.fitToWidth = 1;
  worksheet.pageSetup.fitToHeight = 0;
  worksheet.pageSetup.margins = {
    left: model.page.marginLeft / 72,
    right: model.page.marginRight / 72,
    top: model.page.marginTop / 72,
    bottom: model.page.marginBottom / 72,
    header: 0.1,
    footer: 0.1,
  };
  if (usedRows > 0 && usedColumns > 0) worksheet.pageSetup.printArea = `A1:${worksheet.getColumn(usedColumns).letter}${usedRows}`;
  const footer = footerText(model, request, globals, reportWidth);
  worksheet.headerFooter = { oddFooter: footer, evenFooter: footer };
}

function declaredSectionName(items, context) {
  for (const item of items) {
    if (item.pageName !== null && item.pageName !== undefined && String(item.pageName).trim() !== '') {
      const value = evaluateExpression(item.pageName, scopeOf(item, { context }).context);
      const name = String(value ?? '').replace(/\s+/g, ' ').trim();
      if (name) return name;
    }
  }
  return '';
}

function containsTablix(item) {
  if (item.type === 'Tablix') return true;
  return (item.items || []).some((child) => containsTablix(child));
}

function containsChart(item) {
  if (item.type === 'Chart') return true;
  // A bundled subreport contributes its child body's charts to this render, whether it is invoked from the
  // body canvas or from a tablix cell. The definition is already resolved on the item, so deciding whether
  // a raster workspace is needed does not require the invocation's data.
  const childBody = item.type === 'Subreport' ? (item.resolvedSubreport?.model?.body?.items || []) : [];
  const cellItems = (item.rows || []).flatMap((row) => (row.cells || []).flatMap((cell) => cell.items || []));
  return [...(item.items || []), ...cellItems, ...childBody].some((child) => containsChart(child));
}

async function renderReportExcel(model, request, config, tempDir) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'RDL Converter Service';
  workbook.title = request.outputFileName || model.name || 'Report';
  workbook.calcProperties.fullCalcOnLoad = true;
  const globals = { PageNumber: 1, TotalPages: 1, ReportName: request.outputFileName || model.name, ExecutionTime: new Date(), variables: model.variables || {}, culture: resolveReportCulture(model, { parameters: request.parameters || {} }) };
  const datasets = normalizeDatasets(model, request);
  const context = { parameters: request.parameters || {}, globals, fields: {}, dataset: [], datasets };
  const sections = partitionReportSections(model, context);
  const tablixCache = new Map();
  const measureDoc = new PDFDocument({ autoFirstPage: false });
  // PDFKit's colour/font state is page-backed even when no bytes are being collected. A private measuring
  // page gives XLSX the exact PDF text metrics without producing or embedding a PDF artifact.
  measureDoc.addPage({ size: [1000, 1000], margins: { top: 0, right: 0, bottom: 0, left: 0 } });
  const chartCounter = { value: 0 };
  let rowCount = 0;

  for (let index = 0; index < sections.length; index += 1) {
    const section = sections[index];
    // REPORT worksheets are partitioned only at explicit RDL page breaks. They are continuous Excel
    // sheets rather than PDF pages, but each section still has a stable logical first/middle/last position.
    // Use that position when resolving page-dependent header/body expressions; evaluating every sheet as
    // Page 1 of 1 hides constructs intended for interior sections while still reserving their band height.
    const sectionGlobals = { ...globals, PageNumber: index + 1, TotalPages: sections.length };
    const sectionContext = { ...context, globals: sectionGlobals };
    const firstSectionItem = section[0];
    // REPORT worksheets mirror body coordinate semantics without pretending that each worksheet is a PDF
    // page. Only the first report section preserves the body's leading Top. Sections created by an explicit
    // Start break keep their established local origin, matching PDF page-start and continuation behavior.
    const preserveLeadingBodyGap = index === 0
      && firstSectionItem
      && !/^(Start|StartAndEnd)$/i.test(visiblePageBreak(firstSectionItem, sectionContext));
    const title = declaredSectionName(section, sectionContext)
      || firstVisibleText(section, model, request, sectionGlobals);
    const name = uniqueSheetName(workbook, title, `Section ${index + 1}`);
    const worksheet = workbook.addWorksheet(name);
    const xGrid = reportGrid(model, section, request, sectionGlobals, tablixCache);
    for (let column = 0; column < xGrid.length - 1; column += 1) {
      worksheet.getColumn(column + 1).width = excelWidthFromPoints(xGrid[column + 1] - xGrid[column]);
    }
    const merges = createMergeIndex();
    let cursor = 1;
    let headerBandRows = 0;
    if (model.page.header?.items?.length) {
      headerBandRows = await renderFreeformBand({
        workbook, worksheet, model, items: model.page.header.items, height: model.page.header.height,
        context: sectionContext, config, tempDir, chartCounter, xGrid, startRow: cursor, merges, measureDoc,
      });
      cursor += headerBandRows;
    }
    const detailRegions = [];
    const coordinateSchedulable = section.every((item) => (
      item.type !== 'Rectangle' || !containsTablix(item)
    ));
    if (coordinateSchedulable) {
      const scheduled = await renderCoordinateScheduledSection({
        workbook,
        worksheet,
        model,
        section,
        request,
        globals: sectionGlobals,
        context: sectionContext,
        config,
        tempDir,
        chartCounter,
        xGrid,
        startRow: cursor,
        merges,
        tablixCache,
        measureDoc,
        preserveLeadingBodyGap,
      });
      cursor += scheduled.rowsConsumed;
      detailRegions.push(...scheduled.detailRegions);
    }
    const reportScope = { model, request, globals: sectionGlobals, context: sectionContext };
    const renderSectionItem = async (sourceItem, parentLeft = 0) => {
      const item = { ...sourceItem, left: point(parentLeft + (sourceItem.left || 0)) };
      const scope = scopeOf(item, reportScope);
      if (item.type === 'Rectangle') {
        // A fixed rectangle is a true coordinate container. Rendering all of its children in one band
        // preserves side-by-side charts/text/images. Rectangles containing tablixes retain the flow-aware
        // path below because their materialized height can grow beyond the design-time rectangle.
        if (!containsTablix(item)) {
          const consumed = await renderFreeformBand({
            workbook,
            worksheet,
            model: scope.model,
            items: [{ ...item, top: 0 }],
            height: Math.max(2, item.height || DEFAULT_ROW_POINTS),
            context: scope.context,
            config,
            tempDir,
            chartCounter,
            xGrid,
            startRow: cursor,
            merges,
            measureDoc,
          });
          cursor += consumed;
          return;
        }
        const containerStart = cursor;
        let previousChildBottom = 0;
        for (const child of [...(item.items || [])].sort((a, b) => (
          (a.top || 0) - (b.top || 0) || (a.left || 0) - (b.left || 0) || a.zIndex - b.zIndex
        ))) {
          const gap = Math.max(0, (child.top || 0) - previousChildBottom);
          cursor += addGapRows(worksheet, cursor, gap);
          await renderSectionItem({ ...child, top: 0 }, item.left);
          previousChildBottom = Math.max(previousChildBottom, (child.top || 0) + (child.height || 0));
        }
        const trailing = Math.max(0, (item.height || 0) - previousChildBottom);
        cursor += addGapRows(worksheet, cursor, trailing);
        if (cursor === containerStart) cursor += addGapRows(worksheet, cursor, Math.max(2, item.height || DEFAULT_ROW_POINTS));
        const range = {
          ...gridRange(xGrid, item.left, item.width || 0),
          startRow: containerStart,
          endRow: Math.max(containerStart, cursor - 1),
        };
        paintContainerExtent(worksheet, range, item, scope.context);
        return;
      }
      if (item.type === 'Tablix') {
        const region = renderReportTablix({
          worksheet,
          model: scope.model,
          item,
          request: scope.request,
          globals: scope.globals,
          config,
          xGrid,
          startRow: cursor,
          merges,
          tablixCache,
          measureDoc,
        });
        cursor += region.rowsConsumed;
        if (region.dynamic) detailRegions.push(region);
        return;
      }
      const height = Math.max(2, item.height || DEFAULT_ROW_POINTS);
      const consumed = await renderFreeformBand({
        workbook,
        worksheet,
        model: scope.model,
        items: [{ ...item, top: 0 }],
        height,
        context: scope.context,
        config,
        tempDir,
        chartCounter,
        xGrid,
        startRow: cursor,
        merges,
        measureDoc,
      });
      cursor += consumed;
    };
    const isFreeformCoordinateItem = (candidate) => {
      if (candidate.type === 'Tablix') return false;
      if (candidate.type === 'Rectangle') {
        return (candidate.items || []).every(isFreeformCoordinateItem);
      }
      return true;
    };
    const sameDesignTop = (left, right) => (
      Math.abs((left.top || 0) - (right.top || 0)) <= COINCIDENT_EDGE_TOLERANCE_PT
    );
    const horizontallyDisjoint = (left, right) => (
      (left.left || 0) + (left.width || 0) <= (right.left || 0) + COINCIDENT_EDGE_TOLERANCE_PT
      || (right.left || 0) + (right.width || 0) <= (left.left || 0) + COINCIDENT_EDGE_TOLERANCE_PT
    );
    let previousDesignBottom = null;
    for (let itemIndex = 0; !coordinateSchedulable && itemIndex < section.length; itemIndex += 1) {
      const item = section[itemIndex];
      // Freeform report items are coordinate-positioned. Coincident-top items and horizontally disjoint
      // items whose design-time vertical intervals overlap are peers in one band, matching the canonical
      // PDF layout rule. Without this, Excel's sequential rows insert artificial gaps between side-by-side
      // headings, charts, images, or fixed rectangles merely because they are separate RDL items.
      const peerBand = [item];
      let peerBottom = (item.top || 0) + (item.height || 0);
      if (isFreeformCoordinateItem(item)) {
        for (let peerIndex = itemIndex + 1; peerIndex < section.length; peerIndex += 1) {
          const peer = section[peerIndex];
          if (!isFreeformCoordinateItem(peer)) break;
          const coincidentTop = sameDesignTop(item, peer);
          const overlapsVertically = (peer.top || 0) < peerBottom - COINCIDENT_EDGE_TOLERANCE_PT;
          const independentLane = peerBand.every((candidate) => horizontallyDisjoint(candidate, peer));
          if (!coincidentTop && !(overlapsVertically && independentLane)) break;
          peerBand.push(peer);
          peerBottom = Math.max(peerBottom, (peer.top || 0) + (peer.height || 0));
        }
      }
      if (peerBand.length > 1) {
        const bandTop = item.top || 0;
        const gap = previousDesignBottom === null
          ? (preserveLeadingBodyGap ? Math.max(0, bandTop) : 0)
          : Math.max(0, bandTop - previousDesignBottom);
        cursor += addGapRows(worksheet, cursor, gap);
        const consumed = await renderFreeformBand({
          workbook,
          worksheet,
          model,
          items: peerBand.map((peer) => ({ ...peer, top: (peer.top || 0) - bandTop })),
          height: Math.max(2, peerBottom - bandTop),
          context: sectionContext,
          config,
          tempDir,
          chartCounter,
          xGrid,
          startRow: cursor,
          merges,
          measureDoc,
        });
        cursor += consumed;
        previousDesignBottom = Math.max(previousDesignBottom ?? 0, peerBottom);
        itemIndex += peerBand.length - 1;
        continue;
      }
      const gap = previousDesignBottom === null
        ? (preserveLeadingBodyGap ? Math.max(0, item.top || 0) : 0)
        : Math.max(0, (item.top || 0) - previousDesignBottom);
      cursor += addGapRows(worksheet, cursor, gap);
      await renderSectionItem({ ...item, top: 0 });
      previousDesignBottom = Math.max(previousDesignBottom ?? 0, (item.top || 0) + (item.height || 0));
    }
    const usedRows = Math.max(1, cursor - 1);
    rowCount += usedRows;
    configureReportSheet(
      worksheet,
      model,
      request,
      sectionGlobals,
      usedRows,
      xGrid.length - 1,
      detailRegions,
      headerBandRows,
      xGrid[xGrid.length - 1],
    );
  }
  if (!workbook.worksheets.length) workbook.addWorksheet('Section 1', { views: [{ state: 'normal', showGridLines: false }] });
  const buffer = Buffer.from(await workbook.xlsx.writeBuffer());
  measureDoc.end();
  return {
    buffer,
    pageCount: null,
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    extension: 'xlsx',
    layoutMode: 'report-sections',
    sheetCount: workbook.worksheets.length,
    rowCount,
  };
}

export async function renderExcel(model, request, config, tempDir) {
  const mode = resolveExcelLayoutMode(request);
  if (mode === 'REPORT') {
    let ownedTempDir = null;
    let workingTempDir = tempDir;
    const requiresChartWorkspace = [
      ...(model.page.header?.items || []),
      ...(model.body.items || []),
    ].some((item) => containsChart(item));
    if (!workingTempDir && requiresChartWorkspace) {
      await fs.mkdir(config.tempRoot, { recursive: true, mode: 0o700 });
      await fs.chmod(config.tempRoot, 0o700);
      ownedTempDir = await fs.mkdtemp(path.join(config.tempRoot, 'excel-chart-'));
      await fs.chmod(ownedTempDir, 0o700);
      workingTempDir = ownedTempDir;
    }
    try {
      return await renderReportExcel(model, request, config, workingTempDir);
    } finally {
      if (ownedTempDir) await fs.rm(ownedTempDir, { recursive: true, force: true });
    }
  }
  const rendered = await renderDataExcel(model, request, config, tempDir);
  return {
    ...rendered,
    layoutMode: request.excel?.sheetPerTablix === true ? 'data-per-tablix' : 'data-stacked',
  };
}
