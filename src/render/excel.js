import ExcelJS from 'exceljs';
import ExcelRange from 'exceljs/lib/doc/range.js';
import PDFDocument from 'pdfkit';
import fs from 'node:fs/promises';
import path from 'node:path';
import { ServiceError } from '../errors.js';
import { resolveExcelLayoutMode } from '../excelLayoutMode.js';
import { evaluateExpression } from '../rdl/expression.js';
import { cellBorderStyle, cellText, cellTextbox, color, enforcedBottomBorder, isHidden, matchingChangedGroupOwnerRowBoundary, materializedCellContext, materializedCellVisualSignature, normalizeDatasets, shouldEnforceTablixBottom, styledTextForItem, styleColor, styleSize, styleValue, tablixRows, textForItem } from './common.js';
import { computeCellPlacements } from './tableGrid.js';
import { resolveGridColumns } from './tableLayout.js';
import { materializeChart } from './chartData.js';
import { renderChartPng } from './chartImage.js';
import { measureTextboxHeight } from './pdf.js';
import { DEFAULT_EXCEL_DATE_FORMAT, excelNumberFormat, cellString } from './excelFormat.js';

const SHEET_NAME_FORBIDDEN = /[\\/?*[\]:]/g;
const DEFAULT_ROW_POINTS = 15;
const COINCIDENT_EDGE_TOLERANCE_PT = 0.25;
const EXCEL_LAYOUT_DPI = 96;
const EXCEL_MAX_DIGIT_WIDTH_PX = 7;
const EXCEL_CELL_PADDING_PX = 5;
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
  if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
    return { value: raw, numFmt: DEFAULT_EXCEL_DATE_FORMAT };
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
        name: String(styleValue(style.fontFamily, context, 'Arial')),
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
      name: String(styleValue(item.style?.fontFamily, context, 'Arial')),
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
  const globals = { PageNumber: 1, TotalPages: 1, ReportName: request.outputFileName || model.name, ExecutionTime: new Date(), variables: model.variables || {} };
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
function collectEquivalentXEdges(items, target, context, parentLeft = 0, parentTop = 0) {
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
    };
  }).sort((left, right) => left.left - right.left || left.top - right.top);

  for (let index = 1; index < positioned.length; index += 1) {
    const previous = positioned[index - 1];
    const current = positioned[index];
    const verticallyOverlaps = previous.top < current.bottom && current.top < previous.bottom;
    if (!verticallyOverlaps) continue;
    const previousWidth = visibleBorderWidth(previous.item.style?.borders?.right, context);
    const currentWidth = visibleBorderWidth(current.item.style?.borders?.left, context);
    const tolerance = Math.max(1 / POINT_PRECISION, previousWidth, currentWidth);
    if (Math.abs(previous.right - current.left) <= tolerance) {
      const canonical = point((previous.right + current.left) / 2);
      target.set(previous.right, canonical);
      target.set(current.left, canonical);
    }
  }

  for (const entry of positioned) {
    collectEquivalentXEdges(entry.item.items, target, context, entry.left, entry.top);
  }
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

// PageBreak is a property of every RDL report item, not only of a direct Body child. A rectangle is a
// coordinate container rather than a page unit, so a break declared on one of its children splits the report
// flow *inside* it. REPORT worksheets are partitioned at breaks, so such a container is expanded into its
// children at absolute body coordinates and the partition falls between them. Containers without a nested
// break keep the existing container path untouched.
//
// A container that paints its own fill or border cannot be expanded without losing that paint. That is the
// same construct the PDF renderer refuses to fragment, so it fails closed with the same code rather than
// producing a silently different worksheet.
function expandBreakBearingContainers(items, context, offsetLeft = 0, offsetTop = 0) {
  const expanded = [];
  for (const item of items) {
    if (isHidden(item.hidden, context)) continue;
    // Copy only when a container above actually moved this item, so a report without nested breaks keeps
    // the identical item objects (and therefore the identical tablix layout cache keys) it had before.
    const shifted = offsetLeft || offsetTop
      ? { ...item, left: (item.left || 0) + offsetLeft, top: (item.top || 0) + offsetTop }
      : item;
    if (item.type !== 'Rectangle' || !declaresNestedPageBreak(item, context)) {
      expanded.push(shifted);
      continue;
    }
    if (paintsOwnExtent(item.style, context)) {
      throw new ServiceError(
        'UNSUPPORTED_FEATURE',
        'A page-spanning rectangle with a visible fill or border cannot be safely fragmented',
        422,
        { item: item.name || null },
      );
    }
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
  const items = expandBreakBearingContainers(model.body.items || [], context)
    .sort((left, right) => left.top - right.top || left.left - right.left || left.zIndex - right.zIndex);
  const sections = [];
  let current = [];
  const finish = () => {
    if (current.length) sections.push(current);
    current = [];
  };
  for (const item of items) {
    const location = visiblePageBreak(item, context);
    if (/^(Start|StartAndEnd)$/i.test(location)) finish();
    current.push(item);
    if (/^(End|StartAndEnd)$/i.test(location)) finish();
  }
  finish();
  return sections.length ? sections : [[]];
}

function firstVisibleText(items, model, request, globals) {
  const datasets = normalizeDatasets(model, request);
  const context = { parameters: request.parameters || {}, globals, fields: {}, dataset: [], datasets };
  const visit = (item) => {
    if (item.type === 'Textbox') {
      try {
        const value = String(textForItem(item, context) || '').replace(/\s+/g, ' ').trim();
        if (value) return value;
      } catch { /* a field-scoped textbox is not a stable section title */ }
    }
    if (item.type === 'Tablix') {
      try {
        const { rows } = tablixRows(item, request, globals, model);
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
      const value = visit(child);
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

function excelWidthFromPoints(points) {
  const pixels = Math.max(1, points * (EXCEL_LAYOUT_DPI / 72));
  return Math.max(
    0.2,
    pixels <= 12
      ? pixels / 12
      : (pixels - EXCEL_CELL_PADDING_PX) / EXCEL_MAX_DIGIT_WIDTH_PX,
  );
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

function collectXBoundaries(items, request, globals, model, tablixCache, target, parentLeft = 0) {
  const collectNested = (nested, cellLeft) => {
    const left = point(cellLeft + (nested.item.left || 0));
    target.add(left);
    let cursor = left;
    for (const width of nested.columns || nested.item.columns || []) {
      cursor = point(cursor + width);
      target.add(cursor);
    }
    const placements = computeCellPlacements(nested.rows || [], (nested.columns || []).length);
    for (const [rowIndex, row] of (nested.rows || []).entries()) {
      for (const [cellIndex, cell] of row.cells.entries()) {
        const columnIndex = placements[rowIndex][cellIndex];
        const nestedCellLeft = left + (nested.columns || []).slice(0, columnIndex).reduce((sum, width) => sum + width, 0);
        for (const child of cell.nestedTablixes || []) collectNested(child, nestedCellLeft);
      }
    }
  };
  for (const item of items || []) {
    const left = point(parentLeft + (item.left || 0));
    const right = point(left + (item.width || 0));
    target.add(left);
    target.add(right);
    if (item.type === 'Tablix') {
      const layout = tablixLayout(item, request, globals, model, tablixCache);
      let cursor = left;
      const offsets = [0];
      for (const width of layout.columns) {
        cursor = point(cursor + width);
        target.add(cursor);
        offsets.push(cursor - left);
      }
      const placements = computeCellPlacements(layout.rows, layout.columns.length);
      for (const [rowIndex, row] of layout.rows.entries()) {
        for (const [cellIndex, cell] of row.cells.entries()) {
          const columnIndex = placements[rowIndex][cellIndex];
          const cellLeft = left + offsets[columnIndex];
          for (const nested of cell.nestedTablixes || []) collectNested(nested, cellLeft);
        }
      }
    }
    collectXBoundaries(item.items, request, globals, model, tablixCache, target, left);
  }
}

function reportGrid(model, section, request, globals, tablixCache) {
  const boundaries = new Set([0]);
  collectXBoundaries(model.page.header?.items || [], request, globals, model, tablixCache, boundaries);
  collectXBoundaries(section, request, globals, model, tablixCache, boundaries);
  const aliases = new Map();
  const context = { parameters: request.parameters || {}, globals, fields: {} };
  collectEquivalentXEdges(model.page.header?.items || [], aliases, context);
  collectEquivalentXEdges(section, aliases, context);
  const maximum = Math.max(...boundaries, point(model.page.width - model.page.marginLeft - model.page.marginRight));
  boundaries.add(maximum);
  const values = [...new Set([...boundaries]
    .map((value) => aliases.get(value) ?? value))]
    .filter((value) => value >= 0 && value <= maximum)
    .sort((a, b) => a - b);
  if (values.length < 2) values.push(point(maximum || 100));
  values.aliases = aliases;
  return values;
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
    name: String(styleValue(style?.fontFamily, context, 'Arial')),
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
      cell.border = merged;
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

function collectYBoundaries(items, target, parentTop = 0) {
  for (const item of items || []) {
    const top = point(parentTop + (item.top || 0));
    const bottom = point(top + (item.height || 0));
    target.add(top);
    target.add(bottom);
    collectYBoundaries(item.items, target, top);
  }
}

function freeformRows(worksheet, items, height, startRow) {
  const boundaries = new Set([0, point(height)]);
  collectYBoundaries(items, boundaries);
  const values = splitTallRowIntervals(
    [...boundaries].filter((value) => value >= 0 && value <= point(height)).sort((a, b) => a - b),
  );
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

function addEmbeddedImage(workbook, worksheet, model, item, range, context) {
  if (item.source !== 'Embedded') return;
  const image = model.embeddedImages?.[styleValue(item.value, context, item.value)];
  if (!image?.data) throw new ServiceError('UNSUPPORTED_FEATURE', `Embedded Excel image is unavailable: ${item.name || 'unnamed'}`);
  const buffer = Buffer.from(image.data.replace(/\s+/g, ''), 'base64');
  const id = workbook.addImage({ buffer, extension: imageExtension(buffer) });
  worksheet.addImage(id, {
    tl: { col: range.startCol - 1, row: range.startRow - 1 },
    br: { col: range.endCol, row: range.endRow },
    editAs: 'oneCell',
  });
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
}) {
  if (isHidden(item.hidden, context)) return;
  const left = point(parentLeft + (item.left || 0));
  const top = point(parentTop + (item.top || 0));
  const columns = gridRange(xGrid, left, item.width || 0);
  const rows = rowRange(yGrid, startRow, top, item.height || 0);
  const range = { ...columns, ...rows };
  if (item.type === 'Chart') {
    const data = materializeChart(item, context.datasets, context.parameters, context.globals);
    const png = await renderChartPng(item, data, config, tempDir, context, chartCounter.value++);
    if (!png?.data) throw new ServiceError('RENDER_FAILED', 'Excel chart picture could not be rendered', 500);
    const id = workbook.addImage({ buffer: png.data, extension: 'png' });
    worksheet.addImage(id, {
      tl: { col: range.startCol - 1, row: range.startRow - 1 },
      br: { col: range.endCol, row: range.endRow },
      editAs: 'oneCell',
    });
    return;
  }
  if (item.type === 'Image') {
    addEmbeddedImage(workbook, worksheet, model, item, range, context);
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
      });
    }
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
  const resolved = items.map((item, sourceIndex) => {
    const layout = resolveExcelFreeformLayout(item, context, config, measureDoc);
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
    });
  }
  return yGrid.length - 1;
}

function cellStyle(item, cell, context) {
  const textbox = cellTextbox(cell);
  return {
    textbox,
    style: textbox?.style || item.style,
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

function reportCellBorders(gridOwners, owner, itemStyle, enforceBottomClosure) {
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
      ? excelBorderSide(enforcedBottomBorder(itemStyle), owner.context)
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
  const columnOffsets = [0];
  columns.forEach((width) => columnOffsets.push(point(columnOffsets[columnOffsets.length - 1] + width)));
  const addNestedYBoundaries = (nested, baseTop, target) => {
    let cursor = point(baseTop + (nested.item.top || 0));
    target.add(cursor);
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
        (nested.item.top || 0) + nestedMeasurement.heights.reduce((sum, height) => sum + height, 0));
    }
    if (required <= 0) continue;
    const endRow = Math.min(rows.length - 1, owner.rowIndex + rowSpan - 1);
    const available = measuredHeights.slice(owner.rowIndex, endRow + 1).reduce((sum, height) => sum + height, 0);
    if (required > available) measuredHeights[endRow] += required - available;
  }

  // Replace declared child-row boundaries with their measured CanGrow boundaries before mapping them to
  // physical Excel rows. Otherwise a nested row may grow in measurement but still be rendered into its old
  // one-line interval.
  const addMeasuredNestedYBoundaries = (nested, baseTop, target) => {
    const measurement = measureNestedTablix(nested);
    let cursor = point(baseTop + (nested.item.top || 0));
    target.add(cursor);
    (nested.rows || []).forEach((row, index) => {
      const rowTop = cursor;
      cursor = point(cursor + measurement.heights[index]);
      target.add(cursor);
      for (const cell of row.cells || []) {
        for (const child of cell.nestedTablixes || []) addMeasuredNestedYBoundaries(child, rowTop, target);
      }
    });
  };
  rows.forEach((row, rowIndex) => {
    const boundaries = new Set([0, point(row.height || DEFAULT_ROW_POINTS)]);
    for (const cell of row.cells || []) {
      for (const nested of cell.nestedTablixes || []) addMeasuredNestedYBoundaries(nested, 0, boundaries);
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
  const physicalStarts = [];
  let physicalCursor = startRow;
  rowProfiles.forEach((profile) => {
    physicalStarts.push(physicalCursor);
    physicalCursor += Math.max(1, profile.length - 1);
  });
  rowProfiles.forEach((profile, rowIndex) => {
    for (let index = 0; index < profile.length - 1; index += 1) {
      const height = profile[index + 1] - profile[index];
      if (height > MAX_EXCEL_ROW_HEIGHT) throw new ServiceError('UNSUPPORTED_FEATURE', 'Rendered Excel cell text exceeds the 409-point row-height limit');
      worksheet.getRow(physicalStarts[rowIndex] + index).height = Math.max(2, height);
    }
  });

  const renderNested = (nested, parentCellLeft, parentRowIndex, baseTop = 0) => {
    const nestedRows = nested.rows || [];
    const nestedColumns = nested.columns || nested.item.columns || [];
    const nestedPlacements = computeCellPlacements(nestedRows, nestedColumns.length);
    const nestedOffsets = [0];
    nestedColumns.forEach((width) => nestedOffsets.push(point(nestedOffsets.at(-1) + width)));
    const nestedOwners = nestedRows.map(() => new Array(nestedColumns.length).fill(null));
    const rowOffsets = [point(baseTop + (nested.item.top || 0))];
    const nestedHeights = measureNestedTablix(nested).heights;
    nestedRows.forEach((row, index) => rowOffsets.push(point(rowOffsets.at(-1) + nestedHeights[index])));
    for (const [rowIndex, row] of nestedRows.entries()) {
      for (const [cellIndex, cell] of row.cells.entries()) {
        const start = nestedPlacements[rowIndex][cellIndex];
        const context = materializedCellContext(cell, row, {
          parameters: request.parameters || {}, globals, dataset: datasets[nested.item.datasetName] || [], datasets,
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
        if (range.startRow !== range.endRow || range.startCol !== range.endCol) mergeSafe(worksheet, range, merges);
        const target = worksheet.getCell(range.startRow, range.startCol);
        const { value, numFmt } = excelCellValue(cell, owner.context);
        const display = cellText(cell);
        target.value = typeof value === 'string' ? (richTextValue(owner.textbox, owner.context, display) || value) : value;
        if (numFmt) target.numFmt = numFmt;
        applyFillFontAlignment(target, owner.style || {}, owner.context);
        target.border = reportCellBorders(nestedOwners, owner, nested.item.style, shouldEnforceTablixBottom(nestedRows, nested.item));
        const childCellLeft = left;
        for (const child of cell.nestedTablixes || []) renderNested(child, childCellLeft, parentRowIndex, top);
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
      // The section-wide coordinate grid also contains boundaries from unrelated report items (for example,
      // a logo edge can fall inside a tablix column). A single logical RDL cell may therefore cover several
      // physical Excel columns even when its ColSpan is 1. Merge the occupied Excel region, not merely RDL
      // spans, or the extra grid slices appear as blank gaps inside the table.
      const hasNested = (owner.cell.nestedTablixes || []).length > 0;
      if (!hasNested && (range.startRow !== range.endRow || range.startCol !== range.endCol)) mergeSafe(worksheet, range, merges);
      const target = worksheet.getCell(excelRow, range.startCol);
      const { value, numFmt } = excelCellValue(owner.cell, owner.context);
      const display = cellText(owner.cell);
      target.value = typeof value === 'string' ? (richTextValue(owner.textbox, owner.context, display) || value) : value;
      if (numFmt) target.numFmt = numFmt;
      applyFillFontAlignment(target, owner.style || {}, owner.context);
      if (hasNested) {
        applyRegionStyle(worksheet, range, owner.style || {}, owner.context, { includeBorders: false });
        for (const nested of owner.cell.nestedTablixes || []) renderNested(nested, left, rowIndex);
        // The child grid writes its own cells inside this region, so the enclosing cell's box goes on last
        // and only on the perimeter: applied first it would be overwritten by the child's first row, and
        // applied to the anchor cell it would draw this cell's bottom rule across its top physical row.
        applyRegionBorder(worksheet, range, reportCellBorders(gridOwners, owner, item.style, enforceBottomClosure));
      } else {
        target.border = reportCellBorders(gridOwners, owner, item.style, enforceBottomClosure);
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
    startRow,
    endRow: physicalCursor - 1,
    headerRows,
    dynamic: rows.some((row) => !row.isHeader && !row.isStatic),
    fixedColumnsSplit,
    ...tableRange,
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

function copyPlannedTablix({ worksheet, plan, yGrid, startRow, merges }) {
  const offset = plan.resolvedTop;
  const local = plan.localBoundaries;
  const mergeByMaster = new Map((plan.worksheet.model.merges || []).map((range) => {
    const decoded = decodeMerge(range);
    return [`${decoded.startRow}:${decoded.startCol}`, decoded];
  }));
  const mapStart = (localRow) => startRow + boundaryIndex(yGrid, point(offset + local[localRow - 1]));
  const mapEnd = (localRow) => startRow + boundaryIndex(yGrid, point(offset + local[localRow])) - 1;

  for (let row = 1; row <= plan.region.rowsConsumed; row += 1) {
    for (let column = plan.region.startCol; column <= plan.region.endCol; column += 1) {
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
  const plans = [];
  for (const item of section) {
    const designTop = point((item.top || 0) - sectionOriginTop);
    if (item.type !== 'Tablix') {
      const resolved = resolveExcelFreeformLayout(item, context, config, measureDoc);
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
      model,
      item,
      request,
      globals,
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

  const boundaries = new Set([0]);
  for (const plan of plans) {
    boundaries.add(plan.resolvedTop);
    boundaries.add(point(plan.resolvedTop + plan.occupiedHeight));
    if (plan.item.type === 'Tablix') {
      for (const boundary of plan.localBoundaries) boundaries.add(point(plan.resolvedTop + boundary));
    } else {
      collectYBoundaries([{ ...plan.item, top: plan.resolvedTop }], boundaries);
    }
  }
  const maximum = Math.max(...boundaries);
  boundaries.add(maximum);
  const yGrid = splitTallRowIntervals([...boundaries].filter((value) => value >= 0 && value <= maximum).sort((a, b) => a - b));
  allocateHeightRows(worksheet, yGrid, startRow);

  const detailRegions = [];
  for (const plan of plans) {
    if (plan.item.type === 'Tablix') {
      const region = copyPlannedTablix({ worksheet, plan, yGrid, startRow, merges });
      if (region.dynamic) detailRegions.push(region);
      continue;
    }
    await renderFreeformItem({
      workbook,
      worksheet,
      model,
      item: {
        ...plan.item,
        top: plan.resolvedTop,
        height: plan.occupiedHeight,
      },
      context,
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
      const value = evaluateExpression(item.pageName, context);
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
  return (item.items || []).some((child) => containsChart(child));
}

async function renderReportExcel(model, request, config, tempDir) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'RDL Converter Service';
  workbook.title = request.outputFileName || model.name || 'Report';
  workbook.calcProperties.fullCalcOnLoad = true;
  const globals = { PageNumber: 1, TotalPages: 1, ReportName: request.outputFileName || model.name, ExecutionTime: new Date(), variables: model.variables || {} };
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
    const renderSectionItem = async (sourceItem, parentLeft = 0) => {
      const item = { ...sourceItem, left: point(parentLeft + (sourceItem.left || 0)) };
      if (item.type === 'Rectangle') {
        // A fixed rectangle is a true coordinate container. Rendering all of its children in one band
        // preserves side-by-side charts/text/images. Rectangles containing tablixes retain the flow-aware
        // path below because their materialized height can grow beyond the design-time rectangle.
        if (!containsTablix(item)) {
          const consumed = await renderFreeformBand({
            workbook,
            worksheet,
            model,
            items: [{ ...item, top: 0 }],
            height: Math.max(2, item.height || DEFAULT_ROW_POINTS),
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
        const fill = hex(styleColor(item.style?.backgroundColor, sectionContext, null));
        const borders = item.style?.borders || {};
        for (let row = range.startRow; row <= range.endRow; row += 1) {
          for (let column = range.startCol; column <= range.endCol; column += 1) {
            const target = worksheet.getCell(row, column);
            if (fill && !target.fill?.type && (target.value === null || target.value === undefined)) {
              target.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${fill}` } };
            }
            // A container's declared border ADDS an edge on its own perimeter; it never removes one. The
            // children already wrote their resolved edges into these cells, and a container edge resolving
            // to None must leave them alone. Assigning the resolved side unconditionally erased them: a
            // Rectangle with Border/Style=None wrapping a tablix wiped the tablix's own outer rule wherever
            // the two perimeters coincide — visibly, the closing bottom rule of the table's last row.
            const containerEdge = (side, onEdge) => (
              (onEdge ? excelBorderSide(borders[side], sectionContext) : undefined) || target.border?.[side]
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
        return;
      }
      if (item.type === 'Tablix') {
        const region = renderReportTablix({
          worksheet,
          model,
          item,
          request,
          globals: sectionGlobals,
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
        model,
        items: [{ ...item, top: 0 }],
        height,
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
