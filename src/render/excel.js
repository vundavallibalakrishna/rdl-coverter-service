import ExcelJS from 'exceljs';
import PDFDocument from 'pdfkit';
import { ServiceError } from '../errors.js';
import { evaluateExpression } from '../rdl/expression.js';
import { cellText, cellTextbox, color, enforcedBottomBorder, isHidden, normalizeDatasets, shouldEnforceTablixBottom, styledTextForItem, styleColor, styleSize, styleValue, tablixRows, textForItem } from './common.js';
import { computeCellPlacements } from './tableGrid.js';
import { resolveGridColumns } from './tableLayout.js';
import { materializeChart } from './chartData.js';
import { renderChartPng } from './chartImage.js';
import { measureTextboxHeight } from './pdf.js';
import { DEFAULT_EXCEL_DATE_FORMAT, excelNumberFormat, cellString } from './excelFormat.js';

const SHEET_NAME_FORBIDDEN = /[\\/?*[\]:]/g;
const DEFAULT_ROW_POINTS = 15;

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
      const style = (cell.containerWrapped ? item.style : textbox?.style) || item.style;
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
      const borders = style.borders || {};
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

export function resolveExcelLayoutMode(request = {}) {
  const requested = request.excel?.layoutMode;
  if (requested !== undefined && requested !== null && String(requested).trim() !== '') {
    const mode = String(requested).trim().toUpperCase();
    if (!['REPORT', 'DATA'].includes(mode)) throw new ServiceError('RDL_INVALID', `Unsupported Excel layoutMode: ${requested}`);
    if (mode === 'REPORT' && request.excel?.sheetPerTablix === true) {
      throw new ServiceError('RDL_INVALID', 'excel.sheetPerTablix is only valid with excel.layoutMode DATA');
    }
    return mode;
  }
  // Backward compatibility for callers that already explicitly selected the old per-tablix workbook.
  if (request.excel?.sheetPerTablix === true) return 'DATA';
  return 'REPORT';
}

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
  const items = [...(model.body.items || [])]
    .filter((item) => !isHidden(item.hidden, context))
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
  const pixels = Math.max(1, points * (96 / 72));
  return Math.max(0.2, pixels <= 12 ? pixels / 12 : (pixels - 5) / 7);
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

function applyRegionStyle(worksheet, range, style, context) {
  const borders = style?.borders || {};
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startCol; column <= range.endCol; column += 1) {
      const cell = worksheet.getCell(row, column);
      applyFillFontAlignment(cell, style, context);
      cell.border = {
        top: row === range.startRow ? excelBorderSide(borders.top, context) : undefined,
        bottom: row === range.endRow ? excelBorderSide(borders.bottom, context) : undefined,
        left: column === range.startCol ? excelBorderSide(borders.left, context) : undefined,
        right: column === range.endCol ? excelBorderSide(borders.right, context) : undefined,
      };
    }
  }
}

function rangesOverlap(left, right) {
  return left.startRow <= right.endRow && right.startRow <= left.endRow
    && left.startCol <= right.endCol && right.startCol <= left.endCol;
}

function mergeSafe(worksheet, range, merges, owner = null) {
  if (range.startRow === range.endRow && range.startCol === range.endCol) return;
  const existing = merges.find((candidate) => rangesOverlap(candidate, range));
  if (existing) {
    throw new ServiceError(
      'RDL_INVALID',
      `RDL produced overlapping Excel merged-cell ranges${owner ? ` for ${owner}` : ''} (${existing.startRow},${existing.startCol}:${existing.endRow},${existing.endCol} and ${range.startRow},${range.startCol}:${range.endRow},${range.endCol})`,
    );
  }
  worksheet.mergeCells(range.startRow, range.startCol, range.endRow, range.endCol);
  merges.push(range);
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

function renderFreeformItem({ workbook, worksheet, model, item, context, xGrid, yGrid, startRow, merges, parentLeft = 0, parentTop = 0 }) {
  if (isHidden(item.hidden, context)) return;
  const left = point(parentLeft + (item.left || 0));
  const top = point(parentTop + (item.top || 0));
  const columns = gridRange(xGrid, left, item.width || 0);
  const rows = rowRange(yGrid, startRow, top, item.height || 0);
  const range = { ...columns, ...rows };
  if (item.type === 'Chart') throw new ServiceError('UNSUPPORTED_FEATURE', 'Charts are not supported in Excel REPORT mode without drawings; use excel.layoutMode DATA');
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
    return;
  }
  if (item.type === 'Rectangle') {
    for (const child of [...(item.items || [])].sort((a, b) => a.zIndex - b.zIndex || a.top - b.top || a.left - b.left)) {
      renderFreeformItem({ workbook, worksheet, model, item: child, context, xGrid, yGrid, startRow, merges, parentLeft: left, parentTop: top });
    }
  }
}

function renderFreeformBand({ workbook, worksheet, model, items, height, context, xGrid, startRow, merges }) {
  const yGrid = freeformRows(worksheet, items, height, startRow);
  for (const item of [...items].sort((a, b) => a.zIndex - b.zIndex || a.top - b.top || a.left - b.left)) {
    renderFreeformItem({ workbook, worksheet, model, item, context, xGrid, yGrid, startRow, merges });
  }
  return yGrid.length - 1;
}

function cellStyle(item, cell, context) {
  const textbox = cellTextbox(cell);
  return { textbox, style: (cell.containerWrapped ? item.style : textbox?.style) || item.style, context };
}

function resolvedOwnerBorder(owner, side) {
  const border = owner.style?.borders?.[side];
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
  return {
    top: above === owner ? undefined : resolvedOwnerBorder(owner, 'top') || (above && resolvedOwnerBorder(above, 'bottom')) || undefined,
    bottom,
    left: resolvedOwnerBorder(owner, 'left') || (left && resolvedOwnerBorder(left, 'right')) || undefined,
    right: resolvedOwnerBorder(owner, 'right') || (right && resolvedOwnerBorder(right, 'left')) || undefined,
  };
}

function renderReportTablix({ worksheet, model, item, request, globals, config, xGrid, startRow, merges, tablixCache, measureDoc }) {
  const { rows, columns } = tablixLayout(item, request, globals, model, tablixCache);
  const placements = computeCellPlacements(rows, columns.length);
  const datasets = normalizeDatasets(model, request);
  const enforceBottomClosure = shouldEnforceTablixBottom(rows);
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
    const context = { fields: row.fields || {}, parameters: request.parameters || {}, globals, dataset: datasets[item.datasetName] || [], datasets };
    row.cells.forEach((cell, cellIndex) => {
      const start = placements[rowIndex][cellIndex];
      if (start === undefined || start < 0) return;
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

  // Excel does not auto-fit merged cells. Measure every logical owner once, then grow the final row of its
  // vertical span only when the declared combined row heights cannot contain the text. This preserves the
  // RDL row proportions while keeping merged group labels readable.
  const measuredHeights = rows.map((row, index) => Math.max(
    2,
    row.height || DEFAULT_ROW_POINTS,
    rowProfiles[index].at(-1) || 0,
  ));
  for (const owner of owners) {
    const display = cellText(owner.cell);
    if (!owner.textbox || !display) continue;
    const span = Math.max(1, owner.cell.colSpan || 1);
    const rowSpan = Math.max(1, owner.cell.rowSpan || 1);
    const width = point(columnOffsets[owner.start + span] - columnOffsets[owner.start]);
    const required = measureTextboxHeight(measureDoc, config, owner.textbox, owner.context, display, width)
      + styleSize(owner.style?.paddingTop, owner.context, 2) + styleSize(owner.style?.paddingBottom, owner.context, 2);
    const endRow = Math.min(rows.length - 1, owner.rowIndex + rowSpan - 1);
    const available = measuredHeights.slice(owner.rowIndex, endRow + 1).reduce((sum, height) => sum + height, 0);
    if (required > available) measuredHeights[endRow] += required - available;
  }

  // A logical parent row may contain a child grid with several physical rows. Split only that parent row
  // at the child-grid boundaries; ordinary surrounding cells are vertically merged across the resulting
  // native Excel rows, keeping the workbook editable without flattening the child tablix to text.
  rowProfiles.forEach((profile, rowIndex) => {
    const declared = profile.at(-1) || 0;
    if (measuredHeights[rowIndex] > declared) profile[profile.length - 1] = point(measuredHeights[rowIndex]);
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
    nestedRows.forEach((row) => rowOffsets.push(point(rowOffsets.at(-1) + (row.height || DEFAULT_ROW_POINTS))));
    for (const [rowIndex, row] of nestedRows.entries()) {
      const context = {
        fields: row.fields || {},
        parameters: request.parameters || {},
        globals,
        dataset: row.scopeDataset || datasets[nested.item.datasetName] || [],
        datasets,
      };
      for (const [cellIndex, cell] of row.cells.entries()) {
        const start = nestedPlacements[rowIndex][cellIndex];
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
        target.border = reportCellBorders(nestedOwners, owner, nested.item.style, shouldEnforceTablixBottom(nestedRows));
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
      target.border = reportCellBorders(gridOwners, owner, item.style, enforceBottomClosure);
      if (hasNested) {
        applyRegionStyle(worksheet, range, owner.style || {}, owner.context);
        target.border = reportCellBorders(gridOwners, owner, item.style, enforceBottomClosure);
        for (const nested of owner.cell.nestedTablixes || []) renderNested(nested, left, rowIndex);
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

async function renderReportExcel(model, request, config) {
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
  let rowCount = 0;

  for (let index = 0; index < sections.length; index += 1) {
    const section = sections[index];
    const title = declaredSectionName(section, context) || firstVisibleText(section, model, request, globals);
    const name = uniqueSheetName(workbook, title, `Section ${index + 1}`);
    const worksheet = workbook.addWorksheet(name);
    const xGrid = reportGrid(model, section, request, globals, tablixCache);
    for (let column = 0; column < xGrid.length - 1; column += 1) {
      worksheet.getColumn(column + 1).width = excelWidthFromPoints(xGrid[column + 1] - xGrid[column]);
    }
    const merges = [];
    let cursor = 1;
    let headerBandRows = 0;
    if (model.page.header?.items?.length) {
      headerBandRows = renderFreeformBand({
        workbook, worksheet, model, items: model.page.header.items, height: model.page.header.height,
        context, xGrid, startRow: cursor, merges,
      });
      cursor += headerBandRows;
    }
    const detailRegions = [];
    const renderSectionItem = (sourceItem, parentLeft = 0) => {
      const item = { ...sourceItem, left: point(parentLeft + (sourceItem.left || 0)) };
      if (item.type === 'Rectangle') {
        const containerStart = cursor;
        let previousChildBottom = 0;
        for (const child of [...(item.items || [])].sort((a, b) => (
          (a.top || 0) - (b.top || 0) || (a.left || 0) - (b.left || 0) || a.zIndex - b.zIndex
        ))) {
          const gap = Math.max(0, (child.top || 0) - previousChildBottom);
          cursor += addGapRows(worksheet, cursor, gap);
          renderSectionItem({ ...child, top: 0 }, item.left);
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
        const fill = hex(styleColor(item.style?.backgroundColor, context, null));
        const borders = item.style?.borders || {};
        for (let row = range.startRow; row <= range.endRow; row += 1) {
          for (let column = range.startCol; column <= range.endCol; column += 1) {
            const target = worksheet.getCell(row, column);
            if (fill && !target.fill?.type && (target.value === null || target.value === undefined)) {
              target.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${fill}` } };
            }
            target.border = {
              ...(target.border || {}),
              top: row === range.startRow ? excelBorderSide(borders.top, context) : target.border?.top,
              bottom: row === range.endRow ? excelBorderSide(borders.bottom, context) : target.border?.bottom,
              left: column === range.startCol ? excelBorderSide(borders.left, context) : target.border?.left,
              right: column === range.endCol ? excelBorderSide(borders.right, context) : target.border?.right,
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
          globals,
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
      if (item.type === 'Chart') {
        throw new ServiceError('UNSUPPORTED_FEATURE', 'Charts are not supported in Excel REPORT mode without drawings; use excel.layoutMode DATA');
      }
      const height = Math.max(2, item.height || DEFAULT_ROW_POINTS);
      const consumed = renderFreeformBand({
        workbook,
        worksheet,
        model,
        items: [{ ...item, top: 0 }],
        height,
        context,
        xGrid,
        startRow: cursor,
        merges,
      });
      cursor += consumed;
    };
    let previousDesignBottom = null;
    for (const item of section) {
      const gap = previousDesignBottom === null ? 0 : Math.max(0, (item.top || 0) - previousDesignBottom);
      cursor += addGapRows(worksheet, cursor, gap);
      renderSectionItem({ ...item, top: 0 });
      previousDesignBottom = Math.max(previousDesignBottom ?? 0, (item.top || 0) + (item.height || 0));
    }
    const usedRows = Math.max(1, cursor - 1);
    rowCount += usedRows;
    configureReportSheet(worksheet, model, request, globals, usedRows, xGrid.length - 1, detailRegions, headerBandRows, xGrid[xGrid.length - 1]);
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
  if (mode === 'REPORT') return renderReportExcel(model, request, config);
  const rendered = await renderDataExcel(model, request, config, tempDir);
  return {
    ...rendered,
    layoutMode: request.excel?.sheetPerTablix === true ? 'data-per-tablix' : 'data-stacked',
  };
}
