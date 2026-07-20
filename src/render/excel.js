import ExcelJS from 'exceljs';
import { evaluateExpression } from '../rdl/expression.js';
import { cellText, cellTextbox, color, isHidden, normalizeDatasets, styleColor, styleValue, tablixRows, textForItem } from './common.js';
import { computeCellPlacements } from './tableGrid.js';
import { materializeChart } from './chartData.js';
import { renderChartPng } from './chartImage.js';
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
        size: Number(styleValue(style.fontSize, context, 10)) || 10,
        bold: /bold|[6-9]00/i.test(String(styleValue(style.fontWeight, context, 'Normal'))),
        italic: /italic/i.test(String(styleValue(style.fontStyle, context, 'Normal'))),
        color: { argb: `FF${hex(styleColor(style.color, context, '#000000'), '000000')}` },
      };
      target.alignment = {
        vertical: /bottom/i.test(style.verticalAlign) ? 'bottom' : /middle|center/i.test(style.verticalAlign) ? 'middle' : 'top',
        horizontal: /center/i.test(style.textAlign) ? 'center' : /right/i.test(style.textAlign) ? 'right' : /justify/i.test(style.textAlign) ? 'justify' : 'left',
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
      size: Number(styleValue(item.style?.fontSize, context, 10)) || 10,
      bold: /bold|[6-9]00/i.test(String(styleValue(item.style?.fontWeight, context, 'Normal'))),
      color: { argb: `FF${hex(styleColor(item.style?.color, context, '#000000'), '000000')}` },
    };
    return { rows: 1, chartIndex };
  }
  if (item.type === 'Image') {
    const image = model.embeddedImages?.[item.value];
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
    views: [{ state: 'frozen', ySplit: 0 }],
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

export async function renderExcel(model, request, config, tempDir) {
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
