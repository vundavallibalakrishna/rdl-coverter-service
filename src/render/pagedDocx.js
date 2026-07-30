import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash, randomUUID } from 'node:crypto';
import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  HeightRule,
  ImageRun,
  LineRuleType,
  PageOrientation,
  Packer,
  Paragraph,
  SectionType,
  ShadingType,
  Table,
  TableBorders,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  UnderlineType,
  VerticalAlignTable,
  WidthType,
} from 'docx';
import JSZip from 'jszip';
import { loadConfig } from '../config.js';
import { ServiceError } from '../errors.js';
import { pointsToTwips } from '../units.js';
import { normalizeDatasets } from './common.js';
import { materializeChart } from './chartData.js';
import { renderChartPng } from './chartImage.js';
import { editableFontEmbeddingPermission, resolveFontFile } from './fonts.js';
import { renderPdf } from './pdf.js';
import { validateLayoutTrace } from './layoutTrace.js';
import { validateWindowsWordRequest } from './windowsWordCompatibility.js';

const WORD_MAX_PAGE_POINTS = 22 * 72;
const WORD_MAX_TABLE_COLUMNS = 63;
const WORD_MAX_TABLE_ROWS = 32_767;
const GRID_PRECISION_POINTS = 0.25;
// Word requires a terminal section paragraph after the page canvas table. Reserve a deterministic 2pt
// band for that paragraph and table-border rounding; report content normally ends above the declared
// bottom margin/footer. A PDF item that actually occupies this band fails closed below.
const SECTION_ANCHOR_POINTS = 2;
const GEOMETRY_EPSILON = 0.13;
const NONE_BORDER = Object.freeze({ style: BorderStyle.NONE, size: 0, color: 'auto' });
const VARIANTS = Object.freeze([
  { key: 'regular', bold: false, italic: false, element: 'embedRegular' },
  { key: 'bold', bold: true, italic: false, element: 'embedBold' },
  { key: 'italic', bold: false, italic: true, element: 'embedItalic' },
  { key: 'boldItalic', bold: true, italic: true, element: 'embedBoldItalic' },
]);

function unsupported(message, details) {
  throw new ServiceError('UNSUPPORTED_FEATURE', message, 422, details);
}

function snap(value) {
  return Math.round(Number(value || 0) / GRID_PRECISION_POINTS) * GRID_PRECISION_POINTS;
}

function pointsToDrawingPixels(points) {
  // docx accepts fractional CSS-pixel dimensions and converts them to integer DrawingML EMUs. Rounding
  // here first can make an image larger than its exact-height Word row: for example, 42.5pt rounds from
  // 56.667px to 57px (42.75pt), so Word clips it until the row is manually enlarged. Preserve the point
  // measurement through the EMU conversion instead.
  return (Number(points || 0) / 72) * 96;
}

function cleanColor(value, fallback = '000000') {
  const normalized = String(value || fallback).replace(/^#/, '').trim();
  return /^[0-9a-f]{6}$/i.test(normalized) ? normalized.toUpperCase() : fallback;
}

function uniqueBoundaries(values) {
  const sorted = values.map(snap).sort((left, right) => left - right);
  const result = [];
  for (const value of sorted) {
    if (result.length === 0 || Math.abs(result[result.length - 1] - value) > GEOMETRY_EPSILON) result.push(value);
  }
  return result;
}

function boundaryIndex(boundaries, value) {
  const snapped = snap(value);
  const index = boundaries.findIndex((candidate) => Math.abs(candidate - snapped) <= GEOMETRY_EPSILON);
  if (index < 0) unsupported('PDF layout geometry cannot be represented by a stable Word table grid', { value });
  return index;
}

function positiveOverlap(left, right) {
  const width = Math.min(left.x + left.width, right.x + right.width) - Math.max(left.x, right.x);
  const height = Math.min(left.y + left.height, right.y + right.height) - Math.max(left.y, right.y);
  return width > GEOMETRY_EPSILON && height > GEOMETRY_EPSILON;
}

function contains(outer, inner) {
  return inner.x >= outer.x - GEOMETRY_EPSILON
    && inner.y >= outer.y - GEOMETRY_EPSILON
    && inner.x + inner.width <= outer.x + outer.width + GEOMETRY_EPSILON
    && inner.y + inner.height <= outer.y + outer.height + GEOMETRY_EPSILON;
}

function isEmptyCell(item) {
  return item.kind === 'tablixCell' && !String(item.text || '') && (item.lines || []).length === 0;
}

function alignment(value) {
  const normalized = String(value || '').toLowerCase();
  if (normalized === 'center') return AlignmentType.CENTER;
  if (normalized === 'right') return AlignmentType.RIGHT;
  if (normalized === 'justify') return AlignmentType.JUSTIFIED;
  return AlignmentType.LEFT;
}

function verticalAlignment(value) {
  const normalized = String(value || '').toLowerCase();
  if (/middle|center/.test(normalized)) return VerticalAlignTable.CENTER;
  if (/bottom/.test(normalized)) return VerticalAlignTable.BOTTOM;
  return VerticalAlignTable.TOP;
}

function borderStyle(style) {
  const normalized = String(style || 'Solid').replace(/[\s_-]/g, '').toLowerCase();
  if (normalized === 'double') return BorderStyle.DOUBLE;
  if (normalized === 'dotted') return BorderStyle.DOTTED;
  if (normalized === 'dashed') return BorderStyle.DASHED;
  if (normalized === 'dashdot') return BorderStyle.DOT_DASH;
  if (normalized === 'dashdotdot') return BorderStyle.DOT_DOT_DASH;
  return BorderStyle.SINGLE;
}

function wordBorder(border) {
  if (!border) return NONE_BORDER;
  return {
    style: borderStyle(border.style),
    size: Math.max(1, Math.round(Number(border.width || 1) * 8)),
    color: cleanColor(border.color),
  };
}

function strongerBorder(left, right) {
  if (!left) return right || null;
  if (!right) return left;
  const leftWidth = Number(left.width || 0);
  const rightWidth = Number(right.width || 0);
  return rightWidth >= leftWidth ? right : left;
}

function linesForParagraphs(item) {
  const source = item.lines || [];
  if (source.length === 0) return [new Paragraph({
    spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
    children: [new TextRun({ text: '' })],
  })];

  const groups = [];
  let current = [];
  for (const line of source) {
    current.push(line);
    if (line.paragraphEnd) {
      groups.push(current);
      current = [];
    }
  }
  if (current.length > 0) groups.push(current);

  return groups.map((group) => {
    const first = group[0];
    const last = group[group.length - 1];
    const runs = [];
    group.forEach((line, lineIndex) => {
      const lineRuns = line.runs?.length ? line.runs : [{ text: '', font: {} }];
      lineRuns.forEach((run, runIndex) => {
        const font = run.font || {};
        runs.push(new TextRun({
          text: String(run.text ?? ''),
          break: lineIndex > 0 && runIndex === 0 ? 1 : undefined,
          font: font.family || 'Arial',
          size: Math.max(1, Math.round(Number(font.size || 10) * 2)),
          bold: Boolean(font.bold),
          italics: Boolean(font.italic),
          underline: font.underline ? { type: UnderlineType.SINGLE } : undefined,
          strike: Boolean(font.strike),
          color: cleanColor(font.color),
          characterSpacing: 0,
        }));
      });
    });
    // Trace line.height includes SpaceBefore/SpaceAfter because PDF advances by the complete physical
    // line box. Word models those margins independently on w:spacing, so feeding the total back as the
    // line pitch would count them twice and make the DOCX taller than the canonical PDF.
    const lineHeight = Math.max(...group.map((line) => Number(
      line.contentHeight ?? Math.max(0, Number(line.height || 0) - Number(line.before || 0) - Number(line.after || 0)),
    )), 0.05);
    return new Paragraph({
      alignment: alignment(first.alignment),
      spacing: {
        before: Math.max(0, pointsToTwips(first.before || 0)),
        after: Math.max(0, pointsToTwips(last.after || 0)),
        line: Math.max(1, pointsToTwips(lineHeight)),
        lineRule: LineRuleType.EXACT,
      },
      children: runs.length > 0 ? runs : [new TextRun({ text: '' })],
    });
  });
}

function detectImageType(buffer) {
  if (buffer.length > 8 && buffer[0] === 0x89 && buffer.toString('ascii', 1, 4) === 'PNG') return 'png';
  if (buffer.length > 3 && buffer[0] === 0xFF && buffer[1] === 0xD8) return 'jpg';
  if (buffer.length > 6 && buffer.toString('ascii', 0, 3) === 'GIF') return 'gif';
  if (buffer.length > 2 && buffer[0] === 0x42 && buffer[1] === 0x4D) return 'bmp';
  return null;
}

function naturalImageSize(buffer) {
  if (buffer.length > 24 && buffer.toString('ascii', 1, 4) === 'PNG') {
    return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
  }
  if (buffer.length > 4 && buffer[0] === 0xFF && buffer[1] === 0xD8) {
    let index = 2;
    while (index < buffer.length - 8) {
      if (buffer[index] !== 0xFF) {
        index += 1;
        continue;
      }
      const marker = buffer[index + 1];
      if (marker >= 0xC0 && marker <= 0xCF && ![0xC4, 0xC8, 0xCC].includes(marker)) {
        return { height: buffer.readUInt16BE(index + 5), width: buffer.readUInt16BE(index + 7) };
      }
      index += 2 + buffer.readUInt16BE(index + 2);
    }
  }
  if (buffer.length > 10 && buffer.toString('ascii', 0, 3) === 'GIF') {
    return { width: buffer.readUInt16LE(6), height: buffer.readUInt16LE(8) };
  }
  if (buffer.length > 26 && buffer[0] === 0x42 && buffer[1] === 0x4D) {
    return { width: buffer.readInt32LE(18), height: Math.abs(buffer.readInt32LE(22)) };
  }
  return null;
}

function collectModels(model, result = []) {
  result.push(model);
  for (const child of Object.values(model.subreports || {})) collectModels(child, result);
  return result;
}

function collectItems(items, map) {
  for (const item of items || []) {
    if (item.name && !map.has(item.name)) map.set(item.name, item);
    collectItems(item.items, map);
    for (const row of item.rows || []) {
      for (const cell of row.cells || []) collectItems(cell.items, map);
    }
  }
}

function modelResources(model) {
  const embeddedImages = {};
  const items = new Map();
  for (const current of collectModels(model)) {
    Object.assign(embeddedImages, current.embeddedImages || {});
    collectItems(current.body?.items, items);
    collectItems(current.page?.header?.items, items);
    collectItems(current.page?.footer?.items, items);
  }
  return { embeddedImages, items };
}

async function pictureForItem(item, resources, model, request, config, tempDir, chartIndex) {
  let data;
  let type;
  if (item.kind === 'image') {
    const image = resources.embeddedImages[item.embeddedImage];
    if (!image?.data) unsupported('A PDF-traced embedded image is unavailable to the Word renderer', {
      item: item.itemName,
      embeddedImage: item.embeddedImage,
    });
    data = Buffer.from(image.data.replace(/\s+/g, ''), 'base64');
    type = detectImageType(data);
    if (!type) unsupported('The embedded image format cannot be represented safely in native Word', {
      item: item.itemName,
    });
  } else {
    const chart = resources.items.get(item.itemName);
    if (!chart || !config || !tempDir) {
      unsupported('A PDF-traced chart cannot be materialized as a native Word picture', { item: item.itemName });
    }
    const datasets = normalizeDatasets(model, request);
    const globals = {
      PageNumber: item.pageNumber || 1,
      TotalPages: item.totalPages || 1,
      ExecutionTime: new Date(),
      variables: model.variables || {},
    };
    const chartData = materializeChart(chart, datasets, request.parameters || {}, globals);
    const rendered = await renderChartPng(
      chart,
      chartData,
      config,
      tempDir,
      { datasets, parameters: request.parameters || {}, globals, fields: {}, dataset: [] },
      chartIndex,
    );
    if (!rendered?.data) unsupported('A PDF-traced chart could not be rendered for native Word', {
      item: item.itemName,
    });
    data = rendered.data;
    type = 'png';
  }

  let width = item.width;
  let height = item.height;
  const sizing = String(item.sizing || 'FitProportional');
  if (/^(Clip|AutoSize)$/i.test(sizing)) {
    unsupported(`RDL image sizing '${sizing}' is not safely representable in the page-locked editable Word contract`, {
      item: item.itemName,
    });
  }
  if (!/^Fit$/i.test(sizing)) {
    const natural = naturalImageSize(data);
    if (natural) {
      const scale = Math.min(item.width / natural.width, item.height / natural.height);
      width = natural.width * scale;
      height = natural.height * scale;
    }
  }
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    // An inline picture participates in Word line layout. A one-twip line clips or suppresses pictures
    // that fill their traced cell (most visibly repeated page-header logos). Use the PDF-resolved picture
    // height as the exact native line box; the enclosing table row and cell remain fixed to the same trace.
    spacing: {
      before: 0,
      after: 0,
      line: Math.max(1, pointsToTwips(height)),
      lineRule: LineRuleType.EXACT,
    },
    children: [new ImageRun({
      data,
      type,
      transformation: {
        width: Math.max(1, pointsToDrawingPixels(width)),
        height: Math.max(1, pointsToDrawingPixels(height)),
      },
    })],
  });
}

function lineBorder(line) {
  if (!line?.line || /^none$/i.test(String(line.line.style || 'None'))) return null;
  return {
    style: line.line.style || 'Solid',
    width: Number(line.line?.width || 1),
    color: line.line?.color || '#000000',
  };
}

function edgeMatches(item, side, box) {
  if (side === 'top' || side === 'bottom') {
    const itemY = side === 'top' ? item.y : item.y + item.height;
    const boxY = side === 'top' ? box.y : box.y + box.height;
    return Math.abs(itemY - boxY) <= GEOMETRY_EPSILON
      && item.x <= box.x + GEOMETRY_EPSILON
      && item.x + item.width >= box.x + box.width - GEOMETRY_EPSILON;
  }
  const itemX = side === 'left' ? item.x : item.x + item.width;
  const boxX = side === 'left' ? box.x : box.x + box.width;
  return Math.abs(itemX - boxX) <= GEOMETRY_EPSILON
    && item.y <= box.y + GEOMETRY_EPSILON
    && item.y + item.height >= box.y + box.height - GEOMETRY_EPSILON;
}

function lineMatches(line, side, box) {
  const horizontal = Math.abs(line.height) <= GEOMETRY_EPSILON;
  const vertical = Math.abs(line.width) <= GEOMETRY_EPSILON;
  if (!horizontal && !vertical) return false;
  if ((side === 'top' || side === 'bottom') && horizontal) {
    const y = side === 'top' ? box.y : box.y + box.height;
    return Math.abs(line.y - y) <= GEOMETRY_EPSILON
      && line.x <= box.x + GEOMETRY_EPSILON
      && line.x + line.width >= box.x + box.width - GEOMETRY_EPSILON;
  }
  if ((side === 'left' || side === 'right') && vertical) {
    const x = side === 'left' ? box.x : box.x + box.width;
    return Math.abs(line.x - x) <= GEOMETRY_EPSILON
      && line.y <= box.y + GEOMETRY_EPSILON
      && line.y + line.height >= box.y + box.height - GEOMETRY_EPSILON;
  }
  return false;
}

function resolvedCellBorders(box, owner, decorators, lines) {
  return Object.fromEntries(['top', 'right', 'bottom', 'left'].map((side) => {
    let resolved = owner?.borders?.[side] || null;
    for (const decorator of decorators) {
      if (edgeMatches(decorator, side, box)) resolved = strongerBorder(resolved, decorator.borders?.[side]);
    }
    for (const line of lines) {
      if (lineMatches(line, side, box)) resolved = strongerBorder(resolved, lineBorder(line));
    }
    return [side, wordBorder(resolved)];
  }));
}

function resolvedBackground(box, owner, decorators) {
  if (owner?.backgroundColor) return owner.backgroundColor;
  const containing = decorators
    .filter((item) => item.backgroundColor && contains(item, box))
    .sort((left, right) => Number(left.zIndex || 0) - Number(right.zIndex || 0));
  return containing.at(-1)?.backgroundColor || null;
}

function preparePageGrid(page, {
  items = page.items,
  originY = 0,
  canvasHeight = page.height,
  reserveSectionAnchor = true,
} = {}) {
  if (page.width > WORD_MAX_PAGE_POINTS + GEOMETRY_EPSILON || page.height > WORD_MAX_PAGE_POINTS + GEOMETRY_EPSILON) {
    unsupported('A PDF page exceeds Microsoft Word’s 22-by-22-inch page-size limit', {
      page: page.number,
      widthPt: page.width,
      heightPt: page.height,
    });
  }
  // Snap physical edges, not origins and dimensions independently. Independent rounding can move the
  // derived right/bottom edge by another quarter point and turn two coincident PDF cells into a false
  // overlap in Word (or leave a false gap). This is the same edge-coalescing semantic the PDF trace uses.
  const normalized = items.map((item) => {
    const x = snap(item.x);
    const y = snap(Number(item.y || 0) - originY);
    const right = snap(Number(item.x || 0) + Number(item.width || 0));
    const bottom = snap(Number(item.y || 0) + Number(item.height || 0) - originY);
    return {
      ...item,
      x,
      y,
      width: Math.max(0, right - x),
      height: Math.max(0, bottom - y),
    };
  }).filter((item) => item.width >= 0 && item.height >= 0);
  const maximumCanvasBottom = canvasHeight - (reserveSectionAnchor ? SECTION_ANCHOR_POINTS : 0);
  // The table needs to extend only through the last painted PDF primitive. Empty space below it remains
  // ordinary page space; filling that space with exact-height blank table rows makes Word
  // push the mandatory terminal section paragraph onto a spurious blank page.
  const canvasBottom = Math.max(
    GRID_PRECISION_POINTS,
    normalized.length > 0
      ? Math.min(maximumCanvasBottom, Math.max(...normalized.map((item) => item.y + item.height)))
      : GRID_PRECISION_POINTS,
  );

  for (const item of normalized) {
    if (item.x < -GEOMETRY_EPSILON || item.y < -GEOMETRY_EPSILON
      || item.x + item.width > page.width + GEOMETRY_EPSILON
      || item.y + item.height > maximumCanvasBottom + GEOMETRY_EPSILON) {
      unsupported('A PDF item falls outside the Word page canvas', {
        page: page.number,
        item: item.itemName,
        kind: item.kind,
      });
    }
    if (item.writingMode && !/^default$/i.test(item.writingMode)) {
      unsupported('Rotated or vertical editable text is not safely representable by the page-locked Word renderer', {
        page: page.number,
        item: item.itemName,
        writingMode: item.writingMode,
      });
    }
    if (item.kind === 'line'
      && Math.abs(item.width) > GEOMETRY_EPSILON
      && Math.abs(item.height) > GEOMETRY_EPSILON) {
      unsupported('Diagonal RDL lines are not safely representable as native Word cell borders', {
        page: page.number,
        item: item.itemName,
      });
    }
  }

  const decorators = normalized.filter((item) => item.kind === 'rectangle');
  const lines = normalized.filter((item) => item.kind === 'line');
  const candidates = normalized.filter((item) => ['textbox', 'tablixCell', 'image', 'chart'].includes(item.kind));
  const demoted = new Set();
  for (const candidate of candidates) {
    if (!isEmptyCell(candidate)) continue;
    if (candidates.some((other) => other !== candidate && contains(candidate, other) && positiveOverlap(candidate, other))) {
      demoted.add(candidate);
      decorators.push(candidate);
    }
  }
  const owners = candidates.filter((candidate) => !demoted.has(candidate));
  for (let leftIndex = 0; leftIndex < owners.length; leftIndex += 1) {
    for (let rightIndex = leftIndex + 1; rightIndex < owners.length; rightIndex += 1) {
      const left = owners[leftIndex];
      const right = owners[rightIndex];
      if (positiveOverlap(left, right)) {
        unsupported('Overlapping editable PDF regions cannot be represented safely as native Word cells', {
          page: page.number,
          first: left.itemName,
          second: right.itemName,
        });
      }
    }
  }

  for (const line of lines) {
    for (const owner of owners) {
      if (positiveOverlap({
        x: line.x - GEOMETRY_EPSILON,
        y: line.y - GEOMETRY_EPSILON,
        width: Math.max(line.width, GEOMETRY_EPSILON * 2),
        height: Math.max(line.height, GEOMETRY_EPSILON * 2),
      }, owner)
        && !['top', 'right', 'bottom', 'left'].some((side) => lineMatches(line, side, owner))) {
        unsupported('An RDL line crosses editable content instead of coinciding with a cell edge', {
          page: page.number,
          line: line.itemName,
          item: owner.itemName,
        });
      }
    }
  }

  const xBoundaries = uniqueBoundaries([
    0,
    page.width,
    ...normalized.flatMap((item) => [item.x, item.x + item.width]),
  ]);
  const yBoundaries = uniqueBoundaries([
    0,
    canvasBottom,
    ...normalized.flatMap((item) => [item.y, item.y + item.height]),
  ]);
  if (xBoundaries.length - 1 > WORD_MAX_TABLE_COLUMNS) {
    unsupported('The PDF page requires more than Microsoft Word’s 63 table columns', {
      page: page.number,
      columns: xBoundaries.length - 1,
    });
  }
  if (yBoundaries.length - 1 > WORD_MAX_TABLE_ROWS) {
    unsupported('The PDF page requires more Word table rows than the platform supports', {
      page: page.number,
      rows: yBoundaries.length - 1,
    });
  }

  const placements = owners.map((item) => ({
    item,
    startColumn: boundaryIndex(xBoundaries, item.x),
    endColumn: boundaryIndex(xBoundaries, item.x + item.width),
    startRow: boundaryIndex(yBoundaries, item.y),
    endRow: boundaryIndex(yBoundaries, item.y + item.height),
  }));
  const coverage = Array.from(
    { length: yBoundaries.length - 1 },
    () => Array(xBoundaries.length - 1).fill(null),
  );
  for (const placement of placements) {
    for (let row = placement.startRow; row < placement.endRow; row += 1) {
      for (let column = placement.startColumn; column < placement.endColumn; column += 1) {
        if (coverage[row][column]) {
          unsupported('Two PDF regions resolve to the same native Word grid cell', {
            page: page.number,
            first: coverage[row][column].item.itemName,
            second: placement.item.itemName,
          });
        }
        coverage[row][column] = placement;
      }
    }
  }
  return { page, xBoundaries, yBoundaries, placements, coverage, decorators, lines };
}

function emptyFooterParagraph() {
  return new Paragraph({
    spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
    children: [new TextRun({ text: '' })],
  });
}

function footerLayout(page) {
  const region = page.regions?.footer;
  if (!region || region.height <= GEOMETRY_EPSILON) return null;
  const items = page.items.filter((item) => item.region === 'footer');
  const contentBottom = items.length > 0
    ? Math.max(...items.map((item) => Number(item.y || 0) + Number(item.height || 0)))
    : region.y + region.height;
  const height = Math.max(region.height, contentBottom - region.y);
  return {
    region,
    items,
    height,
    bottomDistance: Math.max(0, page.height - region.y - height),
  };
}

async function nativePageFooter(
  page,
  resources,
  model,
  request,
  config,
  tempDir,
  chartCounter,
) {
  const layout = footerLayout(page);
  if (!layout) return null;

  if (layout.items.length === 0) {
    // Every canonical page owns its footer relationship. An explicit empty footer prevents Word from
    // inheriting visible content from the previous one-page section when the RDL hides a page footer.
    return new Footer({ children: [emptyFooterParagraph()] });
  }

  const grid = preparePageGrid(page, {
    items: layout.items,
    originY: layout.region.y,
    canvasHeight: layout.height,
    reserveSectionAnchor: false,
  });
  return new Footer({
    children: [await pageTable(
      grid,
      resources,
      model,
      request,
      config,
      tempDir,
      chartCounter,
    )],
  });
}

function cellBox(grid, row, column, rowSpan = 1, columnSpan = 1) {
  return {
    x: grid.xBoundaries[column],
    y: grid.yBoundaries[row],
    width: grid.xBoundaries[column + columnSpan] - grid.xBoundaries[column],
    height: grid.yBoundaries[row + rowSpan] - grid.yBoundaries[row],
  };
}

function cellMargins(item) {
  const padding = item?.padding || {};
  return {
    top: Math.max(0, pointsToTwips(padding.top || 0)),
    right: Math.max(0, pointsToTwips(padding.right || 0)),
    bottom: Math.max(0, pointsToTwips(padding.bottom || 0)),
    left: Math.max(0, pointsToTwips(padding.left || 0)),
    marginUnitType: WidthType.DXA,
  };
}

async function tableCellFor(grid, row, column, placement, resources, model, request, config, tempDir, chartCounter) {
  const rowSpan = placement ? placement.endRow - placement.startRow : 1;
  const columnSpan = placement ? placement.endColumn - placement.startColumn : 1;
  const owner = placement?.item || null;
  const box = cellBox(grid, row, column, rowSpan, columnSpan);
  const margins = owner ? cellMargins(owner) : {
    top: 0, right: 0, bottom: 0, left: 0, marginUnitType: WidthType.DXA,
  };
  let children;
  if (owner?.kind === 'image' || owner?.kind === 'chart') {
    const withPage = {
      ...owner,
      pageNumber: grid.page.number,
      totalPages: request.__canonicalPageCount,
    };
    children = [await pictureForItem(
      withPage,
      resources,
      model,
      request,
      config,
      tempDir,
      chartCounter.value++,
    )];
  } else {
    children = owner ? linesForParagraphs(owner) : [new Paragraph({
      spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
      children: [new TextRun({ text: '' })],
    })];
  }
  const background = resolvedBackground(box, owner, grid.decorators);
  return {
    cell: new TableCell({
      width: { size: pointsToTwips(box.width), type: WidthType.DXA },
      columnSpan,
      rowSpan,
      margins,
      verticalAlign: owner?.kind === 'image' || owner?.kind === 'chart'
        ? VerticalAlignTable.CENTER
        : verticalAlignment(owner?.verticalAlign),
      shading: background ? { type: ShadingType.CLEAR, fill: cleanColor(background), color: 'auto' } : undefined,
      borders: resolvedCellBorders(box, owner, grid.decorators, grid.lines),
      children,
    }),
    bottomPaddingTwips: margins.bottom,
  };
}

async function pageTable(grid, resources, model, request, config, tempDir, chartCounter) {
  const rows = [];
  for (let row = 0; row < grid.yBoundaries.length - 1; row += 1) {
    const children = [];
    let maximumBottomPaddingTwips = 0;
    let column = 0;
    while (column < grid.xBoundaries.length - 1) {
      const placement = grid.coverage[row][column];
      if (placement) {
        if (placement.startRow < row) {
          if (placement.startColumn === column) column = placement.endColumn;
          else column += 1;
          continue;
        }
        if (placement.startColumn !== column) {
          column += 1;
          continue;
        }
        const built = await tableCellFor(
          grid,
          row,
          column,
          placement,
          resources,
          model,
          request,
          config,
          tempDir,
          chartCounter,
        );
        children.push(built.cell);
        maximumBottomPaddingTwips = Math.max(maximumBottomPaddingTwips, built.bottomPaddingTwips);
        column = placement.endColumn;
      } else {
        const built = await tableCellFor(
          grid,
          row,
          column,
          null,
          resources,
          model,
          request,
          config,
          tempDir,
          chartCounter,
        );
        children.push(built.cell);
        maximumBottomPaddingTwips = Math.max(maximumBottomPaddingTwips, built.bottomPaddingTwips);
        column += 1;
      }
    }
    const tracedHeightTwips = Math.max(
      1,
      pointsToTwips(grid.yBoundaries[row + 1] - grid.yBoundaries[row]),
    );
    // Microsoft Word adds the largest bottom cell margin to an exact w:trHeight. The PDF trace already
    // includes that padding inside the physical row box, so emitting the unadjusted trace height makes
    // every padded Word row too tall. Near a page boundary, the accumulated surplus causes an otherwise
    // fitting cantSplit row to jump wholesale to the next page. Compensate from the actual cells in this
    // physical row so Word's final height remains the canonical PDF height.
    const wordHeightTwips = tracedHeightTwips - maximumBottomPaddingTwips;
    if (wordHeightTwips < 1) {
      unsupported('A PDF row is too short to preserve its bottom cell padding in an exact Word row', {
        page: grid.page.number,
        row: row + 1,
        heightTwips: tracedHeightTwips,
        bottomPaddingTwips: maximumBottomPaddingTwips,
      });
    }
    rows.push(new TableRow({
      cantSplit: true,
      height: {
        value: wordHeightTwips,
        rule: HeightRule.EXACT,
      },
      children,
    }));
  }
  return new Table({
    width: { size: pointsToTwips(grid.page.width), type: WidthType.DXA },
    columnWidths: grid.xBoundaries.slice(1).map((value, index) => (
      pointsToTwips(value - grid.xBoundaries[index])
    )),
    margins: { top: 0, right: 0, bottom: 0, left: 0 },
    indent: { size: 0, type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    borders: TableBorders.NONE,
    rows,
  });
}

function pageProperties(page, index) {
  const landscape = page.width > page.height;
  const footerDistance = footerLayout(page)?.bottomDistance || 0;
  return {
    type: index === 0 ? undefined : SectionType.NEXT_PAGE,
    page: {
      size: landscape
        ? {
          width: pointsToTwips(page.height),
          height: pointsToTwips(page.width),
          orientation: PageOrientation.LANDSCAPE,
        }
        : {
          width: pointsToTwips(page.width),
          height: pointsToTwips(page.height),
          orientation: PageOrientation.PORTRAIT,
        },
      margin: {
        top: 0,
        right: 0,
        bottom: 0,
        left: 0,
        header: 0,
        footer: pointsToTwips(footerDistance),
        gutter: 0,
      },
    },
  };
}

function consumedFamilies(trace) {
  return [...new Set(trace.pages.flatMap((page) => page.items.flatMap((item) => (
    (item.lines || []).flatMap((line) => (line.runs || []).map((run) => run.font?.family).filter(Boolean))
  ))))];
}

async function embeddedFontFamilies(trace, config) {
  const families = consumedFamilies(trace);
  const result = [];
  for (const family of families) {
    const files = {};
    let missing = null;
    for (const variant of VARIANTS) {
      const file = resolveFontFile(config.fontDir, family, variant.bold, variant.italic);
      if (!file) {
        missing = variant.key;
        break;
      }
      files[variant.key] = {
        file,
        data: await fs.readFile(file),
        ...editableFontEmbeddingPermission(file, family, variant.key),
      };
    }
    if (missing) {
      throw new ServiceError('FONT_MISSING', `Required font is unavailable: ${family}:${missing}`, 503, {
        family,
        variant: missing,
      });
    }
    result.push({ family, files });
  }
  return result;
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function obfuscateFont(data, fontKey) {
  const guidBytes = fontKey
    .replace(/-/g, '')
    .match(/../g)
    .map((hex) => Number.parseInt(hex, 16))
    .reverse();
  const result = Buffer.from(data);
  for (let index = 0; index < Math.min(32, result.length); index += 1) {
    result[index] ^= guidBytes[index % guidBytes.length];
  }
  return result;
}

async function addFontVariants(buffer, embeddedFonts) {
  const zip = await JSZip.loadAsync(buffer);
  if (embeddedFonts.length > 0) {
    const fontTableFile = zip.file('word/fontTable.xml');
    const relationshipsFile = zip.file('word/_rels/fontTable.xml.rels');
    if (!fontTableFile || !relationshipsFile) {
      throw new ServiceError('RENDER_FAILED', 'Word font table packaging is incomplete', 500);
    }
    let fontTable = await fontTableFile.async('string');
    let relationships = await relationshipsFile.async('string');
    const existingIds = [...relationships.matchAll(/\bId="rId(\d+)"/g)].map((match) => Number(match[1]));
    let nextRelationship = Math.max(0, ...existingIds) + 1;
    const existingTargets = [...relationships.matchAll(/Target="fonts\/font(\d+)\.odttf"/g)]
      .map((match) => Number(match[1]));
    let nextFontPart = Math.max(0, ...existingTargets) + 1;

    for (const embedded of embeddedFonts) {
      const name = escapeXml(embedded.family);
      const fontPattern = new RegExp(`(<w:font\\s+w:name="${escapeRegExp(name)}"[^>]*>)([\\s\\S]*?)(</w:font>)`);
      const matched = fontTable.match(fontPattern);
      if (!matched) {
        throw new ServiceError('RENDER_FAILED', `Embedded Word font entry is missing: ${embedded.family}`, 500);
      }
      const existingFontMarkup = matched[2].replace(/<w:embedRegular\b([^>]*)\/>/, (_element, attributes) => {
        const explicitFullFile = /\bw:subsetted=/.test(attributes)
          ? attributes.replace(/\bw:subsetted="[^"]*"/, 'w:subsetted="0"')
          : `${attributes} w:subsetted="0"`;
        return `<w:embedRegular${explicitFullFile}/>`;
      });
      const additions = [];
      for (const variant of VARIANTS.filter((candidate) => candidate.key !== 'regular')) {
        const relationshipId = `rId${nextRelationship++}`;
        const fontPart = `font${nextFontPart++}.odttf`;
        const fontKey = randomUUID().toUpperCase();
        additions.push(`<w:${variant.element} r:id="${relationshipId}" w:fontKey="{${fontKey}}" w:subsetted="0"/>`);
        relationships = relationships.replace(
          '</Relationships>',
          `<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/${fontPart}"/></Relationships>`,
        );
        zip.file(`word/fonts/${fontPart}`, obfuscateFont(embedded.files[variant.key].data, fontKey));
      }
      fontTable = fontTable.replace(fontPattern, `${matched[1]}${existingFontMarkup}${additions.join('')}${matched[3]}`);
    }
    zip.file('word/fontTable.xml', fontTable);
    zip.file('word/_rels/fontTable.xml.rels', relationships);
  }

  // docx assigns wp:docPr id="1" to every independently-created ImageRun. Repeated IDs violate the
  // DrawingML non-visual-property identity contract and can make Word repair the package or suppress
  // repeated images. Normalize IDs across all document stories after construction.
  const drawingParts = Object.keys(zip.files)
    .filter((name) => /^word\/(?:document|header\d+|footer\d+)\.xml$/i.test(name))
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
  let nextDrawingId = 1;
  for (const name of drawingParts) {
    const file = zip.file(name);
    if (!file) continue;
    const xml = await file.async('string');
    zip.file(name, xml.replace(
      /(<wp:docPr\b[^>]*\bid=")\d+(")/g,
      (_match, prefix, suffix) => `${prefix}${nextDrawingId++}${suffix}`,
    ));
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

async function writeInternalArtifacts(tempDir, pdf, trace) {
  if (!tempDir) return null;
  const pdfPath = path.join(tempDir, 'docx-canonical.pdf');
  const tracePath = path.join(tempDir, 'docx-layout-trace.json');
  try {
    await fs.writeFile(pdfPath, pdf, { mode: 0o600 });
    await fs.writeFile(tracePath, JSON.stringify(trace), { mode: 0o600 });
    return { pdfPath, tracePath };
  } catch (error) {
    await Promise.all([pdfPath, tracePath].map((file) => fs.unlink(file).catch(() => {})));
    throw error;
  }
}

async function cleanupInternalArtifacts(files) {
  if (!files) return;
  await Promise.all(Object.values(files).map((file) => fs.unlink(file).catch(() => {})));
}

export async function renderPagedEditableDocx(model, request, config, tempDir) {
  config ||= loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
  validateWindowsWordRequest(request);
  const canonical = await renderPdf(model, request, config, { captureLayoutTrace: true });
  const trace = canonical.layoutTrace;
  try {
    validateLayoutTrace(trace, canonical.pageCount);
  } catch (error) {
    throw new ServiceError('RENDER_FAILED', 'Canonical PDF layout trace is incomplete', 500, {
      cause: error.message,
    });
  }
  const internalFiles = await writeInternalArtifacts(tempDir, canonical.buffer, trace);
  try {
    const embeddedFonts = await embeddedFontFamilies(trace, config);
    const resources = modelResources(model);
    const canonicalRequest = { ...request, __canonicalPageCount: canonical.pageCount };
    const chartCounter = { value: 0 };
    const sections = [];
    for (const [index, page] of trace.pages.entries()) {
      // Footer items must not participate in body flow. Word always inserts a terminal paragraph after
      // the section's body table; keeping footer rows in that table lets the terminal paragraph push the
      // final footer row onto a blank page. A native footer is positioned independently from body flow.
      const bodyGrid = preparePageGrid(page, {
        items: page.items.filter((item) => item.region !== 'footer'),
      });
      const footer = await nativePageFooter(
        page,
        resources,
        model,
        canonicalRequest,
        config,
        tempDir,
        chartCounter,
      );
      sections.push({
        properties: pageProperties(page, index),
        footers: footer ? { default: footer } : undefined,
        children: [await pageTable(
          bodyGrid,
          resources,
          model,
          canonicalRequest,
          config,
          tempDir,
          chartCounter,
        )],
      });
    }
    const document = new Document({
      creator: 'RDL Converter Service',
      title: request.outputFileName || model.name,
      description: 'Windows Word page-locked editable rendering derived from the canonical PDF layout trace',
      compatibilityModeVersion: 15,
      features: { updateFields: false },
      fonts: embeddedFonts.map((embedded) => ({
        name: embedded.family,
        data: embedded.files.regular.data,
      })),
      styles: {
        default: {
          document: {
            run: { font: 'Arial', size: 2 },
            paragraph: {
              spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
            },
          },
        },
      },
      sections,
    });
    let buffer = await Packer.toBuffer(document);
    buffer = await addFontVariants(buffer, embeddedFonts);
    return {
      buffer,
      pageCount: canonical.pageCount,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      extension: 'docx',
      layoutMode: 'windows-paged-editable',
      editableTextRatio: 1,
      canonicalPdfSha256: createHash('sha256').update(canonical.buffer).digest('hex'),
    };
  } finally {
    await cleanupInternalArtifacts(internalFiles);
  }
}
