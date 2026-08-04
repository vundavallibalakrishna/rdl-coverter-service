import { evaluateExpression, formatValue } from '../rdl/expression.js';
import { normalizeRowFields } from '../rdl/fields.js';
import { materializeTablixColumns, materializeTablixRows } from '../rdl/validation.js';
import { normalizeDisplayText, renderMarkupText } from '../rdl/text.js';
import { toPoints } from '../units.js';

// Cross-format labels for an explicitly detected logical row continuation. The feature is opt-in because
// adding explanatory text changes the rendered artifact. Renderers must only emit these labels when they
// own the pagination decision; a viewer-created page break is not detectable while generating the file.
export const CONTINUATION_MARKERS = Object.freeze({
  fromPrevious: 'Continued from previous page',
});

export function continuationMarkersEnabled(request) {
  return request?.pagination?.continuationMarkers === true;
}

export function isHidden(expression, context) {
  const result = evaluateExpression(expression, context);
  return result === true || String(result).toLowerCase() === 'true';
}

// Resolves a size that may be a literal number, a unit string ("1pt"), or a conditional expression
// (=IIF(...,"1pt","0pt")) into points. Border widths and similar sizes can be data-dependent, so they are
// kept as expressions by the parser and resolved here per row/scope.
export function styleSize(value, context, fallback = 0) {
  if (typeof value === 'number') return value;
  const resolved = styleValue(value, context, undefined);
  if (typeof resolved === 'number') return resolved;
  if (resolved === undefined || resolved === null || resolved === '') return fallback;
  try {
    return toPoints(String(resolved), fallback);
  } catch {
    return fallback;
  }
}

export function styledTextForItem(item, context) {
  if (!item.paragraphs) return null;
  return item.paragraphs.map((runs, paragraphIndex) => ({
    style: item.paragraphStyles?.[paragraphIndex] || item.style,
    runs: runs.map((run) => {
      const definition = run && typeof run === 'object' ? run : { value: run, markupType: 'None' };
      const value = /^constant$/i.test(String(definition.evaluationMode || 'Auto'))
        ? definition.value
        : evaluateExpression(definition.value, context);
      const runStyle = definition.style || item.style;
      if (value === null || value === undefined) return { text: '', style: runStyle };
      const format = styleValue(runStyle?.format ?? item.style?.format, context, null);
      const formatted = format ? String(formatValue(value, format)) : String(value);
      return {
        text: normalizeDisplayText(renderMarkupText(formatted, definition.markupType)),
        style: runStyle,
      };
    }),
  }));
}

// Flattens an item's styled paragraphs into text segments, each carrying its run (or paragraph) style, and —
// when a specific text is requested (e.g. a materialized tablix cell's already-flattened value) — slices
// those segments to that text so per-run styling (bold / colour / font) survives the flattening. Returns
// null when the item has no paragraphs, or when the requested text is not a contiguous slice of the item's
// own text (a cell combining several conditional textboxes), so callers fall back to a plain single-style
// path. Shared by the PDF, DOCX, and Excel renderers so all three honour per-run styles identically.
export function styledSegmentsForText(item, context, requestedText) {
  const paragraphs = styledTextForItem(item, context);
  if (!paragraphs) return null;
  const segments = [];
  let fullText = '';
  for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
    for (const run of paragraph.runs) {
      const text = String(run.text ?? '');
      segments.push({ text, style: run.style || item.style, paragraphStyle: paragraph.style || item.style, paragraphIndex });
      fullText += text;
    }
    if (paragraphIndex < paragraphs.length - 1) {
      segments.push({ text: '\n', style: paragraph.style || item.style, paragraphStyle: paragraph.style || item.style, paragraphIndex });
      fullText += '\n';
    }
  }
  const target = requestedText === undefined ? fullText : String(requestedText ?? '');
  if (target.length === 0) return { text: target, segments: [] };
  let start = 0;
  if (target !== fullText) {
    if (fullText.startsWith(target)) start = 0;
    else if (fullText.endsWith(target)) start = fullText.length - target.length;
    else start = fullText.indexOf(target);
    // A materialized cell can combine several conditional textboxes. If the override is not a contiguous
    // slice of this textbox, retain the plain-text path rather than assigning misleading styles.
    if (start < 0) return null;
  }
  const end = start + target.length;
  let offset = 0;
  const sliced = [];
  for (const segment of segments) {
    const segmentStart = offset;
    const segmentEnd = offset + segment.text.length;
    offset = segmentEnd;
    const from = Math.max(start, segmentStart);
    const to = Math.min(end, segmentEnd);
    if (to <= from) continue;
    sliced.push({ ...segment, text: segment.text.slice(from - segmentStart, to - segmentStart) });
  }
  return { text: target, segments: sliced };
}

export function textForItem(item, context) {
  const styled = styledTextForItem(item, context);
  if (styled) return normalizeDisplayText(styled.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n'));
  const value = evaluateExpression(item.value, context);
  if (value === null || value === undefined) return '';
  return normalizeDisplayText(item.style?.format ? String(formatValue(value, styleValue(item.style.format, context))) : String(value));
}

function styleDeclaresVisibleBorder(style) {
  const sides = style?.borders || (style?.border
    ? { top: style.border, right: style.border, bottom: style.border, left: style.border }
    : null);
  return Object.values(sides || {}).some((border) => {
    if (!border || border.style === undefined) return false;
    // Expressions are border intent even when a particular row resolves to None. Literal None on every
    // side is the RDL author's explicit declaration that this item is a borderless layout construct.
    return typeof border.style === 'string' && border.style.startsWith('=')
      ? true
      : !/^none$/i.test(String(border.style));
  });
}

function itemDeclaresVisibleBorder(item) {
  if (!item) return false;
  if (styleDeclaresVisibleBorder(item.style)) return true;
  return (item.items || item.children || []).some(itemDeclaresVisibleBorder);
}

// A populated dynamic tablix is not necessarily a visual data grid: RDL commonly uses data-bound tablixes
// as borderless narrative/layout containers. Enforce a physical closing edge only when the tablix actually
// declares border intent at the table or cell-content level. This preserves fragment closure for bordered
// grids while honoring an explicit Border=None throughout a dynamic prose section.
export function shouldEnforceTablixBottom(rows, tablix) {
  const hasDynamicRows = (rows || []).some((row) => row.isStatic === false);
  if (!hasDynamicRows) return false;
  // A trailing static row is an intentional matrix/footer band. Its own declared edges are authoritative:
  // synthesizing a bottom closure across it turns borderless axis labels and legends into a decorative
  // horizontal rule that SSRS does not draw. Dynamic detail and group-footer rows remain covered below.
  if (rows?.at(-1)?.isStatic === true) return false;
  if (styleDeclaresVisibleBorder(tablix?.style)) return true;
  // Spans are structural grid intent even if individual sides are omitted. Their physical fragments must
  // still close where a merged owner reaches a page/table boundary.
  if ((rows || []).some((row) => (row.cells || []).some((cell) => (
    (cell.rowSpan || 1) > 1 || (cell.colSpan || 1) > 1
  )))) return true;
  return (rows || []).some((row) => (row.cells || []).some((cell) => (
    (cell.items || []).some(itemDeclaresVisibleBorder)
  )));
}

// For data tablixes, the last row must be closed with a bottom border even when the RDL leaves the
// tablix/last-row bottom edge as None (or the row splits across a page). Returns the declared bottom border
// when it is already a visible rule (Solid, or a conditional expression the row may resolve), otherwise a
// Solid border matching another declared side of the same style (so the enforced edge picks up the table's
// own colour/width), falling back to black 1pt. Shared by the PDF (outer fragment bottom) and DOCX (last-row
// cell borders) so both renderers close the table identically.
export function enforcedBottomBorder(style) {
  const sides = style?.borders || (style?.border
    ? { top: style.border, right: style.border, bottom: style.border, left: style.border }
    : null);
  const visible = (border) => border && border.style !== undefined && !/^none$/i.test(String(border.style));
  if (visible(sides?.bottom)) return sides.bottom;
  const template = [sides?.left, sides?.right, sides?.top].find(visible);
  return { style: 'Solid', color: template?.color ?? '#000000', width: template?.width ?? 1 };
}

export function normalizeDatasets(model, request) {
  const context = { parameters: request.parameters || {} };
  return Object.fromEntries(model.datasets.map((dataset) => [
    dataset.name,
    (request.datasets?.[dataset.name] || []).map((row) => normalizeRowFields(row, dataset.fields, context)),
  ]));
}

// A materialized cell owns the row/group scope in which SSRS evaluated it. That scope can differ from the
// physical row currently visiting the cell after row spanning, matrix intersection, nested layout, or page
// fragmentation. Keep style expressions (especially conditional borders) grounded in the cell owner rather
// than accidentally re-evaluating them against an adjacent detail row.
export function materializedCellContext(cell, row, {
  parameters = {}, globals = {}, dataset = [], datasets = {},
} = {}) {
  return {
    fields: cell?.fields ?? row?.fields ?? {},
    parameters,
    globals,
    dataset: cell?.scopeDataset ?? row?.scopeDataset ?? dataset,
    datasets,
    scopes: cell?.scopes ?? row?.scopes ?? {},
  };
}

// Returns the materialized rows AND the grid column widths. For static-column tablixes `columns`
// is identically `item.columns`; for a matrix (dynamic column groups) it is the expanded
// rowHeader + keys×body column array, so both renderers build the same expanded grid.
export function tablixRows(item, request, globals = {}, model) {
  const rows = request.datasets?.[item.datasetName] || [];
  const parameters = request.parameters || {};
  const datasets = model ? normalizeDatasets(model, request) : {};
  return {
    rows: materializeTablixRows(item, rows, parameters, globals, datasets),
    columns: materializeTablixColumns(item, rows, parameters, globals, datasets),
  };
}

// Full CSS/X11 named-colour table. PDFKit only understands hex and a handful of HTML4 names, so any
// named colour an RDL uses (e.g. SlateBlue, DarkOliveGreen, LightBlue) must be resolved to hex here or
// the fill silently fails and the shape renders empty/black.
const NAMED_COLORS = {
  transparent: null, none: null,
  aliceblue: '#f0f8ff', antiquewhite: '#faebd7', aqua: '#00ffff', aquamarine: '#7fffd4', azure: '#f0ffff',
  beige: '#f5f5dc', bisque: '#ffe4c4', black: '#000000', blanchedalmond: '#ffebcd', blue: '#0000ff',
  blueviolet: '#8a2be2', brown: '#a52a2a', burlywood: '#deb887', cadetblue: '#5f9ea0', chartreuse: '#7fff00',
  chocolate: '#d2691e', coral: '#ff7f50', cornflowerblue: '#6495ed', cornsilk: '#fff8dc', crimson: '#dc143c',
  cyan: '#00ffff', darkblue: '#00008b', darkcyan: '#008b8b', darkgoldenrod: '#b8860b', darkgray: '#a9a9a9',
  darkgreen: '#006400', darkgrey: '#a9a9a9', darkkhaki: '#bdb76b', darkmagenta: '#8b008b', darkolivegreen: '#556b2f',
  darkorange: '#ff8c00', darkorchid: '#9932cc', darkred: '#8b0000', darksalmon: '#e9967a', darkseagreen: '#8fbc8f',
  darkslateblue: '#483d8b', darkslategray: '#2f4f4f', darkslategrey: '#2f4f4f', darkturquoise: '#00ced1',
  darkviolet: '#9400d3', deeppink: '#ff1493', deepskyblue: '#00bfff', dimgray: '#696969', dimgrey: '#696969',
  dodgerblue: '#1e90ff', firebrick: '#b22222', floralwhite: '#fffaf0', forestgreen: '#228b22', fuchsia: '#ff00ff',
  gainsboro: '#dcdcdc', ghostwhite: '#f8f8ff', gold: '#ffd700', goldenrod: '#daa520', gray: '#808080',
  green: '#008000', greenyellow: '#adff2f', grey: '#808080', honeydew: '#f0fff0', hotpink: '#ff69b4',
  indianred: '#cd5c5c', indigo: '#4b0082', ivory: '#fffff0', khaki: '#f0e68c', lavender: '#e6e6fa',
  lavenderblush: '#fff0f5', lawngreen: '#7cfc00', lemonchiffon: '#fffacd', lightblue: '#add8e6', lightcoral: '#f08080',
  lightcyan: '#e0ffff', lightgoldenrodyellow: '#fafad2', lightgray: '#d3d3d3', lightgreen: '#90ee90', lightgrey: '#d3d3d3',
  lightpink: '#ffb6c1', lightsalmon: '#ffa07a', lightseagreen: '#20b2aa', lightskyblue: '#87cefa', lightslategray: '#778899',
  lightslategrey: '#778899', lightsteelblue: '#b0c4de', lightyellow: '#ffffe0', lime: '#00ff00', limegreen: '#32cd32',
  linen: '#faf0e6', magenta: '#ff00ff', maroon: '#800000', mediumaquamarine: '#66cdaa', mediumblue: '#0000cd',
  mediumorchid: '#ba55d3', mediumpurple: '#9370db', mediumseagreen: '#3cb371', mediumslateblue: '#7b68ee',
  mediumspringgreen: '#00fa9a', mediumturquoise: '#48d1cc', mediumvioletred: '#c71585', midnightblue: '#191970',
  mintcream: '#f5fffa', mistyrose: '#ffe4e1', moccasin: '#ffe4b5', navajowhite: '#ffdead', navy: '#000080',
  oldlace: '#fdf5e6', olive: '#808000', olivedrab: '#6b8e23', orange: '#ffa500', orangered: '#ff4500',
  orchid: '#da70d6', palegoldenrod: '#eee8aa', palegreen: '#98fb98', paleturquoise: '#afeeee', palevioletred: '#db7093',
  papayawhip: '#ffefd5', peachpuff: '#ffdab9', peru: '#cd853f', pink: '#ffc0cb', plum: '#dda0dd',
  powderblue: '#b0e0e6', purple: '#800080', rebeccapurple: '#663399', red: '#ff0000', rosybrown: '#bc8f8f',
  royalblue: '#4169e1', saddlebrown: '#8b4513', salmon: '#fa8072', sandybrown: '#f4a460', seagreen: '#2e8b57',
  seashell: '#fff5ee', sienna: '#a0522d', silver: '#c0c0c0', skyblue: '#87ceeb', slateblue: '#6a5acd',
  slategray: '#708090', slategrey: '#708090', snow: '#fffafa', springgreen: '#00ff7f', steelblue: '#4682b4',
  tan: '#d2b48c', teal: '#008080', thistle: '#d8bfd8', tomato: '#ff6347', turquoise: '#40e0d0',
  violet: '#ee82ee', wheat: '#f5deb3', white: '#ffffff', whitesmoke: '#f5f5f5', yellow: '#ffff00', yellowgreen: '#9acd32',
};

export function color(value, fallback = '#000000') {
  if (!value || typeof value !== 'string' || value.startsWith('=')) return fallback;
  const key = value.trim().toLowerCase();
  if (key in NAMED_COLORS) return NAMED_COLORS[key];
  return value.trim();
}

export function styleValue(value, context, fallback) {
  if (typeof value !== 'string' || !value.startsWith('=')) return value ?? fallback;
  const evaluated = evaluateExpression(value, context);
  return evaluated ?? fallback;
}

export function styleColor(value, context, fallback = '#000000') {
  const evaluated = styleValue(value, context, fallback);
  if (evaluated === null || evaluated === undefined || evaluated === '') return fallback;
  return color(String(evaluated), fallback);
}

export function cellTextbox(cell) {
  const items = cell.items || [];
  const values = cell.values || [];
  // A cell can hold several textboxes with different styles — e.g. a LimeGreen "☺" and a Red "☹" where only
  // one evaluates. cellText() draws whichever produced text, so the style must come from that same textbox;
  // taking the first one unconditionally rendered the frown in the smiley's green.
  const produced = items.findIndex((item, index) => item.type === 'Textbox' && String(values[index] ?? '').length > 0);
  if (produced >= 0) return items[produced];
  return items.find((item) => item.type === 'Textbox');
}

// Assembles a cell's display text from its materialized per-item values. Empty values
// (e.g. duplicate-suppressed or hidden items) are dropped, and distinct items are placed on
// separate lines rather than concatenated. Both renderers use this so the PDF and DOCX
// derive identical text; line breaks inside a value are preserved for downstream layout.
export function cellText(cell) {
  return (cell.values || [])
    .map((value) => (value === null || value === undefined ? '' : String(value)))
    .filter((value) => value.length > 0)
    .join('\n');
}
