export function asArray(value) {
  if (value === undefined || value === null) return [];
  return Array.isArray(value) ? value : [value];
}

export function textValue(value, fallback = '') {
  if (value === undefined || value === null) return fallback;
  const text = typeof value === 'object' && '#text' in value ? String(value['#text']) : String(value);
  return text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

export function childEntries(container, names) {
  if (!container || typeof container !== 'object') return [];
  return names.flatMap((name) => asArray(container[name]).map((value) => [name, value]));
}

export function parseBoolean(value, fallback = false) {
  if (value === undefined || value === null || value === '') return fallback;
  return String(value).toLowerCase() === 'true';
}

// The report-item types a TablixCell can actually be DRAWN with. The capability catalogue classifies
// element NAMES, so an element can be SUPPORTED at body level yet unhandled inside a cell — that gap is how
// Rectangle-wrapped cells silently rendered blank. Everything a cell renderer consumes must be listed here,
// and anything else is refused (analyze reports TablixCellContent:<Type>; materialization throws) rather
// than being silently dropped to an empty value.
export const RENDERABLE_CELL_ITEMS = new Set(['Textbox']);

// A cell often wraps its content in a Rectangle used purely as a container
// (CellContents > Rectangle > ReportItems > Textbox…). Flatten those away so the real content is seen.
export function flattenCellItems(items) {
  return (items || []).flatMap((item) => (item.type === 'Rectangle' ? flattenCellItems(item.items) : [item]));
}

export function firstDefined(...values) {
  return values.find((value) => value !== undefined && value !== null);
}
