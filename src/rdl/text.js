const NAMED_ENTITIES = Object.freeze({
  amp: '&', apos: "'", gt: '>', lt: '<', nbsp: ' ', quot: '"',
});

function decodeEntity(entity) {
  if (entity.startsWith('#x') || entity.startsWith('#X')) {
    const codePoint = Number.parseInt(entity.slice(2), 16);
    return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : `&${entity};`;
  }
  if (entity.startsWith('#')) {
    const codePoint = Number.parseInt(entity.slice(1), 10);
    return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : `&${entity};`;
  }
  return NAMED_ENTITIES[entity.toLowerCase()] ?? `&${entity};`;
}

export function decodeHtmlEntities(value) {
  return String(value ?? '').replace(/&(#x[0-9a-f]+|#[0-9]+|[a-z][a-z0-9]+);/gi, (_, entity) => decodeEntity(entity));
}

export function htmlToPlainText(value) {
  const source = String(value ?? '')
    .replace(/<\s*(script|style|iframe|object|svg)\b[^>]*>[\s\S]*?<\s*\/\s*\1\s*>/gi, '')
    .replace(/<\s*br\s*\/?\s*>/gi, '\n')
    .replace(/<\s*li\b[^>]*>/gi, '• ')
    .replace(/<\s*\/\s*(?:p|div|li|tr|h[1-6])\s*>/gi, '\n')
    .replace(/<\s*(?:p|div|tr|h[1-6])\b[^>]*>/gi, '')
    .replace(/<\s*\/\s*(?:ul|ol|table)\s*>/gi, '\n')
    .replace(/<[^>]*>/g, '');
  return decodeHtmlEntities(source)
    .replace(/\r\n?/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

export function renderMarkupText(value, markupType = 'None') {
  return /^html$/i.test(String(markupType || 'None')) ? htmlToPlainText(value) : String(value ?? '');
}

// Normalizes text for rendering: expands tabs to a space and strips other C0 control
// characters that have no glyph. Untouched, an SSRS tab (e.g. the "•\tScope..." bullets in
// the Combined Assurance report) renders as a missing-glyph "tofu" box in the PDF and as a
// stray character in Word. Line breaks (\n) are preserved. Deterministic and pure; never
// logs its input.
export function normalizeDisplayText(value) {
  if (typeof value !== 'string') return value;
  const normalizedNewlines = value.replace(/\r\n?/g, '\n');
  let result = '';
  for (const character of normalizedNewlines) {
    if (character === '\t') { result += ' '; continue; }
    const code = character.codePointAt(0);
    if (code < 0x20 && character !== '\n') continue;
    result += character;
  }
  return result;
}
