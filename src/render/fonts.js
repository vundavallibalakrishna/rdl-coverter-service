import fs from 'node:fs';
import path from 'node:path';
import * as fontkit from 'fontkit';
import { ServiceError } from '../errors.js';

const FILE_CANDIDATES = {
  Arial: ['Arial.ttf', 'arial.ttf', 'Arial Regular.ttf'],
  'Arial:bold': ['Arial Bold.ttf', 'arialbd.ttf', 'Arial-Bold.ttf'],
  'Arial:italic': ['Arial Italic.ttf', 'ariali.ttf', 'Arial-Italic.ttf'],
  'Arial:bolditalic': ['Arial Bold Italic.ttf', 'arialbi.ttf', 'Arial-BoldItalic.ttf'],
  'Times New Roman': ['Times New Roman.ttf', 'times.ttf', 'TimesNewRoman.ttf'],
  'Times New Roman:bold': ['Times New Roman Bold.ttf', 'timesbd.ttf', 'TimesNewRoman-Bold.ttf'],
  'Times New Roman:italic': ['Times New Roman Italic.ttf', 'timesi.ttf', 'TimesNewRoman-Italic.ttf'],
  'Times New Roman:bolditalic': ['Times New Roman Bold Italic.ttf', 'timesbi.ttf', 'TimesNewRoman-BoldItalic.ttf'],
  'Segoe UI': ['Segoe UI.ttf', 'segoeui.ttf'],
  'Segoe UI:bold': ['Segoe UI Bold.ttf', 'segoeuib.ttf'],
  'Segoe UI:italic': ['Segoe UI Italic.ttf', 'segoeuii.ttf'],
  'Segoe UI:bolditalic': ['Segoe UI Bold Italic.ttf', 'segoeuiz.ttf'],
  'Segoe UI Emoji': ['Segoe UI Emoji.ttf', 'seguiemj.ttf'],
  'Noto Emoji': ['NotoEmoji-Regular.ttf', 'NotoEmoji.ttf'],
  // Segoe UI Symbol carries the ☺/☹ (U+263A/U+2639) legend glyphs. Windows ships seguisym.ttf — preferred
  // and found first in the licensed font directory. Where it is absent, fall back to another font that
  // actually covers those code points: the base-14 Helvetica fallback does NOT, so the glyphs would
  // otherwise render as garbage rather than smileys.
  'Segoe UI Symbol': ['Segoe UI Symbol.ttf', 'seguisym.ttf', 'Arial Unicode.ttf', 'Apple Symbols.ttf'],

  // ---- Coverage faces -------------------------------------------------------------------------------
  // Reached only through the glyph-coverage ladder below, never by a report declaring them. A report font
  // is used for everything it can draw; these step in only for the characters it cannot (✓ ✗ ☹ →, CJK,
  // Indic, emoji). Liberation Sans/Serif are metric-compatible with Arial/Times New Roman — identical
  // advance widths, so substituting them keeps line breaks and pagination identical to SSRS.
  'Liberation Sans': ['LiberationSans-Regular.ttf', 'Liberation Sans.ttf'],
  'Liberation Sans:bold': ['LiberationSans-Bold.ttf', 'Liberation Sans Bold.ttf'],
  'Liberation Sans:italic': ['LiberationSans-Italic.ttf', 'Liberation Sans Italic.ttf'],
  'Liberation Sans:bolditalic': ['LiberationSans-BoldItalic.ttf', 'Liberation Sans Bold Italic.ttf'],
  'Liberation Serif': ['LiberationSerif-Regular.ttf', 'Liberation Serif.ttf'],
  'Liberation Serif:bold': ['LiberationSerif-Bold.ttf', 'Liberation Serif Bold.ttf'],
  'Liberation Serif:italic': ['LiberationSerif-Italic.ttf', 'Liberation Serif Italic.ttf'],
  'Liberation Serif:bolditalic': ['LiberationSerif-BoldItalic.ttf', 'Liberation Serif Bold Italic.ttf'],
  'Microsoft Sans Serif': ['micross.ttf', 'Microsoft Sans Serif.ttf'],
  'Lucida Sans Unicode': ['l_10646.ttf', 'Lucida Sans Unicode.ttf', 'LucidaSansUnicode.ttf'],
  'Arial Unicode MS': ['ARIALUNI.TTF', 'arialuni.ttf', 'Arial Unicode.ttf', 'Arial Unicode MS.ttf'],
  'DejaVu Sans': ['DejaVuSans.ttf'],
  'DejaVu Sans:bold': ['DejaVuSans-Bold.ttf'],
  'DejaVu Sans:italic': ['DejaVuSans-Oblique.ttf'],
  'DejaVu Sans:bolditalic': ['DejaVuSans-BoldOblique.ttf'],
  // Open faces an operator can simply drop into RDL_FONT_DIR to widen coverage with no code change:
  // Symbols 2 carries the dingbat/geometric ranges on hosts without seguisym.ttf, and the CJK/Devanagari
  // faces cover scripts no Latin family does. Single-face files only — a .ttc collection needs a face
  // name that pdfkit cannot receive through this function's path-only return contract.
  'Noto Sans': ['NotoSans-Regular.ttf'],
  'Noto Sans:bold': ['NotoSans-Bold.ttf'],
  'Noto Sans:italic': ['NotoSans-Italic.ttf'],
  'Noto Sans:bolditalic': ['NotoSans-BoldItalic.ttf'],
  'Noto Sans Symbols 2': ['NotoSansSymbols2-Regular.ttf'],
  'Noto Sans CJK': ['NotoSansCJK-Regular.ttf', 'NotoSansCJKsc-Regular.otf', 'NotoSansSC-Regular.otf', 'NotoSansSC-Regular.ttf'],
  'Noto Sans Devanagari': ['NotoSansDevanagari-Regular.ttf'],
};

const COMPATIBLE_FALLBACKS = Object.freeze({
  // Segoe UI Emoji is commonly a colour font that cannot be subset by every PDF backend. Noto Emoji is
  // an outline font with compatible Unicode semantics; the ordinary UI families cover any Latin tokens
  // that share the same RDL run (for example "Low 😊").
  'Segoe UI Emoji': ['Noto Emoji', 'Segoe UI', 'Arial'],
});

// Ordered stand-ins tried when the declared family IS installed but has no glyph for some character in a
// text run — Arial has no ✓ (U+2713), ✗ (U+2717) or ☹ (U+2639), and no Latin family covers CJK or Indic.
// Distinct from COMPATIBLE_FALLBACKS above, which answers a different question ("the family's file is
// absent — may I substitute a whole different family?") and stays gated behind an explicit opt-in because
// it changes the metrics of every run. A coverage substitution touches only runs the declared font cannot
// draw at all, where the alternative is a failed export, so it needs no opt-in.
//
// Metric-compatible faces come first: Liberation Sans has Arial's advance widths, so when it covers the
// run the page breaks are identical to SSRS. Broad-coverage faces follow.
const COVERAGE_PREFERENCES = Object.freeze({
  Arial: ['Liberation Sans', 'Microsoft Sans Serif', 'Segoe UI', 'Segoe UI Symbol', 'Lucida Sans Unicode', 'Arial Unicode MS'],
  'Times New Roman': ['Liberation Serif', 'Lucida Sans Unicode', 'Arial Unicode MS', 'Segoe UI Symbol', 'Microsoft Sans Serif'],
  'Segoe UI': ['Segoe UI Symbol', 'Arial', 'Liberation Sans', 'Microsoft Sans Serif', 'Lucida Sans Unicode', 'Arial Unicode MS'],
  'Segoe UI Symbol': ['Arial Unicode MS', 'Lucida Sans Unicode', 'Noto Sans Symbols 2', 'Segoe UI'],
});

// Tried after the per-family preference, for every family including ones this service does not otherwise
// know. Ordered widest-coverage-first within each script group.
const GENERIC_COVERAGE = Object.freeze([
  'Segoe UI Symbol', 'Noto Sans Symbols 2', 'Arial Unicode MS', 'Lucida Sans Unicode', 'Microsoft Sans Serif',
  'Liberation Sans', 'DejaVu Sans', 'Noto Sans', 'Noto Sans CJK', 'Noto Sans Devanagari',
  'Segoe UI', 'Arial', 'Times New Roman',
]);

const openedFonts = new Map();

// pdfFont runs once per drawn text run, and every miss now walks a ladder of families. Cache the coverage
// verdict per (file, text) so a repeated column value or header costs one fontkit layout, not one per row.
const COVERAGE_CACHE_LIMIT = 20_000;
const coverageCache = new Map();

// Coverage substitutions performed while rendering, so a silent-but-visible font swap is reportable rather
// than invisible. A render worker is forked per request and exits with it (see worker/runner.js), so this
// is request-scoped state, not a cross-request leak.
const substitutions = new Map();

function styleKey(family, bold, italic) {
  return `${family}${bold || italic ? `:${bold ? 'bold' : ''}${italic ? 'italic' : ''}` : ''}`;
}

function systemDirectories(fontDir) {
  const directories = [fontDir];
  // Windows font locations: machine-wide %SystemRoot%\Fonts (e.g. C:\Windows\Fonts) and the per-user
  // %LOCALAPPDATA%\Microsoft\Windows\Fonts. These env vars are set only on Windows, so on Linux/macOS the
  // branches are skipped entirely (and guarding on the variable also keeps path.join from ever receiving
  // undefined). Windows installs Segoe UI / Arial under exactly the filenames FILE_CANDIDATES expects
  // (segoeui.ttf, arial.ttf, ...), so a normally-installed font now resolves here with no mount required.
  if (process.env.SystemRoot) directories.push(path.join(process.env.SystemRoot, 'Fonts'));
  if (process.env.LOCALAPPDATA) directories.push(path.join(process.env.LOCALAPPDATA, 'Microsoft', 'Windows', 'Fonts'));
  // macOS and Linux system font locations. Non-existent entries are harmlessly skipped by the existsSync
  // check in resolveFontFile, so listing all platforms' paths together is safe.
  directories.push(
    '/Library/Fonts',
    '/System/Library/Fonts',
    '/System/Library/Fonts/Supplemental',
    '/usr/share/fonts/truetype/msttcorefonts',
    // Distro packages install the coverage faces (Liberation, DejaVu, Noto) into per-family
    // subdirectories that a flat probe of the parent would miss. These come after msttcorefonts, so the
    // licensed originals still win for the families that declare them — this only adds places to find a
    // stand-in for characters those originals cannot draw.
    '/usr/share/fonts/truetype/liberation',
    '/usr/share/fonts/truetype/liberation2',
    '/usr/share/fonts/truetype/dejavu',
    '/usr/share/fonts/truetype/noto',
    '/usr/share/fonts/opentype/noto',
    '/usr/share/fonts/truetype',
    '/usr/share/fonts',
    '/usr/local/share/fonts',
  );
  return directories;
}

// A coverage miss now walks a ladder of families across a dozen directories. Unmemoised that is thousands
// of existsSync calls per render, so cache hits and misses alike. The directory list is part of the key
// because it is derived from env vars the tests vary.
const resolutionCache = new Map();

export function resolveFontFile(fontDir, family, bold = false, italic = false) {
  const exact = styleKey(family, bold, italic);
  const candidates = FILE_CANDIDATES[exact] || FILE_CANDIDATES[family] || [];
  const directories = systemDirectories(fontDir);
  const cacheKey = `${directories.join('|')} ${exact}`;
  if (resolutionCache.has(cacheKey)) return resolutionCache.get(cacheKey);

  let resolved = null;
  for (const directory of directories) {
    for (const candidate of candidates) {
      const filePath = path.join(directory, candidate);
      if (fs.existsSync(filePath)) { resolved = filePath; break; }
    }
    if (resolved) break;
  }
  resolutionCache.set(cacheKey, resolved);
  return resolved;
}

function openedFont(file) {
  if (!openedFonts.has(file)) openedFonts.set(file, fontkit.openSync(file));
  return openedFonts.get(file);
}

const USABLE = Object.freeze({ covers: true, embeddable: true });
const UNUSABLE = Object.freeze({ covers: false, embeddable: false });

/**
 * What `file` can do with `text`: draw it (`covers`) and survive being embedded in the PDF (`embeddable`).
 *
 * Both matter, and coverage alone is not enough. PDFKit subsets an embedded TrueType font by decoding every
 * glyph the document used. Colour fonts — COLR/CBDT/sbix, of which Segoe UI Emoji is one — return glyph
 * objects with no `_decode`, so embedding throws `glyph._decode is not a function`. That happens at the end
 * of the render, long after the run chose its font, and surfaces as an opaque RENDER_FAILED 500. fontkit
 * reports coverage for those same glyphs perfectly happily, so a font must be checked for both.
 */
function fontCapability(file, text) {
  if (!text) return USABLE;
  const key = `${file}\u0000${text}`;
  const cached = coverageCache.get(key);
  if (cached !== undefined) return cached;
  let capability;
  try {
    // Whole-string layout rather than per-code-point lookup: it is what actually gets drawn, so ligatures,
    // combining marks and multi-code-point emoji sequences are judged as the clusters they render as.
    const glyphs = openedFont(file).layout(String(text)).glyphs;
    capability = {
      covers: glyphs.every((glyph) => glyph.id !== 0),
      embeddable: glyphs.every((glyph) => typeof glyph._decode === 'function'),
    };
  } catch {
    // Unreadable, truncated or collection (.ttc) file — not usable for this text either way.
    capability = UNUSABLE;
  }
  if (coverageCache.size >= COVERAGE_CACHE_LIMIT) coverageCache.clear();
  coverageCache.set(key, capability);
  return capability;
}

// Draws the text AND can be embedded — the bar a font must clear to be chosen for a run.
function fontUsableFor(file, text) {
  const capability = fontCapability(file, text);
  return capability.covers && capability.embeddable;
}

// Embeddable, whether or not it covers everything. The bar for a last-resort font, where an uncovered
// character becoming .notdef is acceptable but a mid-render embedding crash is not.
function fontEmbeddableFor(file, text) {
  return fontCapability(file, text).embeddable;
}

function recordSubstitution(requested, substituted, reason) {
  const key = `${requested}\u0000${substituted}\u0000${reason}`;
  substitutions.set(key, { requested, substituted, reason, runs: (substitutions.get(key)?.runs || 0) + 1 });
}

/**
 * Coverage substitutions made since the last call, most-used first, and clears the record. Called once per
 * render so the result can travel back to the caller as render metadata.
 */
export function takeFontSubstitutions() {
  const taken = [...substitutions.values()].sort((left, right) => right.runs - left.runs);
  substitutions.clear();
  return taken;
}

function coverageCandidates(family, text) {
  // Text faces before emoji faces. Several characters that report authors use as ordinary text — ☹ ✓ ✗ ★ —
  // are Extended_Pictographic, so leading with an emoji face would set an entire Latin run in it. The
  // symbol faces near the front of the preference lists cover those; a genuine pictograph finds nothing
  // there and falls through to the emoji faces below.
  const candidates = [...(COVERAGE_PREFERENCES[family] || [])];
  if (containsEmoji(text)) candidates.push('Segoe UI Emoji', 'Noto Emoji');
  candidates.push(...GENERIC_COVERAGE);
  return [...new Set(candidates)].filter((candidate) => candidate !== family);
}

// Walks the coverage ladder and returns the first installed font `accept` approves of.
function ladderFont(config, family, bold, italic, text, accept) {
  if (!text) return null;
  for (const candidateFamily of coverageCandidates(family, text)) {
    const file = resolveFontFile(config.fontDir, candidateFamily, bold, italic);
    if (file && accept(file, text)) return { family: candidateFamily, file };
  }
  return null;
}

/**
 * The first installed font that can draw every character of `text` and be embedded, for a declared family
 * that cannot.
 *
 * Full coverage is required rather than best-effort: this returns one font for the whole run, so a face
 * that covers the missing character but not the surrounding text (Noto Emoji has no Latin) would trade one
 * unrenderable character for a whole unrenderable run. When nothing covers the run the caller keeps the
 * declared font, and only the characters it lacks come out as .notdef.
 */
function coveringFont(config, family, bold, italic, text) {
  return ladderFont(config, family, bold, italic, text, fontUsableFor);
}

// Used only when the DECLARED font cannot be embedded at all (it is a colour font). Coverage is no longer
// the question — anything embeddable beats a render that dies during PDF subsetting.
function embeddableFont(config, family, bold, italic, text) {
  return ladderFont(config, family, bold, italic, text, fontEmbeddableFor);
}

function containsEmoji(text) {
  return /\p{Extended_Pictographic}/u.test(String(text || ''));
}

function fallbackFamilies(family, text = '') {
  const families = [...(COMPATIBLE_FALLBACKS[family] || [])];
  // A pictographic character in an otherwise ordinary run still needs a glyph-capable font. This is
  // deliberately coverage-checked below; no family is selected merely because its filename exists.
  if (containsEmoji(text)) families.unshift('Noto Emoji', 'Segoe UI Emoji', 'Segoe UI Symbol');
  return [...new Set(families.filter((candidate) => candidate !== family))];
}

function compatibleFont(config, family, bold, italic, text = '') {
  if (!config.allowCompatibleFontFallbacks) return null;
  for (const candidateFamily of fallbackFamilies(family, text)) {
    const file = resolveFontFile(config.fontDir, candidateFamily, bold, italic);
    // Embeddability is checked here too: this list leads with the emoji faces, and a colour font that
    // covers the text would still abort the render when PDFKit came to subset it.
    if (file && fontUsableFor(file, text)) return { family: candidateFamily, file };
  }
  return null;
}

/**
 * The font file (or base-14 font name) to draw `text` with.
 *
 * Two different conditions used to share one outcome here, and conflating them failed exports that had no
 * business failing:
 *
 *  - The declared family has no file on this host. Substituting changes the advance widths of every run and
 *    therefore every page break, so strict mode still fails closed and a whole-family substitution stays
 *    behind RDL_ALLOW_COMPATIBLE_FONT_FALLBACKS.
 *  - The declared family IS installed but has no glyph for some character in this run — Arial has no ✓,
 *    ✗ or ☹, and no Latin family covers CJK or Indic. This is a property of a few characters, not of the
 *    document, and the declared font still draws every other run at its own metrics. One uncovered
 *    character must not cost the whole export, so the run is drawn in a font that can display it. No opt-in
 *    gates this: the alternative is a 503 for a report that renders fine everywhere else.
 */
export function pdfFont(config, family, bold = false, italic = false, text = '') {
  const supported = ['Arial', 'Times New Roman', 'Segoe UI', 'Segoe UI Emoji', 'Segoe UI Symbol', 'Noto Emoji'];
  const normalized = supported.includes(family) ? family : null;
  if (!normalized) {
    // An unknown family is still worth covering: a real installed font beats both a hard failure and the
    // base-14 approximation below.
    const covering = coveringFont(config, family, bold, italic, text);
    if (covering) {
      recordSubstitution(family, covering.family, 'family-unsupported');
      return covering.file;
    }
    if (config.strictFonts) throw new ServiceError('FONT_MISSING', `Required font is unavailable: ${family}`, 503);
    if (containsEmoji(text)) {
      const fallback = compatibleFont(config, family, bold, italic, text);
      if (fallback) return fallback.file;
      throw new ServiceError('FONT_MISSING', `No embedded font covers the required emoji for: ${family}`, 503);
    }
    return bold ? (italic ? 'Helvetica-BoldOblique' : 'Helvetica-Bold') : (italic ? 'Helvetica-Oblique' : 'Helvetica');
  }

  const file = resolveFontFile(config.fontDir, normalized, bold, italic);
  if (file && fontUsableFor(file, text)) return file;

  if (file) {
    // The declared font is present and simply cannot draw this text. Never fails: worst case the run keeps
    // the declared font and its uncovered characters come out as .notdef, exactly as SSRS renders them.
    const covering = coveringFont(config, normalized, bold, italic, text);
    if (covering) {
      recordSubstitution(normalized, covering.family, 'glyph-coverage');
      return covering.file;
    }
    // Returning the declared TrueType file is safe even for emoji: an embedded font renders an uncovered
    // code point as .notdef, whereas the base-14 path below would write corrupt surrogate bytes.
    if (fontEmbeddableFor(file, text)) {
      if (text) recordSubstitution(normalized, normalized, 'no-covering-font');
      return file;
    }
    // The declared font is itself a colour font, so embedding it would abort the render. Anything the
    // renderer can actually embed is better, even if some characters fall back to .notdef.
    const embeddable = embeddableFont(config, normalized, bold, italic, text);
    if (embeddable) {
      recordSubstitution(normalized, embeddable.family, 'not-embeddable');
      return embeddable.file;
    }
    // Nothing embeddable at all — fall through to the base-14 / fail-closed tail below.
  }

  // The declared family has no file at all — whole-family substitution, unchanged fail-closed semantics.
  const fallback = compatibleFont(config, normalized, bold, italic, text);
  if (fallback) return fallback.file;
  if (config.strictFonts) throw new ServiceError('FONT_MISSING', `Required font is unavailable: ${normalized}`, 503);
  // Base-14 fonts cannot represent supplementary-plane emoji. Failing here prevents the corrupt surrogate
  // bytes that Helvetica otherwise writes while keeping legacy non-strict fallback for ordinary text.
  if (containsEmoji(text)) {
    throw new ServiceError('FONT_MISSING', `No embedded font covers the required emoji for: ${normalized}`, 503);
  }
  if (normalized === 'Times New Roman') return bold ? (italic ? 'Times-BoldItalic' : 'Times-Bold') : (italic ? 'Times-Italic' : 'Times-Roman');
  return bold ? (italic ? 'Helvetica-BoldOblique' : 'Helvetica-Bold') : (italic ? 'Helvetica-Oblique' : 'Helvetica');
}

export function checkFonts(config, families = ['Arial', 'Times New Roman']) {
  const missing = families.flatMap((family) => [
    [false, false, 'regular'], [true, false, 'bold'], [false, true, 'italic'], [true, true, 'bolditalic'],
  ].filter(([bold, italic]) => !resolveFontFile(config.fontDir, family, bold, italic)).map(([, , style]) => `${family}:${style}`));
  return { ready: !config.strictFonts || missing.length === 0, missing, strict: config.strictFonts };
}

const VARIANTS = [[false, false, 'regular'], [true, false, 'bold'], [false, true, 'italic'], [true, true, 'bolditalic']];

// Per-family availability of the actual licensed font files on THIS server, for the families a report
// declares. `available` means every variant resolves to a file we could embed in a PDF; when it is false
// the PDF silently falls back (Helvetica/Times base-14) in non-strict mode or fails closed in strict mode,
// and a named-font DOCX/Excel defers the family to whatever the opening application substitutes — so this
// surfaces the exact "the declared font is not on the render host" condition that otherwise passes silently.
// Generic: driven only by the RDL's declared family list, nothing report-specific.
export function fontAvailability(config, families = []) {
  return [...new Set(families)].map((family) => {
    const missingVariants = VARIANTS
      .filter(([bold, italic]) => !resolveFontFile(config.fontDir, family, bold, italic))
      .map(([, , style]) => style);
    const fallback = compatibleFont(config, family, false, false);
    return {
      family,
      available: missingVariants.length === 0,
      missingVariants,
      compatibleFallback: fallback?.family || null,
      renderable: missingVariants.length === 0 || Boolean(fallback),
      // strict render would reject a report that consumes an unavailable family (PDF must embed it).
      blocksStrictRender: config.strictFonts && missingVariants.length > 0 && !fallback,
    };
  });
}
