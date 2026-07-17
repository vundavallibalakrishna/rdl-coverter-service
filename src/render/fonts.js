import fs from 'node:fs';
import path from 'node:path';
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
  // Segoe UI Symbol carries the ☺/☹ (U+263A/U+2639) legend glyphs. Windows ships seguisym.ttf — preferred
  // and found first in the licensed font directory. Where it is absent, fall back to another font that
  // actually covers those code points: the base-14 Helvetica fallback does NOT, so the glyphs would
  // otherwise render as garbage rather than smileys.
  'Segoe UI Symbol': ['Segoe UI Symbol.ttf', 'seguisym.ttf', 'Arial Unicode.ttf', 'Apple Symbols.ttf'],
};

function styleKey(family, bold, italic) {
  return `${family}${bold || italic ? `:${bold ? 'bold' : ''}${italic ? 'italic' : ''}` : ''}`;
}

function systemDirectories(fontDir) {
  return [
    fontDir,
    '/Library/Fonts',
    '/System/Library/Fonts',
    '/System/Library/Fonts/Supplemental',
    '/usr/share/fonts/truetype/msttcorefonts',
    '/usr/share/fonts/truetype/msttcorefonts',
  ];
}

export function resolveFontFile(fontDir, family, bold = false, italic = false) {
  const exact = styleKey(family, bold, italic);
  const candidates = FILE_CANDIDATES[exact] || FILE_CANDIDATES[family] || [];
  for (const directory of systemDirectories(fontDir)) {
    for (const candidate of candidates) {
      const filePath = path.join(directory, candidate);
      if (fs.existsSync(filePath)) return filePath;
    }
  }
  return null;
}

export function pdfFont(config, family, bold = false, italic = false) {
  const normalized = ['Arial', 'Times New Roman', 'Segoe UI', 'Segoe UI Symbol'].includes(family) ? family : null;
  if (!normalized) {
    if (config.strictFonts) throw new ServiceError('FONT_MISSING', `Required font is unavailable: ${family}`, 503);
    return bold ? (italic ? 'Helvetica-BoldOblique' : 'Helvetica-Bold') : (italic ? 'Helvetica-Oblique' : 'Helvetica');
  }
  const file = resolveFontFile(config.fontDir, normalized, bold, italic);
  if (file) return file;
  if (config.strictFonts) throw new ServiceError('FONT_MISSING', `Required font is unavailable: ${normalized}`, 503);
  if (normalized === 'Times New Roman') return bold ? (italic ? 'Times-BoldItalic' : 'Times-Bold') : (italic ? 'Times-Italic' : 'Times-Roman');
  return bold ? (italic ? 'Helvetica-BoldOblique' : 'Helvetica-Bold') : (italic ? 'Helvetica-Oblique' : 'Helvetica');
}

export function checkFonts(config, families = ['Arial', 'Times New Roman']) {
  const missing = families.flatMap((family) => [
    [false, false, 'regular'], [true, false, 'bold'], [false, true, 'italic'], [true, true, 'bolditalic'],
  ].filter(([bold, italic]) => !resolveFontFile(config.fontDir, family, bold, italic)).map(([, , style]) => `${family}:${style}`));
  return { ready: !config.strictFonts || missing.length === 0, missing, strict: config.strictFonts };
}
