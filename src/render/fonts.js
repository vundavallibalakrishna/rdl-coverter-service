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
  );
  return directories;
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
    return {
      family,
      available: missingVariants.length === 0,
      missingVariants,
      // strict render would reject a report that consumes an unavailable family (PDF must embed it).
      blocksStrictRender: config.strictFonts && missingVariants.length > 0,
    };
  });
}
