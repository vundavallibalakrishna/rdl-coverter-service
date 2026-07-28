// The render host can lack a font a report declares (e.g. Segoe UI). The PDF path then silently falls back
// (non-strict) or fails closed (strict); a named-font DOCX/Excel defers to the viewer. These tests cover the
// surfacing that makes that condition visible via fontAvailability(), /v1/analyze, and /readyz — generic,
// driven only by the declared family list, never a specific report.
import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { buildApp } from '../src/app.js';
import { fontAvailability, pdfFont, resolveFontFile } from '../src/render/fonts.js';
import { loadConfig } from '../src/config.js';

// Set an env var for the duration of fn(), restoring (or unsetting) it afterwards even on failure.
async function withEnv(name, value, fn) {
  const had = Object.prototype.hasOwnProperty.call(process.env, name);
  const prev = process.env[name];
  process.env[name] = value;
  try { return await fn(); } finally {
    if (had) process.env[name] = prev; else delete process.env[name];
  }
}

const ALL_VARIANTS = ['bold', 'bolditalic', 'italic', 'regular'];

test('fontAvailability marks an absent declared family unavailable with every variant missing', () => {
  const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
  const [entry] = fontAvailability(config, ['Definitely Missing Face']);
  assert.equal(entry.family, 'Definitely Missing Face');
  assert.equal(entry.available, false);
  assert.deepEqual([...entry.missingVariants].sort(), ALL_VARIANTS);
  assert.equal(entry.blocksStrictRender, false); // non-strict: falls back rather than rejecting
});

test('an unavailable declared family blocks strict rendering (fail-closed parity with the PDF embed path)', () => {
  const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'true' });
  const [entry] = fontAvailability(config, ['Definitely Missing Face']);
  assert.equal(entry.available, false);
  assert.equal(entry.blocksStrictRender, true);
});

test('fontAvailability marks a fully-mounted family available with no missing variants', async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-fonts-'));
  // The four Windows Segoe UI filenames the resolver looks for, one per variant.
  for (const file of ['segoeui.ttf', 'segoeuib.ttf', 'segoeuii.ttf', 'segoeuiz.ttf']) {
    await fs.writeFile(path.join(dir, file), 'stub');
  }
  const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_FONT_DIR: dir });
  const [entry] = fontAvailability(config, ['Segoe UI']);
  assert.equal(entry.available, true);
  assert.deepEqual(entry.missingVariants, []);
  await fs.rm(dir, { recursive: true, force: true });
});

test('emoji never falls through to Helvetica when no embedded font covers the glyph', async () => {
  const fontDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-empty-fonts-'));
  const config = loadConfig({
    ...process.env,
    RDL_FONT_DIR: fontDir,
    RDL_STRICT_FONTS: 'false',
    RDL_ALLOW_COMPATIBLE_FONT_FALLBACKS: 'true',
  });
  assert.throws(
    () => pdfFont(config, 'Segoe UI Emoji', false, false, '😊'),
    (error) => error?.code === 'FONT_MISSING' && /emoji/i.test(error.message),
  );
  await fs.rm(fontDir, { recursive: true, force: true });
});

test('compatible emoji fallback is explicit and visible in font availability', async () => {
  const fontDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-emoji-fallback-'));
  const fallback = path.join(fontDir, 'NotoEmoji-Regular.ttf');
  await fs.writeFile(fallback, 'resolver fixture');
  const disabled = loadConfig({
    ...process.env,
    RDL_FONT_DIR: fontDir,
    RDL_STRICT_FONTS: 'true',
    RDL_ALLOW_COMPATIBLE_FONT_FALLBACKS: 'false',
  });
  assert.throws(() => pdfFont(disabled, 'Segoe UI Emoji'), (error) => error?.code === 'FONT_MISSING');

  const enabled = loadConfig({
    ...process.env,
    RDL_FONT_DIR: fontDir,
    RDL_STRICT_FONTS: 'true',
    RDL_ALLOW_COMPATIBLE_FONT_FALLBACKS: 'true',
  });
  assert.equal(pdfFont(enabled, 'Segoe UI Emoji'), fallback);
  const [availability] = fontAvailability(enabled, ['Segoe UI Emoji']);
  assert.equal(availability.available, false);
  assert.equal(availability.compatibleFallback, 'Noto Emoji');
  assert.equal(availability.renderable, true);
  assert.equal(availability.blocksStrictRender, false);
  await fs.rm(fontDir, { recursive: true, force: true });
});

test('a fallback filename alone is insufficient when its font does not cover the emoji', async () => {
  const fontDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-invalid-emoji-fallback-'));
  await fs.writeFile(path.join(fontDir, 'NotoEmoji-Regular.ttf'), 'not a font');
  const config = loadConfig({
    ...process.env,
    RDL_FONT_DIR: fontDir,
    RDL_STRICT_FONTS: 'false',
    RDL_ALLOW_COMPATIBLE_FONT_FALLBACKS: 'true',
  });
  assert.throws(() => pdfFont(config, 'Segoe UI Emoji', false, false, '😡'), (error) => error?.code === 'FONT_MISSING');
  await fs.rm(fontDir, { recursive: true, force: true });
});

test('a font installed in the Windows %SystemRoot%\\Fonts directory resolves without mounting', async () => {
  const sysRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-winroot-'));
  const winFonts = path.join(sysRoot, 'Fonts');
  await fs.mkdir(winFonts);
  await fs.writeFile(path.join(winFonts, 'segoeui.ttf'), 'stub'); // the exact name Windows ships
  const emptyFontDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-emptyfonts-'));
  const resolved = await withEnv('SystemRoot', sysRoot, () => resolveFontFile(emptyFontDir, 'Segoe UI', false, false));
  assert.equal(resolved, path.join(winFonts, 'segoeui.ttf'));
  await fs.rm(sysRoot, { recursive: true, force: true });
  await fs.rm(emptyFontDir, { recursive: true, force: true });
});

test('a font installed in the per-user %LOCALAPPDATA% fonts directory resolves', async () => {
  const localAppData = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-localapp-'));
  const perUserFonts = path.join(localAppData, 'Microsoft', 'Windows', 'Fonts');
  await fs.mkdir(perUserFonts, { recursive: true });
  await fs.writeFile(path.join(perUserFonts, 'arialbd.ttf'), 'stub'); // Arial Bold, Windows filename
  const emptyFontDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-emptyfonts-'));
  const resolved = await withEnv('LOCALAPPDATA', localAppData, () => resolveFontFile(emptyFontDir, 'Arial', true, false));
  assert.equal(resolved, path.join(perUserFonts, 'arialbd.ttf'));
  await fs.rm(localAppData, { recursive: true, force: true });
  await fs.rm(emptyFontDir, { recursive: true, force: true });
});

test('the configured fontDir still wins over a Windows system font of the same family', async () => {
  // Precedence matters: a mounted/deliberately-chosen font must take priority over the OS copy.
  const sysRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-winroot-'));
  await fs.mkdir(path.join(sysRoot, 'Fonts'));
  await fs.writeFile(path.join(sysRoot, 'Fonts', 'segoeui.ttf'), 'system');
  const fontDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-mount-'));
  await fs.writeFile(path.join(fontDir, 'segoeui.ttf'), 'mounted');
  const resolved = await withEnv('SystemRoot', sysRoot, () => resolveFontFile(fontDir, 'Segoe UI', false, false));
  assert.equal(resolved, path.join(fontDir, 'segoeui.ttf')); // fontDir is searched first
  await fs.rm(sysRoot, { recursive: true, force: true });
  await fs.rm(fontDir, { recursive: true, force: true });
});

test('/v1/analyze reports fontAvailability for the fonts the report declares', async (context) => {
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-fa-analyze-'));
  const config = loadConfig({ ...process.env, RDL_TEMP_ROOT: tempRoot, RDL_STRICT_FONTS: 'false' });
  const app = await buildApp({ config, logger: false });
  context.after(async () => { await app.close(); await fs.rm(tempRoot, { recursive: true, force: true }); });
  const fixture = await fs.readFile(new URL('./fixtures/basic.rdl', import.meta.url));
  const response = await app.inject({ method: 'POST', url: '/v1/analyze', payload: { rdlBase64: fixture.toString('base64') } });
  const body = response.json();
  assert.equal(Array.isArray(body.fontAvailability), true);
  // Every declared font is reported with the availability shape.
  assert.deepEqual([...new Set(body.fonts)].sort(), body.fontAvailability.map((f) => f.family).sort());
  for (const entry of body.fontAvailability) {
    assert.equal(typeof entry.available, 'boolean');
    assert.equal(Array.isArray(entry.missingVariants), true);
  }
});

test('/readyz lists a font catalogue including Segoe UI without changing the readiness gate', async (context) => {
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-fa-ready-'));
  const config = loadConfig({ ...process.env, RDL_TEMP_ROOT: tempRoot, RDL_STRICT_FONTS: 'false' });
  const app = await buildApp({ config, logger: false });
  context.after(async () => { await app.close(); await fs.rm(tempRoot, { recursive: true, force: true }); });
  const body = (await app.inject({ method: 'GET', url: '/readyz' })).json();
  assert.equal(typeof body.checks.fonts.ready, 'boolean'); // gate field preserved
  const catalogued = body.checks.fonts.catalogue.map((f) => f.family);
  assert.equal(catalogued.includes('Segoe UI'), true);
  assert.equal(catalogued.includes('Arial'), true);
});
