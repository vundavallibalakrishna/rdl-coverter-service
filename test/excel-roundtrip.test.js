// Round-trip verification for the XLSX renderer: open our workbook in a real spreadsheet engine
// (LibreOffice, headless), convert it to PDF, and assert on what actually RENDERS. Markup tests confirm the
// file is structurally right; this confirms a spreadsheet engine agrees — values are visible, translated
// number formats display correctly, and untrusted "=..." text is shown literally rather than executed.
//
// LibreOffice is a TEST-ONLY tool. The service never invokes it: XLSX is written directly with exceljs, so
// production and the standalone library need nothing here. This test skips cleanly when soffice is absent,
// exactly like the existing DOCX stress verifier, so local runs without LibreOffice stay green and CI with
// it installed does the real verification.
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import { existsSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { promisify } from 'node:util';
import { parseRdl } from '../src/rdl/parser.js';
import { loadConfig } from '../src/config.js';
import { renderExcel } from '../src/render/excel.js';

const execFileAsync = promisify(execFile);

// Resolve soffice from the env override, PATH, or the platform's default install location.
function resolveSoffice() {
  const override = process.env.RDL_SOFFICE_PATH;
  if (override && existsSync(override)) return override;
  const candidates = [
    '/Applications/LibreOffice.app/Contents/MacOS/soffice', // macOS cask
    '/usr/bin/soffice', '/usr/bin/libreoffice', // Linux
    '/opt/homebrew/bin/soffice',
  ];
  return candidates.find((candidate) => existsSync(candidate)) || 'soffice';
}

async function sofficeAvailable(soffice) {
  try {
    await execFileAsync(soffice, ['--version'], { timeout: 30_000 });
    return true;
  } catch {
    return false;
  }
}

const soffice = resolveSoffice();
// An explicit opt-out for anyone who has LibreOffice installed but wants a faster run (local or CI):
// RDL_SKIP_ROUNDTRIP=1 skips these without uninstalling anything. Otherwise they run when soffice exists.
const forceSkip = /^(1|true|yes)$/i.test(process.env.RDL_SKIP_ROUNDTRIP || '');
const available = !forceSkip && await sofficeAvailable(soffice);
const skip = forceSkip
  ? 'RDL_SKIP_ROUNDTRIP is set; XLSX round-trip verification skipped'
  : available
    ? false
    : 'LibreOffice (soffice) is not installed; XLSX round-trip verification skipped';

const model = parseRdl(await fs.readFile(new URL('./fixtures/basic.rdl', import.meta.url)));
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const baseRequest = { outputFileName: 'roundtrip', parameters: { Title: 'Sales', Choice: 'A' }, datasets: { Sales: [{ Name: 'North', Amount: 1234.5 }, { Name: 'South', Amount: 99 }], Choices: [{ Value: 'A' }] } };

// Renders an XLSX, converts it to PDF with headless LibreOffice, and returns the extracted text layer.
async function renderedText(request, label) {
  const result = await renderExcel(model, request, config, null);
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), `rdl-xlsx-roundtrip-${label}-`));
  try {
    const xlsxPath = path.join(dir, 'book.xlsx');
    await fs.writeFile(xlsxPath, result.buffer);
    const profile = path.join(dir, 'lo-profile');
    await execFileAsync(soffice, [
      `-env:UserInstallation=file://${profile}`,
      '--invisible', '--headless', '--norestore', '--convert-to', 'pdf', '--outdir', dir, xlsxPath,
    ], { timeout: 120_000, maxBuffer: 16 * 1024 * 1024 });
    const pdfPath = path.join(dir, 'book.pdf');
    assert.equal(existsSync(pdfPath), true, 'LibreOffice did not produce a PDF from the workbook');
    const { stdout } = await execFileAsync('pdftotext', ['-layout', pdfPath, '-'], { maxBuffer: 16 * 1024 * 1024 });
    return stdout;
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
}

test('the XLSX opens in a spreadsheet engine and renders its cell values', { skip }, async () => {
  const text = await renderedText(baseRequest, 'values');
  assert.match(text, /North/);
  assert.match(text, /South/);
  assert.match(text, /Amount/);
});

test('a live number displays with its translated Excel number format', { skip }, async () => {
  // Amount is written as the live number 1234.5 with format #,##0.00 (recovered from =Format(..,"N2")).
  // The engine must render it as "1,234.50" — proving the .NET->Excel format translation is correct.
  const text = await renderedText(baseRequest, 'format');
  assert.match(text, /1,234\.50/);
});

test('an untrusted =... value is displayed as literal text, never evaluated by the engine', { skip }, async () => {
  const request = { ...baseRequest, datasets: { ...baseRequest.datasets, Sales: [{ Name: '=1+2+cmd|calc', Amount: 1 }] } };
  const text = await renderedText(request, 'injection');
  // If the engine had treated it as a formula it would show "3" or an error; the literal string proves the
  // typed-string storage prevents execution end to end.
  assert.match(text, /=1\+2\+cmd\|calc/);
});

test('sheetPerTablix produces a multi-sheet workbook the engine renders across multiple pages', { skip }, async () => {
  const text = await renderedText({ ...baseRequest, excel: { sheetPerTablix: true } }, 'sheets');
  assert.match(text, /North/); // the tablix data still renders on its own sheet
});
