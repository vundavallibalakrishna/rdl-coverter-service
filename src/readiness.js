import fs from 'node:fs/promises';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { checkFonts, fontAvailability } from './render/fonts.js';
import PDFDocument from 'pdfkit';

const execFileAsync = promisify(execFile);

export async function readiness(config) {
  const checks = {};
  try {
    await fs.mkdir(config.tempRoot, { recursive: true, mode: 0o700 });
    await fs.chmod(config.tempRoot, 0o700);
    const probe = path.join(config.tempRoot, `.probe-${process.pid}`);
    await fs.writeFile(probe, 'ready', { mode: 0o600 });
    await fs.rm(probe, { force: true });
    checks.temporaryStorage = { ready: true };
  } catch {
    checks.temporaryStorage = { ready: false };
  }
  // `ready` still gates only on the always-required base families (Arial, Times New Roman). The catalogue is
  // informational: it shows which other licensed faces the host actually has, so ops can see at a glance that
  // an optional-but-report-consumed family such as Segoe UI is present before a report that needs it arrives.
  checks.fonts = {
    ...checkFonts(config),
    catalogue: fontAvailability(config, [
      'Arial', 'Times New Roman', 'Segoe UI', 'Segoe UI Emoji', 'Segoe UI Symbol', 'Noto Emoji',
    ]),
  };
  try {
    await execFileAsync(config.pdftoppmPath, ['-v'], { timeout: 5_000, maxBuffer: 64 * 1024 });
    checks.pdftoppm = { ready: true };
  } catch (error) {
    // pdftoppm prints its version to stderr and may return either 0 or 1 depending on build.
    checks.pdftoppm = { ready: Boolean(error.stderr && /pdftoppm/i.test(error.stderr)) };
  }
  try {
    const pdfReady = await new Promise((resolve) => {
      const doc = new PDFDocument({ autoFirstPage: false });
      const chunks = [];
      doc.on('data', (chunk) => chunks.push(chunk));
      doc.once('error', () => resolve(false));
      doc.once('end', () => resolve(Buffer.concat(chunks).subarray(0, 4).toString() === '%PDF'));
      doc.addPage().font('Helvetica').text('readiness');
      doc.end();
    });
    checks.pdfRenderer = { ready: pdfReady };
  } catch {
    checks.pdfRenderer = { ready: false };
  }
  return { ready: Object.values(checks).every((check) => check.ready), checks };
}
