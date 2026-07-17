import fs from 'node:fs/promises';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { PNG } from 'pngjs';
import { ServiceError } from '../errors.js';

const execFileAsync = promisify(execFile);

// Rasterizes a PDF buffer to one or more validated PNG buffers via Poppler. Shared by the visual-DOCX
// renderer (multi-page) and the editable-DOCX chart embedder (single page). Temp files live under the
// caller's already-isolated request tempDir and are removed with it.
export async function rasterizePdf(pdfBuffer, config, tempDir, prefix, { dpi = 300, singleFile = false } = {}) {
  const pdfPath = path.join(tempDir, `${prefix}.pdf`);
  const outputPrefix = path.join(tempDir, prefix);
  await fs.writeFile(pdfPath, pdfBuffer, { mode: 0o600 });
  const args = ['-png', '-r', String(dpi), ...(singleFile ? ['-singlefile'] : []), pdfPath, outputPrefix];
  try {
    await execFileAsync(config.pdftoppmPath, args, { timeout: config.renderTimeoutMs, maxBuffer: 1024 * 1024 });
  } catch (error) {
    throw new ServiceError('RENDER_FAILED', 'PDF rasterization failed', 500, { reason: error.code || 'PDFTOPPM_FAILED' });
  }
  const files = singleFile
    ? [`${prefix}.png`]
    : (await fs.readdir(tempDir))
      .filter((name) => new RegExp(`^${prefix}-\\d+\\.png$`).test(name))
      .sort((left, right) => Number(left.match(/(\d+)/)[1]) - Number(right.match(/(\d+)/)[1]));
  const pages = [];
  for (const file of files) {
    const data = await fs.readFile(path.join(tempDir, file));
    const png = PNG.sync.read(data); // Validate the raster before packaging it.
    pages.push({ data, width: png.width, height: png.height });
  }
  return pages;
}
