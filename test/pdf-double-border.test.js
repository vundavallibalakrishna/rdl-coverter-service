// A Double border style must render as two parallel strands (line + gap + line), matching SSRS and the
// editable-DOCX double rule. PDFKit has no double-line primitive, so the PDF renderer draws the two strands
// explicitly. This is a generic construct test on a synthetic RDL: a box with Double borders yields TWO
// horizontal strands per edge, where an otherwise identical Solid box yields one.
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { promisify } from 'node:util';
import { PNG } from 'pngjs';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';

const execFileAsync = promisify(execFile);
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '30000' });
const DPI = 144;

// A single body box (3in wide, 1in tall) at 1in,1in with a thick border on all sides. A vertical scan down
// the box's horizontal centre crosses only the top and bottom edges — never the left/right ones or any text.
const rdlFor = (borderStyle) => `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <ReportSections><ReportSection><Body><ReportItems>
  <Textbox Name="t"><Paragraphs><Paragraph><TextRuns><TextRun><Value> </Value></TextRun></TextRuns></Paragraph></Paragraphs>
    <Top>1in</Top><Left>1in</Left><Width>3in</Width><Height>1in</Height>
    <Style>
      <TopBorder><Color>Black</Color><Style>${borderStyle}</Style><Width>6pt</Width></TopBorder>
      <BottomBorder><Color>Black</Color><Style>${borderStyle}</Style><Width>6pt</Width></BottomBorder>
      <LeftBorder><Color>Black</Color><Style>${borderStyle}</Style><Width>6pt</Width></LeftBorder>
      <RightBorder><Color>Black</Color><Style>${borderStyle}</Style><Width>6pt</Width></RightBorder>
    </Style></Textbox>
 </ReportItems><Height>3in</Height><Style/></Body>
 <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth></Page></ReportSection></ReportSections></Report>`;

// Runs of consecutive dark rows at column x, as [startY, endY] pairs.
function darkRunsInColumn(png, x) {
  const runs = [];
  let start = -1;
  for (let y = 0; y < png.height; y += 1) {
    const idx = (png.width * y + x) << 2;
    const dark = png.data[idx] < 100 && png.data[idx + 1] < 100 && png.data[idx + 2] < 100;
    if (dark && start < 0) start = y;
    else if (!dark && start >= 0) { runs.push([start, y - 1]); start = -1; }
  }
  if (start >= 0) runs.push([start, png.height - 1]);
  return runs;
}

// A double border is two dark strands separated by a small light gap. Detect the signature: two dark runs
// whose gap is small (<= 8px) — the "line + gap + line" pattern — as opposed to a single solid rule, whose
// only close-neighbour would be the far side of the box (a gap much larger than 8px).
function hasDoubleStrandSignature(png, x) {
  const runs = darkRunsInColumn(png, x);
  for (let i = 1; i < runs.length; i += 1) {
    const gap = runs[i][0] - runs[i - 1][1] - 1;
    if (gap >= 1 && gap <= 8) return true;
  }
  return false;
}

async function rasterCentreColumn(borderStyle, context) {
  const model = parseRdl(rdlFor(borderStyle));
  const { buffer } = await renderPdf(model, { outputFileName: borderStyle, parameters: {}, datasets: {} }, config);
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-dbl-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, 'r.pdf');
  const prefix = path.join(tempDir, 'p');
  await fs.writeFile(pdfPath, buffer);
  await execFileAsync(config.pdftoppmPath, ['-f', '1', '-singlefile', '-png', '-r', String(DPI), pdfPath, prefix]);
  const png = PNG.sync.read(await fs.readFile(`${prefix}.png`));
  const centre = Math.round(2.5 * DPI); // box horizontal centre = 1in + 1.5in
  return { png, centre };
}

test('a Solid box edge is a single rule (no close double-strand pattern)', async (context) => {
  const { png, centre } = await rasterCentreColumn('Solid', context);
  assert.equal([-3, -1, 0, 1, 3].some((dx) => hasDoubleStrandSignature(png, centre + dx)), false);
});

test('a Double box edge renders as two strands with a small gap (the double-rule signature)', async (context) => {
  const { png, centre } = await rasterCentreColumn('Double', context);
  assert.equal([-3, -1, 0, 1, 3].some((dx) => hasDoubleStrandSignature(png, centre + dx)), true);
});
