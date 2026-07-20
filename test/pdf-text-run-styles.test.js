import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { promisify } from 'node:util';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';

const execFileAsync = promisify(execFile);
const rdl = `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <Body><ReportItems>
    <Textbox Name="AcrossParagraphs"><CanGrow>true</CanGrow><Paragraphs>
      <Paragraph><TextRuns><TextRun><Value>DIVISION</Value><Style><FontFamily>Arial</FontFamily><FontSize>10pt</FontSize><FontWeight>Bold</FontWeight></Style></TextRun></TextRuns></Paragraph>
      <Paragraph><TextRuns><TextRun><Value>PROCESS</Value><Style><FontFamily>Arial</FontFamily><FontSize>10pt</FontSize><FontWeight>Normal</FontWeight></Style></TextRun></TextRuns></Paragraph>
    </Paragraphs><Top>0.1in</Top><Left>0.1in</Left><Width>2in</Width><Height>0.5in</Height><Style/></Textbox>
    <Textbox Name="WithinParagraph"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns>
      <TextRun><Value>NORMAL </Value><Style><FontFamily>Arial</FontFamily><FontSize>10pt</FontSize><FontWeight>Normal</FontWeight></Style></TextRun>
      <TextRun><Value>BOLD</Value><Style><FontFamily>Arial</FontFamily><FontSize>10pt</FontSize><FontWeight>Bold</FontWeight></Style></TextRun>
    </TextRuns></Paragraph></Paragraphs><Top>0.7in</Top><Left>0.1in</Left><Width>2in</Width><Height>0.3in</Height><Style/></Textbox>
    <Textbox Name="EmptyFirst"><CanGrow>true</CanGrow><Paragraphs>
      <Paragraph><TextRuns><TextRun><Value/><Style><FontFamily>Arial</FontFamily><FontSize>6pt</FontSize><FontWeight>Normal</FontWeight><Color>Black</Color></Style></TextRun></TextRuns></Paragraph>
      <Paragraph><TextRuns><TextRun><Value>LATE BOLD</Value><Style><FontFamily>Arial</FontFamily><FontSize>10pt</FontSize><FontWeight>Bold</FontWeight><Color>White</Color></Style></TextRun></TextRuns></Paragraph>
    </Paragraphs><Top>1.1in</Top><Left>0.1in</Left><Width>2in</Width><Height>0.5in</Height><Style><BackgroundColor>Blue</BackgroundColor></Style></Textbox>
  </ReportItems><Height>2in</Height></Body>
  <Width>3in</Width><Page><PageWidth>3in</PageWidth><PageHeight>2in</PageHeight><TopMargin>0in</TopMargin><BottomMargin>0in</BottomMargin><LeftMargin>0in</LeftMargin><RightMargin>0in</RightMargin></Page>
</Report>`;

test('preserves and renders font styles at each RDL text-run boundary', async (context) => {
  const model = parseRdl(rdl);
  const across = model.body.items.find((item) => item.name === 'AcrossParagraphs');
  const within = model.body.items.find((item) => item.name === 'WithinParagraph');
  const emptyFirst = model.body.items.find((item) => item.name === 'EmptyFirst');

  assert.deepEqual(across.paragraphs.map((paragraph) => paragraph[0].style.fontWeight), ['Bold', 'Normal']);
  assert.deepEqual(within.paragraphs[0].map((run) => run.style.fontWeight), ['Normal', 'Bold']);
  assert.equal(emptyFirst.paragraphs[1][0].style.color, 'White');
  assert.equal(emptyFirst.paragraphs[1][0].style.fontWeight, 'Bold');

  const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
  const rendered = await renderPdf(model, {}, config);
  const tempRoot = path.resolve('tmp');
  await fs.mkdir(tempRoot, { recursive: true });
  const pdfPath = path.join(tempRoot, `pdf-text-run-styles-${process.pid}.pdf`);
  context.after(() => fs.rm(pdfPath, { force: true }));
  await fs.writeFile(pdfPath, rendered.buffer, { mode: 0o600 });

  const [{ stdout: xml }, { stdout: fonts }] = await Promise.all([
    execFileAsync('pdftohtml', ['-xml', '-hidden', '-nodrm', '-i', '-stdout', pdfPath]),
    execFileAsync('pdffonts', [pdfPath]),
  ]);
  assert.match(xml, /<b>DIVISION<\/b>/);
  assert.match(xml, />PROCESS<\/text>/);
  assert.doesNotMatch(xml, /<b>PROCESS<\/b>/);
  assert.match(xml, />NORMAL <\/text>/);
  assert.match(xml, /<b>BOLD<\/b>/);
  assert.match(xml, /<fontspec id="2"[^>]*color="#ffffff"/);
  assert.match(xml, /font="2"><b>LATE BOLD<\/b>/);
  assert.match(fonts, /Arial-BoldMT\s+CID TrueType\s+Identity-H\s+yes\s+yes\s+yes/);
  assert.match(fonts, /ArialMT\s+CID TrueType\s+Identity-H\s+yes\s+yes\s+yes/);
});
