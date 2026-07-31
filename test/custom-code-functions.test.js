import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import test from 'node:test';
import { promisify } from 'node:util';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { analyzeRdl, parseRdl } from '../src/rdl/parser.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

const execFileAsync = promisify(execFile);
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

function customFunctionRdl() {
  return `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields>
  <Field Name="Key"><DataField>Key</DataField></Field><Field Name="Value"><DataField>Value</DataField></Field>
 </Fields><Query><CommandText>metadata only</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Textbox Name="Level"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun>
   <Value>=Code.GetPercentLevel(0.6)</Value><Style><FontFamily>Arial</FontFamily></Style>
  </TextRun></TextRuns></Paragraph></Paragraphs><Top>0in</Top><Left>0in</Left><Height>0.3in</Height><Width>3in</Width><Style/></Textbox>
  <Textbox Name="Distinct"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun>
   <Value>=Join(Code.GetDistinct(LookupSet("All", Fields!Key.Value, Fields!Value.Value, "D")), ", ")</Value>
   <Style><FontFamily>Arial</FontFamily></Style>
  </TextRun></TextRuns></Paragraph></Paragraphs><Top>0.4in</Top><Left>0in</Left><Height>0.3in</Height><Width>3in</Width><Style/></Textbox>
 </ReportItems><Height>1in</Height><Style/></Body><Width>3in</Width>
 <Page><PageHeight>4in</PageHeight><PageWidth>4in</PageWidth><LeftMargin>0.25in</LeftMargin><RightMargin>0.25in</RightMargin><TopMargin>0.25in</TopMargin><BottomMargin>0.25in</BottomMargin></Page>
 </ReportSection></ReportSections>
 <Code>Public Function GetPercentLevel(totalScore As Double) As String : Return "not executed" : End Function
 Function GetDistinct(input As Object()) As Object() : Return Nothing : End Function</Code>
</Report>`;
}

const request = {
  outputFileName: 'native-custom-functions',
  parameters: {},
  datasets: { D: [
    { Key: 'All', Value: 'Alpha' },
    { Key: 'All', Value: 'Alpha' },
    { Key: 'All', Value: 'Beta' },
  ] },
};

test('GetPercentLevel and GetDistinct pass capability analysis and render identically across PDF, editable DOCX, and XLSX', async () => {
  const rdl = customFunctionRdl();
  const analysis = analyzeRdl(rdl);
  assert.equal(analysis.compatible, true);
  assert.ok(analysis.capabilities.expressions.detected
    .some((entry) => entry.name === 'Code.GetPercentLevel' && entry.status === 'SUPPORTED'));
  assert.ok(analysis.capabilities.expressions.detected
    .some((entry) => entry.name === 'Code.GetDistinct' && entry.status === 'SUPPORTED'));
  const model = parseRdl(rdl);

  const pdf = await renderPdf(model, request, config);
  const pdfPath = new URL(`../tmp/native-custom-functions-${process.pid}.pdf`, import.meta.url);
  await fs.writeFile(pdfPath, pdf.buffer);
  try {
    const extracted = await execFileAsync('pdftotext', [pdfPath.pathname, '-']);
    assert.match(extracted.stdout, /Level 3/);
    assert.match(extracted.stdout, /Alpha, Beta/);
  } finally {
    await fs.unlink(pdfPath).catch(() => {});
  }

  const docx = await renderEditableDocx(model, request, config);
  const documentXml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  const nativeText = [...documentXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
    .map((match) => match[1])
    .join(' ');
  const normalizedNativeText = nativeText.replace(/\s+/g, ' ').trim();
  assert.match(normalizedNativeText, /Level 3/);
  assert.match(normalizedNativeText, /Alpha, Beta/);

  const xlsx = await renderExcel(model, request, config, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(xlsx.buffer);
  const values = [];
  for (const worksheet of workbook.worksheets) {
    worksheet.eachRow((row) => row.eachCell((cell) => values.push(cell.value)));
  }
  assert.ok(values.includes('Level 3'));
  assert.ok(values.includes('Alpha, Beta'));
});
