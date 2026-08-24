import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseDateValue } from '../src/rdl/dateValue.js';
import { evaluateExpression, formatValue } from '../src/rdl/expression.js';
import { formatNet } from '../src/rdl/format.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

// An RDL DateTime is a wall-clock value: SSRS renders exactly the date and time it was given and never
// applies a time-zone conversion. Every formatter here reads a Date through its UTC accessors, so a value
// with no time-zone designator must be parsed as UTC too. Parsed as local time (the JavaScript default for
// that form) a midnight value renders as the PREVIOUS day on any host east of UTC, and the same inputs
// render differently on different servers — which also breaks the determinism invariant.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000' });

test('an unzoned date-time keeps its wall clock instead of shifting into the host time zone', () => {
  assert.equal(parseDateValue('2026-05-15T00:00:00').toISOString(), '2026-05-15T00:00:00.000Z');
  assert.equal(parseDateValue('2026-05-15T00:00:00.000').toISOString(), '2026-05-15T00:00:00.000Z');
  assert.equal(parseDateValue('2026-05-15 09:30:00').toISOString(), '2026-05-15T09:30:00.000Z');
  assert.equal(parseDateValue('2026-12-31T23:59:59').toISOString(), '2026-12-31T23:59:59.000Z');
  // A date-only value already parsed as UTC and must keep doing so.
  assert.equal(parseDateValue('2026-05-15').toISOString(), '2026-05-15T00:00:00.000Z');
  // A value that names a real instant keeps the standard parse.
  assert.equal(parseDateValue('2026-05-15T00:00:00Z').toISOString(), '2026-05-15T00:00:00.000Z');
  assert.equal(parseDateValue('2026-05-15T00:00:00+05:30').toISOString(), '2026-05-14T18:30:00.000Z');
  assert.equal(parseDateValue(''), null);
  assert.equal(parseDateValue('not a date'), null);
});

test('the parsed wall clock is identical on every host time zone', (context) => {
  const original = process.env.TZ;
  context.after(() => { process.env.TZ = original; });
  // Node re-reads TZ per Date operation, so this exercises the real failure: on a UTC build machine the
  // legacy local-time parse looked correct and only shifted the day once deployed elsewhere.
  for (const zone of ['UTC', 'Asia/Kolkata', 'America/Los_Angeles', 'Pacific/Kiritimati']) {
    process.env.TZ = zone;
    assert.equal(parseDateValue('2026-05-15T00:00:00').toISOString(), '2026-05-15T00:00:00.000Z', zone);
    assert.equal(formatNet('2026-05-15T00:00:00', 'dd/MM/yyyy'), '15/05/2026', zone);
    assert.equal(String(formatValue('2026-06-05T00:00:00', 'dd/MM/yyyy')), '05/06/2026', zone);
  }
});

test('every formatting path renders the supplied day, not the day before', () => {
  const midnight = '2026-05-15T00:00:00';
  assert.equal(formatNet(midnight, 'dd/MM/yyyy'), '15/05/2026');
  assert.equal(formatNet(midnight, 'MMMM d, yyyy', 'en-US'), 'May 15, 2026');
  assert.equal(String(formatValue(midnight, 'dd/MM/yyyy')), '15/05/2026');
  assert.equal(String(formatValue(midnight, null)), '15/05/2026 00:00:00');
  // The culture-dependent standard specifiers resolve the same day in either ordering.
  assert.match(String(formatValue(midnight, 'd', 'en-US')), /5\/15\/(20)?26/);
  assert.equal(String(formatValue(midnight, 'd', 'en-GB')), '15/05/2026');
  assert.equal(String(formatValue(midnight, 'D', 'en-US')), 'May 15, 2026');
  // Expression date functions read the same wall clock.
  assert.equal(evaluateExpression(`=Day(CDate("${midnight}"))`, {}), 15);
  assert.equal(evaluateExpression(`=Month(CDate("${midnight}"))`, {}), 5);
  assert.equal(evaluateExpression(`=Year(CDate("${midnight}"))`, {}), 2026);
  assert.equal(String(evaluateExpression(`=Format(CDate("${midnight}"), "dd/MM/yyyy")`, {})), '15/05/2026');
});

const rdl = `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <Body><ReportItems>
    <Tablix Name="Dates"><TablixBody>
      <TablixColumns><TablixColumn><Width>3in</Width></TablixColumn></TablixColumns>
      <TablixRows><TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents>
        <Textbox Name="DueCell"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun>
          <Value>=Fields!Due.Value</Value>
          <Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize><Format>dd/MM/yyyy</Format></Style>
        </TextRun></TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border></Style></Textbox>
      </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
    </TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="Details"/></TablixMember></TablixMembers></TablixRowHierarchy>
    <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.25in</Height><Width>3in</Width><Style/>
    </Tablix>
  </ReportItems><Height>1in</Height></Body><Width>3in</Width>
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="Due"><DataField>Due</DataField><rd:TypeName xmlns:rd="http://schemas.microsoft.com/SQLServer/reporting/reportdesigner">System.DateTime</rd:TypeName></Field>
  </Fields></DataSet></DataSets>
  <Page><PageWidth>4in</PageWidth><PageHeight>3in</PageHeight><TopMargin>0.25in</TopMargin>
    <BottomMargin>0.25in</BottomMargin><LeftMargin>0.25in</LeftMargin><RightMargin>0.25in</RightMargin></Page>
</Report>`;

const request = {
  outputFileName: 'wall-clock',
  parameters: {},
  datasets: { D: [{ Due: '2026-05-15T00:00:00' }, { Due: '2026-06-05T00:00:00' }] },
};

test('PDF, editable DOCX, and XLSX all render the supplied wall-clock day', async (context) => {
  const model = parseRdl(rdl);
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-wall-clock-'));
  context.after(() => fs.rm(tempRoot, { recursive: true, force: true }));
  const ownedConfig = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000', RDL_TEMP_ROOT: tempRoot });

  const pdf = await renderPdf(model, request, ownedConfig, { captureLayoutTrace: true });
  const pdfText = pdf.layoutTrace.pages
    .flatMap((page) => (page.tablixFragments || []).flatMap((fragment) => (fragment.cells || []).map((cell) => cell.text)))
    .filter(Boolean);
  assert.deepEqual(pdfText, ['15/05/2026', '05/06/2026']);

  const docx = await renderEditableDocx(model, request, ownedConfig);
  const documentXml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  const nativeText = [...documentXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((match) => match[1]).join('|');
  assert.match(nativeText, /15\/05\/2026/);
  assert.match(nativeText, /05\/06\/2026/);

  // XLSX writes a live typed date: the serial itself must carry the supplied day.
  const xlsx = await renderExcel(model, { ...request, excel: { layoutMode: 'REPORT' } }, ownedConfig, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(xlsx.buffer);
  const dates = [];
  workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => {
    if (cell.value instanceof Date) dates.push(cell.value.toISOString().slice(0, 10));
  }));
  assert.deepEqual(dates, ['2026-05-15', '2026-06-05']);
});
