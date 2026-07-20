// Integration tests for the XLSX renderer. The point of Excel export over PDF/DOCX is live, computable
// values, so these assert that numbers are real numbers with format codes, styling and merges survive, and
// untrusted values are stored as typed strings that Excel cannot execute.
import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { parseRdl } from '../src/rdl/parser.js';
import { loadConfig } from '../src/config.js';
import { renderExcel, resolveExcelLayoutMode } from '../src/render/excel.js';
import { renderDocument, OUTPUTS } from '../src/render/index.js';

const model = parseRdl(await fs.readFile(new URL('./fixtures/basic.rdl', import.meta.url)));
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const request = { outputFileName: 'Sales', parameters: { Title: 'Sales', Choice: 'A' }, datasets: { Sales: [{ Name: 'North', Amount: 1234.5 }, { Name: 'South', Amount: 99 }], Choices: [{ Value: 'A' }] } };

function groupedRdl() {
  return `<?xml version="1.0"?><Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D"><Fields><Field Name="Region"><DataField>Region</DataField></Field><Field Name="Name"><DataField>Name</DataField></Field><Field Name="Amount"><DataField>Amount</DataField></Field></Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems><Tablix Name="T"><TablixBody>
  <TablixColumns><TablixColumn><Width>1in</Width></TablixColumn><TablixColumn><Width>1in</Width></TablixColumn></TablixColumns>
  <TablixRows><TablixRow><Height>0.25in</Height><TablixCells>
  <TablixCell><CellContents><Textbox Name="n"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Name.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell>
  <TablixCell><CellContents><Textbox Name="a"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Amount.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell>
  </TablixCells></TablixRow></TablixRows></TablixBody>
  <TablixColumnHierarchy><TablixMembers><TablixMember/><TablixMember/></TablixMembers></TablixColumnHierarchy>
  <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="Regions"><GroupExpressions><GroupExpression>=Fields!Region.Value</GroupExpression></GroupExpressions></Group>
  <TablixHeader><Size>0.75in</Size><CellContents><Textbox Name="r"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Region.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixHeader><FixedData>true</FixedData>
  <TablixMembers><TablixMember><Group Name="Details"/></TablixMember></TablixMembers></TablixMember></TablixMembers></TablixRowHierarchy>
  <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.75in</Height><Width>2.75in</Width><Style/></Tablix></ReportItems><Height>3in</Height><Style/></Body>
  <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth></Page></ReportSection></ReportSections></Report>`;
}

async function load(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb.worksheets[0];
}
const findCell = (ws, text) => {
  let found = null;
  ws.eachRow((row) => row.eachCell((cell) => { if (cell.value === text) found = cell; }));
  return found;
};

test('XLSX is a registered output and dispatches through renderDocument', async () => {
  assert.equal(OUTPUTS.has('XLSX'), true);
  const result = await renderDocument(model, { ...request, output: 'XLSX' }, config, null);
  assert.equal(result.extension, 'xlsx');
  assert.equal(result.mimeType, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  assert.equal(result.buffer.subarray(0, 2).toString(), 'PK');
  assert.equal(result.layoutMode, 'report-sections');
});

test('Excel layout mode is case-insensitive and preserves sheetPerTablix compatibility', () => {
  assert.equal(resolveExcelLayoutMode({}), 'REPORT');
  assert.equal(resolveExcelLayoutMode({ excel: { layoutMode: 'data' } }), 'DATA');
  assert.equal(resolveExcelLayoutMode({ excel: { layoutMode: 'rEpOrT' } }), 'REPORT');
  assert.equal(resolveExcelLayoutMode({ excel: { sheetPerTablix: true } }), 'DATA');
  assert.throws(
    () => resolveExcelLayoutMode({ excel: { layoutMode: 'REPORT', sheetPerTablix: true } }),
    (error) => error.code === 'RDL_INVALID',
  );
  assert.throws(() => resolveExcelLayoutMode({ excel: { layoutMode: 'canvas' } }), (error) => error.code === 'RDL_INVALID');
});

test('renders a valid workbook that reloads with the expected header and data text', async () => {
  const result = await renderExcel(model, request, config, null);
  const ws = await load(result.buffer);
  assert.match(ws.name, /Sales/);
  assert.ok(findCell(ws, 'Name'));
  assert.ok(findCell(ws, 'Amount'));
  assert.ok(findCell(ws, 'North'));
  assert.ok(findCell(ws, 'South'));
});

test('a Format()-wrapped numeric field becomes a live number with a translated format code', async () => {
  // basic.rdl renders Amount as =Format(Fields!Amount.Value, "N2"). It must arrive as the number 1234.5,
  // not the string "1,234.50", so it stays summable — with an N2-equivalent Excel format for display.
  const ws = await load((await renderExcel(model, request, config, null)).buffer);
  let amount = null;
  ws.eachRow((row) => row.eachCell((cell) => { if (cell.value === 1234.5) amount = cell; }));
  assert.ok(amount, 'expected a live numeric 1234.5 cell');
  assert.equal(typeof amount.value, 'number');
  assert.equal(amount.numFmt, '#,##0.00');
});

test('a text field stays text', async () => {
  const ws = await load((await renderExcel(model, request, config, null)).buffer);
  assert.equal(typeof findCell(ws, 'North').value, 'string');
});

test('cell fills and borders from the RDL survive into the workbook', async () => {
  const ws = await load((await renderExcel(model, request, config, null)).buffer);
  const header = findCell(ws, 'Amount'); // basic.rdl gives the header a #dddddd background and a solid border
  assert.equal(header.fill?.pattern, 'solid');
  assert.equal(header.fill?.fgColor?.argb, 'FFDDDDDD');
  assert.equal(header.border?.bottom?.style, 'thin');
  assert.equal(header.font?.bold, true);
});

test('valid horizontal header spans are merged', async () => {
  const merged = structuredClone(model);
  const tablix = merged.body.items.find((item) => item.type === 'Tablix');
  tablix.rows[0].cells[0].colSpan = 2;
  tablix.rows[0].cells.splice(1, 1); // a valid span owns both columns; there is no covered placeholder
  const result = await renderExcel(merged, request, config, null);
  const sheetXml = await (await JSZip.loadAsync(result.buffer)).file('xl/worksheets/sheet1.xml').async('string');
  assert.match(sheetXml, /<mergeCell ref="[A-Z]+\d+:[A-Z]+\d+"/);
});

test('section boundaries inside a logical tablix cell do not create blank Excel gaps', async () => {
  const subdivided = structuredClone(model);
  subdivided.body.items.push({
    type: 'Rectangle', name: 'BoundaryOnly', top: 32, left: 100, width: 10, height: 5,
    zIndex: 0, hidden: 'false', style: {}, pageBreak: null, items: [],
  });
  const ws = await load((await renderExcel(subdivided, request, config, null)).buffer);
  const nameCell = findCell(ws, 'Name');
  assert.ok(nameCell);
  assert.ok(
    (ws.model.merges || []).some((range) => ws.getCell(range.split(':')[0]).value === 'Name'),
    'expected the subdivided logical Name cell to own its full horizontal region',
  );
});

test('overlapping tablix merge plans fail closed', async () => {
  const overlapping = structuredClone(model);
  overlapping.body.items.find((item) => item.type === 'Tablix').rows[0].cells[0].colSpan = 2;
  await assert.rejects(renderExcel(overlapping, request, config, null), (error) => error.code === 'RDL_INVALID');
});

test('sheetPerTablix puts each tablix on its own worksheet with its own columns', async () => {
  const single = await renderExcel(model, request, config, null);
  assert.equal(single.sheetCount, 1); // default: everything stacked on one sheet

  const perTablix = await renderExcel(model, { ...request, excel: { sheetPerTablix: true } }, config, null);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(perTablix.buffer);
  const names = wb.worksheets.map((ws) => ws.name);
  assert.ok(names.includes('Table 1'), `expected a per-tablix sheet, got ${names.join(', ')}`);
  // basic.rdl has a free-form title textbox, so its non-tablix content lands on a leading Overview sheet.
  assert.ok(names.includes('Overview'));
  const table = wb.getWorksheet('Table 1');
  assert.ok(findCell(table, 'North'), 'the tablix data belongs on its own sheet');
});

test('REPORT columns come from RDL coordinates and do not autofit to content', async () => {
  const ordinary = await load((await renderExcel(model, request, config, null)).buffer);
  const long = {
    ...request,
    datasets: { ...request.datasets, Sales: [{ Name: 'A very long value that would materially widen an autofit column', Amount: 1 }] },
  };
  const expanded = await load((await renderExcel(model, long, config, null)).buffer);
  assert.deepEqual(
    ordinary.columns.map((column) => column.width),
    expanded.columns.map((column) => column.width),
  );
  assert.ok(ordinary.columns.every((column) => typeof column.width === 'number' && column.width > 0));
});

test('explicit page breaks create stable section worksheets and normal views have no zero split', async () => {
  const sectioned = structuredClone(model);
  sectioned.body.items.find((item) => item.type === 'Tablix').pageBreak = { location: 'Start', disabled: 'false' };
  const result = await renderExcel(sectioned, request, config, null);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(result.buffer);
  assert.deepEqual(wb.worksheets.map((sheet) => sheet.name), ['Sales', 'Name']);
  assert.equal(result.layoutMode, 'report-sections');
  for (const sheet of wb.worksheets) {
    assert.ok(sheet.views.every((view) => view.state !== 'frozen' || (view.xSplit || 0) > 0 || (view.ySplit || 0) > 0));
    assert.match(sheet.pageSetup.printArea, /^A1:[A-Z]+\d+$/);
  }
});

test('only declared repeating header rows produce a valid frozen pane and AutoFilter', async () => {
  const repeating = structuredClone(model);
  const tablix = repeating.body.items.find((item) => item.type === 'Tablix');
  tablix.rowMembers[0].repeatOnNewPage = true;
  const result = await renderExcel(repeating, request, config, null);
  const ws = await load(result.buffer);
  const frozen = ws.views.find((view) => view.state === 'frozen');
  assert.ok((frozen?.ySplit || 0) > 0);
  assert.ok(ws.autoFilter);
});

test('multiple RDL text runs remain native Excel rich text', async () => {
  const rich = structuredClone(model);
  const title = rich.body.items.find((item) => item.type === 'Textbox');
  title.value = 'Sales Report';
  title.paragraphs = [[
    { value: 'Sales ', markupType: 'None', style: { ...title.style, fontWeight: 'Normal', color: '#112233' } },
    { value: 'Report', markupType: 'None', style: { ...title.style, fontWeight: 'Bold', color: '#CC0000' } },
  ]];
  const ws = await load((await renderExcel(rich, request, config, null)).buffer);
  let cell = null;
  ws.eachRow((row) => row.eachCell((candidate) => {
    if (candidate.value?.richText?.map((run) => run.text).join('') === 'Sales Report') cell = candidate;
  }));
  assert.ok(cell);
  assert.notEqual(cell.value.richText[0].font.bold, true);
  assert.equal(cell.value.richText[1].font.bold, true);
  assert.equal(cell.value.richText[1].font.color.argb, 'FFCC0000');
});

test('vertical group spans become repeated editable values and FixedData freezes a real column split', async () => {
  const grouped = parseRdl(groupedRdl());
  const groupedRequest = { outputFileName: 'Grouped', parameters: {}, datasets: { D: [
    { Region: 'East', Name: 'A', Amount: 1 },
    { Region: 'East', Name: 'B', Amount: 2 },
    { Region: 'West', Name: 'C', Amount: 3 },
  ] } };
  const result = await renderExcel(grouped, groupedRequest, config, null);
  const ws = await load(result.buffer);
  const eastCells = [];
  ws.eachRow((row) => row.eachCell((cell) => { if (cell.value === 'East') eastCells.push(cell.address); }));
  assert.equal(eastCells.length, 2, `expected repeated group values, got ${eastCells.join(', ')}`);
  assert.equal((ws.model.merges || []).some((range) => range.includes(`${eastCells[0]}:`)), false);
  assert.ok((ws.views.find((view) => view.state === 'frozen')?.xSplit || 0) > 0);
});

test('REPORT drawings contain pictures only for declared embedded RDL images', async () => {
  const withLogo = structuredClone(model);
  withLogo.embeddedImages = {
    Logo: {
      mimeType: 'image/png',
      data: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+X2NDVQAAAABJRU5ErkJggg==',
    },
  };
  withLogo.body.items.push({
    type: 'Image', name: 'Logo', source: 'Embedded', value: 'Logo', top: 90, left: 7.2, width: 36, height: 18,
    zIndex: 0, hidden: 'false', style: {}, pageBreak: null,
  });
  const result = await renderExcel(withLogo, request, config, null);
  const zip = await JSZip.loadAsync(result.buffer);
  const drawingNames = Object.keys(zip.files).filter((name) => /^xl\/drawings\/drawing\d+\.xml$/.test(name));
  assert.equal(drawingNames.length, 1);
  const drawing = await zip.file(drawingNames[0]).async('string');
  assert.match(drawing, /<xdr:pic>/);
  assert.doesNotMatch(drawing, /<xdr:sp>/);
  assert.equal(Object.keys(zip.files).some((name) => name.startsWith('xl/charts/')), false);
});

test('visible charts fail closed in REPORT mode instead of disappearing', async () => {
  const charted = structuredClone(model);
  charted.body.items.push({ type: 'Chart', name: 'C', top: 90, left: 0, width: 100, height: 100, zIndex: 0, hidden: 'false', pageBreak: null });
  await assert.rejects(renderExcel(charted, request, config, null), (error) => error.code === 'UNSUPPORTED_FEATURE');
});

test('an untrusted value beginning with = is stored as a typed string, never a live formula', async () => {
  const evil = { ...request, datasets: { ...request.datasets, Sales: [{ Name: '=1+2+cmd|calc', Amount: 1 }] } };
  const result = await renderExcel(model, evil, config, null);
  const sheetXml = await (await JSZip.loadAsync(result.buffer)).file('xl/worksheets/sheet1.xml').async('string');
  // The dangerous string must appear only inside a string cell/sharedString, and there must be no formula
  // element anywhere in the sheet.
  assert.doesNotMatch(sheetXml, /<f>/);
  const ws = await load(result.buffer);
  const cell = findCell(ws, '=1+2+cmd|calc');
  assert.ok(cell, 'value should round-trip verbatim (no apostrophe corruption)');
  assert.equal(typeof cell.value, 'string');
});
