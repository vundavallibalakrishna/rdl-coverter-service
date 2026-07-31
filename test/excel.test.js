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

function chartReportRdl() {
  const chart = (name, left, title) => `<Chart Name="${name}">
    <ChartCategoryHierarchy><ChartMembers><ChartMember><Group Name="${name}Categories"><GroupExpressions>
      <GroupExpression>=Fields!Category.Value</GroupExpression>
    </GroupExpressions></Group><Label>=Fields!Category.Value</Label></ChartMember></ChartMembers></ChartCategoryHierarchy>
    <ChartSeriesHierarchy><ChartMembers><ChartMember><Label>Amount</Label></ChartMember></ChartMembers></ChartSeriesHierarchy>
    <ChartData><ChartSeriesCollection><ChartSeries Name="AmountSeries"><ChartDataPoints><ChartDataPoint>
      <ChartDataPointValues><Y>=Sum(Fields!Amount.Value)</Y></ChartDataPointValues>
      <ChartDataLabel><UseValueAsLabel>true</UseValueAsLabel><Visible>true</Visible><Style/></ChartDataLabel>
      <Style/>
    </ChartDataPoint></ChartDataPoints><Type>Column</Type></ChartSeries></ChartSeriesCollection></ChartData>
    <ChartAreas><ChartArea Name="Default"><ChartCategoryAxes><ChartAxis Name="Category"><Style/></ChartAxis></ChartCategoryAxes>
      <ChartValueAxes><ChartAxis Name="Value"><Style/></ChartAxis></ChartValueAxes></ChartArea></ChartAreas>
    <ChartTitles><ChartTitle Name="Title"><Caption>${title}</Caption><Style/></ChartTitle></ChartTitles>
    <DataSetName>D</DataSetName><Top>0.75in</Top><Left>${left}in</Left><Width>3in</Width><Height>2in</Height>
    <Style><Border><Style>Solid</Style><Color>Black</Color><Width>1pt</Width></Border></Style>
  </Chart>`;
  return `<?xml version="1.0"?><Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D"><Fields><Field Name="Category"><DataField>Category</DataField></Field>
    <Field Name="Amount"><DataField>Amount</DataField></Field></Fields></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Textbox Name="Heading"><CanGrow>true</CanGrow><KeepTogether>true</KeepTogether>
      <Paragraphs><Paragraph><TextRuns><TextRun><Value>Dashboard Overview</Value><Style><FontWeight>Bold</FontWeight></Style></TextRun></TextRuns></Paragraph></Paragraphs>
      <Top>0in</Top><Left>0in</Left><Height>0.3in</Height><Width>6.2in</Width><Style/>
    </Textbox>
    <Textbox Name="LeftBand"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun>
      <Value>Risk Matrix with Count of Risks</Value>
    </TextRun></TextRuns></Paragraph></Paragraphs><Top>0.4in</Top><Left>0in</Left><Height>0.3in</Height><Width>3in</Width><Style/></Textbox>
    <Textbox Name="RightBand"><CanGrow>true</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun>
      <Value>Action Status</Value>
    </TextRun></TextRuns></Paragraph></Paragraphs><Top>0.4in</Top><Left>3.2in</Left><Height>0.3in</Height><Width>3in</Width><Style/></Textbox>
    ${chart('LeftChart', 0, 'By Division')}
    ${chart('RightChart', 3.2, 'By Department')}
  </ReportItems><Height>2.75in</Height><Style/></Body><Width>6.2in</Width>
  <Page><PageHeight>4in</PageHeight><PageWidth>6.5in</PageWidth><LeftMargin>0.1in</LeftMargin>
    <RightMargin>0.1in</RightMargin><TopMargin>0.1in</TopMargin><BottomMargin>0.1in</BottomMargin></Page>
  </ReportSection></ReportSections></Report>`;
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

test('REPORT mode honors CanGrow for wrapped cell text and preserves a fixed CanGrow=false row', async () => {
  const growingModel = structuredClone(model);
  const tablix = growingModel.body.items.find((item) => item.type === 'Tablix');
  tablix.columns[0] = 45.5;
  tablix.width = tablix.columns.reduce((sum, width) => sum + width, 0);
  const detailRow = tablix.rows.at(-1);
  const textbox = detailRow.cells[0].items.find((item) => item.type === 'Textbox');
  textbox.canGrow = true;
  textbox.style.fontFamily = 'Arial';
  textbox.style.fontSize = 8;
  textbox.style.paddingLeft = 2;
  textbox.style.paddingRight = 2;
  for (const run of textbox.paragraphs.flat()) {
    run.style.fontFamily = 'Arial';
    run.style.fontSize = 8;
  }
  const growingRequest = {
    ...request,
    datasets: { ...request.datasets, Sales: [{ Name: 'establishment', Amount: 1 }] },
  };
  const growingSheet = await load((await renderExcel(growingModel, growingRequest, config, null)).buffer);
  const growingCell = findCell(growingSheet, 'establishment');
  assert.ok(growingCell);
  assert.equal(growingCell.alignment?.wrapText, true);
  assert.ok(
    growingSheet.getRow(growingCell.row).height > detailRow.height,
    `expected CanGrow row to exceed its declared ${detailRow.height}pt height`,
  );

  const fixedModel = structuredClone(growingModel);
  fixedModel.body.items.find((item) => item.type === 'Tablix')
    .rows.at(-1).cells[0].items.find((item) => item.type === 'Textbox').canGrow = false;
  const fixedSheet = await load((await renderExcel(fixedModel, growingRequest, config, null)).buffer);
  const fixedCell = findCell(fixedSheet, 'establishment');
  assert.ok(fixedCell);
  assert.equal(fixedCell.alignment?.wrapText, true);
  assert.equal(fixedSheet.getRow(fixedCell.row).height, detailRow.height);
});

test('REPORT mode splits a CanGrow cell beyond Excel row-height limits without losing editability', async () => {
  const tallModel = structuredClone(model);
  const tablix = tallModel.body.items.find((item) => item.type === 'Tablix');
  tablix.columns[0] = 72;
  tablix.width = tablix.columns.reduce((sum, width) => sum + width, 0);
  const textbox = tablix.rows.at(-1).cells[0].items.find((item) => item.type === 'Textbox');
  textbox.canGrow = true;
  textbox.style.fontFamily = 'Arial';
  textbox.style.fontSize = 8;
  const tallText = Array.from({ length: 160 }, (_, index) => `Line ${String(index + 1).padStart(3, '0')}`).join('\n');
  const tallSheet = await load((await renderExcel(tallModel, {
    ...request,
    datasets: { ...request.datasets, Sales: [{ Name: tallText, Amount: 1 }] },
  }, config, null)).buffer);
  const foundCell = findCell(tallSheet, tallText);
  assert.ok(foundCell, 'expected the complete value to remain in one editable cell');
  const tallCell = foundCell.master || foundCell;
  const physicalRows = [];
  tallSheet.eachRow({ includeEmpty: true }, (row) => {
    assert.ok((row.height || 0) <= 409, `row ${row.number} exceeds Excel's 409-point limit`);
    if (row.number >= tallCell.row && (row.height || 0) > 0) physicalRows.push(row);
  });
  assert.ok(physicalRows.reduce((sum, row) => sum + row.height, 0) > 409);
  const spanningMerge = (tallSheet.model.merges || []).find((range) => range.startsWith(`${tallCell.address}:`));
  assert.ok(spanningMerge, 'expected the tall editable cell to span multiple native Excel rows');
  const endRow = Number.parseInt(/\d+$/.exec(spanningMerge)?.[0] || '0', 10);
  assert.ok(endRow > tallCell.row);
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

test('REPORT mode coalesces adjacent free-form edges that overlap only by their shared border stroke', async () => {
  const adjacent = structuredClone(model);
  const source = adjacent.body.items.find((item) => item.type === 'Textbox');
  const item = (name, value, left, width) => {
    const textbox = structuredClone(source);
    textbox.name = name;
    textbox.value = value;
    textbox.top = 0;
    textbox.left = left;
    textbox.width = width;
    textbox.height = 20;
    textbox.style.border = { style: 'Solid', color: '#000000', width: 1 };
    textbox.style.borders = Object.fromEntries(['top', 'right', 'bottom', 'left']
      .map((side) => [side, { style: 'Solid', color: '#000000', width: 1 }]));
    textbox.paragraphs = [[{ value, markupType: 'None', style: textbox.style }]];
    return textbox;
  };
  adjacent.page.header = {
    height: 20,
    printOnFirstPage: true,
    printOnLastPage: true,
    items: [{
      type: 'Rectangle', name: 'HeaderBand', top: 0, left: 0, width: 140, height: 20,
      zIndex: 0, hidden: 'false', style: source.style, pageBreak: null, items: [
        item('LeftHeader', 'Left header', 0, 70.75),
        item('RightHeader', 'Right header', 69.75, 70.25),
      ],
    }],
  };

  const ws = await load((await renderExcel(adjacent, request, config, null)).buffer);
  assert.ok(findCell(ws, 'Left header'));
  assert.ok(findCell(ws, 'Right header'));
});

test('REPORT mode treats a one-quarter-point accumulated band edge as the declared band boundary', async () => {
  const quantized = structuredClone(model);
  const imageEdge = structuredClone(quantized.body.items.find((item) => item.type === 'Textbox'));
  imageEdge.name = 'BandEdge';
  imageEdge.value = 'Band edge';
  imageEdge.paragraphs = [[{ value: 'Band edge', markupType: 'None', style: imageEdge.style }]];
  imageEdge.top = 2.414976377952756;
  imageEdge.left = 0;
  imageEdge.width = 100;
  imageEdge.height = 49.1700188976378;
  quantized.page.header = {
    height: 51.58499527559055,
    printOnFirstPage: true,
    printOnLastPage: true,
    items: [imageEdge],
  };

  const ws = await load((await renderExcel(quantized, request, config, null)).buffer);
  assert.ok(findCell(ws, 'Band edge'));
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

test('a declared RDL PageName is preferred for the native section worksheet name', async () => {
  const named = structuredClone(model);
  named.body.items.find((item) => item.type === 'Tablix').pageName = '=Parameters!Title.Value & " detail"';
  const result = await renderExcel(named, request, config, null);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(result.buffer);
  assert.equal(wb.worksheets[0].name, 'Sales detail');
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

test('REPORT mode preserves vertical group spans as native merges and closes the final grid edge', async () => {
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
  const eastMasters = new Set(eastCells.map((address) => ws.getCell(address).master.address));
  assert.equal(eastMasters.size, 1, `expected one merged group owner, got ${eastCells.join(', ')}`);
  const eastMaster = [...eastMasters][0];
  const zip = await JSZip.loadAsync(result.buffer);
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string');
  const mergeRefs = [...sheetXml.matchAll(/<mergeCell ref="([A-Z]+\d+:[A-Z]+\d+)"/g)].map((match) => match[1]);
  const eastMerge = mergeRefs.find((range) => range.startsWith(`${eastMaster}:`));
  assert.ok(eastMerge, `the grouped risk/header value must own a vertical merged range: ${eastMaster} / ${mergeRefs.join(', ')}`);
  const [, endAddress] = eastMerge.split(':');
  assert.notEqual(endAddress.replace(/\d+/g, ''), '');
  assert.notEqual(Number(endAddress.match(/\d+/)?.[0]), Number(eastMaster.match(/\d+/)?.[0]));
  const finalRow = ws.getRow(ws.rowCount);
  for (let column = 1; column <= ws.columnCount; column += 1) {
    assert.equal(finalRow.getCell(column).border?.bottom?.style, 'thin', `missing final bottom border in column ${column}`);
  }
  assert.ok((ws.views.find((view) => view.state === 'frozen')?.xSplit || 0) > 0);
});

test('REPORT embedded images use pictures without emitting layout shapes or native chart parts', async () => {
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

test('REPORT mode anchors visible charts as pictures while preserving RDL section naming and peer geometry', async () => {
  const charted = parseRdl(chartReportRdl());
  const rendered = await renderExcel(charted, {
    outputFileName: 'Dashboard',
    parameters: {},
    datasets: { D: [{ Category: 'A', Amount: 2 }, { Category: 'B', Amount: 3 }] },
    excel: { layoutMode: 'REPORT' },
  }, config, null);
  assert.equal(rendered.layoutMode, 'report-sections');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(rendered.buffer);
  assert.deepEqual(workbook.worksheets.map((worksheet) => worksheet.name), ['Dashboard Overview']);
  assert.equal(workbook.worksheets[0].getImages().length, 2);
  const leftBand = findCell(workbook.worksheets[0], 'Risk Matrix with Count of Risks');
  const rightBand = findCell(workbook.worksheets[0], 'Action Status');
  assert.ok(leftBand && rightBand);
  assert.equal(leftBand.row, rightBand.row, 'coincident side-by-side freeform headings must share their worksheet row');

  const zip = await JSZip.loadAsync(rendered.buffer);
  const drawing = await zip.file('xl/drawings/drawing1.xml').async('string');
  assert.equal((drawing.match(/<xdr:pic>/g) || []).length, 2);
  assert.equal(Object.keys(zip.files).some((name) => name.startsWith('xl/charts/')), false);

  const anchors = [...drawing.matchAll(/<xdr:from><xdr:col>(\d+)<\/xdr:col><xdr:colOff>\d+<\/xdr:colOff><xdr:row>(\d+)<\/xdr:row>/g)]
    .map((match) => ({ column: Number(match[1]), row: Number(match[2]) }));
  assert.equal(anchors.length, 2);
  assert.equal(anchors[0].row, anchors[1].row, 'side-by-side chart peers must share their top row');
  assert.notEqual(anchors[0].column, anchors[1].column);
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
