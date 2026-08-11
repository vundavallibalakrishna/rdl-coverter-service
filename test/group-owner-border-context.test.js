import assert from 'node:assert/strict';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { matchingChangedGroupOwnerRowBoundary, styleColor, styleValue, tablixRows } from '../src/render/common.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

// A dynamic parent row header is evaluated once against the first row of its group and then spans all
// child detail rows. SSRS keeps that owner scope for expression-backed styles. Re-evaluating the merged
// cell against a later child row changes Number from 1 to 2 and silently removes this separator.
const model = parseRdl(`<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="Risk"><DataField>Risk</DataField></Field>
    <Field Name="Number"><DataField>Number</DataField></Field>
    <Field Name="Rating"><DataField>Rating</DataField></Field>
    <Field Name="Detail"><DataField>Detail</DataField></Field>
  </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="GroupedRisk">
      <TablixBody>
        <TablixColumns><TablixColumn><Width>2in</Width></TablixColumn></TablixColumns>
        <TablixRows><TablixRow><Height>0.35in</Height><TablixCells>
          <TablixCell><CellContents><Textbox Name="DetailCell"><CanGrow>true</CanGrow>
            <Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Detail.Value</Value><Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs>
            <Style><Border><Style>Solid</Style><Color>Black</Color><Width>1pt</Width></Border></Style>
          </Textbox></CellContents></TablixCell>
        </TablixCells></TablixRow></TablixRows>
      </TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers><TablixMember>
        <Group Name="RiskGroup"><GroupExpressions><GroupExpression>=Fields!Risk.Value</GroupExpression></GroupExpressions></Group>
        <TablixHeader><Size>1in</Size><CellContents><Textbox Name="RiskOwner"><CanGrow>true</CanGrow>
          <Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Risk.Value</Value><Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs>
          <Style><Border><Style>Solid</Style><Color>Black</Color><Width>1pt</Width></Border></Style>
        </Textbox></CellContents></TablixHeader>
        <TablixMembers><TablixMember>
          <Group Name="RatingGroup"><GroupExpressions><GroupExpression>=Fields!Rating.Value</GroupExpression></GroupExpressions></Group>
          <TablixHeader><Size>1in</Size><CellContents><Textbox Name="RatingOwner"><CanGrow>true</CanGrow>
            <Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Rating.Value</Value><Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs>
            <Style>
              <Border><Style>None</Style></Border>
              <TopBorder><Style>=IIF(Fields!Number.Value = 1, "Solid", "None")</Style><Color>=IIF(Fields!Number.Value = 1, "#123456", "None")</Color><Width>2pt</Width></TopBorder>
              <LeftBorder><Style>Solid</Style><Color>Black</Color><Width>1pt</Width></LeftBorder>
              <RightBorder><Style>Solid</Style><Color>Black</Color><Width>1pt</Width></RightBorder>
              <BackgroundColor>Yellow</BackgroundColor>
            </Style>
          </Textbox></CellContents></TablixHeader>
          <TablixMembers><TablixMember><Group Name="Details"/></TablixMember></TablixMembers>
        </TablixMember></TablixMembers>
      </TablixMember></TablixMembers></TablixRowHierarchy>
      <DataSetName>D</DataSetName><Top>0.2in</Top><Left>0.2in</Left><Width>3in</Width><Height>0.7in</Height><Style/>
    </Tablix>
  </ReportItems><Height>2in</Height><Style/></Body><Width>7in</Width>
  <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`);

const request = {
  outputFileName: 'group-owner-border-context',
  parameters: {},
  datasets: { D: [
    { Risk: 'R1', Number: '001', Rating: 6, Detail: 'First action' },
    { Risk: 'R1', Number: '002', Rating: 6, Detail: 'Second action' },
  ] },
};

function tracedRatingCell(pdf, text, rowSpan) {
  return pdf.layoutTrace.pages[0].items.find((item) => (
    item.kind === 'tablixCell' && item.text === String(text) && item.rowSpan === rowSpan
  ));
}

async function docxCellXml(docxBuffer, text) {
  const documentXml = await (await JSZip.loadAsync(docxBuffer)).file('word/document.xml').async('string');
  const textIndex = documentXml.indexOf(`>${text}</w:t>`);
  assert.ok(textIndex >= 0, `expected ${text} in editable DOCX`);
  return documentXml.slice(
    documentXml.lastIndexOf('<w:tc>', textIndex),
    documentXml.indexOf('</w:tc>', textIndex) + '</w:tc>'.length,
  );
}

async function excelRatingCell(buffer, value) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  let ratingCell = null;
  workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => {
    if ((cell.value === value || cell.value === String(value)) && cell.fill?.fgColor?.argb === 'FFFFFF00') {
      ratingCell = cell;
    }
  }));
  assert.ok(ratingCell, `expected ${value} rating cell in XLSX`);
  return ratingCell;
}

test('a row-spanned group owner keeps its first-row context for conditional borders', () => {
  const tablix = model.body.items.find((item) => item.name === 'GroupedRisk');
  const rows = tablixRows(tablix, request, {}, model).rows;
  const owner = rows[0].cells[1];
  const textbox = owner.items[0];
  const context = { fields: owner.fields, dataset: owner.scopeDataset, scopes: owner.scopes };
  assert.equal(owner.rowSpan, 2);
  assert.equal(owner.fields.Number, '001');
  assert.equal(styleValue(textbox.style.borders.top.style, context), 'Solid');
  assert.equal(styleColor(textbox.style.borders.top.color, context), '#123456');
});

test('an omitted merged-owner top remains absent without a prior visual owner', async () => {
  const bridgedRequest = {
    ...request,
    datasets: { D: request.datasets.D.map((row) => ({ ...row, Number: '002' })) },
  };
  const pdf = await renderPdf(model, bridgedRequest, config, { captureLayoutTrace: true });
  const tracedOwner = pdf.layoutTrace.pages[0].items.find((item) => (
    item.kind === 'tablixCell' && item.text === '6' && item.rowSpan === 2
  ));
  assert.equal(tracedOwner?.borders?.top, null);

  const docx = await renderEditableDocx(model, bridgedRequest, config);
  const documentXml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  const ratingTextIndex = documentXml.indexOf('>6</w:t>');
  assert.ok(ratingTextIndex >= 0);
  const ratingCellXml = documentXml.slice(
    documentXml.lastIndexOf('<w:tc>', ratingTextIndex),
    documentXml.indexOf('</w:tc>', ratingTextIndex) + '</w:tc>'.length,
  );
  assert.match(ratingCellXml, /<w:top w:val="none" w:color="auto" w:sz="0"\/>/);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load((await renderExcel(model, bridgedRequest, config, null)).buffer);
  let ratingCell = null;
  workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => {
    if ((cell.value === 6 || cell.value === '6') && cell.fill?.fgColor?.argb === 'FFFFFF00') ratingCell = cell;
  }));
  assert.ok(ratingCell);
  assert.equal(ratingCell.border?.top, undefined);
});

test('an unchanged merged owner does not gain an internal row boundary', async () => {
  const unchangedRequest = {
    ...request,
    datasets: { D: [
      { Risk: 'R1', Number: '002', Rating: 6, Detail: 'Previous group' },
      { Risk: 'R2', Number: '002', Rating: 6, Detail: 'First action' },
      { Risk: 'R2', Number: '003', Rating: 6, Detail: 'Second action' },
    ] },
  };
  const pdf = await renderPdf(model, unchangedRequest, config, { captureLayoutTrace: true });
  const tracedOwner = pdf.layoutTrace.pages[0].items.find((item) => (
    item.kind === 'tablixCell' && item.text === '6' && item.rowSpan === 2
  ));
  assert.equal(tracedOwner?.borders?.top, null);
});

test('an unchanged single-row owner does not gain an internal row boundary', async () => {
  const unchangedRequest = {
    ...request,
    datasets: { D: [
      { Risk: 'R1', Number: '002', Rating: 6, Detail: 'Previous group' },
      { Risk: 'R2', Number: '002', Rating: 6, Detail: 'Current group' },
    ] },
  };
  const pdf = await renderPdf(model, unchangedRequest, config, { captureLayoutTrace: true });
  const tracedOwners = pdf.layoutTrace.pages[0].items.filter((item) => (
    item.kind === 'tablixCell' && item.text === '6' && item.rowSpan === 1
  ));
  assert.equal(tracedOwners.length, 2);
  assert.equal(tracedOwners[1]?.borders?.top, null);
});

test('PDF, editable DOCX, and XLSX bridge every changed group-owner span transition', async () => {
  const cases = [
    {
      name: 'single-to-single',
      target: 6,
      rowSpan: 1,
      rows: [
        { Risk: 'R1', Number: '002', Rating: 'Not Assessed', Detail: 'Previous group' },
        { Risk: 'R2', Number: '002', Rating: 6, Detail: 'Current group' },
      ],
    },
    {
      name: 'single-to-merged',
      target: 6,
      rowSpan: 2,
      rows: [
        { Risk: 'R1', Number: '002', Rating: 'Not Assessed', Detail: 'Previous group' },
        { Risk: 'R2', Number: '002', Rating: 6, Detail: 'First action' },
        { Risk: 'R2', Number: '003', Rating: 6, Detail: 'Second action' },
      ],
    },
    {
      name: 'merged-to-single',
      target: 7,
      rowSpan: 1,
      rows: [
        { Risk: 'R1', Number: '002', Rating: 'Not Assessed', Detail: 'First previous action' },
        { Risk: 'R1', Number: '003', Rating: 'Not Assessed', Detail: 'Second previous action' },
        { Risk: 'R2', Number: '002', Rating: 7, Detail: 'Current group' },
      ],
    },
  ];

  for (const scenario of cases) {
    const changedRequest = { ...request, datasets: { D: scenario.rows } };
    const pdf = await renderPdf(model, changedRequest, config, { captureLayoutTrace: true });
    const tracedOwner = tracedRatingCell(pdf, scenario.target, scenario.rowSpan);
    assert.deepEqual(
      tracedOwner?.borders?.top,
      { style: 'Solid', width: 1, color: '#000000' },
      `${scenario.name} PDF boundary`,
    );

    const docx = await renderEditableDocx(model, changedRequest, config);
    assert.match(
      await docxCellXml(docx.buffer, scenario.target),
      /<w:top w:val="single" w:color="000000" w:sz="8"\/>/,
      `${scenario.name} editable DOCX boundary`,
    );

    const ratingCell = await excelRatingCell(
      (await renderExcel(model, changedRequest, config, null)).buffer,
      scenario.target,
    );
    assert.equal(ratingCell.border?.top?.style, 'thin', `${scenario.name} XLSX boundary style`);
    assert.equal(ratingCell.border?.top?.color?.argb, 'FF000000', `${scenario.name} XLSX boundary color`);
  }
});

test('boundary inference rejects detail cells and mismatched neighbouring edges', () => {
  const header = (rowIndex, visual, border = null) => ({
    rowIndex,
    cell: { isRowHeader: true },
    visual,
    border,
  });
  const detail = (rowIndex, visual, border = null) => ({
    rowIndex,
    cell: { isRowHeader: false },
    visual,
    border,
  });
  const black = { style: 'thin', color: { argb: 'FF000000' } };
  const red = { style: 'thin', color: { argb: 'FFFF0000' } };
  const infer = (owner, above, leftBorder, rightBorder) => matchingChangedGroupOwnerRowBoundary(
    owner,
    above,
    { rowIndex: owner.rowIndex, border: leftBorder },
    { rowIndex: owner.rowIndex, border: rightBorder },
    (candidate) => candidate?.border,
    (border) => `${border.style}|${border.color.argb}`,
    (candidate) => candidate.visual,
  );

  assert.equal(infer(detail(1, 'new'), header(0, 'old'), black, black), null);
  assert.equal(infer(header(1, 'new'), detail(0, 'old'), black, black), null);
  assert.equal(infer(header(1, 'new'), header(0, 'old'), black, red), null);
  assert.equal(infer(header(1, 'same'), header(0, 'same'), black, black), null);
  assert.equal(infer(header(1, 'new'), header(0, 'old'), black, black), black);
});

test('an explicit single-row owner edge remains authoritative across PDF, editable DOCX, and XLSX', async () => {
  const explicitRequest = {
    ...request,
    datasets: { D: [
      { Risk: 'R1', Number: '002', Rating: 'Not Assessed', Detail: 'Previous group' },
      { Risk: 'R2', Number: '001', Rating: 6, Detail: 'Current group' },
    ] },
  };
  const pdf = await renderPdf(model, explicitRequest, config, { captureLayoutTrace: true });
  assert.deepEqual(
    tracedRatingCell(pdf, 6, 1)?.borders?.top,
    { style: 'Solid', width: 2, color: '#123456' },
  );

  const docx = await renderEditableDocx(model, explicitRequest, config);
  assert.match(await docxCellXml(docx.buffer, 6), /<w:top w:val="single" w:color="123456" w:sz="16"\/>/);

  const ratingCell = await excelRatingCell(
    (await renderExcel(model, explicitRequest, config, null)).buffer,
    6,
  );
  assert.equal(ratingCell.border?.top?.style, 'thin');
  assert.equal(ratingCell.border?.top?.color?.argb, 'FF123456');
});

test('PDF, editable DOCX, and XLSX preserve the conditional top border of a grouped owner cell', async () => {
  const pdf = await renderPdf(model, request, config, { captureLayoutTrace: true });
  const tracedOwner = pdf.layoutTrace.pages[0].items.find((item) => (
    item.kind === 'tablixCell' && item.text === '6' && item.rowSpan === 2
  ));
  assert.deepEqual(tracedOwner?.borders?.top, { style: 'Solid', width: 2, color: '#123456' });

  const docx = await renderEditableDocx(model, request, config);
  const documentXml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  assert.match(documentXml, /<w:top w:val="single" w:color="123456" w:sz="16"\/>/);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load((await renderExcel(model, request, config, null)).buffer);
  let ratingCell = null;
  workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => {
    if ((cell.value === 6 || cell.value === '6') && cell.fill?.fgColor?.argb === 'FFFFFF00') ratingCell = cell;
  }));
  assert.ok(ratingCell, 'the grouped rating cell must remain native and editable');
  assert.equal(ratingCell.border?.top?.style, 'thin');
  assert.equal(ratingCell.border?.top?.color?.argb, 'FF123456');
});
