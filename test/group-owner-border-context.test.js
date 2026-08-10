import assert from 'node:assert/strict';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { styleColor, styleValue, tablixRows } from '../src/render/common.js';
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

test('matching row-boundary neighbours close a vertically merged owner whose own top is None', async () => {
  const bridgedRequest = {
    ...request,
    datasets: { D: request.datasets.D.map((row) => ({ ...row, Number: '002' })) },
  };
  const pdf = await renderPdf(model, bridgedRequest, config, { captureLayoutTrace: true });
  const tracedOwner = pdf.layoutTrace.pages[0].items.find((item) => (
    item.kind === 'tablixCell' && item.text === '6' && item.rowSpan === 2
  ));
  assert.deepEqual(tracedOwner?.borders?.top, { style: 'Solid', width: 1, color: '#000000' });

  const docx = await renderEditableDocx(model, bridgedRequest, config);
  const documentXml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  assert.match(documentXml, /<w:top w:val="single" w:color="000000" w:sz="8"\/>/);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load((await renderExcel(model, bridgedRequest, config, null)).buffer);
  let ratingCell = null;
  workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => {
    if ((cell.value === 6 || cell.value === '6') && cell.fill?.fgColor?.argb === 'FFFFFF00') ratingCell = cell;
  }));
  assert.ok(ratingCell);
  assert.equal(ratingCell.border?.top?.style, 'thin');
  assert.equal(ratingCell.border?.top?.color?.argb, 'FF000000');
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
