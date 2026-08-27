// An XLSX worksheet is assembled by several passes, and a later pass discarded what an earlier one had
// resolved. A Rectangle paints its own outline over the region its children occupy; that pass ASSIGNED
// each perimeter side, so a Rectangle declaring Border/Style=None wrote `undefined` over the edges its
// children had already placed there — erasing an enclosed tablix's own outer rule wherever the two
// perimeters coincide, visibly the closing bottom rule of the table's last row. A container border must
// add an edge, never remove one.
//
// This is XLSX-specific: PDF strokes borders as geometry from the same resolved model, and DOCX inherits
// the canonical PDF trace. The PDF assertion below is the counterexample proving that.
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { promisify } from 'node:util';
import ExcelJS from 'exceljs';
import { PNG } from 'pngjs';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

const execFileAsync = promisify(execFile);
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const pdfTextTool = /pdftoppm(?:\.exe)?$/i.test(config.pdftoppmPath)
  ? path.join(path.dirname(config.pdftoppmPath), process.platform === 'win32' ? 'pdftotext.exe' : 'pdftotext')
  : 'pdftotext';
const request = { outputFileName: 'excel-container-border', parameters: {}, datasets: { Rows: [{ Label: 'only' }] } };

const SOLID = '<Border><Style>Solid</Style><Color>Black</Color><Width>1pt</Width></Border>';

function cellTextbox(name, value) {
  return `<Textbox Name="${name}">
    <Paragraphs><Paragraph><TextRuns><TextRun><Value>${value}</Value></TextRun></TextRuns></Paragraph></Paragraphs>
    <Style><FontFamily>Arial</FontFamily>${SOLID}</Style>
  </Textbox>`;
}

// `rectangleBorder` is the only difference between the two border scenarios: a container that declares no
// border must leave the enclosed table's own rule alone, and one that declares a border must add its own.
function report({ rectangleBorder = '<Border><Style>None</Style></Border>', heading = 'H' } = {}) {
  return Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSources><DataSource Name="S"><DataSourceReference>/x</DataSourceReference></DataSource></DataSources>
  <DataSets><DataSet Name="Rows">
    <Fields><Field Name="Label"><DataField>Label</DataField><TypeName>System.String</TypeName></Field></Fields>
    <Query><DataSourceName>S</DataSourceName><CommandText>ignored</CommandText></Query>
  </DataSet></DataSets>
  <ReportSections><ReportSection>
    <Body>
      <ReportItems>
        <Rectangle Name="Wrapper">
          <ReportItems>
            <Tablix Name="Grid">
              <TablixBody>
                <TablixColumns><TablixColumn><Width>2in</Width></TablixColumn><TablixColumn><Width>2in</Width></TablixColumn></TablixColumns>
                <TablixRows>
                  <TablixRow><Height>0.25in</Height><TablixCells>
                    <TablixCell><CellContents>${cellTextbox('HeadA', 'HEAD_A')}</CellContents></TablixCell>
                    <TablixCell><CellContents>${cellTextbox('HeadB', 'HEAD_B')}</CellContents></TablixCell>
                  </TablixCells></TablixRow>
                  <TablixRow><Height>0.25in</Height><TablixCells>
                    <TablixCell><CellContents>${cellTextbox('LastA', 'LAST_A')}</CellContents></TablixCell>
                    <TablixCell><CellContents>${cellTextbox('LastB', 'LAST_B')}</CellContents></TablixCell>
                  </TablixCells></TablixRow>
                </TablixRows>
              </TablixBody>
              <TablixColumnHierarchy><TablixMembers><TablixMember/><TablixMember/></TablixMembers></TablixColumnHierarchy>
              <TablixRowHierarchy><TablixMembers><TablixMember/><TablixMember/></TablixMembers></TablixRowHierarchy>
              <DataSetName>Rows</DataSetName>
              <Top>0in</Top><Left>0in</Left><Height>0.5in</Height><Width>4in</Width>
              <Style><FontFamily>Arial</FontFamily>${SOLID}</Style>
            </Tablix>
          </ReportItems>
          <Top>0in</Top><Left>0in</Left><Height>0.5in</Height><Width>4in</Width>
          <Style>${rectangleBorder}</Style>
        </Rectangle>
        <Textbox Name="Heading">
          <CanGrow>true</CanGrow>
          <Paragraphs><Paragraph><TextRuns><TextRun><Value>${heading}</Value></TextRun></TextRuns></Paragraph></Paragraphs>
          <Top>0.7in</Top><Left>0in</Left><Height>0.2in</Height><Width>1.6in</Width>
          <Style><FontFamily>Arial</FontFamily><FontSize>11pt</FontSize><FontWeight>Bold</FontWeight></Style>
        </Textbox>
      </ReportItems>
      <Height>1.2in</Height><Style/>
    </Body>
    <Width>7in</Width>
    <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`, 'utf8');
}

async function worksheet(options) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-xlsx-border-'));
  try {
    const rendered = await renderExcel(parseRdl(report(options)), { ...request, excelLayoutMode: 'REPORT' }, config, tempDir);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(rendered.buffer);
    return workbook.worksheets[0];
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

function locate(sheet, text) {
  let found = null;
  sheet.eachRow((row, rowNumber) => row.eachCell((cell, columnNumber) => {
    const value = cell.value?.richText
      ? cell.value.richText.map((run) => run.text).join('')
      : String(cell.value ?? '');
    if (!found && value.includes(text)) found = { rowNumber, columnNumber, cell, row };
  }));
  assert.ok(found, `expected a cell containing ${text}`);
  return found;
}

test('a borderless container leaves the enclosed table its own closing rules', async () => {
  const sheet = await worksheet();
  const first = locate(sheet, 'HEAD_A');
  const last = locate(sheet, 'LAST_A');
  assert.ok(last.rowNumber > first.rowNumber, 'the table must occupy at least two worksheet rows');
  // The tablix and the wrapping rectangle share a perimeter, so these are exactly the edges the container
  // pass used to overwrite.
  assert.ok(last.cell.border?.bottom, 'the last table row must keep its declared bottom rule');
  assert.ok(first.cell.border?.top, 'the first table row must keep its declared top rule');
  assert.ok(last.cell.border?.left && last.cell.border?.right);
});

test('a container that declares its own border still adds it', async () => {
  const sheet = await worksheet({ rectangleBorder: '<Border><Style>Solid</Style><Color>Red</Color><Width>1pt</Width></Border>' });
  const last = locate(sheet, 'LAST_A');
  assert.ok(last.cell.border?.bottom, 'a declared container border must be present, not suppressed');
});

test('a picture flush with a bordered Rectangle is inset so it cannot cover the container rule', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-xlsx-picture-border-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const png = new PNG({ width: 2, height: 2 });
  png.data.fill(0xFF);
  const imageData = PNG.sync.write(png).toString('base64');
  const pictureReport = Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <EmbeddedImages><EmbeddedImage Name="Tile"><MIMEType>image/png</MIMEType><ImageData>${imageData}</ImageData></EmbeddedImage></EmbeddedImages>
  <ReportSections><ReportSection><Body><ReportItems>
    <Rectangle Name="Frame"><ReportItems>
      <Image Name="Picture"><Source>Embedded</Source><Value>Tile</Value><Sizing>Fit</Sizing>
        <Top>0pt</Top><Left>0pt</Left><Height>100pt</Height><Width>200pt</Width><Style/></Image>
    </ReportItems><Top>0pt</Top><Left>0pt</Left><Height>100pt</Height><Width>200pt</Width>
      <Style><Border><Style>Solid</Style><Color>Black</Color><Width>2pt</Width></Border></Style></Rectangle>
  </ReportItems><Height>120pt</Height><Style/></Body><Width>220pt</Width>
  <Page><PageWidth>300pt</PageWidth><PageHeight>300pt</PageHeight><TopMargin>10pt</TopMargin><BottomMargin>10pt</BottomMargin><LeftMargin>10pt</LeftMargin><RightMargin>10pt</RightMargin></Page>
  </ReportSection></ReportSections>
</Report>`, 'utf8');
  const rendered = await renderExcel(parseRdl(pictureReport), { outputFileName: 'picture-frame', parameters: {}, datasets: {}, excelLayoutMode: 'REPORT' }, config, tempDir);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(rendered.buffer);
  const sheet = workbook.worksheets[0];
  const [picture] = sheet.getImages();
  assert.ok(picture, 'the embedded image must be present');
  assert.ok(picture.range.tl.col > 0, 'left parent border remains outside the floating picture');
  assert.ok(picture.range.tl.row > 0, 'top parent border remains outside the floating picture');
  assert.ok(picture.range.br.col < 1, 'right parent border remains outside the floating picture');
  assert.ok(picture.range.br.row < 1, 'bottom parent border remains outside the floating picture');
  const anchor = sheet.getCell(1, 1);
  assert.ok(anchor.border?.top && anchor.border?.left, 'the parent Rectangle outline is retained in the worksheet cells');
});

test('PDF is unaffected: it strokes both rules as geometry either way', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-xlsx-border-pdf-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const pdfPath = path.join(tempDir, 'borders.pdf');
  const rendered = await renderPdf(parseRdl(report()), request, config);
  await fs.writeFile(pdfPath, rendered.buffer);
  const { stdout } = await execFileAsync(pdfTextTool, ['-layout', pdfPath, '-']);
  assert.match(stdout, /HEAD_A/);
  assert.match(stdout, /LAST_A/);
});
