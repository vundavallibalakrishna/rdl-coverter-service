// Report Builder emits `<FontFamily></FontFamily>` for a style that names no family. RDL treats an empty
// element exactly like an omitted one - the family is inherited, and at the top of the chain that is the
// report default - but the parser's textValue() only substitutes its fallback for a MISSING element, so a
// blank one reached the model as a declared family literally named "".
//
// That is a shared normalization defect, not a per-format one, and it broke every output built on the model:
// strict PDF font resolution demanded a font with no name and failed the whole export with
// `FONT_MISSING: Required font is unavailable: ` (503), and the Word/Excel writers put an empty typeface
// into their runs. Fixing it in the parser fixes PDF, DOCX_EDITABLE, DOCX_VISUAL (which rasterizes that
// same canonical PDF) and XLSX at once; the tests below assert the format-appropriate evidence for each.
//
// The counterexamples matter as much as the fix: a family that IS named must still be carried through
// verbatim and must still fail closed under strict fonts when the host lacks it. Blank means "inherit",
// never "substitute whatever is convenient".
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { promisify } from 'node:util';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';
import { resolveFontFile } from '../src/render/fonts.js';

const execFileAsync = promisify(execFile);
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const strictConfig = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'true' });
const request = { outputFileName: 'blank-font-family', parameters: {}, datasets: { Rows: [{ Label: 'only' }] } };

// `family` is the only variable: the blank element is the defect, a named one and an absent one are the
// counterexamples that prove the rule is about blankness rather than about this report.
function report({ family = '', defaultFamily = null } = {}) {
  const declaration = family === null ? '' : `<FontFamily>${family}</FontFamily>`;
  const reportDefault = defaultFamily === null ? '' : `<df:DefaultFontFamily>${defaultFamily}</df:DefaultFontFamily>`;
  return Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition"
        xmlns:df="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition/defaultfontfamily">
  ${reportDefault}
  <DataSources><DataSource Name="S"><DataSourceReference>/x</DataSourceReference></DataSource></DataSources>
  <DataSets><DataSet Name="Rows">
    <Fields><Field Name="Label"><DataField>Label</DataField><TypeName>System.String</TypeName></Field></Fields>
    <Query><DataSourceName>S</DataSourceName><CommandText>ignored</CommandText></Query>
  </DataSet></DataSets>
  <ReportSections><ReportSection>
    <Body>
      <ReportItems>
        <Tablix Name="Grid">
          <TablixBody>
            <TablixColumns><TablixColumn><Width>3in</Width></TablixColumn></TablixColumns>
            <TablixRows>
              <TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents>
                <Textbox Name="Only">
                  <Paragraphs><Paragraph><TextRuns><TextRun><Value>BLANK_FAMILY_CELL</Value></TextRun></TextRuns></Paragraph></Paragraphs>
                  <Style>${declaration}<FontSize>11pt</FontSize></Style>
                </Textbox>
              </CellContents></TablixCell></TablixCells></TablixRow>
            </TablixRows>
          </TablixBody>
          <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
          <TablixRowHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixRowHierarchy>
          <DataSetName>Rows</DataSetName>
          <Top>0in</Top><Left>0in</Left><Height>0.25in</Height><Width>3in</Width>
          <Style>${declaration}</Style>
        </Tablix>
      </ReportItems>
      <Height>1in</Height><Style/>
    </Body>
    <Width>7in</Width>
    <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`, 'utf8');
}

function onlyTextbox(model) {
  return model.body.items.find((item) => item.type === 'Tablix').rows[0].cells[0].items
    .find((item) => item.type === 'Textbox');
}

test('a blank FontFamily normalizes to the report default, never to an unnamed family', () => {
  const model = parseRdl(report({ family: '' }));
  assert.equal(model.defaultFontFamily, 'Arial');
  assert.equal(onlyTextbox(model).style.fontFamily, 'Arial');
  assert.equal(model.body.items.find((item) => item.type === 'Tablix').style.fontFamily, 'Arial');
});

test('a whitespace-only FontFamily is blank too', () => {
  assert.equal(onlyTextbox(parseRdl(report({ family: '   ' }))).style.fontFamily, 'Arial');
});

test('a blank FontFamily inherits the report DefaultFontFamily when one is declared', () => {
  const model = parseRdl(report({ family: '', defaultFamily: 'Times New Roman' }));
  assert.equal(model.defaultFontFamily, 'Times New Roman');
  assert.equal(onlyTextbox(model).style.fontFamily, 'Times New Roman');
});

test('a blank df:DefaultFontFamily falls back to Arial rather than becoming the unnamed default', () => {
  const model = parseRdl(report({ family: '', defaultFamily: '' }));
  assert.equal(model.defaultFontFamily, 'Arial');
  assert.equal(onlyTextbox(model).style.fontFamily, 'Arial');
});

test('counterexample: a declared family is still carried through verbatim', () => {
  assert.equal(onlyTextbox(parseRdl(report({ family: 'Times New Roman' }))).style.fontFamily, 'Times New Roman');
});

test('counterexample: an omitted Style/FontFamily behaves exactly like a blank one', () => {
  assert.equal(onlyTextbox(parseRdl(report({ family: null }))).style.fontFamily, 'Arial');
});

test('a blank FontFamily never enters the declared-font catalogue strict checks are driven from', () => {
  // "" in this list is what made /readyz and the render font check demand a font with no name.
  assert.deepEqual(parseRdl(report({ family: '' })).fonts, ['Arial']);
  assert.deepEqual(parseRdl(report({ family: 'Times New Roman' })).fonts.sort(), ['Arial', 'Times New Roman']);
});

test('PDF: a blank FontFamily renders under strict fonts instead of failing FONT_MISSING', async (context) => {
  if (!resolveFontFile(strictConfig.fontDir, 'Arial', false, false)) {
    context.skip('this host has no Arial to resolve the inherited default to');
    return;
  }
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-blank-font-pdf-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  // Previously: ServiceError FONT_MISSING "Required font is unavailable: " (503), with an empty family name.
  const rendered = await renderPdf(parseRdl(report({ family: '' })), request, strictConfig);
  const pdfPath = path.join(tempDir, 'blank-family.pdf');
  await fs.writeFile(pdfPath, rendered.buffer);
  const { stdout } = await execFileAsync('pdftotext', ['-layout', pdfPath, '-']);
  assert.match(stdout, /BLANK_FAMILY_CELL/);
});

test('DOCX_EDITABLE: the run names the inherited family, not an empty typeface', async () => {
  const rendered = await renderEditableDocx(parseRdl(report({ family: '' })), request, config);
  const xml = await (await JSZip.loadAsync(rendered.buffer)).file('word/document.xml').async('string');
  assert.match(xml, /BLANK_FAMILY_CELL/);
  // Word falls back to its own default for an empty ascii/hAnsi, which is not the family the RDL inherits.
  assert.doesNotMatch(xml, /w:rFonts[^>]*w:ascii=""/);
  assert.match(xml, /w:rFonts[^>]*w:ascii="Arial"/);
});

test('XLSX: the cell font names the inherited family, not an empty typeface', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-blank-font-xlsx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderExcel(parseRdl(report({ family: '' })), { ...request, excelLayoutMode: 'REPORT' }, config, tempDir);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(rendered.buffer);
  let found = null;
  workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => {
    const text = cell.value?.richText ? cell.value.richText.map((run) => run.text).join('') : String(cell.value ?? '');
    if (!found && text.includes('BLANK_FAMILY_CELL')) found = cell;
  }));
  assert.ok(found, 'expected the rendered cell');
  assert.equal(found.font?.name, 'Arial');
});
