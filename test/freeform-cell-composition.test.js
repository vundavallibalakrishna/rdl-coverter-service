// A TablixCell's CellContents holds exactly one RDL report item. When that item is a Rectangle, its
// children are free-form: each carries its own Top/Left/Width/Height inside the cell and SSRS draws it
// exactly there. Materialization flattens the Rectangle away, so the cell arrives at the renderers holding
// several positioned children.
//
// The PDF renderer treated the first of those children as "the cell's textbox" and stretched it over the
// whole cell. Two things followed. The label was drawn at the cell origin instead of its declared offset,
// which is simply the wrong place. And the traced cell then ENCLOSED the nested data region positioned
// beside the label, which is an overlap no native grid can represent: the page-locked Word renderer
// rejected the whole report with `UNSUPPORTED_FEATURE: Overlapping editable PDF regions cannot be
// represented safely as native Word cells`.
//
// The fix belongs in the canonical PDF layout, which is the shared authority: the composition's children
// keep their declared rects and the cell rectangle carries only fill and borders. DOCX_EDITABLE and
// DOCX_VISUAL inherit it from the trace and the canonical PDF respectively; XLSX builds its own grid from
// the same model and is asserted separately.
//
// The counterexample is the load-bearing half of this file: the ubiquitous single-textbox-in-a-Rectangle
// cell is NOT a composition — that child does fill its container — and must keep rendering exactly as it
// always has.
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

const execFileAsync = promisify(execFile);
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const request = { outputFileName: 'freeform-cell', parameters: {}, datasets: { Rows: [{ Label: 'only' }] } };

// Declared inside the cell's Rectangle, in points. The label deliberately sits inset from the cell origin
// on both axes so "drawn where it was declared" and "stretched to the cell" are distinguishable.
const LABEL_LEFT = 18;
const LABEL_TOP = 4;
const LABEL_WIDTH = 144;
const LABEL_HEIGHT = 14;
const NESTED_LEFT = 18;
const NESTED_TOP = 26;

function nestedTablix() {
  return `<Tablix Name="Inner">
    <TablixBody>
      <TablixColumns><TablixColumn><Width>1in</Width></TablixColumn></TablixColumns>
      <TablixRows><TablixRow><Height>0.2in</Height><TablixCells><TablixCell><CellContents>
        <Textbox Name="InnerCell">
          <Paragraphs><Paragraph><TextRuns><TextRun><Value>NESTED_CELL</Value></TextRun></TextRuns></Paragraph></Paragraphs>
          <Style><FontFamily>Arial</FontFamily><FontSize>8pt</FontSize></Style>
        </Textbox>
      </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
    </TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixRowHierarchy>
    <Top>${NESTED_TOP}pt</Top><Left>${NESTED_LEFT}pt</Left><Height>0.2in</Height><Width>1in</Width>
    <Style><FontFamily>Arial</FontFamily></Style>
  </Tablix>`;
}

// `composition` is the only difference between the two scenarios: a Rectangle holding a positioned label
// beside a nested data region, versus the same Rectangle wrapping that label alone.
function report({ composition = true } = {}) {
  const contents = composition
    ? `<Textbox Name="Label">
         <Paragraphs><Paragraph><TextRuns><TextRun><Value>LEGEND_LABEL</Value></TextRun></TextRuns></Paragraph></Paragraphs>
         <Top>${LABEL_TOP}pt</Top><Left>${LABEL_LEFT}pt</Left><Height>${LABEL_HEIGHT}pt</Height><Width>${LABEL_WIDTH}pt</Width>
         <Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize><Border><Style>None</Style></Border></Style>
       </Textbox>
       ${nestedTablix()}`
    : `<Textbox Name="Label">
         <Paragraphs><Paragraph><TextRuns><TextRun><Value>LEGEND_LABEL</Value></TextRun></TextRuns></Paragraph></Paragraphs>
         <Top>0pt</Top><Left>0pt</Left><Height>0.7in</Height><Width>4in</Width>
         <Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize><Border><Style>None</Style></Border></Style>
       </Textbox>`;
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
        <Tablix Name="Outer">
          <TablixBody>
            <TablixColumns><TablixColumn><Width>4in</Width></TablixColumn></TablixColumns>
            <TablixRows><TablixRow><Height>0.7in</Height><TablixCells><TablixCell><CellContents>
              <Rectangle Name="Holder">
                <ReportItems>${contents}</ReportItems>
                <Top>0in</Top><Left>0in</Left><Height>0.7in</Height><Width>4in</Width>
                <Style><Border><Style>None</Style></Border></Style>
              </Rectangle>
            </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
          </TablixBody>
          <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
          <TablixRowHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixRowHierarchy>
          <DataSetName>Rows</DataSetName>
          <Top>0in</Top><Left>0in</Left><Height>0.7in</Height><Width>4in</Width>
          <Style><FontFamily>Arial</FontFamily><Border><Style>Solid</Style><Color>Black</Color><Width>1pt</Width></Border></Style>
        </Tablix>
      </ReportItems>
      <Height>1.5in</Height><Style/>
    </Body>
    <Width>7in</Width>
    <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`, 'utf8');
}

async function trace(options) {
  const rendered = await renderPdf(parseRdl(report(options)), request, config, { captureLayoutTrace: true });
  return { rendered, page: rendered.layoutTrace.pages[0] };
}

const named = (page, name) => page.items.filter((entry) => entry.itemName === name);
const rightOf = (entry) => entry.x + entry.width;
const bottomOf = (entry) => entry.y + entry.height;

test('a composition draws its label at the rect it declares inside the cell', async () => {
  const { page } = await trace();
  const [label] = named(page, 'Label');
  const [cell] = named(page, 'InnerCell');
  assert.ok(label, 'the label must still be traced');
  assert.ok(cell, 'the nested data region must still be traced');
  // The cell origin: the tablix sits at the body origin, so the enclosing cell starts at the page margin.
  const cellOriginX = Math.min(label.x, cell.x) - Math.min(LABEL_LEFT, NESTED_LEFT);
  assert.ok(Math.abs(label.x - (cellOriginX + LABEL_LEFT)) < 0.5, `label x ${label.x}`);
  assert.ok(Math.abs(label.width - LABEL_WIDTH) < 0.5, `label width ${label.width}`);
  assert.ok(Math.abs(label.height - LABEL_HEIGHT) < 0.5, `label height ${label.height}`);
  // Previously the label was stretched to the whole 4in cell and swallowed the nested region below it.
  assert.ok(bottomOf(label) <= cell.y + 0.5, 'the label must end above the nested region, not enclose it');
});

test('a composition never encloses the data region positioned beside it', async () => {
  const { page } = await trace();
  const [label] = named(page, 'Label');
  const [cell] = named(page, 'InnerCell');
  const encloses = label.x <= cell.x && label.y <= cell.y
    && rightOf(label) >= rightOf(cell) && bottomOf(label) >= bottomOf(cell);
  assert.equal(encloses, false, 'the label must not contain the nested cell');
});

test('counterexample: a Rectangle wrapping one textbox still fills the cell', async () => {
  const { page } = await trace({ composition: false });
  const [label] = named(page, 'Label');
  assert.ok(label, 'the wrapped textbox must still be traced');
  // 4in cell. The single-child idiom is not a composition and keeps the cell-filling geometry it always had.
  assert.ok(Math.abs(label.width - 288) < 1, `wrapped textbox width ${label.width}`);
  assert.equal(label.kind, 'tablixCell', 'a wrapped single textbox is still the cell itself');
});

test('DOCX_EDITABLE: a composition renders as native content instead of failing closed', async () => {
  // Previously: UNSUPPORTED_FEATURE "Overlapping editable PDF regions cannot be represented safely as
  // native Word cells".
  const rendered = await renderEditableDocx(parseRdl(report()), request, config);
  const xml = await (await JSZip.loadAsync(rendered.buffer)).file('word/document.xml').async('string');
  assert.match(xml, /LEGEND_LABEL/);
  assert.match(xml, /NESTED_CELL/);
  assert.match(xml, /<w:tbl>/, 'the page must still be built from native Word tables');
  assert.doesNotMatch(xml, /<w:drawing>/, 'no part of the composition may become a picture');
});

test('PDF: both the label and the nested region remain selectable text', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-freeform-cell-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const { rendered } = await trace();
  const pdfPath = path.join(tempDir, 'composition.pdf');
  await fs.writeFile(pdfPath, rendered.buffer);
  const { stdout } = await execFileAsync('pdftotext', ['-layout', pdfPath, '-']);
  assert.match(stdout, /LEGEND_LABEL/);
  assert.match(stdout, /NESTED_CELL/);
});

test('XLSX: a composition places both children without colliding merged ranges', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-freeform-cell-xlsx-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const rendered = await renderExcel(parseRdl(report()), { ...request, excelLayoutMode: 'REPORT' }, config, tempDir);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(rendered.buffer);
  const texts = [];
  workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => {
    texts.push(cell.value?.richText ? cell.value.richText.map((run) => run.text).join('') : String(cell.value ?? ''));
  }));
  assert.ok(texts.some((value) => value.includes('LEGEND_LABEL')), 'the label must reach the worksheet');
  assert.ok(texts.some((value) => value.includes('NESTED_CELL')), 'the nested region must reach the worksheet');
});
