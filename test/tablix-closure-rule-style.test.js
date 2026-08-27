import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { enforcedBottomBorder, tablixRows } from '../src/render/common.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

// The synthetic closing rule a data tablix gets when its own bottom edge is None must LOOK like the rule
// the table draws everywhere else. It was hardcoded to Solid, so the common RDL idiom — a Border=None
// tablix whose cells all declare the same Dotted rule — closed each page fragment with a solid black bar.
// Where the page footer carries its own rule just below the body boundary the two fused into one thick
// line, which is how the defect surfaced. The style is a shared model semantic, so PDF, Word and Excel all
// have to close the table the same way.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000' });

// A tablix that declares Border=None on itself and carries its grid entirely on its cells, plus a page
// footer whose own rule sits just inside the footer band — the geometry that made the fused bar visible.
const rdl = (cellRule, tablixBorder = '<Border><Style>None</Style></Border>') => `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields><Field Name="V"><DataField>V</DataField></Field></Fields>
   <Query><CommandText>never executed</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="T"><TablixBody><TablixColumns><TablixColumn><Width>3in</Width></TablixColumn></TablixColumns>
   <TablixRows>
    <TablixRow><Height>0.3in</Height><TablixCells><TablixCell><CellContents>
      <Textbox Name="h"><Paragraphs><Paragraph><TextRuns><TextRun><Value>COLHDR</Value>
        <Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs>
        <Style><Border><Style>None</Style></Border><BottomBorder>${cellRule}</BottomBorder></Style></Textbox>
    </CellContents></TablixCell></TablixCells></TablixRow>
    <TablixRow><Height>0.3in</Height><TablixCells><TablixCell><CellContents>
      <Textbox Name="d"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!V.Value</Value>
        <Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs>
        <Style><Border><Style>None</Style></Border><BottomBorder>${cellRule}</BottomBorder></Style></Textbox>
    </CellContents></TablixCell></TablixCells></TablixRow>
   </TablixRows></TablixBody>
   <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
   <TablixRowHierarchy><TablixMembers><TablixMember/><TablixMember>
     <Group Name="g"><GroupExpressions><GroupExpression>=Fields!V.Value</GroupExpression></GroupExpressions></Group>
   </TablixMember></TablixMembers></TablixRowHierarchy>
   <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Width>3in</Width>
   <Style>${tablixBorder}</Style></Tablix>
 </ReportItems><Height>9in</Height><Style/></Body>
 <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth>
   <TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin>
   <LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin>
   <PageFooter><Height>0.3in</Height><PrintOnFirstPage>true</PrintOnFirstPage><PrintOnLastPage>true</PrintOnLastPage>
     <ReportItems><Line Name="FooterRule"><Top>0.0125in</Top><Left>0in</Left><Height>0in</Height><Width>7.5in</Width>
       <Style><Border><Style>Solid</Style><Color>#000000</Color><Width>1pt</Width></Border></Style></Line></ReportItems>
     <Style/></PageFooter>
 </Page></ReportSection></ReportSections></Report>`;

const DOTTED = '<Style>Dotted</Style><Color>Black</Color><Width>1pt</Width>';
const SOLID = '<Style>Solid</Style><Color>Black</Color><Width>1pt</Width>';

// Enough rows to force several page fragments, so the closure under test is a page cut and not only the
// table's true end.
const request = (rowCount = 120) => ({
  outputFileName: 'tablix-closure-rule-style',
  parameters: {},
  datasets: { D: Array.from({ length: rowCount }, (_, index) => ({ V: `Row ${index}` })) },
  excel: { layoutMode: 'REPORT' },
});

const fragmentBottoms = (trace) => trace.pages.flatMap((page) => (page.items || []).filter((item) => (
  item.traceRole === 'resolvedTablixFragmentBorder' && item.fragmentSide === 'bottom'
)));

async function ownedConfig(context, prefix) {
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), prefix));
  context.after(() => fs.rm(tempRoot, { recursive: true, force: true }));
  return {
    tempRoot,
    config: loadConfig({
      ...process.env,
      RDL_STRICT_FONTS: 'false',
      RDL_RENDER_TIMEOUT_MS: '60000',
      RDL_TEMP_ROOT: tempRoot,
    }),
  };
}

test('the synthesized closure adopts the rule the table itself draws', () => {
  const model = parseRdl(rdl(DOTTED));
  const tablix = model.body.items.find((item) => item.type === 'Tablix');
  const { rows } = tablixRows(tablix, request(3), { PageNumber: 1, TotalPages: 1 }, model);
  const closure = enforcedBottomBorder(tablix.style, rows, tablix);
  assert.equal(closure.style, 'Dotted', 'a dotted grid must not close with a solid bar');
  assert.equal(String(closure.width), '1');

  // Counterexample: the same construct with a solid grid still closes solid.
  const solidModel = parseRdl(rdl(SOLID));
  const solidTablix = solidModel.body.items.find((item) => item.type === 'Tablix');
  const solidRows = tablixRows(solidTablix, request(3), { PageNumber: 1, TotalPages: 1 }, solidModel).rows;
  assert.equal(enforcedBottomBorder(solidTablix.style, solidRows, solidTablix).style, 'Solid');

  // A tablix that declares a side of its OWN keeps that side's line style, not a forced Solid.
  const dashed = { borders: { bottom: { style: 'None' }, left: { style: 'Dashed', color: '#123456', width: 2 } } };
  assert.deepEqual(enforcedBottomBorder(dashed), { style: 'Dashed', color: '#123456', width: 2 });

  // Nothing declared anywhere still falls back to a plain black rule.
  assert.deepEqual(enforcedBottomBorder({}, [], null), { style: 'Solid', color: '#000000', width: 1 });
});

test('every PDF page fragment closes with the table’s own rule, clear of the footer rule', async () => {
  const model = parseRdl(rdl(DOTTED));
  const rendered = await renderPdf(model, request(), config, { captureLayoutTrace: true });
  const bottoms = fragmentBottoms(rendered.layoutTrace);
  assert.ok(bottoms.length > 1, `expected a closure on each of several page fragments, got ${bottoms.length}`);
  for (const bottom of bottoms) {
    assert.equal(bottom.line.style, 'Dotted', 'a page fragment must close with the table’s own rule');
  }
  // The footer's own rule stays a separate, solid line: fusing the two into one bar is the reported defect.
  const footerRules = rendered.layoutTrace.pages
    .flatMap((page) => (page.items || []).filter((item) => item.itemName === 'FooterRule'));
  assert.ok(footerRules.length > 0, 'the report footer draws its own rule');
  for (const rule of footerRules) assert.equal(rule.line.style, 'Solid');
});

test('editable DOCX closes the fragment with the same rule the PDF drew', async (context) => {
  const owned = await ownedConfig(context, 'rdl-closure-rule-docx-');
  const docx = await renderEditableDocx(parseRdl(rdl(DOTTED)), request(), owned.config);
  const xml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  const styles = new Map();
  for (const match of xml.matchAll(/<w:(?:top|bottom)\b[^>]*w:val="([a-z]+)"/g)) {
    styles.set(match[1], (styles.get(match[1]) || 0) + 1);
  }
  assert.ok((styles.get('dotted') || 0) > 0, 'the table grid must reach Word as dotted borders');
  // Nothing in this table's body is solid, so a solid horizontal cell rule can only be the old closure.
  assert.equal(styles.get('single') || 0, 0, `unexpected solid horizontal rules in the Word table: ${[...styles]}`);
});

test('XLSX closes the table with the same rule too', async (context) => {
  const owned = await ownedConfig(context, 'rdl-closure-rule-xlsx-');
  const xlsx = await renderExcel(parseRdl(rdl(DOTTED)), request(), owned.config, owned.tempRoot);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(xlsx.buffer);
  const sheet = workbook.worksheets[0];
  const styles = new Set();
  let lastBordered = 0;
  sheet.eachRow((row, rowNumber) => row.eachCell((cell) => {
    for (const side of ['top', 'bottom']) {
      if (!cell.border?.[side]) continue;
      styles.add(cell.border[side].style);
      lastBordered = Math.max(lastBordered, rowNumber);
    }
  }));
  assert.ok(lastBordered > 0, 'the table draws horizontal rules in Excel');
  assert.deepEqual([...styles], ['dotted'], `expected only dotted rules, got ${[...styles]}`);
});
