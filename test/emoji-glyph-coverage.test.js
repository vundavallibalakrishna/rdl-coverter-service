// A pictographic character the declared family cannot draw must degrade to .notdef — exactly as SSRS
// renders it on a host without that glyph — not fail the export. Font selection lives in one shared layer,
// so the same defect reached PDF (applyFont/measureTextboxHeight), DOCX_EDITABLE (the canonical layout
// trace), DOCX_VISUAL (rasterized from that PDF) and XLSX (shared textbox measurement) alike.
//
// Two independent triggers are pinned here:
//  - no installed face covers the pictograph, so the coverage ladder ends empty;
//  - the run also carries a layout control (a newline or tab), which no font has a glyph for, so judging
//    the raw string reported EVERY multi-line textbox as uncovered and defeated the ladder even on a host
//    that does have an emoji face.
import assert from 'node:assert/strict';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';
import { pdfFont, resolveFontFile, takeFontSubstitutions } from '../src/render/fonts.js';

// U+1FAFF is Extended_Pictographic and reserved for a future emoji, so no shipping font covers it on any
// platform. That makes "the declared family cannot draw this pictograph, and neither can anything else on
// the host" reproducible in CI without depending on which emoji faces the machine happens to have.
const UNCOVERED_PICTOGRAPH = '\u{1FAFF}';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

// The caption carries a literal newline inside one TextRun: the second trigger needs the pictograph and a
// layout control to reach the font layer in the same string, which is exactly what the canonical layout
// trace and textbox measurement ask for.
const rdl = `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="Rating"><DataField>Rating</DataField><TypeName>System.String</TypeName></Field>
  </Fields><Query><CommandText>never executed</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Textbox Name="Caption"><Top>0in</Top><Left>0in</Left><Width>3in</Width><Height>0.5in</Height><CanGrow>true</CanGrow>
      <Paragraphs><Paragraph><TextRuns><TextRun><Value>Legend ${UNCOVERED_PICTOGRAPH}
Second line ${UNCOVERED_PICTOGRAPH}</Value>
        <Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs>
      <Style/>
    </Textbox>
    <Tablix Name="Grid"><Top>0.75in</Top><Left>0in</Left>
      <TablixBody>
        <TablixColumns><TablixColumn><Width>3in</Width></TablixColumn></TablixColumns>
        <TablixRows><TablixRow><Height>0.3in</Height><TablixCells><TablixCell><CellContents>
          <Textbox Name="RatingCell"><CanGrow>true</CanGrow>
            <Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Rating.Value</Value>
              <Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs>
            <Style><Border><Style>Solid</Style><Color>Black</Color><Width>1pt</Width></Border></Style>
          </Textbox>
        </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
      </TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="Details"/></TablixMember></TablixMembers></TablixRowHierarchy>
      <DataSetName>D</DataSetName><Height>0.3in</Height><Width>3in</Width>
    </Tablix>
  </ReportItems><Height>1.5in</Height><Style/></Body><Width>3in</Width>
  <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`;

const rows = [{ Rating: `High ${UNCOVERED_PICTOGRAPH}` }, { Rating: 'Low' }];

function prepared(output) {
  return {
    model: parseRdl(Buffer.from(rdl, 'utf8')),
    request: { output, outputFileName: 'emoji-glyph-coverage', excelLayoutMode: 'REPORT', parameters: {}, datasets: { D: rows } },
  };
}

// ---- Font selection -------------------------------------------------------------------------------

test('an installed family that cannot draw a pictograph keeps the declared font instead of failing', () => {
  const arial = resolveFontFile(config.fontDir, 'Arial', false, false);
  if (!arial) return; // no Arial on this host: the whole-family path, covered elsewhere
  takeFontSubstitutions();
  // Strict mode on: strictness is about a whole family being absent, which shifts every advance width and
  // therefore every page break. One character the present family lacks is not that, and must not fail.
  const strict = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'true', RDL_ALLOW_COMPATIBLE_FONT_FALLBACKS: 'false' });
  assert.equal(pdfFont(strict, 'Arial', false, false, `Legend ${UNCOVERED_PICTOGRAPH}`), arial);
  // Degrading is reported rather than silent — the API surfaces this as X-Font-Substitutions.
  const [reported] = takeFontSubstitutions();
  assert.equal(reported.reason, 'no-covering-font');
  assert.equal(reported.requested, 'Arial');
});

test('layout controls are not glyphs: a multi-line run is not treated as uncovered', () => {
  const arial = resolveFontFile(config.fontDir, 'Arial', false, false);
  if (!arial) return;
  takeFontSubstitutions();
  assert.equal(pdfFont(config, 'Arial', false, false, 'Line 1\nLine 2\tEnd'), arial);
  // Nothing was substituted, because nothing was missing: no font on earth has a glyph for U+000A.
  assert.deepEqual(takeFontSubstitutions(), []);
});

test('a newline in the run does not change which font a pictograph resolves to', () => {
  if (!resolveFontFile(config.fontDir, 'Arial', false, false)) return;
  // Whatever this host has — a covering emoji face or nothing at all — the answer must be the same for the
  // single-line and multi-line forms of the same content. It was not: the newline defeated every candidate.
  const single = pdfFont(config, 'Arial', false, false, 'High \u{1F600}');
  const multiline = pdfFont(config, 'Arial', false, false, 'High \u{1F600}\nDetail');
  takeFontSubstitutions();
  assert.equal(multiline, single);
});

// ---- Renderer impact ------------------------------------------------------------------------------
// One shared font layer owns the defect, so each output is asserted against its own mechanics: PDF drawing
// commands, Word OOXML runs, and Excel cell values are not interchangeable evidence.

test('PDF renders the report and keeps the surrounding text in the canonical trace', async () => {
  const { model, request } = prepared('PDF');
  const pdf = await renderPdf(model, request, config, { captureLayoutTrace: true });
  takeFontSubstitutions();
  assert.ok(pdf.pageCount >= 1);
  const texts = pdf.layoutTrace.pages.flatMap((page) => page.items).map((item) => item.text || '');
  assert.ok(texts.some((text) => text.includes('Legend')));
  assert.ok(texts.some((text) => text.includes('High')));
});

test('DOCX_EDITABLE builds from that trace with native text rather than failing the export', async () => {
  const { model, request } = prepared('DOCX_EDITABLE');
  const docx = await renderEditableDocx(model, request, config);
  takeFontSubstitutions();
  const documentXml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  assert.ok(/<w:t[ >]/.test(documentXml)); // native Word text, not a page image
  assert.ok(documentXml.includes('Legend'));
  assert.ok(documentXml.includes('High'));
});

test('XLSX writes the same content as live cells', async () => {
  const { model, request } = prepared('XLSX');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load((await renderExcel(model, request, config, null)).buffer);
  takeFontSubstitutions();
  const values = [];
  workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => values.push(String(cell.value ?? ''))));
  assert.ok(values.some((value) => value.includes('Legend')));
  assert.ok(values.some((value) => value.includes('High')));
});
