// A border width can be a conditional expression (=IIF(Fields!Row_Number_1.Value = 1, "1pt", "0pt")) — a
// common SSRS pattern for row-dependent separators. The parser used to call toPoints() on it eagerly and
// threw "Invalid RDL size", so the whole report failed to parse. It must instead keep the expression and let
// the renderers resolve it per row (0pt = the side is intentionally absent).
import assert from 'node:assert/strict';
import test from 'node:test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { parseRdl } from '../src/rdl/parser.js';
import { styleSize } from '../src/render/common.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { loadConfig } from '../src/config.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const CONDITIONAL = '=IIF(Fields!rn.Value = 1, "1pt", "0pt")';

test('an RDL with a conditional border width parses instead of throwing, keeping the expression', () => {
  const rdl = `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields><Field Name="rn"><DataField>rn</DataField></Field></Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Textbox Name="t"><Paragraphs><Paragraph><TextRuns><TextRun><Value>hi</Value></TextRun></TextRuns></Paragraph></Paragraphs>
    <Top>0in</Top><Left>0in</Left><Width>1in</Width><Height>0.25in</Height>
    <Style><TopBorder><Color>Black</Color><Style>Solid</Style><Width>${CONDITIONAL}</Width></TopBorder></Style></Textbox>
 </ReportItems><Height>1in</Height><Style/></Body>
 <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth></Page></ReportSection></ReportSections></Report>`;
  const model = parseRdl(rdl);
  const textbox = model.body.items.find((item) => item.type === 'Textbox');
  assert.equal(textbox.style.borders.top.width, CONDITIONAL); // kept raw, not converted/thrown
});

test('styleSize resolves a conditional width to points, honoring the row context', () => {
  const ctx = (rn) => ({ fields: { rn }, parameters: {}, globals: {}, dataset: [], datasets: {} });
  assert.equal(styleSize(CONDITIONAL, ctx(1), 1), 1); // "1pt"
  assert.equal(styleSize(CONDITIONAL, ctx(2), 1), 0); // "0pt" -> absent
  assert.equal(styleSize(3, ctx(1), 1), 3); // a literal number passes through unchanged
  assert.equal(styleSize('2pt', ctx(1), 1), 2);
});

test('a conditional width of 0 hides the border in editable DOCX; 1 draws it', async () => {
  const base = parseRdl(await fs.readFile(new URL('./fixtures/basic.rdl', import.meta.url)));
  const request = { parameters: { Title: 'T', Choice: 'A' }, datasets: { Sales: [{ Name: 'North', Amount: 1 }], Choices: [{ Value: 'A' }] }, outputFileName: 'cbw' };
  const none = { style: 'None', color: '#000000', width: 1 };
  // Strip every other border in the tablix so the only top border that can appear comes from our target
  // cell — otherwise the fixture's own bordered header/data cells pollute a whole-document assertion.
  const prepare = (widthExpr) => {
    const model = structuredClone(base);
    const tablix = model.body.items.find((item) => item.type === 'Tablix');
    tablix.style.borders = { top: { ...none }, right: { ...none }, bottom: { ...none }, left: { ...none } };
    for (const row of tablix.rows) {
      for (const cell of row.cells) {
        for (const item of cell.items) {
          if (item.style) item.style.borders = { top: { ...none }, right: { ...none }, bottom: { ...none }, left: { ...none } };
        }
      }
    }
    const target = tablix.rows[tablix.rows.length - 1].cells[0].items.find((item) => item.type === 'Textbox');
    target.style.borders.top = { style: 'Solid', color: '#000000', width: widthExpr };
    return model;
  };

  const drawnXml = await (await JSZip.loadAsync((await renderEditableDocx(prepare('=IIF(1=1,"1pt","0pt")'), request, config)).buffer)).file('word/document.xml').async('string');
  assert.equal((drawnXml.match(/<w:top w:val="single"/g) || []).length, 1); // exactly our one cell

  const hiddenXml = await (await JSZip.loadAsync((await renderEditableDocx(prepare('=IIF(1=1,"0pt","1pt")'), request, config)).buffer)).file('word/document.xml').async('string');
  assert.equal((hiddenXml.match(/<w:top w:val="single"/g) || []).length, 0); // 0pt -> no top border anywhere
});
