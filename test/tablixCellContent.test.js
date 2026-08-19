import assert from 'node:assert/strict';
import test from 'node:test';
import { analyzeRdl, parseRdl } from '../src/rdl/parser.js';
import { materializeTablixRows } from '../src/rdl/validation.js';

// A tablix cell is its own rendering context: the capability catalogue classifies element NAMES, so an
// element can be SUPPORTED at body level yet unhandled inside a cell. That gap let a Rectangle-wrapped cell
// render blank while analyze still reported compatible:true — silently wrong output, which is exactly what
// the fail-closed design exists to prevent. Cell content must therefore be rendered or refused, never dropped.

const TEXTBOX = '<Textbox Name="t"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!N.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox>';

function reportWithCell(inner) {
  return `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <EmbeddedImages><EmbeddedImage Name="L"><MIMEType>image/png</MIMEType><ImageData>AA==</ImageData></EmbeddedImage></EmbeddedImages>
 <DataSets><DataSet Name="D"><Fields><Field Name="N"><DataField>N</DataField></Field></Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="T"><TablixBody>
    <TablixColumns><TablixColumn><Width>1in</Width></TablixColumn></TablixColumns>
    <TablixRows><TablixRow><Height>0.25in</Height><TablixCells>
      <TablixCell><CellContents>${inner}</CellContents></TablixCell>
    </TablixCells></TablixRow></TablixRows></TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="G"/></TablixMember></TablixMembers></TablixRowHierarchy>
    <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>1in</Height><Width>1in</Width><Style/></Tablix>
 </ReportItems><Height>3in</Height><Style/></Body>
 <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth></Page></ReportSection></ReportSections></Report>`;
}
const tablixOf = (rdl) => parseRdl(rdl).body.items.find((item) => item.type === 'Tablix');

test('a cell wrapping its textbox in a container Rectangle renders the textbox, not an empty cell', () => {
  const rdl = reportWithCell(`<Rectangle><ReportItems>${TEXTBOX}</ReportItems><Style/></Rectangle>`);
  assert.equal(analyzeRdl(rdl).compatible, true);
  const rows = materializeTablixRows(tablixOf(rdl), [{ N: 'value' }], {}, {}, {});
  assert.deepEqual(rows.map((row) => row.cells.map((cell) => (cell.values || []).join(''))), [['value']]);
});

test('a List / canvas cell (line, image drawn item-by-item) is supported, not refused', () => {
  // A Rectangle-flattened free-form cell draws its positioned items directly. Line and Image are rendered
  // in place, so a cell containing them is compatible and materializes without throwing.
  for (const [type, inner] of [
    ['Image', '<Image Name="i"><Source>Embedded</Source><Value>L</Value><Sizing>Fit</Sizing><Style/></Image>'],
    ['Line', '<Line Name="ln"><Style/></Line>'],
  ]) {
    const analysis = analyzeRdl(reportWithCell(inner));
    assert.equal(analysis.compatible, true, `${type} in a cell must now be supported`);
    assert.equal(analysis.blockingErrors.some(({ feature }) => feature === `TablixCellContent:${type}`), false);
    assert.doesNotThrow(() => materializeTablixRows(tablixOf(reportWithCell(inner)), [{ N: 'v' }], {}, {}, {}));
  }
});

test('genuinely unsupported cell content still fails closed instead of rendering blank', () => {
  const inner = '<GaugePanel Name="g"><Style/></GaugePanel>';
  const analysis = analyzeRdl(reportWithCell(inner));
  assert.equal(analysis.compatible, false, 'a GaugePanel in a cell must not be reported compatible');
  // Defence in depth: reaching materialization anyway must throw rather than yield an empty cell.
  assert.throws(
    () => materializeTablixRows(tablixOf(reportWithCell(inner)), [{ N: 'v' }], {}, {}, {}),
    /Tablix cell content is not supported/,
  );
});

test('a plain textbox cell stays compatible', () => {
  assert.equal(analyzeRdl(reportWithCell(TEXTBOX)).compatible, true);
});
