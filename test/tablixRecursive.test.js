import assert from 'node:assert/strict';
import test from 'node:test';
import JSZip from 'jszip';
import { analyzeRdl, parseRdl } from '../src/rdl/parser.js';
import { materializeTablixRows, needsAdvancedMaterialization } from '../src/rdl/validation.js';
import { renderEditableDocx } from '../src/render/docx.js';

// A recursive (parent/child) row group: Group/Parent references the parent id, so the flat list is
// rendered expanded as a hierarchy, indented by recursion depth.
function recursiveRdl() {
  return `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields><Field Name="Id"><DataField>Id</DataField></Field><Field Name="Pid"><DataField>Pid</DataField></Field><Field Name="Name"><DataField>Name</DataField></Field></Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="R"><TablixBody>
    <TablixColumns><TablixColumn><Width>3in</Width></TablixColumn></TablixColumns>
    <TablixRows><TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents><Textbox Name="n"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Name.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell></TablixCells></TablixRow></TablixRows></TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers><TablixMember>
      <Group Name="Tree"><GroupExpressions><GroupExpression>=Fields!Id.Value</GroupExpression></GroupExpressions><Parent>=Fields!Pid.Value</Parent></Group>
    </TablixMember></TablixMembers></TablixRowHierarchy>
    <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.25in</Height><Width>3in</Width><Style/></Tablix>
 </ReportItems><Height>3in</Height><Style/></Body><Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth></Page></ReportSection></ReportSections></Report>`;
}

const tablixOf = (rdl) => parseRdl(rdl).body.items.find((item) => item.type === 'Tablix');
const rows = [
  { Id: '1', Pid: '', Name: 'CEO' },
  { Id: '2', Pid: '1', Name: 'VP-A' },
  { Id: '3', Pid: '2', Name: 'Eng' },
  { Id: '4', Pid: '1', Name: 'VP-B' },
];

test('a recursive parent group is compatible (parsed, not fail-closed)', () => {
  assert.equal(analyzeRdl(recursiveRdl()).compatible, true);
  const tablix = tablixOf(recursiveRdl());
  assert.equal(tablix.rowMembers[0].group.parent, '=Fields!Pid.Value');
  assert.equal(needsAdvancedMaterialization(tablix), true);
});

test('a recursive parent group renders expanded, hierarchically ordered, and indented by depth', () => {
  const tablix = tablixOf(recursiveRdl());
  const materialized = materializeTablixRows(tablix, rows, {}, {}, {});
  assert.deepEqual(
    materialized.map((row) => ({ name: row.cells[0].values.join(''), indent: row.indentLevel })),
    [
      { name: 'CEO', indent: 0 },
      { name: 'VP-A', indent: 1 },
      { name: 'Eng', indent: 2 },
      { name: 'VP-B', indent: 1 },
    ],
  );
});

test('recursive indentation is applied in editable DOCX', async () => {
  const model = parseRdl(recursiveRdl());
  const result = await renderEditableDocx(model, { outputFileName: 'tree', parameters: {}, datasets: { D: rows } });
  const zip = await JSZip.loadAsync(result.buffer);
  const documentXml = await zip.file('word/document.xml').async('string');
  // PDF recursive indentation is captured as increased left cell padding in the page-locked grid.
  const leftMargins = [...documentXml.matchAll(/<w:left w:type="dxa" w:w="(\d+)"\/>/g)]
    .map((match) => Number(match[1]));
  assert.equal(Math.max(...leftMargins) > Math.min(...leftMargins), true);
  for (const name of ['CEO', 'VP-A', 'Eng', 'VP-B']) assert.ok(documentXml.includes(name));
});
