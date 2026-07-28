import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import { parseRdl } from '../src/rdl/parser.js';
import { materializeTablixColumns, needsAdvancedMaterialization } from '../src/rdl/validation.js';
import { MISSING_SAMPLES, hasSamples, samplePath } from '../scripts/lib/samples.js';

const COMBINED_ASSURANCE = 'Combined Assurance Reports Excel.rdl';

// Guard: tablixes WITHOUT dynamic column groups, subtotal rows, or recursive parents must take the
// unchanged flat materialization path and keep item.columns as their grid. This protects byte-identical
// output for every existing detail/static/merged-header report (the parity oracle).
test('static and detail tablixes stay on the flat path with unchanged grid columns', () => {
  const basic = parseRdl(fs.readFileSync(new URL('./fixtures/basic.rdl', import.meta.url)));
  for (const tablix of basic.body.items.filter((item) => item.type === 'Tablix')) {
    assert.equal(needsAdvancedMaterialization(tablix), false);
    assert.equal(tablix.hasColumnGroups, false);
    assert.equal(materializeTablixColumns(tablix, [{ Name: 'A', Amount: 1 }], {}, {}, {}), tablix.columns);
  }
});

test('the Combined Assurance matrix and detail tablixes are NOT rerouted to the advanced path', {
  skip: hasSamples(COMBINED_ASSURANCE) ? false : MISSING_SAMPLES,
}, () => {
  const model = parseRdl(fs.readFileSync(samplePath(COMBINED_ASSURANCE)));
  for (const tablix of model.body.items.filter((item) => item.type === 'Tablix')) {
    assert.equal(needsAdvancedMaterialization(tablix), false);
    assert.equal(tablix.hasColumnGroups, false);
    assert.equal(materializeTablixColumns(tablix, [], {}, {}, {}), tablix.columns);
  }
});

// A plain group>Details tablix (a group whose leaf is the detail row) is NOT a subtotal-row tablix and
// must stay on the flat path.
test('a plain grouped detail tablix is not treated as an advanced (subtotal) tablix', () => {
  const rdl = `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields><Field Name="Cat"><DataField>Cat</DataField></Field></Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="T"><TablixBody><TablixColumns><TablixColumn><Width>2in</Width></TablixColumn></TablixColumns>
    <TablixRows><TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents><Textbox Name="d"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Cat.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell></TablixCells></TablixRow></TablixRows></TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="CatGroup"><GroupExpressions><GroupExpression>=Fields!Cat.Value</GroupExpression></GroupExpressions></Group>
      <TablixMembers><TablixMember><Group Name="Details"/></TablixMember></TablixMembers></TablixMember></TablixMembers></TablixRowHierarchy>
    <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.25in</Height><Width>2in</Width><Style/></Tablix>
 </ReportItems><Height>3in</Height><Style/></Body><Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth></Page></ReportSection></ReportSections></Report>`;
  const tablix = parseRdl(rdl).body.items.find((item) => item.type === 'Tablix');
  assert.equal(needsAdvancedMaterialization(tablix), false);
});

test('a grouped multi-column row-header chain uses per-instance advanced materialization', () => {
  const leaf = { children: [] };
  const fourthHeader = { header: { size: 40 }, children: [leaf] };
  const thirdHeader = { header: { size: 40 }, children: [fourthHeader] };
  const secondHeader = { header: { size: 40 }, children: [thirdHeader] };
  const firstHeader = { header: { size: 40 }, children: [secondHeader] };
  const groupedHeaderChain = {
    group: { name: 'DetailGroup', expressions: ['=Fields!Id.Value'] },
    children: [firstHeader],
  };
  const tablix = {
    hasColumnGroups: false,
    rowMembers: [groupedHeaderChain],
    rowMemberPaths: [[groupedHeaderChain, firstHeader, secondHeader, thirdHeader, fourthHeader, leaf]],
  };

  assert.equal(needsAdvancedMaterialization(tablix), true);
});
