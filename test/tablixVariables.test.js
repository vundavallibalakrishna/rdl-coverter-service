import assert from 'node:assert/strict';
import test from 'node:test';
import { analyzeRdl, parseRdl } from '../src/rdl/parser.js';
import { evaluateExpression } from '../src/rdl/expression.js';

// Report/group Variables, DomainScope and NaturalGroup were previously fail-closed. Variables now resolve
// via globals.variables in the current scope; DomainScope/NaturalGroup are acknowledged (metadata) so a
// report that declares them still renders.

function reportWith(groupInner, reportVars = '') {
  return `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">${reportVars}
 <DataSets><DataSet Name="D"><Fields>
   <Field Name="R"><DataField>R</DataField></Field>
   <Field Name="A"><DataField>A</DataField></Field>
 </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="T"><TablixBody>
    <TablixColumns><TablixColumn><Width>1in</Width></TablixColumn></TablixColumns>
    <TablixRows><TablixRow><Height>0.25in</Height><TablixCells>
      <TablixCell><CellContents><Textbox Name="d"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Variables!V.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell>
    </TablixCells></TablixRow></TablixRows></TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers><TablixMember>
      ${groupInner}
      <TablixMembers><TablixMember><Group Name="Det"/></TablixMember></TablixMembers>
    </TablixMember></TablixMembers></TablixRowHierarchy>
    <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>1in</Height><Width>1in</Width><Style/></Tablix>
 </ReportItems><Height>3in</Height><Style/></Body>
 <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth></Page></ReportSection></ReportSections></Report>`;
}
const GROUP = '<Group Name="G"><GroupExpressions><GroupExpression>=Fields!R.Value</GroupExpression></GroupExpressions>'
  + '<Variables><Variable Name="V"><Value>=Sum(Fields!A.Value)</Value></Variable></Variables>'
  + '<DomainScope>ColG</DomainScope><NaturalGroup/></Group>';

test('a group Variable resolves as Variables!Name.Value in the current scope', () => {
  const context = { fields: { Amt: 200 }, globals: { variables: { Tax: '=Fields!Amt.Value * 0.15' } } };
  assert.equal(evaluateExpression('=Variables!Tax.Value', context), 30);
});

test('a Variable may reference another Variable', () => {
  const context = { globals: { variables: { Tax: '=15', Label: '="Q" & Variables!Tax.Value' } } };
  assert.equal(evaluateExpression('=Variables!Label.Value', context), 'Q15');
});

test('an unknown Variable is null and a self-referencing Variable cannot recurse forever', () => {
  assert.equal(evaluateExpression('=Variables!Missing.Value', { globals: { variables: {} } }), null);
  assert.equal(evaluateExpression('=Variables!Loop.Value', { globals: { variables: { Loop: '=Variables!Loop.Value + 1' } } }), 1);
});

test('Variables, DomainScope and NaturalGroup no longer fail closed', () => {
  const analysis = analyzeRdl(reportWith(GROUP, '<Variables><Variable Name="RV"><Value>=1</Value></Variable></Variables>'));
  assert.equal(analysis.compatible, true);
  assert.deepEqual(analysis.blockingErrors, []);
});

test('report and group variable definitions are collected into model.variables', () => {
  const model = parseRdl(reportWith(GROUP, '<Variables><Variable Name="RV"><Value>=1</Value></Variable></Variables>'));
  assert.equal(model.variables.RV, '=1');
  assert.equal(model.variables.V, '=Sum(Fields!A.Value)');
});
