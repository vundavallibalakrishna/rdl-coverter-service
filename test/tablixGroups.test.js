import assert from 'node:assert/strict';
import test from 'node:test';
import { analyzeRdl, parseRdl } from '../src/rdl/parser.js';
import { materializeTablixRows } from '../src/rdl/validation.js';

// Minimal single-column grouped tablix: one row group (CatGroup) over a detail row. `bodyCell` is the
// detail cell expression; extra options inject member/group attributes for the specific test.
function groupedRdl({ bodyCell, groupPageBreak = '', memberHidden = '', groupFilter = '' }) {
  return `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields>
   <Field Name="Cat"><DataField>Cat</DataField></Field>
   <Field Name="Amount"><DataField>Amount</DataField></Field>
   <Field Name="Keep"><DataField>Keep</DataField></Field>
 </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="T"><TablixBody>
    <TablixColumns><TablixColumn><Width>2in</Width></TablixColumn></TablixColumns>
    <TablixRows><TablixRow><Height>0.25in</Height>${memberHidden ? '' : ''}<TablixCells>
      <TablixCell><CellContents><Textbox Name="d"><Paragraphs><Paragraph><TextRuns><TextRun><Value>${bodyCell}</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell>
    </TablixCells></TablixRow></TablixRows></TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers><TablixMember>
      <Group Name="CatGroup"><GroupExpressions><GroupExpression>=Fields!Cat.Value</GroupExpression></GroupExpressions>${groupPageBreak ? `<PageBreak><BreakLocation>${groupPageBreak}</BreakLocation></PageBreak>` : ''}${groupFilter ? `<Filters><Filter><FilterExpression>=Fields!Keep.Value</FilterExpression><Operator>Equal</Operator><FilterValues><FilterValue>yes</FilterValue></FilterValues></Filter></Filters>` : ''}</Group>
      ${memberHidden ? `<Visibility><Hidden>${memberHidden}</Hidden></Visibility>` : ''}
      <TablixMembers><TablixMember><Group Name="Details"/></TablixMember></TablixMembers>
    </TablixMember></TablixMembers></TablixRowHierarchy>
    <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.25in</Height><Width>2in</Width><Style/></Tablix>
 </ReportItems><Height>3in</Height><Style/></Body>
 <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth></Page></ReportSection></ReportSections></Report>`;
}

function nestedDuplicateRdl() {
  return `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields>
   <Field Name="Parent"><DataField>Parent</DataField></Field>
   <Field Name="Child"><DataField>Child</DataField></Field>
 </Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="T"><TablixBody>
   <TablixColumns><TablixColumn><Width>2in</Width></TablixColumn></TablixColumns>
   <TablixRows><TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents>
    <Textbox Name="status"><HideDuplicates>ChildGroup</HideDuplicates><Paragraphs><Paragraph><TextRuns><TextRun><Value>Not Assessed</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox>
   </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
  </TablixBody>
  <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
  <TablixRowHierarchy><TablixMembers><TablixMember>
   <Group Name="ParentGroup"><GroupExpressions><GroupExpression>=Fields!Parent.Value</GroupExpression></GroupExpressions></Group>
   <TablixMembers><TablixMember>
    <Group Name="ChildGroup"><GroupExpressions><GroupExpression>=Fields!Child.Value</GroupExpression></GroupExpressions></Group>
    <TablixMembers><TablixMember><Group Name="Details"/></TablixMember></TablixMembers>
   </TablixMember></TablixMembers>
  </TablixMember></TablixMembers></TablixRowHierarchy>
  <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.25in</Height><Width>2in</Width><Style/></Tablix>
 </ReportItems><Height>3in</Height><Style/></Body>
 <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth></Page></ReportSection></ReportSections></Report>`;
}
const tablixOf = (rdl) => parseRdl(rdl).body.items.find((item) => item.type === 'Tablix');
const values = (rows) => rows.map((row) => row.cells.map((cell) => (cell.values || []).join('')).join('|'));

test('a no-scope aggregate scopes to the innermost group instance, not the whole dataset', () => {
  const t = tablixOf(groupedRdl({ bodyCell: '=Sum(Fields!Amount.Value)' }));
  const rows = materializeTablixRows(t, [{ Cat: 'A', Amount: 10 }, { Cat: 'A', Amount: 20 }, { Cat: 'B', Amount: 5 }], {}, {}, {});
  assert.deepEqual(values(rows), ['30', '30', '5']); // per-group (A=30, B=5), not dataset total 35
});

test('a named scope resolves to that group instance', () => {
  const t = tablixOf(groupedRdl({ bodyCell: '=Sum(Fields!Amount.Value, "CatGroup")' }));
  const rows = materializeTablixRows(t, [{ Cat: 'A', Amount: 10 }, { Cat: 'A', Amount: 20 }, { Cat: 'B', Amount: 5 }], {}, {}, {});
  assert.deepEqual(values(rows), ['30', '30', '5']);
});

test('a group filter removes rows outside the group', () => {
  const t = tablixOf(groupedRdl({ bodyCell: '=Fields!Cat.Value', groupFilter: 'yes' }));
  const rows = materializeTablixRows(t, [{ Cat: 'A', Keep: 'yes' }, { Cat: 'B', Keep: 'no' }, { Cat: 'C', Keep: 'yes' }], {}, {}, {});
  assert.deepEqual(values(rows), ['A', 'C']); // B filtered out
});

test('a static member Visibility.Hidden removes the group rows', () => {
  const t = tablixOf(groupedRdl({ bodyCell: '=Fields!Cat.Value', memberHidden: 'true' }));
  const rows = materializeTablixRows(t, [{ Cat: 'A' }, { Cat: 'B' }], {}, {}, {});
  assert.equal(rows.length, 0);
});

test('a group page break tags the first row of each new instance', () => {
  const t = tablixOf(groupedRdl({ bodyCell: '=Fields!Cat.Value', groupPageBreak: 'Start' }));
  const rows = materializeTablixRows(t, [{ Cat: 'A' }, { Cat: 'A' }, { Cat: 'B' }], {}, {}, {});
  assert.deepEqual(rows.map((row) => row.pageBreakBefore), [false, false, true]); // break before B
});

test('HideDuplicates resets for the same child value in a different parent group instance', () => {
  const t = tablixOf(nestedDuplicateRdl());
  const rows = materializeTablixRows(t, [
    { Parent: 'A', Child: null },
    { Parent: 'A', Child: null },
    { Parent: 'B', Child: null },
    { Parent: 'B', Child: null },
  ], {}, {}, {});
  assert.deepEqual(values(rows), ['Not Assessed', '', 'Not Assessed', '']);
});

test('HideIfNoRows, NoRowsMessage and ToggleItem no longer fail closed', () => {
  const inject = (snippet) => parseRdl(groupedRdl({ bodyCell: '=Fields!Cat.Value' }).replace('</TablixRowHierarchy>', `</TablixRowHierarchy>${snippet}`));
  // These constructs, injected into a valid tablix report, must keep the report compatible.
  const withHideIfNoRows = groupedRdl({ bodyCell: '=Fields!Cat.Value' }).replace('<TablixMember>\n      <Group Name="CatGroup"', '<TablixMember>\n      <HideIfNoRows>true</HideIfNoRows>\n      <Group Name="CatGroup"');
  assert.equal(analyzeRdl(withHideIfNoRows).compatible, true);
  const withNoRowsMessage = groupedRdl({ bodyCell: '=Fields!Cat.Value' }).replace('<DataSetName>D</DataSetName>', '<NoRowsMessage>Nothing here</NoRowsMessage><DataSetName>D</DataSetName>');
  assert.equal(analyzeRdl(withNoRowsMessage).compatible, true);
  const withToggle = groupedRdl({ bodyCell: '=Fields!Cat.Value' }).replace('<Textbox Name="d">', '<Textbox Name="d"><Visibility><ToggleItem>d</ToggleItem></Visibility>');
  assert.equal(analyzeRdl(withToggle).compatible, true);
});
