import assert from 'node:assert/strict';
import test from 'node:test';
import { parseRdl } from '../src/rdl/parser.js';
import { materializeTablixColumns, materializeTablixRows, needsAdvancedMaterialization } from '../src/rdl/validation.js';

// A single row group (CatGroup) with a static group-header leaf, a Details leaf, and a static
// group-footer leaf. The header/footer are STATIC members (no <Group>) nested inside the group, so
// SSRS emits them once per group instance — a header-once/footer-once subtotal row.
function subtotalRdl() {
  const bodyRow = (name, expression) => `<TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents><Textbox Name="${name}"><Paragraphs><Paragraph><TextRuns><TextRun><Value>${expression}</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style/></Textbox></CellContents></TablixCell></TablixCells></TablixRow>`;
  return `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields><Field Name="Cat"><DataField>Cat</DataField></Field><Field Name="Amount"><DataField>Amount</DataField></Field></Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="T"><TablixBody>
    <TablixColumns><TablixColumn><Width>2in</Width></TablixColumn></TablixColumns>
    <TablixRows>
      ${bodyRow('h', '=Fields!Cat.Value')}
      ${bodyRow('d', '=Fields!Amount.Value')}
      ${bodyRow('f', '=Sum(Fields!Amount.Value)')}
    </TablixRows></TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers>
      <TablixMember><Group Name="CatGroup"><GroupExpressions><GroupExpression>=Fields!Cat.Value</GroupExpression></GroupExpressions></Group>
      <TablixMembers>
        <TablixMember/>
        <TablixMember><Group Name="Details"/></TablixMember>
        <TablixMember/>
      </TablixMembers></TablixMember>
    </TablixMembers></TablixRowHierarchy>
    <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.25in</Height><Width>2in</Width><Style/></Tablix>
 </ReportItems><Height>3in</Height><Style/></Body><Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth></Page></ReportSection></ReportSections></Report>`;
}

const tablixOf = (rdl) => parseRdl(rdl).body.items.find((item) => item.type === 'Tablix');
const rowsOf = (rows) => rows.map((row) => ({ role: row.role, value: row.cells.map((cell) => (cell.values || []).join('')).join('|') }));

test('a grouped tablix emits header-once and footer-once subtotal rows scoped to the group instance', () => {
  const tablix = tablixOf(subtotalRdl());
  assert.equal(needsAdvancedMaterialization(tablix), true);
  const rows = materializeTablixRows(tablix, [
    { Cat: 'A', Amount: 10 }, { Cat: 'A', Amount: 20 }, { Cat: 'B', Amount: 5 },
  ], {}, {}, {});
  assert.deepEqual(rowsOf(rows), [
    { role: 'header', value: 'A' },
    { role: 'detail', value: '10' },
    { role: 'detail', value: '20' },
    { role: 'footer', value: '30' }, // group A sum, NOT the dataset total 35
    { role: 'header', value: 'B' },
    { role: 'detail', value: '5' },
    { role: 'footer', value: '5' }, // group B sum
  ]);
  // The footer subtotal is the per-group Sum, never the whole-dataset total.
  assert.equal(rows.filter((row) => row.role === 'footer').map((row) => row.cells[0].values.join('')).join(','), '30,5');
});

test('the subtotal tablix keeps its static grid columns (no matrix expansion)', () => {
  const tablix = tablixOf(subtotalRdl());
  assert.deepEqual(materializeTablixColumns(tablix, [{ Cat: 'A', Amount: 10 }], {}, {}, {}), tablix.columns);
});
