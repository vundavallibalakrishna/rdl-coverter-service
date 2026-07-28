// Completeness + behaviour for the capability model's expression-property dimension (Phase 5). This is the
// structural fix for "we found expression-property bugs one at a time": the catalogue now enumerates which
// properties are ExpressionType, and /v1/analyze reports which an RDL actually uses and whether we handle
// them. These tests enforce that the catalogue stays complete and that every catalogued property is handled.
import assert from 'node:assert/strict';
import test from 'node:test';
import { EXPRESSION_PROPERTIES } from '../src/rdl/capabilities.js';
import { analyzeRdl } from '../src/rdl/parser.js';

// A NEW unhandled entry (a property catalogued but silently unhandled) must be added here consciously —
// otherwise this test catches it.
const KNOWN_UNHANDLED = [];

test('the only unhandled expression properties are the documented ones', () => {
  const unhandled = Object.entries(EXPRESSION_PROPERTIES).filter(([, spec]) => !spec.handled).map(([k]) => k).sort();
  assert.deepEqual(unhandled, [...KNOWN_UNHANDLED].sort(), 'a new expression property is catalogued but unhandled — handle it or document it');
});

test('the catalogue covers the full Style + Border expression surface', () => {
  // Canonical expression-capable properties for our supported surface. Adding a Style property later without
  // cataloguing it here (and handling it) trips this test.
  const required = [
    'Style.Color', 'Style.BackgroundColor', 'Style.FontFamily', 'Style.FontSize', 'Style.FontWeight',
    'Style.FontStyle', 'Style.TextDecoration', 'Style.TextAlign', 'Style.VerticalAlign', 'Style.Format',
    'Style.PaddingLeft', 'Style.PaddingRight', 'Style.PaddingTop', 'Style.PaddingBottom', 'Style.LineHeight',
    'Paragraph.SpaceBefore', 'Paragraph.SpaceAfter',
    'Visibility.Hidden', 'Image.Value', 'Image.Sizing',
  ];
  for (const owner of ['Border', 'TopBorder', 'RightBorder', 'BottomBorder', 'LeftBorder']) {
    required.push(`${owner}.Style`, `${owner}.Color`, `${owner}.Width`);
  }
  for (const key of required) assert.ok(EXPRESSION_PROPERTIES[key], `missing expression-property entry: ${key}`);
});

test('/analyze reports the expression-driven properties an RDL uses, with value types', () => {
  const rdl = `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields><Field Name="rn"><DataField>rn</DataField></Field></Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Tablix Name="T"><TablixBody><TablixColumns><TablixColumn><Width>2in</Width></TablixColumn></TablixColumns>
   <TablixRows><TablixRow><Height>0.3in</Height><TablixCells><TablixCell><CellContents>
     <Textbox Name="t"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!rn.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs>
       <Style><FontSize>=IIF(Fields!rn.Value=0,"12pt","9pt")</FontSize>
         <TopBorder><Style>=IIF(Fields!rn.Value=0,"Solid","None")</Style><Width>=IIF(Fields!rn.Value=0,"1pt","0pt")</Width></TopBorder></Style>
       <Visibility><Hidden>=IIF(Fields!rn.Value=0,False,True)</Hidden></Visibility></Textbox>
   </CellContents></TablixCell></TablixCells></TablixRow></TablixRows></TablixBody>
   <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
   <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="G"/></TablixMember></TablixMembers></TablixRowHierarchy>
   <DataSetName>D</DataSetName><Top>0in</Top><Left>0in</Left><Height>1in</Height><Width>2in</Width><Style/></Tablix>
 </ReportItems><Height>3in</Height><Style/></Body><Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth></Page></ReportSection></ReportSections></Report>`;
  const props = analyzeRdl(rdl).capabilities.expressions.properties;
  const seen = Object.fromEntries(props.map((p) => [p.property, p.valueType]));
  assert.equal(seen['Style.FontSize'], 'RdlSize');
  assert.equal(seen['TopBorder.Style'], 'Enum');
  assert.equal(seen['TopBorder.Width'], 'RdlSize');
  assert.equal(seen['Visibility.Hidden'], 'Boolean');
  assert.equal(props.every((p) => p.handled), true); // all handled -> a clean report
});
