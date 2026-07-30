import assert from 'node:assert/strict';
import test from 'node:test';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

const rdl = `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets><DataSet Name="D"><Fields><Field Name="V"><DataField>V</DataField></Field></Fields><Query><CommandText>x</CommandText></Query></DataSet></DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Rectangle Name="Container"><ReportItems>
   <Textbox Name="Title"><CanGrow>false</CanGrow><KeepTogether>true</KeepTogether>
    <Paragraphs><Paragraph><TextRuns><TextRun><Value>CONTAINER_TITLE</Value></TextRun></TextRuns></Paragraph></Paragraphs>
    <Top>0.05in</Top><Left>0in</Left><Height>0.2in</Height><Width>3in</Width><ZIndex>2</ZIndex><Style/>
   </Textbox>
   <Tablix Name="GrowingTable"><TablixBody><TablixColumns><TablixColumn><Width>3in</Width></TablixColumn></TablixColumns>
    <TablixRows>
     <TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents><Textbox Name="Header"><Paragraphs><Paragraph><TextRuns><TextRun><Value>HEADER</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border></Style></Textbox></CellContents></TablixCell></TablixCells></TablixRow>
     <TablixRow><Height>0.35in</Height><TablixCells><TablixCell><CellContents><Textbox Name="Detail"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!V.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border></Style></Textbox></CellContents></TablixCell></TablixCells></TablixRow>
    </TablixRows></TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers><TablixMember/><TablixMember><Group Name="Details"><GroupExpressions><GroupExpression>=Fields!V.Value</GroupExpression></GroupExpressions></Group></TablixMember></TablixMembers></TablixRowHierarchy>
    <DataSetName>D</DataSetName><Top>0.3in</Top><Left>0in</Left><Height>0.6in</Height><Width>3in</Width><ZIndex>0</ZIndex><Style/>
   </Tablix>
   <Textbox Name="LegendA"><CanGrow>false</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun><Value>LEGEND_A</Value></TextRun></TextRuns></Paragraph></Paragraphs><Top>0.95in</Top><Left>0in</Left><Height>0.2in</Height><Width>1in</Width><ZIndex>1</ZIndex><Style/></Textbox>
   <Textbox Name="LegendB"><CanGrow>false</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun><Value>LEGEND_B</Value></TextRun></TextRuns></Paragraph></Paragraphs><Top>0.95in</Top><Left>1in</Left><Height>0.2in</Height><Width>1in</Width><ZIndex>3</ZIndex><Style/></Textbox>
   <Textbox Name="LegendC"><CanGrow>false</CanGrow><Paragraphs><Paragraph><TextRuns><TextRun><Value>LEGEND_C</Value></TextRun></TextRuns></Paragraph></Paragraphs><Top>0.9502in</Top><Left>2in</Left><Height>0.2in</Height><Width>1in</Width><ZIndex>4</ZIndex><Style/></Textbox>
  </ReportItems><KeepTogether>true</KeepTogether><Top>0in</Top><Left>0in</Left><Height>1.2in</Height><Width>3in</Width><Style/></Rectangle>
 </ReportItems><Height>3.5in</Height><Style/></Body>
 <Page><PageHeight>4in</PageHeight><PageWidth>4in</PageWidth><LeftMargin>0.25in</LeftMargin><RightMargin>0.25in</RightMargin><TopMargin>0.25in</TopMargin><BottomMargin>0.25in</BottomMargin></Page>
 </ReportSection></ReportSections>
</Report>`;

test('a growing tablix displaces later rectangle peers while same-band peers remain aligned', async () => {
  const model = parseRdl(rdl);
  const rows = Array.from({ length: 15 }, (_, index) => ({ V: `ROW_${String(index + 1).padStart(2, '0')}` }));
  const rendered = await renderPdf(model, {
    parameters: {},
    datasets: { D: rows },
  }, config, { captureLayoutTrace: true });

  assert.ok(rendered.pageCount > 1);
  const pages = rendered.layoutTrace.pages;
  const titlePages = pages
    .map((page, index) => page.items.some((item) => item.itemName === 'Title') ? index : -1)
    .filter((index) => index >= 0);
  assert.deepEqual(titlePages, [0], 'the title remains before the table instead of being repainted after pagination');

  const finalPageIndex = pages.findIndex((page) => page.items.some((item) => item.text === 'ROW_15'));
  assert.ok(finalPageIndex > 0);
  const finalItems = pages[finalPageIndex].items;
  const finalDetail = finalItems.find((item) => item.kind === 'tablixCell' && item.text === 'ROW_15');
  const legends = ['LegendA', 'LegendB', 'LegendC'].map((name) => (
    finalItems.find((item) => item.itemName === name)
  ));
  assert.ok(finalDetail);
  assert.ok(legends.every(Boolean), 'all same-band legend peers follow the completed table');
  assert.ok(Math.max(...legends.map((item) => item.y)) - Math.min(...legends.map((item) => item.y)) <= 0.25);
  const tableToLegendGap = Math.min(...legends.map((item) => item.y))
    - (finalDetail.y + finalDetail.height);
  assert.ok(
    Math.abs(tableToLegendGap - 3.6) <= 0.25,
    'the RDL gap between the table and legend is preserved after the table grows across pages',
  );
});
