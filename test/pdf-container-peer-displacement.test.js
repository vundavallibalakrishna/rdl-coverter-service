import assert from 'node:assert/strict';
import test from 'node:test';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderEditableDocx } from '../src/render/docx.js';
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

const pageAdvanceRdl = `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <ReportSections><ReportSection><Body><ReportItems>
  <Rectangle Name="Container"><ReportItems>
   <Textbox Name="Prelude"><CanGrow>false</CanGrow><KeepTogether>true</KeepTogether>
    <Paragraphs><Paragraph><TextRuns><TextRun><Value>PRELUDE</Value></TextRun></TextRuns></Paragraph></Paragraphs>
    <Top>0in</Top><Left>0in</Left><Height>2.5in</Height><Width>3in</Width><Style/>
   </Textbox>
   <Textbox Name="Mandate"><CanGrow>true</CanGrow><KeepTogether>true</KeepTogether>
    <Paragraphs><Paragraph><TextRuns><TextRun><Value>MANDATE</Value></TextRun></TextRuns></Paragraph></Paragraphs>
    <Top>2.55in</Top><Left>0in</Left><Height>1in</Height><Width>3in</Width><Style/>
   </Textbox>
   <Textbox Name="Approach"><CanGrow>true</CanGrow><KeepTogether>true</KeepTogether>
    <Paragraphs><Paragraph><TextRuns><TextRun><Value>APPROACH</Value></TextRun></TextRuns></Paragraph></Paragraphs>
    <Top>3.6in</Top><Left>0in</Left><Height>1.1in</Height><Width>3in</Width><Style/>
   </Textbox>
  </ReportItems><Top>0in</Top><Left>0in</Left><Height>4.7in</Height><Width>3in</Width><Style/></Rectangle>
 </ReportItems><Height>4.7in</Height><Style/></Body>
 <Page><PageHeight>4in</PageHeight><PageWidth>4in</PageWidth>
  <LeftMargin>0.25in</LeftMargin><RightMargin>0.25in</RightMargin>
  <TopMargin>0.25in</TopMargin><BottomMargin>0.25in</BottomMargin>
 </Page></ReportSection></ReportSections>
</Report>`;

const sideBySideRdl = `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <DataSets>
  <DataSet Name="LeftData"><Fields><Field Name="V"><DataField>V</DataField></Field></Fields><Query><CommandText>x</CommandText></Query></DataSet>
  <DataSet Name="RightData"><Fields><Field Name="V"><DataField>V</DataField></Field></Fields><Query><CommandText>x</CommandText></Query></DataSet>
 </DataSets>
 <ReportSections><ReportSection><Body><ReportItems>
  <Rectangle Name="SideBySideContainer"><ReportItems>
   <Tablix Name="LeftTable"><TablixBody><TablixColumns><TablixColumn><Width>1.7in</Width></TablixColumn></TablixColumns>
    <TablixRows>
     <TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents><Textbox Name="LeftHeader"><Paragraphs><Paragraph><TextRuns><TextRun><Value>LEFT_HEADER</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border></Style></Textbox></CellContents></TablixCell></TablixCells></TablixRow>
     <TablixRow><Height>0.3in</Height><TablixCells><TablixCell><CellContents><Textbox Name="LeftDetail"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!V.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border></Style></Textbox></CellContents></TablixCell></TablixCells></TablixRow>
    </TablixRows></TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers>
     <TablixMember><KeepWithGroup>After</KeepWithGroup><RepeatOnNewPage>true</RepeatOnNewPage></TablixMember>
     <TablixMember><Group Name="LeftDetails"><GroupExpressions><GroupExpression>=Fields!V.Value</GroupExpression></GroupExpressions></Group></TablixMember>
    </TablixMembers></TablixRowHierarchy>
    <DataSetName>LeftData</DataSetName><Top>0in</Top><Left>0in</Left><Height>0.55in</Height><Width>1.7in</Width><Style/>
   </Tablix>
   <Tablix Name="RightTable"><TablixBody><TablixColumns><TablixColumn><Width>1.7in</Width></TablixColumn></TablixColumns>
    <TablixRows>
     <TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents><Textbox Name="RightHeader"><Paragraphs><Paragraph><TextRuns><TextRun><Value>RIGHT_HEADER</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border></Style></Textbox></CellContents></TablixCell></TablixCells></TablixRow>
     <TablixRow><Height>0.3in</Height><TablixCells><TablixCell><CellContents><Textbox Name="RightDetail"><Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!V.Value</Value></TextRun></TextRuns></Paragraph></Paragraphs><Style><Border><Style>Solid</Style></Border></Style></Textbox></CellContents></TablixCell></TablixCells></TablixRow>
    </TablixRows></TablixBody>
    <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
    <TablixRowHierarchy><TablixMembers>
     <TablixMember><KeepWithGroup>After</KeepWithGroup><RepeatOnNewPage>true</RepeatOnNewPage></TablixMember>
     <TablixMember><Group Name="RightDetails"><GroupExpressions><GroupExpression>=Fields!V.Value</GroupExpression></GroupExpressions></Group></TablixMember>
    </TablixMembers></TablixRowHierarchy>
    <DataSetName>RightData</DataSetName><Top>0in</Top><Left>1.8in</Left><Height>0.55in</Height><Width>1.7in</Width><Style/>
   </Tablix>
  </ReportItems><Top>0in</Top><Left>0in</Left><Height>0.6in</Height><Width>3.5in</Width><Style/></Rectangle>
 </ReportItems><Height>3in</Height><Style/></Body>
 <Page><PageHeight>3in</PageHeight><PageWidth>4in</PageWidth>
  <LeftMargin>0.2in</LeftMargin><RightMargin>0.2in</RightMargin>
  <TopMargin>0.2in</TopMargin><BottomMargin>0.2in</BottomMargin>
 </Page></ReportSection></ReportSections>
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

test('a rectangle continues from the final child position after that child advances a page', async () => {
  const rendered = await renderPdf(parseRdl(pageAdvanceRdl), {
    parameters: {},
    datasets: {},
  }, config, { captureLayoutTrace: true });

  const pageFor = (itemName) => rendered.layoutTrace.pages.findIndex((page) => (
    page.items.some((item) => item.itemName === itemName)
  ));
  assert.equal(pageFor('Prelude'), 0);
  assert.equal(pageFor('Mandate'), 1);
  assert.equal(
    pageFor('Approach'),
    1,
    'the next child uses the preceding child final-page end position rather than its stale prior-page Y',
  );
});

test('a rectangle still advances a later child when it genuinely cannot fit after a page advance', async () => {
  const nonFittingRdl = pageAdvanceRdl.replace(
    '<Top>3.6in</Top><Left>0in</Left><Height>1.1in</Height>',
    '<Top>3.6in</Top><Left>0in</Left><Height>3.4in</Height>',
  );
  const rendered = await renderPdf(parseRdl(nonFittingRdl), {
    parameters: {},
    datasets: {},
  }, config, { captureLayoutTrace: true });

  const pageFor = (itemName) => rendered.layoutTrace.pages.findIndex((page) => (
    page.items.some((item) => item.itemName === itemName)
  ));
  assert.equal(pageFor('Mandate'), 1);
  assert.equal(pageFor('Approach'), 2);
});

test('side-by-side growing tablixes share canonical page fragments and feed page-locked Word', async () => {
  const model = parseRdl(sideBySideRdl);
  const request = {
    parameters: {},
    datasets: {
      LeftData: Array.from({ length: 20 }, (_, index) => ({
        V: `LEFT_${String(index + 1).padStart(2, '0')}`,
      })),
      RightData: Array.from({ length: 8 }, (_, index) => ({
        V: `RIGHT_${String(index + 1).padStart(2, '0')}`,
      })),
    },
  };
  const canonical = await renderPdf(model, request, config, { captureLayoutTrace: true });

  assert.ok(canonical.pageCount >= 3);
  const pageForText = (text) => canonical.layoutTrace.pages.findIndex((page) => (
    page.items.some((item) => item.text === text)
  ));
  assert.ok(pageForText('LEFT_20') > pageForText('RIGHT_08'));
  assert.ok(
    canonical.layoutTrace.pages[1].items.some((item) => item.text === 'LEFT_HEADER')
      && canonical.layoutTrace.pages[1].items.some((item) => item.text === 'RIGHT_HEADER'),
    'each still-active tablix repeats its own header on the shared continuation page',
  );
  for (const prefix of ['LEFT', 'RIGHT']) {
    const count = prefix === 'LEFT' ? 20 : 8;
    for (let index = 1; index <= count; index += 1) {
      const marker = `${prefix}_${String(index).padStart(2, '0')}`;
      assert.equal(
        canonical.layoutTrace.pages.flatMap((page) => page.items)
          .filter((item) => item.text === marker).length,
        1,
        `${marker} must occur in exactly one synchronized PDF fragment`,
      );
    }
  }

  const editable = await renderEditableDocx(model, request, config);
  const documentXml = await (
    await JSZip.loadAsync(editable.buffer)
  ).file('word/document.xml').async('string');
  assert.equal(editable.pageCount, canonical.pageCount);
  assert.equal((documentXml.match(/<w:tbl>/g) || []).length, canonical.pageCount);
  assert.match(documentXml, /LEFT_20/);
  assert.match(documentXml, /RIGHT_08/);
});

test('page-spanning synchronized containers fail closed when their own paint cannot be layered safely', async () => {
  const painted = sideBySideRdl.replace(
    '<Height>0.6in</Height><Width>3.5in</Width><Style/></Rectangle>',
    '<Height>0.6in</Height><Width>3.5in</Width><Style><BackgroundColor>#eeeeee</BackgroundColor></Style></Rectangle>',
  );
  await assert.rejects(
    renderPdf(parseRdl(painted), {
      parameters: {},
      datasets: {
        LeftData: Array.from({ length: 20 }, (_, index) => ({ V: `LEFT_${index}` })),
        RightData: Array.from({ length: 8 }, (_, index) => ({ V: `RIGHT_${index}` })),
      },
    }, config),
    (error) => error.code === 'UNSUPPORTED_FEATURE'
      && /page-spanning rectangle with a visible fill or border/.test(error.message),
  );
});
