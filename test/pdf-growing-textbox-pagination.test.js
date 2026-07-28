import assert from 'node:assert/strict';
import test from 'node:test';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const lines = Array.from({ length: 70 }, (_, index) => `LINE_${String(index + 1).padStart(2, '0')}`).join('\n');
const rdl = `<?xml version="1.0"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
 <ReportSections><ReportSection><Body><ReportItems>
  <Textbox Name="Growing"><CanGrow>true</CanGrow><KeepTogether>true</KeepTogether><Paragraphs><Paragraph><TextRuns><TextRun>
   <Value>${lines}</Value><Style><FontFamily>Arial</FontFamily><FontSize>10pt</FontSize></Style>
  </TextRun></TextRuns></Paragraph></Paragraphs><Top>1.5in</Top><Left>0in</Left><Height>0.2in</Height><Width>3in</Width>
  <Style><PaddingTop>2pt</PaddingTop><PaddingBottom>2pt</PaddingBottom></Style></Textbox>
 </ReportItems><Height>3in</Height><Style/></Body><Width>3in</Width>
 <Page><PageWidth>4in</PageWidth><PageHeight>4in</PageHeight><TopMargin>0.25in</TopMargin><BottomMargin>0.25in</BottomMargin><LeftMargin>0.25in</LeftMargin><RightMargin>0.25in</RightMargin>
  <PageFooter><Height>0.4in</Height><PrintOnFirstPage>true</PrintOnFirstPage><PrintOnLastPage>true</PrintOnLastPage><ReportItems>
   <Textbox Name="Footer"><Paragraphs><Paragraph><TextRuns><TextRun><Value>FOOTER</Value></TextRun></TextRuns></Paragraph></Paragraphs><Top>0in</Top><Left>0in</Left><Height>0.2in</Height><Width>3in</Width><Style/></Textbox>
  </ReportItems><Style/></PageFooter>
 </Page></ReportSection></ReportSections>
</Report>`;

test('a growing free-form textbox paginates before the footer band', async () => {
  const result = await renderPdf(parseRdl(rdl), {
    outputFileName: 'growing-textbox',
    parameters: {},
    datasets: {},
  }, config);
  assert.ok(result.pageCount >= 3, `expected a multipage textbox, got ${result.pageCount} page(s)`);
  assert.ok(result.buffer.length > 1000);
});
