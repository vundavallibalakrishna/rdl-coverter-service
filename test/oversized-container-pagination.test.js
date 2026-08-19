// A body band that does not fit the remainder of the current page used to be deferred wholesale to a fresh
// page. That is right only when the band actually fits a fresh page. A container taller than the entire
// printable body cannot be rescued by moving it — it will overflow wherever it starts — so deferring it
// bought nothing and threw away the rest of the page, leaving a blank band the height of the remainder.
// A report whose first band is a short banner followed by a full-height container therefore rendered a
// nearly empty first page.
//
// The rule is the same one the growable-textbox flow already applies to its own oversized blocks: fill the
// remainder and fragment across pages.
import assert from 'node:assert/strict';
import test from 'node:test';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderPdf } from '../src/render/pdf.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

// 300pt page, 10pt margins, no header/footer => 280pt of printable body.
// Banner occupies 40pt; the container that follows is 400pt tall, taller than a whole empty page.
const PRINTABLE_BODY_POINTS = 280;

function block(name, top, text) {
  return `<Textbox Name="${name}"><Top>${top}pt</Top><Left>0pt</Left><Width>200pt</Width><Height>20pt</Height>
    <Paragraphs><Paragraph><TextRuns><TextRun><Value>${text}</Value>
      <Style><FontFamily>Arial</FontFamily></Style></TextRun></TextRuns></Paragraph></Paragraphs>
    <Style/></Textbox>`;
}

// The container declares no fill and no border, so it is a pure layout container and may be fragmented.
const rdl = `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <ReportSections><ReportSection><Body><ReportItems>
    <Rectangle Name="Banner"><Top>0pt</Top><Left>0pt</Left><Width>200pt</Width><Height>40pt</Height>
      <ReportItems>${block('BannerText', 5, 'Banner')}</ReportItems><Style/>
    </Rectangle>
    <Rectangle Name="TallContainer"><Top>40pt</Top><Left>0pt</Left><Width>200pt</Width><Height>400pt</Height>
      <ReportItems>
        ${block('First', 0, 'First block')}
        ${block('Second', 100, 'Second block')}
        ${block('Third', 200, 'Third block')}
        ${block('Fourth', 350, 'Fourth block')}
      </ReportItems><Style/>
    </Rectangle>
  </ReportItems><Height>440pt</Height><Style/></Body><Width>200pt</Width>
  <Page><PageHeight>300pt</PageHeight><PageWidth>400pt</PageWidth>
    <LeftMargin>10pt</LeftMargin><RightMargin>10pt</RightMargin>
    <TopMargin>10pt</TopMargin><BottomMargin>10pt</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`;

const model = parseRdl(Buffer.from(rdl, 'utf8'));
const request = { outputFileName: 'oversized-container', parameters: {}, datasets: {} };

function pageTexts(page) {
  return page.items.flatMap((item) => [
    item.text || '',
    ...(item.lines || []).flatMap((line) => (line.runs || []).map((run) => run.text || '')),
  ]).join(' ');
}

test('a container taller than the page starts on the page it stands on instead of blanking it', async () => {
  const pdf = await renderPdf(model, { ...request, output: 'PDF' }, config, { captureLayoutTrace: true });
  const first = pageTexts(pdf.layoutTrace.pages[0]);
  assert.ok(first.includes('Banner'), 'the banner belongs on page 1');
  // The defect: page 1 held the banner and nothing else, because the 400pt container was deferred whole.
  assert.ok(first.includes('First block'), 'the oversized container must begin on page 1');
  assert.ok(first.includes('Second block'), 'and keep filling the remainder of page 1');
});

test('the content that genuinely does not fit still moves to a later page', async () => {
  const pdf = await renderPdf(model, { ...request, output: 'PDF' }, config, { captureLayoutTrace: true });
  assert.ok(pdf.pageCount >= 2);
  const later = pdf.layoutTrace.pages.slice(1).map(pageTexts).join(' ');
  assert.ok(later.includes('Fourth block'), 'the block past the page boundary must flow onward');
  assert.ok(!pageTexts(pdf.layoutTrace.pages[0]).includes('Fourth block'));
});

test('no page is left with an empty printable body while content is still pending', async () => {
  const pdf = await renderPdf(model, { ...request, output: 'PDF' }, config, { captureLayoutTrace: true });
  // Every page except the last must carry content reaching a meaningful part of the printable body.
  for (const page of pdf.layoutTrace.pages.slice(0, -1)) {
    const painted = page.items.filter((item) => item.region !== 'footer' && item.region !== 'header');
    const lowest = Math.max(0, ...painted.map((item) => Number(item.y || 0) + Number(item.height || 0)));
    assert.ok(
      lowest > PRINTABLE_BODY_POINTS / 2,
      `page ${page.number} stops at ${lowest.toFixed(1)}pt of ${PRINTABLE_BODY_POINTS}pt printable body`,
    );
  }
});

// DOCX_EDITABLE is built from the canonical PDF trace, so it inherits the corrected pagination rather than
// re-deriving it. Asserted through Word's own mechanics: the first section must carry the container's text.
test('DOCX_EDITABLE inherits the corrected first page from the canonical trace', async () => {
  const docx = await renderEditableDocx(model, { ...request, output: 'DOCX_EDITABLE' }, config);
  const documentXml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  // Styled text is laid out token by token, so each word is its own w:t run: match single tokens.
  const firstSection = documentXml.split('<w:sectPr')[0];
  const runs = [...firstSection.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((match) => match[1]);
  assert.ok(runs.includes('Banner'));
  assert.ok(runs.includes('First'), 'Word page 1 must not be blank below the banner');
  assert.ok(runs.includes('Second'), 'and must keep filling the remainder of Word page 1');
});
