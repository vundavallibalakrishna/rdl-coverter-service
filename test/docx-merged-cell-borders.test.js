// The page-locked Word canvas is one table per PDF page, so every item's edges become grid lines for the
// whole page width. An item therefore routinely spans several grid rows purely because some unrelated item
// elsewhere on the page starts part-way down it, and WordprocessingML expresses that as a vertical merge.
//
// The merged region's horizontal rules belong to its outer bands: the top to the first, the bottom to the
// last. Repeating the item's own top and bottom on every band draws a rule *inside* the cell at each grid
// row it crosses, and drops the item's background on every band but the first. Both are invisible while the
// bands are comfortably tall, but a band only as tall as the strokes themselves - which is what a
// neighbour starting a point below this item produces - renders the interior rule as a second border
// immediately under the real one, with the fill starting below it.
//
// This is Word-format-specific: it is a property of the OOXML `w:vMerge` representation. PDF strokes each
// item's borders once as geometry and has no merge concept; XLSX merges a range and distributes the
// resolved edges around its perimeter. The PDF assertion below is the counterexample.
import assert from 'node:assert/strict';
import test from 'node:test';
import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderPdf } from '../src/render/pdf.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const request = { outputFileName: 'docx-merged-cell-borders', parameters: {}, datasets: {} };

// `Boxed` is a bordered, shaded box. `Splitter` sits well clear of it horizontally but starts and ends
// inside its vertical span, so the shared page grid cuts `Boxed` into three bands it never asked for -
// and its first band is deliberately thinner than its own top padding. `Control` is the same box with
// nothing beside it, so it keeps one band and must be left exactly as it was.
const report = Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <ReportSections><ReportSection>
    <Body>
      <ReportItems>
        <Textbox Name="Boxed">
          <Paragraphs><Paragraph><TextRuns><TextRun><Value>BOXED</Value>
            <Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize><Color>White</Color></Style>
          </TextRun></TextRuns></Paragraph></Paragraphs>
          <Top>0in</Top><Left>0in</Left><Height>0.5in</Height><Width>2in</Width>
          <Style>
            <BackgroundColor>Black</BackgroundColor>
            <Border><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></Border>
            <PaddingTop>3pt</PaddingTop>
          </Style>
        </Textbox>
        <Textbox Name="Splitter">
          <Paragraphs><Paragraph><TextRuns><TextRun><Value>SPLITTER</Value>
            <Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style>
          </TextRun></TextRuns></Paragraph></Paragraphs>
          <Top>0.0139in</Top><Left>3in</Left><Height>0.2in</Height><Width>1in</Width>
          <Style><Border><Style>None</Style></Border></Style>
        </Textbox>
        <Textbox Name="Control">
          <Paragraphs><Paragraph><TextRuns><TextRun><Value>CONTROL</Value>
            <Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize><Color>White</Color></Style>
          </TextRun></TextRuns></Paragraph></Paragraphs>
          <Top>1in</Top><Left>0in</Left><Height>0.4in</Height><Width>2in</Width>
          <Style>
            <BackgroundColor>Black</BackgroundColor>
            <Border><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></Border>
            <PaddingTop>3pt</PaddingTop>
          </Style>
        </Textbox>
      </ReportItems>
      <Height>2in</Height><Style/>
    </Body>
    <Width>7in</Width>
    <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`, 'utf8');

function borderStyle(cell, side) {
  const value = cell['w:tcPr']?.['w:tcBorders']?.[`w:${side}`]?.['@w:val'];
  return value === undefined ? 'absent' : value;
}

function cellText(cell) {
  return [].concat(cell['w:p'] || []).flatMap((paragraph) => (
    [].concat(paragraph['w:r'] || []).map((run) => run['w:t']?.['#text'] ?? run['w:t'] ?? '')
  )).join('');
}

let cached = null;
async function pageTable() {
  if (!cached) {
    const rendered = await renderEditableDocx(parseRdl(report), request, config);
    const zip = await JSZip.loadAsync(rendered.buffer);
    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '@',
      isArray: (name) => ['w:tbl', 'w:tr', 'w:tc', 'w:p', 'w:r'].includes(name),
    });
    cached = parser.parse(await zip.file('word/document.xml').async('string'))['w:document']['w:body']['w:tbl'][0];
  }
  return cached;
}

// Walks the page table and returns the vertical-merge bands of the column that holds `text`, top to bottom.
async function mergeBands(text) {
  const table = await pageTable();
  const bands = [];
  let column = null;
  for (const row of table['w:tr'] || []) {
    let cursor = 0;
    for (const cell of [].concat(row['w:tc'] || [])) {
      const span = Number(cell['w:tcPr']?.['w:gridSpan']?.['@w:val'] || 1);
      const merge = cell['w:tcPr']?.['w:vMerge'];
      if (column === null && cellText(cell).includes(text)) column = cursor;
      if (column !== null && cursor === column && merge) {
        bands.push({
          merge: merge['@w:val'] || 'continue',
          top: borderStyle(cell, 'top'),
          bottom: borderStyle(cell, 'bottom'),
          left: borderStyle(cell, 'left'),
          right: borderStyle(cell, 'right'),
          shading: cell['w:tcPr']?.['w:shd']?.['@w:fill'] ?? null,
        });
      }
      cursor += span;
    }
  }
  return bands;
}

// Returns where the cell holding `text` puts its top padding: on the cell, or in the paragraph flow.
async function topPadding(text) {
  const table = await pageTable();
  for (const row of table['w:tr'] || []) {
    for (const cell of [].concat(row['w:tc'] || [])) {
      if (!cellText(cell).includes(text)) continue;
      return {
        cellMarginTwips: Number(cell['w:tcPr']?.['w:tcMar']?.['w:top']?.['@w:w'] ?? 0),
        paragraphBeforeTwips: Number([].concat(cell['w:p'])[0]?.['w:pPr']?.['w:spacing']?.['@w:before'] ?? 0),
      };
    }
  }
  throw new Error(`no cell contains ${text}`);
}

test('a vertically merged cell draws its top rule once, its bottom rule once, and no interior rule', async () => {
  const bands = await mergeBands('BOXED');
  assert.ok(bands.length >= 3, `the neighbour must split the box into several grid bands: ${bands.length}`);
  assert.equal(bands[0].merge, 'restart');
  assert.ok(bands.slice(1).every((band) => band.merge === 'continue'));

  assert.equal(bands[0].top, 'single', 'the first band carries the item’s real top rule');
  assert.equal(bands.at(-1).bottom, 'single', 'the last band carries the item’s real bottom rule');
  // Everything between them is interior to one report item and must be free of horizontal rules.
  assert.equal(bands[0].bottom, 'none', 'the first band must not close the cell it is still inside');
  assert.equal(bands.at(-1).top, 'none', 'the last band must not reopen it');
  for (const band of bands.slice(1, -1)) {
    assert.equal(band.top, 'none');
    assert.equal(band.bottom, 'none');
  }
  // The sides run the full height of the merge, so every band keeps them.
  for (const band of bands) {
    assert.equal(band.left, 'single');
    assert.equal(band.right, 'single');
  }
});

test('a vertically merged cell keeps its background on every band', async () => {
  const bands = await mergeBands('BOXED');
  for (const [index, band] of bands.entries()) {
    assert.equal(band.shading, '000000', `band ${index} lost the declared BackgroundColor`);
  }
});

test('a top margin too tall for its first band moves into the paragraph flow, keeping the fill continuous', async () => {
  // Word cannot start a cell's fill inside a band shorter than the cell's top margin, so it resumes the
  // fill in the next band and strands the top border above an unfilled strip - the extra rule.
  const split = await topPadding('BOXED');
  assert.equal(split.cellMarginTwips, 0, 'the margin cannot stay on a cell whose first band is thinner');
  assert.equal(split.paragraphBeforeTwips, 60, 'and the same 3pt offset must be carried by the content');

  // The identical box with nothing splitting it keeps the ordinary representation untouched.
  const whole = await topPadding('CONTROL');
  assert.equal(whole.cellMarginTwips, 60, 'an unsplit cell keeps its padding as a real cell margin');
  assert.equal(whole.paragraphBeforeTwips, 0, 'and must not gain paragraph space that would double it');
});

test('PDF is unaffected: it strokes the box once and has no merge concept', async () => {
  const captured = await renderPdf(parseRdl(report), request, config, { captureLayoutTrace: true });
  const boxed = captured.layoutTrace.pages[0].items.find((item) => item.itemName === 'Boxed');
  const splitter = captured.layoutTrace.pages[0].items.find((item) => item.itemName === 'Splitter');
  assert.ok(splitter.y > boxed.y && splitter.y < boxed.y + boxed.height, 'the neighbour starts inside the box');
  // One item, one set of four resolved edges - the split exists only in the Word grid.
  assert.equal(boxed.borders.top.style, 'Solid');
  assert.equal(boxed.borders.bottom.style, 'Solid');
});
