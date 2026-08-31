// A data region nested in a tablix cell is CONTENT of that cell, so SSRS paints the containing region's
// grid line after it: every RPL item draws its background, then its children, then its own border. The PDF
// renderer collects every border edge and strokes the collinear pieces of each identical line as one merged
// run, and it used to stroke those runs in the order each LINE was first seen. A column rule first seen on
// the header row was therefore stroked before a nested region's edges on the detail rows below, so a nested
// cell's White border — the standard RDL idiom for suppressing an interior rule — repainted the column rule
// white for exactly the height of the nested region, leaving a rule that was black where the outer cell
// extended past the nested region and white where it did not.
//
// The construct is generic: an outer tablix whose cell hosts a nested tablix shorter than the outer row,
// where the nested cell declares a White edge coincident with a black grid line. Nothing here depends on a
// report, a colour pair, or a coordinate: the assertions resolve their geometry from the layout trace.
import assert from 'node:assert/strict';
import test from 'node:test';
import zlib from 'node:zlib';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { PDFDocument, PDFName } from 'pdf-lib';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

const blackBox = (name, value) => `<Textbox Name="${name}"><CanGrow>true</CanGrow>
  <Paragraphs><Paragraph><TextRuns><TextRun><Value>${value}</Value>
    <Style><FontFamily>Arial</FontFamily><FontSize>8pt</FontSize></Style></TextRun></TextRuns></Paragraph></Paragraphs>
  <Style><Border><Style>Solid</Style></Border>
    <TopBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></TopBorder>
    <BottomBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></BottomBorder>
    <LeftBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></LeftBorder>
    <RightBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></RightBorder></Style>
</Textbox>`;

// `white` is the declared colour of the nested cell's suppressed edges. It is passed through verbatim, so
// the same fixture exercises a literal colour and an expression that resolves to one.
const rdlFor = (white) => `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="D">
    <Fields>
      <Field Name="Key"><DataField>Key</DataField><TypeName>System.String</TypeName></Field>
      <Field Name="Lead"><DataField>Lead</DataField><TypeName>System.String</TypeName></Field>
      <Field Name="Inner"><DataField>Inner</DataField><TypeName>System.String</TypeName></Field>
      <Field Name="Wide"><DataField>Wide</DataField><TypeName>System.String</TypeName></Field>
    </Fields>
    <Query><CommandText>never executed</CommandText></Query>
  </DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Tablix Name="Outer">
      <TablixBody>
        <TablixColumns>
          <TablixColumn><Width>1in</Width></TablixColumn>
          <TablixColumn><Width>1in</Width></TablixColumn>
          <TablixColumn><Width>2in</Width></TablixColumn>
        </TablixColumns>
        <TablixRows>
          <TablixRow><Height>0.25in</Height><TablixCells>
            <TablixCell><CellContents>${blackBox('H1', 'Lead')}</CellContents></TablixCell>
            <TablixCell><CellContents>${blackBox('H2', 'Nested')}</CellContents></TablixCell>
            <TablixCell><CellContents>${blackBox('H3', 'Wide')}</CellContents></TablixCell>
          </TablixCells></TablixRow>
          <TablixRow><Height>0.25in</Height><TablixCells>
            <TablixCell><CellContents>${blackBox('LeadText', '=Fields!Lead.Value')}</CellContents></TablixCell>
            <TablixCell><CellContents>
              <Tablix Name="Inner">
                <TablixBody>
                  <TablixColumns><TablixColumn><Width>1in</Width></TablixColumn></TablixColumns>
                  <TablixRows><TablixRow><Height>0.25in</Height><TablixCells><TablixCell><CellContents>
                    <Textbox Name="InnerText"><CanGrow>true</CanGrow>
                      <Paragraphs><Paragraph><TextRuns><TextRun><Value>=Fields!Inner.Value</Value>
                        <Style><FontFamily>Arial</FontFamily><FontSize>8pt</FontSize></Style></TextRun></TextRuns></Paragraph></Paragraphs>
                      <Style>
                        <Border><Color>White</Color><Style>Solid</Style></Border>
                        <TopBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></TopBorder>
                        <BottomBorder><Color>${white}</Color><Style>Solid</Style><Width>1pt</Width></BottomBorder>
                        <LeftBorder><Color>${white}</Color><Style>Solid</Style><Width>1pt</Width></LeftBorder>
                        <RightBorder><Color>${white}</Color><Style>Solid</Style><Width>1pt</Width></RightBorder>
                      </Style>
                    </Textbox>
                  </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
                </TablixBody>
                <TablixColumnHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixColumnHierarchy>
                <TablixRowHierarchy><TablixMembers><TablixMember><Group Name="InnerGroup"><GroupExpressions><GroupExpression>=Fields!Inner.Value</GroupExpression></GroupExpressions></Group></TablixMember></TablixMembers></TablixRowHierarchy>
                <Height>0.25in</Height><Width>1in</Width>
                <Style>
                  <Border><Style>Solid</Style></Border>
                  <TopBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></TopBorder>
                  <BottomBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></BottomBorder>
                  <LeftBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></LeftBorder>
                  <RightBorder><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></RightBorder>
                </Style>
              </Tablix>
            </CellContents></TablixCell>
            <TablixCell><CellContents>${blackBox('WideText', '=Fields!Wide.Value')}</CellContents></TablixCell>
          </TablixCells></TablixRow>
        </TablixRows>
      </TablixBody>
      <TablixColumnHierarchy><TablixMembers><TablixMember/><TablixMember/><TablixMember/></TablixMembers></TablixColumnHierarchy>
      <TablixRowHierarchy><TablixMembers>
        <TablixMember/>
        <TablixMember><Group Name="OuterGroup"><GroupExpressions><GroupExpression>=Fields!Key.Value</GroupExpression></GroupExpressions></Group></TablixMember>
      </TablixMembers></TablixRowHierarchy>
      <DataSetName>D</DataSetName><Height>0.5in</Height><Width>4in</Width>
      <Style><Border><Style>None</Style></Border></Style>
    </Tablix>
  </ReportItems><Height>3in</Height><Style/></Body><Width>4in</Width>
  <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`;

// The wide column's value wraps over several lines, so the outer row is far taller than the nested region
// it hosts — the state that exposed the defect.
const TALL = 'Tall wrapped value that grows the outer row far beyond its nested region. '.repeat(3);
const datasets = {
  D: [
    { Key: 'A', Lead: 'Lead A', Inner: 'Inner A', Wide: TALL },
    { Key: 'B', Lead: 'Lead B', Inner: 'Inner B', Wide: TALL },
  ],
};
const request = (output = 'PDF') => ({ output, outputFileName: 'nested-border-order', parameters: {}, datasets });

// Every stroked line on a page, in the order the content stream paints them. PDFKit writes each merged
// border run as a colour, an `x y m`, an `x y l` and an `S`.
async function strokesOnFirstPage(buffer) {
  const page = (await PDFDocument.load(buffer)).getPages()[0];
  const contents = page.node.context.lookup(page.node.get(PDFName.of('Contents')));
  let data = contents.constructor.name === 'PDFArray'
    ? Buffer.concat(contents.asArray().map((ref) => Buffer.from(page.node.context.lookup(ref).getContents())))
    : Buffer.from(contents.getContents());
  try { data = zlib.inflateSync(data); } catch { /* already flat */ }
  const strokes = [];
  let colour = null;
  let move = null;
  for (const raw of data.toString('latin1').split('\n')) {
    const line = raw.trim();
    let match = line.match(/^([\d.]+) ([\d.]+) ([\d.]+) SCN$/);
    if (match) { colour = match.slice(1, 4).map(Number).join(','); continue; }
    match = line.match(/^([-\d.]+) ([-\d.]+) m$/);
    if (match) { move = [Number(match[1]), Number(match[2])]; continue; }
    match = line.match(/^([-\d.]+) ([-\d.]+) l$/);
    if (match && move) strokes.push({ colour, x1: move[0], y1: move[1], x2: Number(match[1]), y2: Number(match[2]) });
  }
  return strokes;
}

// The colour actually visible on a vertical rule at `x`, level with `y`: the LAST stroke painted there.
function visibleVerticalColourAt(strokes, x, y) {
  const covering = strokes.filter((stroke) => Math.abs(stroke.x1 - stroke.x2) < 0.01
    && Math.abs(stroke.x1 - x) <= 0.6
    && Math.min(stroke.y1, stroke.y2) <= y
    && Math.max(stroke.y1, stroke.y2) >= y);
  return covering.at(-1)?.colour ?? null;
}

// Geometry of one nested region and the outer cell hosting it, read out of the layout trace so the
// assertions never hard-code a coordinate.
function bands(trace) {
  const items = trace.pages.flatMap((page) => page.items);
  const nested = items.find((item) => item.kind === 'tablixCell' && item.itemName === 'InnerText');
  const host = items.find((item) => item.kind === 'tablixCell'
    && item.itemName === null
    && Math.abs(item.x - nested.x) < 0.01
    && Math.abs(item.y - nested.y) < 0.01
    && item.height > nested.height + 1);
  assert.ok(host, 'the fixture must place the nested region in an outer cell taller than it');
  return {
    leftEdge: nested.x,
    rightEdge: nested.x + nested.width,
    insideNested: nested.y + (nested.height / 2),
    belowNested: nested.y + nested.height + ((host.height - nested.height) / 2),
  };
}

for (const [label, white] of [['a literal White', 'White'], ['an expression-backed White', '=IIf(True, "White", "Black")']]) {
  test(`PDF: ${label} nested-cell edge never repaints the containing tablix column rule`, async () => {
    const rendered = await renderPdf(parseRdl(rdlFor(white)), request(), config, { captureLayoutTrace: true });
    const strokes = await strokesOnFirstPage(rendered.buffer);
    const {
      leftEdge, rightEdge, insideNested, belowNested,
    } = bands(rendered.layoutTrace);

    // The symptom was a rule that changed colour partway down one outer cell. Both column rules must show
    // the same colour level with the nested region and level with the empty remainder below it.
    for (const [side, x] of [['left', leftEdge], ['right', rightEdge]]) {
      const across = visibleVerticalColourAt(strokes, x, insideNested);
      const below = visibleVerticalColourAt(strokes, x, belowNested);
      assert.equal(across, '0,0,0', `the ${side} column rule went white across the nested region`);
      assert.equal(below, '0,0,0', `the ${side} column rule is not black below the nested region`);
      assert.equal(across, below, `the ${side} column rule changes colour partway down one cell`);
    }

    // The declared White edges are still painted — the rule is about paint ORDER, not about dropping a
    // border the RDL asked for. Suppressing them instead would hide this defect and break any report that
    // paints White over a fill.
    assert.ok(
      strokes.some((stroke) => stroke.colour === '1,1,1'),
      'the nested cell declared White edges and they must still be stroked',
    );
    // ...and every one of them is painted before the containing region's rules, never after.
    const lastWhite = strokes.findLastIndex((stroke) => stroke.colour === '1,1,1');
    const lastBlack = strokes.findLastIndex((stroke) => stroke.colour === '0,0,0');
    assert.ok(lastBlack > lastWhite, 'a containing region must stroke its rules after the nested region');
  });
}

// DOCX_VISUAL is not exercised here: it rasterizes this very PDF at 300 DPI, one page image per PDF page,
// so it inherits the canonical stroke order verbatim and has no border model of its own.

test('editable DOCX keeps the containing tablix column rule on the nested cell', async () => {
  const docx = await renderEditableDocx(parseRdl(rdlFor('White')), request('DOCX_EDITABLE'), config);
  const documentXml = await (await JSZip.loadAsync(docx.buffer)).file('word/document.xml').async('string');
  // Word splits a value over several runs, so the cell's text is the concatenation of its `w:t` nodes.
  const cellText = (cell) => [...cell.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((match) => match[1]).join('');
  const nestedCells = [...documentXml.matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)]
    .map((match) => match[0])
    .filter((cell) => /^Inner [AB]$/.test(cellText(cell)));
  assert.equal(nestedCells.length, 2, 'both nested rows must reach Word as native cells');
  for (const cell of nestedCells) {
    // Word has no paint order: one colour per shared edge. The containing column rule must be the colour
    // written there, not the nested cell's suppressed White.
    for (const side of ['left', 'right']) {
      const border = cell.match(new RegExp(`<w:${side} w:val="([a-z]+)" w:color="([0-9A-Fa-f]{6}|auto)"`));
      assert.ok(border, `the nested cell must declare a ${side} border`);
      assert.equal(border[2].toUpperCase(), '000000', `Word wrote a ${side} column rule of ${border[2]}`);
    }
  }
});

test('XLSX keeps the containing tablix column rule on the nested cell', async () => {
  const rendered = await renderExcel(parseRdl(rdlFor('White')), request('XLSX'), config, null);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(rendered.buffer);
  const seen = [];
  workbook.worksheets[0].eachRow((row) => row.eachCell((cell) => {
    if (!/^Inner [AB]$/.test(String(cell.value ?? ''))) return;
    seen.push({
      left: cell.border?.left?.color?.argb ?? null,
      right: cell.border?.right?.color?.argb ?? null,
    });
  }));
  assert.equal(seen.length, 2, 'both nested rows must reach Excel as cells');
  // Excel, like Word, stores one colour per edge: the containing column rule wins.
  for (const borders of seen) assert.deepEqual(borders, { left: 'FF000000', right: 'FF000000' });
});
