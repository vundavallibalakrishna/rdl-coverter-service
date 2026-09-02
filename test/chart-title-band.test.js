import assert from 'node:assert/strict';
import test from 'node:test';
import { inflateSync } from 'node:zlib';
import { PDFDocument } from 'pdf-lib';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderPdf } from '../src/render/pdf.js';

// How much of the chart a docked title occupies.
//
// An SSRS chart title is docked to a side of the chart and occupies that whole band: its area — and so the
// BackgroundColor the RDL declares for it — spans the chart's content width for a top/bottom title, and
// `Position` (TopLeft / TopCenter / TopRight …) places the caption inside that band.
//
// This renderer sized the title's box to the caption instead, so a declared title background rendered as a
// small tab floating above the plot rather than the full-width header bar the report asks for.
//
// A title with no background is unaffected: the caption lands in the same place either way, which is what
// the counterexample below pins down.

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false', RDL_RENDER_TIMEOUT_MS: '60000' });

const CHART_WIDTH_IN = 6;
const CHART_HEIGHT_IN = 3;

const report = ({ position = 'TopCenter', background = '#4f81bd', textAlign = 'General' } = {}) => Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSets><DataSet Name="DS"><Fields>
    <Field Name="C"><DataField>C</DataField><TypeName>System.String</TypeName></Field>
    <Field Name="V"><DataField>V</DataField><TypeName>System.Int32</TypeName></Field>
  </Fields><Query><CommandText>never executed</CommandText></Query></DataSet></DataSets>
  <ReportSections><ReportSection><Body><ReportItems>
    <Chart Name="C1">
      <ChartCategoryHierarchy><ChartMembers><ChartMember>
        <Group Name="Cat"><GroupExpressions><GroupExpression>=Fields!C.Value</GroupExpression></GroupExpressions></Group>
      </ChartMember></ChartMembers></ChartCategoryHierarchy>
      <ChartSeriesHierarchy><ChartMembers><ChartMember/></ChartMembers></ChartSeriesHierarchy>
      <ChartData><ChartSeriesCollection><ChartSeries Name="S1">
        <ChartDataPoints><ChartDataPoint><ChartDataPointValues><Y>=Sum(Fields!V.Value)</Y></ChartDataPointValues></ChartDataPoint></ChartDataPoints>
        <Type>Column</Type>
      </ChartSeries></ChartSeriesCollection></ChartData>
      <ChartAreas><ChartArea Name="A1"><ChartCategoryAxes><ChartAxis/></ChartCategoryAxes>
        <ChartValueAxes><ChartAxis/></ChartValueAxes></ChartArea></ChartAreas>
      <ChartTitles><ChartTitle Name="T1"><Caption>CHART_TITLE</Caption>
        <Position>${position}</Position>
        <Style>${background ? `<BackgroundColor>${background}</BackgroundColor>` : ''}
          <FontFamily>Arial</FontFamily><FontSize>9pt</FontSize><TextAlign>${textAlign}</TextAlign><Color>White</Color></Style>
      </ChartTitle></ChartTitles>
      <ChartLegend Name="L1"><Hidden>true</Hidden></ChartLegend>
      <DataSetName>DS</DataSetName>
      <Top>0in</Top><Left>0in</Left><Width>${CHART_WIDTH_IN}in</Width><Height>${CHART_HEIGHT_IN}in</Height><Style/>
    </Chart>
  </ReportItems><Height>4in</Height><Style/></Body><Width>7in</Width>
  <Page><PageHeight>6in</PageHeight><PageWidth>7.5in</PageWidth><LeftMargin>0.25in</LeftMargin><RightMargin>0.25in</RightMargin>
    <TopMargin>0.25in</TopMargin><BottomMargin>0.25in</BottomMargin></Page>
  </ReportSection></ReportSections></Report>`, 'utf8');

const rows = [{ C: 'A', V: 3 }, { C: 'B', V: 5 }];
const request = { outputFileName: 'chart-title-band', parameters: {}, datasets: { DS: rows } };

// The chart is drawn straight into the page, so its title band is a filled rectangle in the page's own
// content stream.
const contentStream = async (definition) => {
  const rendered = await renderPdf(parseRdl(definition), request, config);
  const pdf = await PDFDocument.load(rendered.buffer);
  const contents = pdf.getPages()[0].node.Contents();
  const raw = contents.asArray
    ? Buffer.concat(contents.asArray().map((ref) => Buffer.from(pdf.context.lookup(ref).getContents())))
    : Buffer.from(contents.getContents());
  try { return inflateSync(raw).toString('latin1'); } catch { return raw.toString('latin1'); }
};

// Every filled rectangle painted in the title's declared colour.
const titleBands = async (definition) => {
  const stream = await contentStream(definition);
  const bands = [];
  // "<r> <g> <b> scn ... x y w h re f" — the fill colour is set immediately before the rectangle.
  const pattern = /([\d.]+) ([\d.]+) ([\d.]+) scn\s*(?:[^]{0,40}?)([\d.-]+) ([\d.-]+) ([\d.-]+) ([\d.-]+) re\s*f/g;
  for (const match of stream.matchAll(pattern)) {
    const [red, green, blue] = [match[1], match[2], match[3]].map(Number);
    // #4f81bd
    if (Math.abs(red - 0x4f / 255) < 0.01 && Math.abs(green - 0x81 / 255) < 0.01 && Math.abs(blue - 0xbd / 255) < 0.01) {
      bands.push({ x: Number(match[4]), y: Number(match[5]), width: Number(match[6]), height: Number(match[7]) });
    }
  }
  return bands;
};

const CHART_WIDTH_PT = CHART_WIDTH_IN * 72;

test('a top-docked title fills the width of the chart it is docked to', async () => {
  const bands = await titleBands(report());
  assert.equal(bands.length, 1, 'the title paints exactly one background band');
  const [band] = bands;
  assert.ok(
    band.width > CHART_WIDTH_PT - 20,
    `the band must span the chart (${CHART_WIDTH_PT}pt wide), got ${band.width.toFixed(1)}pt`,
  );
  assert.ok(band.width <= CHART_WIDTH_PT, 'and must not exceed it');
  assert.ok(band.height < 30, 'while keeping the title band its own height');
});

test('a bottom-docked title fills the same band at the other end', async () => {
  const [top] = await titleBands(report({ position: 'TopLeft' }));
  const [bottom] = await titleBands(report({ position: 'BottomRight' }));
  assert.ok(Math.abs(top.width - bottom.width) < 0.5, 'both bands span the chart width');
  assert.ok(Math.abs(top.x - bottom.x) < 0.5, 'and start at the same edge');
  assert.ok(bottom.y > top.y, 'the bottom title sits below the top one');
});

test('the Position alignment places the caption inside the band', async () => {
  // The caption is drawn immediately after its own background band, so the first text placement following
  // that fill is the caption — no other chart text can come between them.
  const captionX = async (definition) => {
    const stream = await contentStream(definition);
    const band = stream.search(/0\.309\d* 0\.505\d* 0\.741\d* scn/);
    assert.ok(band >= 0, 'the title band must be painted');
    const after = stream.slice(band);
    const placement = /1 0 0 1 ([\d.-]+) ([\d.-]+) Tm/.exec(after);
    assert.ok(placement, 'the caption must follow its band');
    return Number(placement[1]);
  };
  const left = await captionX(report({ position: 'TopLeft' }));
  const centre = await captionX(report({ position: 'TopCenter' }));
  const right = await captionX(report({ position: 'TopRight' }));
  assert.ok(left < centre && centre < right, `caption must move with Position, got ${left}, ${centre}, ${right}`);
});

test('a title with no declared background changes nothing', async () => {
  // The counterexample: with no BackgroundColor there is no band, and the caption still lands where the
  // chart's own alignment puts it — widening an invisible box must not move any content.
  const bands = await titleBands(report({ background: null }));
  assert.deepEqual(bands, [], 'nothing is painted for a background-less title');
});
