// An RDL container and the item it holds routinely declare edges that differ by a fraction of a point,
// because an inch-valued container width and an inch-valued child width rarely land on the same quarter
// point. The fixed-layout PDF strokes both rules and the two strokes overlap into one visible line. The
// renderers that rebuild the page as a native grid cannot do that: they have to materialize the gap as a
// real band, and a band narrower than the certified 0.5pt geometry tolerance is not drawable at that
// width. Word and Excel keep both bounding cells and paint a border on each, so one canonical rule became
// a visible double line and the perpendicular rule lost its corner join.
//
// The rule is a property of the grid quantizer, not of any report: boundaries closer than the tolerance
// are one grid line, boundaries at or beyond it stay distinct. Both facts are asserted below against the
// same synthetic report parameterised only by the container's extra width.
//
// Renderer impact: fixed and tested in DOCX_EDITABLE and XLSX (both build an absolute native grid).
// PDF is the layout authority and is deliberately unchanged - the assertion below proves its trace still
// carries both distinct edges. DOCX_VISUAL rasterizes that unchanged canonical PDF, so it inherits it.
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { loadConfig } from '../src/config.js';
import { parseRdl } from '../src/rdl/parser.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';
import { buildGridBoundaries, GRID_BOUNDARY_TOLERANCE_POINTS } from '../src/render/gridBoundaries.js';

const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });
const request = {
  outputFileName: 'subtolerance-grid-boundary',
  parameters: {},
  datasets: { Rows: [{ Label: 'only' }] },
};

const TABLIX_WIDTH_IN = 3.76456;
const SUB_TOLERANCE_EXTRA_IN = 0.00365; // 0.26pt: below the tolerance, so the two edges are one grid line
const ABOVE_TOLERANCE_EXTRA_IN = 0.01389; // 1.00pt: above the tolerance, so both edges must survive

function boxedCell(name, value, sides) {
  const border = (side) => `<${side}Border><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></${side}Border>`;
  return `<Textbox Name="${name}">
    <CanGrow>true</CanGrow>
    <Paragraphs><Paragraph><TextRuns><TextRun><Value>${value}</Value>
      <Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style>
    </TextRun></TextRuns></Paragraph></Paragraphs>
    <Style><Border><Style>None</Style></Border>${sides.map(border).join('')}</Style>
  </Textbox>`;
}

// The container is the only variable: everything else is a plain two-by-two tablix whose outer cell
// borders draw one box, exactly the construct that produced the doubled Word and Excel rules.
function report(containerExtraInches) {
  const containerWidth = (TABLIX_WIDTH_IN + containerExtraInches).toFixed(5);
  return Buffer.from(`<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <DataSources><DataSource Name="S"><DataSourceReference>/x</DataSourceReference></DataSource></DataSources>
  <DataSets><DataSet Name="Rows">
    <Fields><Field Name="Label"><DataField>Label</DataField><TypeName>System.String</TypeName></Field></Fields>
    <Query><DataSourceName>S</DataSourceName><CommandText>ignored</CommandText></Query>
  </DataSet></DataSets>
  <ReportSections><ReportSection>
    <Body>
      <ReportItems>
        <Rectangle Name="Frame">
          <ReportItems>
            <Tablix Name="Grid">
              <TablixBody>
                <TablixColumns>
                  <TablixColumn><Width>1.04465in</Width></TablixColumn>
                  <TablixColumn><Width>2.71991in</Width></TablixColumn>
                </TablixColumns>
                <TablixRows>
                  <TablixRow><Height>0.25in</Height><TablixCells>
                    <TablixCell><CellContents>${boxedCell('TopLeft', 'TOP_LEFT', ['Top', 'Left'])}</CellContents></TablixCell>
                    <TablixCell><CellContents>${boxedCell('TopRight', 'TOP_RIGHT', ['Top', 'Right'])}</CellContents></TablixCell>
                  </TablixCells></TablixRow>
                  <TablixRow><Height>0.25in</Height><TablixCells>
                    <TablixCell><CellContents>${boxedCell('EndLeft', 'END_LEFT', ['Bottom', 'Left'])}</CellContents></TablixCell>
                    <TablixCell><CellContents>${boxedCell('EndRight', 'END_RIGHT', ['Bottom', 'Right'])}</CellContents></TablixCell>
                  </TablixCells></TablixRow>
                </TablixRows>
              </TablixBody>
              <TablixColumnHierarchy><TablixMembers><TablixMember/><TablixMember/></TablixMembers></TablixColumnHierarchy>
              <TablixRowHierarchy><TablixMembers><TablixMember/><TablixMember/></TablixMembers></TablixRowHierarchy>
              <DataSetName>Rows</DataSetName>
              <Top>0in</Top><Left>0in</Left><Height>0.5in</Height><Width>${TABLIX_WIDTH_IN}in</Width>
              <Style><FontFamily>Arial</FontFamily><Border><Style>None</Style></Border></Style>
            </Tablix>
          </ReportItems>
          <Top>0in</Top><Left>0in</Left><Height>0.5in</Height><Width>${containerWidth}in</Width>
          <Style><Border><Color>Black</Color><Style>Solid</Style><Width>1pt</Width></Border></Style>
        </Rectangle>
      </ReportItems>
      <Height>1in</Height><Style/>
    </Body>
    <Width>7in</Width>
    <Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth><LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin></Page>
  </ReportSection></ReportSections>
</Report>`, 'utf8');
}

async function wordGridColumns(containerExtraInches) {
  const rendered = await renderEditableDocx(parseRdl(report(containerExtraInches)), request, config);
  const zip = await JSZip.loadAsync(rendered.buffer);
  const document = await zip.file('word/document.xml').async('string');
  const grid = document.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/)[0];
  return [...grid.matchAll(/<w:gridCol\s+w:w="(\d+)"\s*\/>/g)].map((match) => Number(match[1]));
}

async function excelColumnCount(containerExtraInches) {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-grid-boundary-'));
  try {
    const rendered = await renderExcel(
      parseRdl(report(containerExtraInches)),
      { ...request, excelLayoutMode: 'REPORT' },
      config,
      tempDir,
    );
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(rendered.buffer);
    return workbook.worksheets[0].columnCount;
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

test('the grid quantizer merges sub-tolerance boundaries and keeps everything else distinct', () => {
  const tolerance = GRID_BOUNDARY_TOLERANCE_POINTS;
  // The canvas extent is structural, so the trailing cluster keeps the outermost coordinate.
  const merged = buildGridBoundaries([0, 100, 307, 307.25]);
  assert.deepEqual(merged.boundaries, [0, 100, 307.25]);
  assert.equal(merged.indexOf(307), merged.indexOf(307.25), 'both collapsed edges resolve to one grid line');
  // An interior cluster keeps the coordinate the most edges already sit on, so the fewest items move.
  const interior = buildGridBoundaries([0, 200, 60, 60.25, 60.25, 60.25]);
  assert.deepEqual(interior.boundaries, [0, 60.25, 200]);
  const distinct = buildGridBoundaries([0, 100, 307, 308]);
  assert.deepEqual(distinct.boundaries, [0, 100, 307, 308]);
  // No collapse may move an edge by the tolerance or more, otherwise geometry leaves certification range.
  for (const value of [307, 307.25]) {
    assert.ok(Math.abs(merged.boundaries[merged.indexOf(value)] - value) < tolerance);
  }
});

test('a span whose two ends anchor the same item is never collapsed away', () => {
  // A 0.25pt-wide report item would otherwise lose the only grid column it can occupy.
  const protectedGrid = buildGridBoundaries([0, 100, 60, 60.25], { protectedSpans: [[60, 60.25]] });
  assert.deepEqual(protectedGrid.boundaries, [0, 60, 60.25, 100]);
  assert.notEqual(protectedGrid.indexOf(60), protectedGrid.indexOf(60.25));
});

test('PDF keeps both declared edges: it is the layout authority and is not quantized', async () => {
  const captured = await renderPdf(parseRdl(report(SUB_TOLERANCE_EXTRA_IN)), request, config, {
    captureLayoutTrace: true,
  });
  const page = captured.layoutTrace.pages[0];
  const container = page.items.find((item) => item.itemName === 'Frame');
  const cell = page.items.find((item) => item.itemName === 'TopRight');
  const containerRight = container.x + container.width;
  const cellRight = cell.x + cell.width;
  assert.ok(containerRight > cellRight, 'the container is declared wider than the tablix it holds');
  assert.ok(containerRight - cellRight < 0.5, 'and by less than the certified geometry tolerance');
});

test('Word collapses the sub-tolerance container gap into one grid line', async () => {
  const columns = await wordGridColumns(SUB_TOLERANCE_EXTRA_IN);
  const toleranceTwips = GRID_BOUNDARY_TOLERANCE_POINTS * 20;
  const slivers = columns.filter((width) => width > 0 && width < toleranceTwips);
  assert.deepEqual(slivers, [], `Word cannot draw a table band narrower than the tolerance: ${columns}`);
});

test('Word still separates a container gap at or above the tolerance', async () => {
  const [collapsed, separated] = await Promise.all([
    wordGridColumns(SUB_TOLERANCE_EXTRA_IN),
    wordGridColumns(ABOVE_TOLERANCE_EXTRA_IN),
  ]);
  assert.equal(
    separated.length,
    collapsed.length + 1,
    'a one-point container gap is a real band and must keep its own Word column',
  );
  assert.ok(separated.some((width) => width === 20), 'and that band keeps its exact one-point width');
});

test('Excel collapses the sub-tolerance container gap and keeps a real one', async () => {
  const [collapsed, separated] = await Promise.all([
    excelColumnCount(SUB_TOLERANCE_EXTRA_IN),
    excelColumnCount(ABOVE_TOLERANCE_EXTRA_IN),
  ]);
  assert.equal(
    separated,
    collapsed + 1,
    'the sub-tolerance gap must not become its own worksheet column while a one-point gap still does',
  );
});
