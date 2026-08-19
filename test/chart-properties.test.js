import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import JSZip from 'jszip';
import { PNG } from 'pngjs';
import { loadConfig } from '../src/config.js';
import { analyzeRdl, parseRdl } from '../src/rdl/parser.js';
import { normalizeDatasets } from '../src/render/common.js';
import { materializeChart } from '../src/render/chartData.js';
import { renderEditableDocx } from '../src/render/docx.js';
import { renderExcel } from '../src/render/excel.js';
import { renderPdf } from '../src/render/pdf.js';
import { rasterizePdf } from '../src/render/raster.js';
import { renderVisualDocx } from '../src/render/visualDocx.js';

const chartXml = ({ name, type, subtype = '', property, value, top, left = 0.1 }) => `
<Chart Name="${name}">
  <ChartCategoryHierarchy><ChartMembers><ChartMember><Group Name="${name}Category"><GroupExpressions>
    <GroupExpression>=Fields!Category.Value</GroupExpression>
  </GroupExpressions></Group><Label>=Fields!Category.Value</Label></ChartMember></ChartMembers></ChartCategoryHierarchy>
  <ChartSeriesHierarchy><ChartMembers><ChartMember><Label>Value</Label></ChartMember></ChartMembers></ChartSeriesHierarchy>
  <ChartData><ChartSeriesCollection><ChartSeries Name="ValueSeries"><ChartDataPoints><ChartDataPoint>
    <ChartDataPointValues><Y>=Sum(Fields!Amount.Value)</Y></ChartDataPointValues>
    <ChartDataLabel><UseValueAsLabel>true</UseValueAsLabel><Visible>true</Visible><Position>Outside</Position><Style/></ChartDataLabel>
    <Style/>
  </ChartDataPoint></ChartDataPoints><Type>${type}</Type>${subtype ? `<Subtype>${subtype}</Subtype>` : ''}
    <CustomProperties><CustomProperty><Name>${property}</Name><Value>${value}</Value></CustomProperty></CustomProperties>
  </ChartSeries></ChartSeriesCollection></ChartData>
  <ChartAreas><ChartArea Name="Default">
    <ChartCategoryAxes><ChartAxis Name="Primary"><Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize></Style>
      <LabelsAutoFitDisabled>true</LabelsAutoFitDisabled></ChartAxis></ChartCategoryAxes>
    <ChartValueAxes><ChartAxis Name="Primary"><Style><FontFamily>Arial</FontFamily><FontSize>8pt</FontSize></Style>
      <LabelsAutoFitDisabled>true</LabelsAutoFitDisabled></ChartAxis></ChartValueAxes>
  </ChartArea></ChartAreas>
  <DataSetName>D</DataSetName><Top>${top}in</Top><Left>${left}in</Left><Width>4in</Width><Height>2in</Height><Style/>
</Chart>`;

const rdl = (firstProperty = 'PointWidth') => `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <Body><ReportItems>
    ${chartXml({ name: 'Columns', type: 'Column', property: firstProperty, value: '0.5', top: 0.1 })}
    ${chartXml({ name: 'Pie', type: 'Shape', subtype: 'ExplodedDoughnut', property: 'PieLineColor', value: 'Black', top: 2.2 })}
  </ReportItems><Height>4.4in</Height></Body><Width>4.2in</Width>
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="Category"><DataField>Category</DataField></Field>
    <Field Name="Amount"><DataField>Amount</DataField></Field>
  </Fields></DataSet></DataSets>
  <Page><PageWidth>4.5in</PageWidth><PageHeight>5in</PageHeight><TopMargin>0.1in</TopMargin>
    <BottomMargin>0.1in</BottomMargin><LeftMargin>0.1in</LeftMargin><RightMargin>0.1in</RightMargin></Page>
</Report>`;

const sideBySideRdl = `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <Body><ReportItems>
    ${chartXml({ name: 'LeftChart', type: 'Column', property: 'PointWidth', value: '0.5', top: 0.1, left: 0.1 })}
    ${chartXml({ name: 'RightChart', type: 'Shape', subtype: 'Pie', property: 'PieLineColor', value: 'Black', top: 0.1, left: 4.2 })}
  </ReportItems><Height>2.2in</Height></Body><Width>8.2in</Width>
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="Category"><DataField>Category</DataField></Field>
    <Field Name="Amount"><DataField>Amount</DataField></Field>
  </Fields></DataSet></DataSets>
  <Page><PageWidth>8.5in</PageWidth><PageHeight>3in</PageHeight><TopMargin>0.1in</TopMargin>
    <BottomMargin>0.1in</BottomMargin><LeftMargin>0.1in</LeftMargin><RightMargin>0.1in</RightMargin></Page>
</Report>`;

const chartTitlePositionRdl = (position = null, textAlign = 'Left') => {
  const chart = chartXml({
    name: 'TitlePositionChart', type: 'Shape', subtype: 'Pie', property: 'PieLineColor', value: 'Black', top: 0.1,
  }).replace(
    '  <DataSetName>D</DataSetName>',
    `  <ChartTitles><ChartTitle Name="Title"><Caption>TITLE_POSITION_PROBE</Caption>
      ${position ? `<Position>${position}</Position>` : ''}
      <Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize><FontWeight>Bold</FontWeight>
        <TextAlign>${textAlign}</TextAlign></Style>
    </ChartTitle></ChartTitles>
  <DataSetName>D</DataSetName>`,
  );
  return `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <Body><ReportItems>${chart}</ReportItems><Height>2.2in</Height></Body><Width>4.2in</Width>
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="Category"><DataField>Category</DataField></Field>
    <Field Name="Amount"><DataField>Amount</DataField></Field>
  </Fields></DataSet></DataSets>
  <Page><PageWidth>4.4in</PageWidth><PageHeight>2.4in</PageHeight><TopMargin>0.1in</TopMargin>
    <BottomMargin>0.1in</BottomMargin><LeftMargin>0.1in</LeftMargin><RightMargin>0.1in</RightMargin></Page>
</Report>`;
};

const styledDoughnutRdl = (subtype = 'ExplodedDoughnut', legendPosition = 'RightCenter', pieStartAngle = null) => `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <Body><ReportItems>
    <Chart Name="StyledDoughnut">
      <ChartCategoryHierarchy><ChartMembers><ChartMember><Group Name="Category"><GroupExpressions>
        <GroupExpression>=Fields!Category.Value</GroupExpression>
      </GroupExpressions></Group><Label>=Fields!Category.Value</Label></ChartMember></ChartMembers></ChartCategoryHierarchy>
      <ChartSeriesHierarchy><ChartMembers><ChartMember><Label>Count</Label></ChartMember></ChartMembers></ChartSeriesHierarchy>
      <ChartData><ChartSeriesCollection><ChartSeries Name="Count"><ChartDataPoints><ChartDataPoint>
        <ChartDataPointValues><Y>=Sum(Fields!Amount.Value)</Y></ChartDataPointValues>
        <ChartDataLabel><UseValueAsLabel>true</UseValueAsLabel><Visible>true</Visible><Style><FontFamily>Arial</FontFamily></Style></ChartDataLabel>
        <Style/>
      </ChartDataPoint></ChartDataPoints><Type>Shape</Type><Subtype>${subtype}</Subtype>
        ${pieStartAngle === null ? '' : `<CustomProperties><CustomProperty><Name>PieStartAngle</Name><Value>${pieStartAngle}</Value></CustomProperty></CustomProperties>`}
      </ChartSeries></ChartSeriesCollection></ChartData>
      <ChartAreas><ChartArea Name="Default"><Style><BackgroundColor>#f2f2f2</BackgroundColor>
        <Border><Style>Solid</Style><Color>#808080</Color><Width>1pt</Width></Border></Style></ChartArea></ChartAreas>
      <ChartLegends><ChartLegend Name="Default"><Position>${legendPosition}</Position><Layout>Column</Layout>
        <Style><BackgroundColor>#ffffff</BackgroundColor><FontFamily>Arial</FontFamily><FontSize>8pt</FontSize>
          <Border><Style>Solid</Style><Color>#404040</Color><Width>1pt</Width></Border></Style>
      </ChartLegend></ChartLegends>
      <ChartTitles><ChartTitle Name="Default"><Caption>By Department</Caption><Position>TopLeft</Position>
        <Style><FontFamily>Arial</FontFamily><FontSize>9pt</FontSize><TextAlign>Left</TextAlign></Style>
      </ChartTitle></ChartTitles>
      <Palette>Custom</Palette><ChartCustomPaletteColors>
        <ChartCustomPaletteColor>#01b8aa</ChartCustomPaletteColor>
        <ChartCustomPaletteColor>#374649</ChartCustomPaletteColor>
      </ChartCustomPaletteColors>
      <DataSetName>D</DataSetName><Top>0.1in</Top><Left>0.1in</Left><Width>5.8in</Width><Height>3.5in</Height>
      <Style><BackgroundColor>#ffffff</BackgroundColor>
        <Border><Style>Solid</Style><Color>#000000</Color><Width>1pt</Width></Border></Style>
    </Chart>
  </ReportItems><Height>3.7in</Height></Body><Width>6in</Width>
  <DataSets><DataSet Name="D"><Fields>
    <Field Name="Category"><DataField>Category</DataField></Field>
    <Field Name="Amount"><DataField>Amount</DataField></Field>
  </Fields></DataSet></DataSets>
  <Page><PageWidth>6.2in</PageWidth><PageHeight>4in</PageHeight><TopMargin>0.1in</TopMargin>
    <BottomMargin>0.1in</BottomMargin><LeftMargin>0.1in</LeftMargin><RightMargin>0.1in</RightMargin></Page>
</Report>`;

const request = { parameters: {}, datasets: { D: [{ Category: 'A', Amount: 2 }, { Category: 'B', Amount: 3 }] } };
const config = loadConfig({ ...process.env, RDL_STRICT_FONTS: 'false' });

test('parses documented axis and series chart properties into the normalized model', () => {
  const model = parseRdl(rdl());
  const [columns, pie] = model.body.items;
  assert.equal(columns.categoryAxis.labelsAutoFitDisabled, 'true');
  assert.equal(columns.valueAxis.labelsAutoFitDisabled, 'true');
  assert.equal(columns.seriesDefs[0].customProperties.PointWidth, '0.5');
  assert.equal(pie.seriesDefs[0].customProperties.PieLineColor, 'Black');
  assert.equal(pie.seriesDefs[0].dataLabel.position, 'Outside');
  const data = materializeChart(pie, normalizeDatasets(model, request), {}, {});
  assert.equal(data.series[0].points[0].labelPosition, 'Outside');
});

test('an unknown chart custom property remains fail-closed', () => {
  const analysis = analyzeRdl(rdl('InventedProperty'));
  assert.equal(analysis.compatible, false);
  assert.deepEqual(analysis.blockingErrors, [{ code: 'UNSUPPORTED_FEATURE', feature: 'ChartProperty:InventedProperty' }]);
});

test('documented PointWidth, disabled auto-fit, and PieLineColor render to a valid selectable PDF', async () => {
  const result = await renderPdf(parseRdl(rdl()), request, config);
  assert.equal(result.buffer.subarray(0, 4).toString(), '%PDF');
  assert.equal(result.pageCount, 1);
  assert.ok(result.buffer.length > 2_000);
});

test('fixed charts with the same RDL Top render side by side on one PDF page', async () => {
  const result = await renderPdf(parseRdl(sideBySideRdl), request, config);
  assert.equal(result.pageCount, 1);
  assert.ok(result.buffer.length > 2_000);
});

test('chart title Position anchors its auto-sized box independently of TextAlign in every rendered format', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-chart-title-position-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const centeredModel = parseRdl(chartTitlePositionRdl());
  const centeredChart = centeredModel.body.items[0];
  assert.equal(centeredChart.title.position, 'TopCenter', 'an omitted Position uses the RDL TopCenter default');
  assert.equal(centeredChart.title.style.textAlign, 'Left', 'TextAlign remains an independent text-box style');

  const inkCenter = (buffer, { left, right, top, bottom }) => {
    const png = PNG.sync.read(buffer);
    let minimum = Infinity;
    let maximum = -Infinity;
    for (let y = Math.max(0, Math.floor(top)); y < Math.min(png.height, Math.ceil(bottom)); y += 1) {
      for (let x = Math.max(0, Math.floor(left)); x < Math.min(png.width, Math.ceil(right)); x += 1) {
        const pixel = (y * png.width + x) * 4;
        if (png.data[pixel] < 100 && png.data[pixel + 1] < 100 && png.data[pixel + 2] < 100) {
          minimum = Math.min(minimum, x);
          maximum = Math.max(maximum, x);
        }
      }
    }
    assert.ok(Number.isFinite(minimum) && Number.isFinite(maximum), 'the chart-title probe must render as dark text');
    return (minimum + maximum) / 2;
  };
  const mediaPng = async (buffer, pattern) => {
    const zip = await JSZip.loadAsync(buffer);
    const name = Object.keys(zip.files).find((entry) => pattern.test(entry));
    assert.ok(name);
    return zip.file(name).async('nodebuffer');
  };
  const chartRegion = (model, dpi) => {
    const chartItem = model.body.items[0];
    const scale = dpi / 72;
    const left = (model.page.marginLeft + chartItem.left) * scale;
    const top = (model.page.marginTop + chartItem.top) * scale;
    return {
      left,
      right: left + chartItem.width * scale,
      top: top + 4 * scale,
      bottom: top + 28 * scale,
      expected: left + chartItem.width * scale / 2,
    };
  };

  const centeredPdf = await renderPdf(centeredModel, request, config);
  const [centeredPage] = await rasterizePdf(centeredPdf.buffer, config, tempDir, 'centered-title-pdf', {
    dpi: 150, singleFile: true,
  });
  const pdfRegion = chartRegion(centeredModel, 150);
  assert.ok(Math.abs(inkCenter(centeredPage.data, pdfRegion) - pdfRegion.expected) <= 3);

  for (const [position, textAlign, expectedSide] of [
    ['TopLeft', 'Center', 'left'],
    ['TopRight', 'Left', 'right'],
  ]) {
    const model = parseRdl(chartTitlePositionRdl(position, textAlign));
    const rendered = await renderPdf(model, request, config);
    const [page] = await rasterizePdf(rendered.buffer, config, tempDir, `title-${expectedSide}`, {
      dpi: 150, singleFile: true,
    });
    const region = chartRegion(model, 150);
    const center = inkCenter(page.data, region);
    assert.ok(expectedSide === 'left' ? center < region.expected - 100 : center > region.expected + 100,
      `${position} must anchor the title box on the ${expectedSide}, regardless of TextAlign=${textAlign}`);
  }

  const editable = await renderEditableDocx(centeredModel, request, config, tempDir);
  const editableChart = await mediaPng(editable.buffer, /^word\/media\/.+\.png$/);
  const editablePng = PNG.sync.read(editableChart);
  const embeddedRegion = { left: 0, right: editablePng.width, top: 8, bottom: 60 };
  assert.ok(Math.abs(inkCenter(editableChart, embeddedRegion) - editablePng.width / 2) <= 3);

  const excel = await renderExcel(centeredModel, { ...request, excelLayoutMode: 'REPORT' }, config, tempDir);
  const excelChart = await mediaPng(excel.buffer, /^xl\/media\/.+\.png$/);
  const excelPng = PNG.sync.read(excelChart);
  assert.ok(Math.abs(inkCenter(excelChart, {
    left: 0, right: excelPng.width, top: 8, bottom: 60,
  }) - excelPng.width / 2) <= 3);

  const visual = await renderVisualDocx(centeredModel, request, config, tempDir);
  const visualPage = await mediaPng(visual.buffer, /^word\/media\/.+\.png$/);
  const visualRegion = chartRegion(centeredModel, 300);
  assert.ok(Math.abs(inkCenter(visualPage, visualRegion) - visualRegion.expected) <= 4);
});

test('normalizes exploded doughnut, custom palette, chart rectangles, title, and legend semantics', () => {
  const model = parseRdl(styledDoughnutRdl());
  const chart = model.body.items[0];
  assert.equal(chart.chartType, 'doughnut');
  assert.equal(chart.exploded, true);
  assert.equal(chart.palette, 'Custom');
  assert.deepEqual(chart.customPaletteColors, ['#01b8aa', '#374649']);
  assert.equal(chart.legend.position, 'RightCenter');
  assert.equal(chart.legend.layout, 'Column');
  assert.equal(chart.legend.style.backgroundColor, '#ffffff');
  assert.equal(chart.title.position, 'TopLeft');
  assert.equal(chart.title.style.textAlign, 'Left');
  assert.equal(chart.chartArea.style.backgroundColor, '#f2f2f2');
  assert.equal(chart.style.borders.top.style, 'Solid');

  const data = materializeChart(chart, normalizeDatasets(model, request), {}, {});
  assert.deepEqual(data.series[0].points.map((point) => point.color), ['#01b8aa', '#374649']);

  const plain = parseRdl(styledDoughnutRdl('Doughnut')).body.items[0];
  assert.equal(plain.chartType, 'doughnut');
  assert.equal(plain.exploded, false);

  const rotated = parseRdl(styledDoughnutRdl('Doughnut', 'RightCenter', '270')).body.items[0];
  assert.equal(rotated.seriesDefs[0].customProperties.PieStartAngle, '270');
});

test('exploded doughnut changes slice geometry while preserving declared palette and styled rectangles', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-exploded-doughnut-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const exploded = await renderPdf(parseRdl(styledDoughnutRdl()), request, config);
  const plain = await renderPdf(parseRdl(styledDoughnutRdl('Doughnut')), request, config);
  const [explodedPage] = await rasterizePdf(exploded.buffer, config, tempDir, 'exploded', { dpi: 72, singleFile: true });
  const [plainPage] = await rasterizePdf(plain.buffer, config, tempDir, 'plain', { dpi: 72, singleFile: true });
  const explodedPng = PNG.sync.read(explodedPage.data);
  const plainPng = PNG.sync.read(plainPage.data);
  assert.equal(explodedPng.width, plainPng.width);
  assert.equal(explodedPng.height, plainPng.height);

  let changed = 0;
  let pacificTeal = 0;
  let pacificSlate = 0;
  let chartAreaGray = 0;
  for (let offset = 0; offset < explodedPng.data.length; offset += 4) {
    const [r, g, b] = explodedPng.data.subarray(offset, offset + 3);
    if (r === 1 && g === 184 && b === 170) pacificTeal += 1;
    if (r === 55 && g === 70 && b === 73) pacificSlate += 1;
    if (r === 242 && g === 242 && b === 242) chartAreaGray += 1;
    if (explodedPng.data[offset] !== plainPng.data[offset]
      || explodedPng.data[offset + 1] !== plainPng.data[offset + 1]
      || explodedPng.data[offset + 2] !== plainPng.data[offset + 2]) changed += 1;
  }
  assert.ok(changed > 500, `expected exploded slice offsets to alter the raster, changed=${changed}`);
  assert.ok(pacificTeal > 500, `expected first custom palette colour, pixels=${pacificTeal}`);
  assert.ok(pacificSlate > 500, `expected second custom palette colour, pixels=${pacificSlate}`);
  assert.ok(chartAreaGray > 500, `expected styled chart-area rectangle, pixels=${chartAreaGray}`);
});

test('pie start angle defaults to SSRS zero degrees and honors the documented custom property', async (context) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-pie-start-angle-'));
  context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  const defaultAngle = await renderPdf(parseRdl(styledDoughnutRdl('Doughnut')), request, config);
  const topAngle = await renderPdf(parseRdl(styledDoughnutRdl('Doughnut', 'RightCenter', '270')), request, config);
  const [defaultPage] = await rasterizePdf(defaultAngle.buffer, config, tempDir, 'default-angle', { dpi: 72, singleFile: true });
  const [topPage] = await rasterizePdf(topAngle.buffer, config, tempDir, 'top-angle', { dpi: 72, singleFile: true });
  const defaultPng = PNG.sync.read(defaultPage.data);
  const topPng = PNG.sync.read(topPage.data);
  let changed = 0;
  for (let offset = 0; offset < defaultPng.data.length; offset += 4) {
    if (defaultPng.data[offset] !== topPng.data[offset]
      || defaultPng.data[offset + 1] !== topPng.data[offset + 1]
      || defaultPng.data[offset + 2] !== topPng.data[offset + 2]) changed += 1;
  }
  assert.ok(changed > 500, `expected PieStartAngle to rotate the doughnut, changed=${changed}`);
});

test('all documented legend positions render without changing the page contract', async () => {
  const positions = [
    'RightTop', 'TopLeft', 'TopCenter', 'TopRight', 'LeftTop', 'LeftCenter',
    'LeftBottom', 'RightCenter', 'RightBottom', 'BottomRight', 'BottomCenter', 'BottomLeft',
  ];
  for (const position of positions) {
    const result = await renderPdf(parseRdl(styledDoughnutRdl('Doughnut', position)), request, config);
    assert.equal(result.pageCount, 1, position);
    assert.equal(result.buffer.subarray(0, 4).toString(), '%PDF', position);
  }
});

test('standalone editable DOCX chart rendering owns and cleans its temporary workspace', async (context) => {
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'rdl-chart-owned-workspace-'));
  context.after(() => fs.rm(tempRoot, { recursive: true, force: true }));
  const ownedWorkspaceConfig = loadConfig({
    ...process.env,
    RDL_STRICT_FONTS: 'false',
    RDL_TEMP_ROOT: tempRoot,
  });
  const result = await renderEditableDocx(parseRdl(rdl()), request, ownedWorkspaceConfig);
  assert.equal(result.buffer.subarray(0, 2).toString(), 'PK');
  assert.deepEqual(await fs.readdir(tempRoot), []);
});
