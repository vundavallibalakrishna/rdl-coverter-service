import assert from 'node:assert/strict';
import test from 'node:test';
import { loadConfig } from '../src/config.js';
import { analyzeRdl, parseRdl } from '../src/rdl/parser.js';
import { normalizeDatasets } from '../src/render/common.js';
import { materializeChart } from '../src/render/chartData.js';
import { renderPdf } from '../src/render/pdf.js';

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
