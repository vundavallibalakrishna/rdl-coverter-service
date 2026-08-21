import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import { analyzeRdl, dropCoveredPlaceholders, parseRdl } from '../src/rdl/parser.js';
import { materializeTablixRows } from '../src/rdl/validation.js';
import { computeCellPlacements } from '../src/render/tableGrid.js';
import { rdl2016SchemaCatalogue } from '../src/rdl/capabilities.js';
import { evaluateExpression } from '../src/rdl/expression.js';
import { prepareTablixData, resolveParameterLabels, validateRenderInput } from '../src/rdl/validation.js';
import { toPoints } from '../src/units.js';
import { MISSING_SAMPLES, hasSamples, samplePath } from '../scripts/lib/samples.js';

const COMBINED_ASSURANCE = 'Combined Assurance Reports Excel.rdl';
import { htmlToPlainText } from '../src/rdl/text.js';

const fixture = fs.readFileSync(new URL('./fixtures/basic.rdl', import.meta.url));

test('analyzes datasets without returning query text', () => {
  const analysis = analyzeRdl(fixture);
  assert.equal(analysis.namespace.endsWith('/2016/01/reportdefinition'), true);
  assert.equal(analysis.compatible, true);
  assert.equal(analysis.capabilities.schema.defaultStatus, 'REJECTED');
  assert.equal(analysis.capabilities.summary.REJECTED, 0);
  assert.deepEqual(analysis.capabilities.expressions.detected, [{ name: 'Format', status: 'SUPPORTED' }]);
  assert.deepEqual(analysis.datasets.map((item) => [item.name, item.requiredForRendering, item.parameterOnly]), [
    ['Sales', true, false],
    ['Choices', false, true],
  ]);
  assert.equal(JSON.stringify(analysis).includes('select secret'), false);
});

test('normalizes supported 2008, 2010, and 2016 namespaces into the same page model', () => {
  for (const version of ['2008', '2010', '2016']) {
    const adapted = fixture.toString().replace('/2016/01/reportdefinition', `/${version}/01/reportdefinition`);
    const parsed = parseRdl(adapted);
    assert.equal(parsed.page.width, 612);
    assert.equal(parsed.body.width, 504);
    assert.equal(parsed.body.items.find((item) => item.type === 'Tablix').rowMembers[1].group.name, 'Details');
  }
});

test('rejects DTD and external entity declarations', () => {
  const malicious = fixture.toString().replace('<Report ', '<!DOCTYPE Report [<!ENTITY xxe SYSTEM "file:///etc/passwd">]><Report ');
  assert.throws(() => parseRdl(malicious), (error) => error.code === 'RDL_INVALID');
});

test('rejects excessive XML depth and unknown required namespaces', () => {
  assert.throws(() => parseRdl(fixture, { maxXmlDepth: 2 }), (error) => error.code === 'RDL_INVALID');
  assert.throws(() => parseRdl(fixture, { maxXmlNodes: 2 }), (error) => error.code === 'RDL_INVALID');
  const unknown = fixture.toString().replace('<Report ', '<Report MustUnderstand="bad" xmlns:bad="urn:unknown" ');
  assert.throws(() => parseRdl(unknown), (error) => error.code === 'RDL_INVALID');
});

test('reports unsupported expressions and report items as blocking compatibility errors', () => {
  // Aggregate() is deliberately kept fail-closed (its server-computed scope can't be reconstructed from
  // caller rows), so it stays a stable example of an unsupported expression function.
  const expression = fixture.toString().replace('=Parameters!Title.Value', '=Aggregate(Fields!Name.Value)');
  const expressionAnalysis = analyzeRdl(expression);
  assert.equal(expressionAnalysis.compatible, false);
  assert.equal(expressionAnalysis.blockingErrors.some(({ feature }) => feature === 'ExpressionFunction:Aggregate'), true);

  // The common chart types are supported, but an unsupported chart type (e.g. Range) must still fail closed.
  const rangeChart = '<Chart Name="UnsupportedChart"><ChartData><ChartSeriesCollection><ChartSeries Name="S"><ChartDataPoints><ChartDataPoint><ChartDataPointValues><Y>=1</Y></ChartDataPointValues></ChartDataPoint></ChartDataPoints><Type>Range</Type></ChartSeries></ChartSeriesCollection></ChartData></Chart>';
  const chart = fixture.toString().replace('<Textbox Name="TitleBox">', `${rangeChart}<Textbox Name="TitleBox">`);
  assert.equal(analyzeRdl(chart).blockingErrors.some(({ feature }) => feature === 'ChartType:Range'), true);
});

test('classifies every encountered XML path and fails closed for unconsumed elements and attributes', () => {
  const metadata = fixture.toString().replace('<Description>Basic Report</Description>', '<Description>Basic Report</Description><Author>Test</Author>');
  assert.equal(analyzeRdl(metadata).compatible, true);

  const element = fixture.toString().replace('<CanGrow>true</CanGrow>', '<CanGrow>true</CanGrow><UnsupportedTextboxProp>x</UnsupportedTextboxProp>');
  const elementAnalysis = analyzeRdl(element);
  assert.equal(elementAnalysis.compatible, false);
  assert.equal(elementAnalysis.capabilities.rejected.some(({ path }) => path.endsWith(".Textbox.UnsupportedTextboxProp")), true);

  const attribute = fixture.toString().replace('<Textbox Name="TitleBox">', '<Textbox Name="TitleBox" UnsupportedAttribute="true">');
  const attributeAnalysis = analyzeRdl(attribute);
  assert.equal(attributeAnalysis.compatible, false);
  assert.equal(attributeAnalysis.capabilities.rejected.some(({ path }) => path.endsWith('.Textbox.@UnsupportedAttribute')), true);

  const tablixHeader = fixture.toString().replace('<Group Name="Details"/>', '<Group Name="Details"/><TablixHeader><Size>1in</Size></TablixHeader>');
  assert.equal(analyzeRdl(tablixHeader).compatible, true);
  // Column-hierarchy TablixHeader now backs dynamic column groups (matrix) and is supported, not blocking.
  const columnHeader = fixture.toString().replace('<TablixColumnHierarchy><TablixMembers><TablixMember/>', '<TablixColumnHierarchy><TablixMembers><TablixMember><TablixHeader><Size>1in</Size></TablixHeader></TablixMember>');
  assert.equal(analyzeRdl(columnHeader).compatible, true);

  const pageName = fixture.toString().replace('<Tablix Name="SalesTable">', '<Tablix Name="SalesTable"><PageName>Sales detail</PageName>');
  assert.equal(analyzeRdl(pageName).compatible, true);
  assert.equal(parseRdl(pageName).body.items.find((item) => item.type === 'Tablix').pageName, 'Sales detail');
});

test('catalogues every element and attribute declared by the published 2016 schema', () => {
  const catalogue = rdl2016SchemaCatalogue();
  assert.equal(catalogue.entries.filter(({ kind }) => kind === 'ELEMENT').length, 691);
  assert.equal(catalogue.entries.filter(({ kind }) => kind === 'ATTRIBUTE').length, 4);
  assert.equal(catalogue.entries.length, 695);
  assert.equal(Object.values(catalogue.summary).reduce((sum, count) => sum + count, 0), 695);
  assert.equal(catalogue.entries.every(({ status }) => ['SUPPORTED', 'METADATA_ONLY', 'REJECTED'].includes(status)), true);
  assert.equal(catalogue.entries.find(({ name, kind }) => name === 'Textbox' && kind === 'ELEMENT').status, 'SUPPORTED');
  assert.equal(catalogue.entries.find(({ name, kind }) => name === 'DataSourceReference' && kind === 'ELEMENT').status, 'METADATA_ONLY');
  assert.equal(catalogue.entries.find(({ name, kind }) => name === 'TablixHeader' && kind === 'ELEMENT').status, 'SUPPORTED');
});

test('converts RDL units without rounding geometry', () => {
  assert.equal(toPoints('1in'), 72);
  assert.equal(toPoints('2.54cm'), 72);
  assert.equal(toPoints('25.4mm'), 72);
});

test('normalizes safe HTML text and client formatting properties without executing markup', () => {
  assert.equal(htmlToPlainText('<p>Hello <strong>world</strong></p><ul><li>One</li><li>Two</li></ul><script>alert(1)</script>'), 'Hello world\n• One\n• Two');
  const extended = fixture.toString()
    .replace('<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">', '<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition" xmlns:df="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition/defaultfontfamily" MustUnderstand="df"><df:DefaultFontFamily>Segoe UI</df:DefaultFontFamily>')
    .replace('<CanGrow>true</CanGrow>', '<CanGrow>true</CanGrow><KeepTogether>true</KeepTogether><ZIndex>4</ZIndex><HideDuplicates>Sales</HideDuplicates>')
    .replace('<Value>=Parameters!Title.Value</Value>', '<Value>=Parameters!Title.Value</Value><MarkupType>HTML</MarkupType>')
    .replace('<Style><FontFamily>Arial</FontFamily><FontSize>14pt</FontSize><FontWeight>Bold</FontWeight></Style>', '<Style><FontSize>14pt</FontSize><FontWeight>Bold</FontWeight><TopBorder><Style>Solid</Style><Width>2pt</Width></TopBorder></Style>');
  const parsed = parseRdl(extended);
  const title = parsed.body.items.find((item) => item.name === 'TitleBox');
  assert.equal(parsed.unsupported.length, 0);
  assert.equal(parsed.defaultFontFamily, 'Segoe UI');
  assert.equal(title.keepTogether, true);
  assert.equal(title.zIndex, 4);
  assert.equal(title.hideDuplicates, 'Sales');
  assert.equal(title.paragraphs[0][0].markupType, 'HTML');
  assert.equal(title.style.borders.top.style, 'Solid');
  assert.equal(title.style.borders.top.width, 2);
});

test('evaluates certified SSRS expressions without eval', () => {
  const context = { fields: { Amount: 1234.5 }, parameters: { Show: true }, globals: { PageNumber: 2 }, dataset: [{ Amount: 5 }, { Amount: 7 }] };
  assert.equal(evaluateExpression('=IIF(Parameters!Show.Value, Format(Fields!Amount.Value, "N2"), "hidden")', context), '1,234.50');
  assert.equal(evaluateExpression('=Globals!PageNumber < 4', context), true);
  assert.equal(evaluateExpression('=Sum(Fields!Amount.Value)', context), 12);
  assert.equal(evaluateExpression('=Format("2026-07-15T00:00:00Z", "y")', context), 'July 2026');
});

test('requires rendering datasets and exact data fields but not parameter lookup datasets', () => {
  const model = parseRdl(fixture);
  const request = { parameters: { Title: 'Sales', Choice: 'A' }, datasets: { Sales: [{ Name: 'North', Amount: 10 }] } };
  assert.equal(validateRenderInput(model, request, { maxRows: 100 }).totalRows, 1);
  assert.throws(() => validateRenderInput(model, { ...request, datasets: {} }, { maxRows: 100 }), (error) => error.code === 'DATASET_MISSING');
  assert.throws(() => validateRenderInput(model, { ...request, datasets: { Sales: [{ Name: 'North' }] } }, { maxRows: 100 }), (error) => error.code === 'FIELD_MISSING');
  const defaults = validateRenderInput(model, { parameters: { Choice: 'A' }, datasets: request.datasets }, { maxRows: 100 });
  assert.equal(defaults.parameters.Title, 'Sales');
});

test('resolves Parameters!Name.Label from static and dataset-backed valid values', () => {
  const labels = resolveParameterLabels([
    { name: 'Division', multiValue: true, lookupDataset: 'Divisions', lookupValueField: 'Id', lookupLabelField: 'Name', staticValidValues: [] },
    { name: 'Mode', multiValue: false, staticValidValues: [{ value: 'S', label: 'Summary' }] },
  ], { Division: ['403', '326'], Mode: 'S' }, {
    Divisions: [{ Id: 403, Name: 'Company Secretariat' }, { Id: 326, Name: 'CoSec' }],
  });
  const parameters = { Division: ['403', '326'], Mode: 'S' };
  Object.defineProperty(parameters, '__rdlParameterLabels', { value: labels, enumerable: false });
  assert.deepEqual(evaluateExpression('=Parameters!Division.Label', { parameters }), ['Company Secretariat', 'CoSec']);
  assert.equal(evaluateExpression('=Parameters!Mode.Label', { parameters }), 'Summary');
  assert.deepEqual(evaluateExpression('=Parameters!Division.Value', { parameters }), ['403', '326']);
});

test('applies tablix filters and stable multi-column sorting to exact DataField rows', () => {
  const tablix = parseRdl(fixture).body.items.find((item) => item.type === 'Tablix');
  const configured = {
    ...tablix,
    filters: [{ expression: '=Fields!Amount.Value', operator: 'GreaterThan', values: ['10'] }],
    sortExpressions: [
      { value: '=Fields!Amount.Value', direction: 'Descending' },
      { value: '=Fields!Name.Value', direction: 'Ascending' },
    ],
  };
  const rows = prepareTablixData(configured, [
    { Name: 'A', Amount: 11 }, { Name: 'B', Amount: 9 }, { Name: 'C', Amount: 30 }, { Name: 'D', Amount: 11 },
  ], {}, {});
  assert.deepEqual(rows.map(({ Name, Amount }) => [Name, Amount]), [['C', 30], ['A', 11], ['D', 11]]);
});

test('RunningValue accumulates a running aggregate through the current row and keeps other data functions fail-closed', () => {
  const dataset = [{ id: 'a' }, { id: 'b' }, { id: 'a' }, { id: 'c' }];
  const running = dataset.map((row) => evaluateExpression('=RunningValue(Fields!id.Value, countdistinct, nothing)', { fields: row, dataset }));
  assert.deepEqual(running, [1, 2, 2, 3]);
  const runningSum = [{ n: 10 }, { n: 5 }, { n: 2 }].map((row, index, rows) => evaluateExpression('=RunningValue(Fields!n.Value, Sum, Nothing)', { fields: rows[index], dataset: rows }));
  assert.deepEqual(runningSum, [10, 15, 17]);
  // Functions outside the supported catalogue must still fail closed.
  assert.throws(() => evaluateExpression('=Aggregate(Fields!id.Value)', { dataset }), /Unsupported RDL expression function/);
  assert.throws(() => evaluateExpression('=RunningValue(Fields!id.Value, Median, Nothing)', { fields: dataset[0], dataset }), /Unsupported RunningValue aggregate/);
});

test('RunningValue and RowNumber treat Nothing as the outermost processed scope across nested groups', () => {
  const rows = [
    { category: 'A', id: 'x' },
    { category: 'A', id: 'y' },
    { category: 'B', id: 'y' },
    { category: 'B', id: 'z' },
  ];
  const contexts = rows.map((fields) => {
    const categoryRows = rows.filter((row) => row.category === fields.category);
    return {
      fields,
      dataset: [fields],
      outermostDataset: rows,
      datasets: { Sales: rows },
      scopes: {
        Sales: rows,
        RiskTable: rows,
        Category: categoryRows,
        Details: [fields],
      },
      tablixDatasetName: 'Sales',
      tablixName: 'RiskTable',
    };
  });

  assert.deepEqual(
    contexts.map((context) => evaluateExpression('=RunningValue(Fields!id.Value, CountDistinct, Nothing)', context)),
    [1, 2, 2, 3],
  );
  assert.deepEqual(
    contexts.map((context) => evaluateExpression('=RunningValue(Fields!id.Value, CountDistinct, "Category")', context)),
    [1, 2, 1, 2],
  );
  assert.deepEqual(
    contexts.map((context) => evaluateExpression('=RunningValue(Fields!id.Value, CountDistinct, "RiskTable")', context)),
    [1, 2, 2, 3],
  );
  assert.deepEqual(
    contexts.map((context) => evaluateExpression('=RowNumber(Nothing)', context)),
    [1, 2, 3, 4],
  );
});

test('tablix materialization keeps RunningValue Nothing outside one-row nested group scopes after sorting', () => {
  const model = parseRdl(fixture);
  const tablix = model.body.items.find((item) => item.type === 'Tablix');
  const detailMember = tablix.rowMembers[1];
  detailMember.group.expressions = ['=Fields!Amount.Value'];
  const numberTextbox = tablix.rows[1].cells[0].items[0];
  numberTextbox.value = '=RunningValue(Fields!Amount.Value, CountDistinct, Nothing)';
  numberTextbox.paragraphs[0][0].value = numberTextbox.value;
  tablix.sortExpressions = [{ value: '=Fields!Amount.Value', direction: 'Ascending' }];
  const source = [
    { Name: 'thirty', Amount: 30 },
    { Name: 'ten', Amount: 10 },
    { Name: 'twenty-a', Amount: 20 },
    { Name: 'twenty-b', Amount: 20 },
  ];

  // A leaf member that IS a group emits ONE row per group instance (SSRS), so the three distinct
  // Amounts render as three rows in ascending sort order. RunningValue(..., Nothing) must still count
  // across the whole region rather than restart inside each one-row nested group scope.
  const materialized = materializeTablixRows(tablix, source, {}, {}, { Sales: source });
  assert.deepEqual(
    materialized.filter((row) => !row.isHeader).map((row) => row.cells[0].values[0]),
    ['1', '2', '3'],
  );
});

test('evaluates the newly supported expression functions', () => {
  const dataset = [{ n: 2 }, { n: 4 }, { n: 4 }, { n: 6 }];
  const at = (index) => ({ fields: dataset[index], dataset, datasets: {} });
  assert.equal(evaluateExpression('=IsNothing(Nothing)', {}), true);
  assert.equal(evaluateExpression('=IsNothing("x")', {}), false);
  assert.equal(evaluateExpression('=Switch(1=2,"a",1=1,"b")', {}), 'b');
  assert.equal(evaluateExpression('=Choose(2,"a","b","c")', {}), 'b');
  assert.equal(evaluateExpression('=Last(Fields!n.Value)', at(0)), 6);
  assert.equal(evaluateExpression('=CountRows(Nothing)', at(0)), 4);
  assert.equal(evaluateExpression('=RowNumber(Nothing)', at(2)), 3);
  assert.equal(evaluateExpression('=Previous(Fields!n.Value)', at(2)), 4);
  assert.equal(evaluateExpression('=Previous(Fields!n.Value)', at(0)), null);
  assert.equal(evaluateExpression('=VarP(Fields!n.Value)', at(0)), 2);
  assert.equal(evaluateExpression('=StDevP(Fields!n.Value)', at(0)), Math.sqrt(2));
  assert.equal(Math.round(evaluateExpression('=Var(Fields!n.Value)', at(0)) * 1000), 2667);
  assert.equal(evaluateExpression('=Join(Parameters!p.Value, "-")', { parameters: { p: ['a', 'b', 'c'] } }), 'a-b-c');
});

test('maps chart Type and Subtype to the supported chart kinds and stack mode', () => {
  const chartXml = (type, subtype) => `<Chart Name="C"><ChartData><ChartSeriesCollection><ChartSeries Name="S"><ChartDataPoints><ChartDataPoint><ChartDataPointValues><Y>=1</Y></ChartDataPointValues></ChartDataPoint></ChartDataPoints><Type>${type}</Type>${subtype ? `<Subtype>${subtype}</Subtype>` : ''}</ChartSeries></ChartSeriesCollection></ChartData><DataSetName>Sales</DataSetName></Chart>`;
  const chartOf = (type, subtype) => parseRdl(fixture.toString().replace('<Textbox Name="TitleBox">', `${chartXml(type, subtype)}<Textbox Name="TitleBox">`)).body.items.find((item) => item.type === 'Chart');
  const kindOf = (type, subtype) => chartOf(type, subtype).chartType;
  assert.equal(kindOf('Shape', 'Doughnut'), 'doughnut');
  assert.equal(kindOf('Shape', 'ExplodedDoughnut'), 'doughnut');
  assert.equal(kindOf('Shape', 'Pie'), 'pie');
  assert.equal(kindOf('Shape', ''), 'pie');
  assert.equal(kindOf('Line', 'Smooth'), 'line');
  assert.equal(kindOf('Area', ''), 'area');
  assert.equal(kindOf('Scatter', ''), 'scatter');
  assert.equal(kindOf('Point', ''), 'scatter');
  assert.equal(kindOf('Bar', ''), 'bar');
  assert.equal(kindOf('Column', ''), 'column');
  // Stack mode from subtype (bar/column/area only).
  assert.equal(chartOf('Column', 'Stacked').stacked, 'stacked');
  assert.equal(chartOf('Column', 'PercentStacked').stacked, 'percent');
  assert.equal(chartOf('Area', 'Stacked').stacked, 'stacked');
  assert.equal(chartOf('Column', 'Plain').stacked, 'none');
});

test('evaluates cross-dataset lookups and date functions', () => {
  const datasets = { Prices: [{ Sku: 'A', Price: 10 }, { Sku: 'B', Price: 20 }, { Sku: 'A', Price: 15 }] };
  const ctx = (fields) => ({ fields, datasets, dataset: [] });
  const lookup = '=Lookup(Fields!Sku.Value, Fields!Sku.Value, Fields!Price.Value, "Prices")';
  assert.equal(evaluateExpression(lookup, ctx({ Sku: 'A' })), 10); // first match
  assert.equal(evaluateExpression(lookup, ctx({ Sku: 'Z' })), null); // no match
  assert.deepEqual(evaluateExpression('=LookupSet(Fields!Sku.Value, Fields!Sku.Value, Fields!Price.Value, "Prices")', ctx({ Sku: 'A' })), [10, 15]);
  assert.deepEqual(evaluateExpression('=MultiLookup(Parameters!p.Value, Fields!Sku.Value, Fields!Price.Value, "Prices")', { parameters: { p: ['A', 'B'] }, datasets }), [10, 20]);

  assert.equal(evaluateExpression('=CDate("2026-03-15")', {}) instanceof Date, true);
  assert.equal(evaluateExpression('=DatePart("yyyy", "2026-03-15")', {}), 2026);
  assert.equal(evaluateExpression('=DatePart("q", "2026-03-15")', {}), 1);
  assert.equal(evaluateExpression('=Format(DateAdd("m", 2, "2026-03-15"), "yyyy-MM-dd")', {}), '2026-05-15');
  assert.equal(evaluateExpression('=DateDiff("d", "2026-03-15", "2026-03-25")', {}), 10);
  assert.equal(evaluateExpression('=DateDiff("m", "2026-01-31", "2026-03-01")', {}), 2);
});

test('interactive-only features are metadata-only and no longer block rendering', () => {
  const withAction = fixture.toString().replace(
    '<CanGrow>true</CanGrow>',
    '<CanGrow>true</CanGrow><ActionInfo><Actions><Action><Hyperlink>=Fields!Name.Value</Hyperlink></Action></Actions></ActionInfo>',
  );
  const analysis = analyzeRdl(withAction);
  assert.equal(analysis.compatible, true);
  assert.equal(analysis.capabilities.rejected.some(({ path }) => path.includes('ActionInfo')), false);
  // Genuinely unsupported report items must still fail closed.
  const withSubreport = fixture.toString().replace('<Textbox Name="TitleBox">', '<Subreport Name="Sub"/><Textbox Name="TitleBox">');
  assert.equal(analyzeRdl(withSubreport).compatible, false);
});

test('analysis exposes subreport dependencies without resolving or rendering them', () => {
  const subreport = `
    <Subreport Name="ChildReport">
      <ReportName>/Reports/Child Detail</ReportName>
      <Parameters>
        <Parameter Name="EntityId"><Value>=Fields!Id.Value</Value></Parameter>
        <Parameter Name="Mode"><Value>Summary</Value></Parameter>
      </Parameters>
      <KeepTogether>true</KeepTogether>
      <MergeTransactions>true</MergeTransactions>
      <OmitBorderOnPageBreak>true</OmitBorderOnPageBreak>
      <Top>0.5in</Top><Left>0in</Left><Width>7in</Width><Height>1in</Height>
    </Subreport>`;
  const analysis = analyzeRdl(fixture.toString().replace('<Textbox Name="TitleBox">', `${subreport}<Textbox Name="TitleBox">`));

  assert.equal(analysis.compatible, false);
  assert.equal(analysis.features.subreports, 1);
  assert.deepEqual(analysis.subreports, [{
    name: 'ChildReport',
    reportName: '/Reports/Child Detail',
    parameters: [
      { name: 'EntityId', value: '=Fields!Id.Value' },
      { name: 'Mode', value: 'Summary' },
    ],
    keepTogether: true,
    mergeTransactions: true,
    omitBorderOnPageBreak: true,
  }]);
  assert.equal(analysis.blockingErrors.some(({ feature }) => feature === 'Subreport'), true);
});

test('unused embedded code is metadata-only, allowlisted native mappings work, and unknown Code.* stays fail-closed', () => {
  const definition = `
    <Code>
      Public Function Paint(ByVal value As Integer) As String
        Return "#00ff00"
      End Function
    </Code>`;
  const unused = fixture.toString().replace('<ReportSections>', `${definition}<ReportSections>`);
  const unusedAnalysis = analyzeRdl(unused);
  assert.equal(unusedAnalysis.compatible, true);
  assert.equal(unusedAnalysis.capabilities.expressions.detected.some((entry) => entry.name === 'Code.*'), false);

  const invoked = unused.replace(
    '<BackgroundColor>#dddddd</BackgroundColor>',
    '<BackgroundColor>=Code.Paint(Fields!Amount.Value)</BackgroundColor>',
  );
  const invokedAnalysis = analyzeRdl(invoked);
  assert.equal(invokedAnalysis.compatible, false);
  assert.equal(invokedAnalysis.blockingErrors.some(({ feature }) => feature === 'CustomCode'), true);
  assert.equal(invokedAnalysis.capabilities.expressions.detected.some((entry) => entry.name === 'Code.*' && entry.status === 'REJECTED'), true);

  const calculateColor = unused.replace(
    '<BackgroundColor>#dddddd</BackgroundColor>',
    '<BackgroundColor>=Code.CalculateColor(CStr(Fields!Amount.Value), CStr(Parameters!Choice.Value))</BackgroundColor>',
  );
  const calculateColorAnalysis = analyzeRdl(calculateColor);
  assert.equal(calculateColorAnalysis.compatible, true);
  assert.ok(calculateColorAnalysis.capabilities.expressions.detected
    .some((entry) => entry.name === 'Code.CalculateColor' && entry.status === 'SUPPORTED'));
});

test('Code.CalculateColor uses the fixed service-owned 5x5 risk heat-map mapping', () => {
  const color = (y, x) => evaluateExpression(`=Code.CalculateColor(CStr(${y}), CStr(${x}))`);
  assert.equal(color(1, 1), 'Green');
  assert.equal(color(5, 1), '#ffff00');
  assert.equal(color(3, 2), '#ffff00');
  assert.equal(color(5, 2), '#FFA500');
  assert.equal(color(4, 3), '#FFA500');
  assert.equal(color(3, 4), '#FFA500');
  assert.equal(color(5, 4), '#ff0000');
  assert.equal(color(1, 5), '#ffff00');
  assert.equal(color(2, 5), '#FFA500');
  assert.equal(color(4, 5), '#ff0000');
  assert.equal(color(1, 0), '#ff0000');
  assert.throws(() => evaluateExpression('=Code.NotAllowlisted(1, 2)'), /Unsupported RDL expression/);
});

test('allowlisted maturity-level and distinct-array custom functions use fixed native contracts', () => {
  const level = (score) => evaluateExpression(`=Code.GetPercentLevel(${score})`);
  assert.equal(level(0), 'Level 1');
  assert.equal(level(0.2), 'Level 1');
  assert.equal(level(0.2001), 'Level 2');
  assert.equal(level(0.4), 'Level 2');
  assert.equal(level(0.6), 'Level 3');
  assert.equal(level(0.8), 'Level 4');
  assert.equal(level(0.8001), 'Level 5');
  assert.equal(evaluateExpression('=Code.GetPercentLevel(Nothing)'), 'Level 1');

  const context = {
    parameters: {},
    fields: {},
    datasets: { D: [
      { Key: 'All', Value: 'Alpha' },
      { Key: 'All', Value: 'Alpha' },
      { Key: 'All', Value: 'Beta' },
    ] },
  };
  assert.deepEqual(
    evaluateExpression('=Code.GetDistinct(LookupSet("All", Fields!Key.Value, Fields!Value.Value, "D"))', context),
    ['Alpha', 'Beta'],
  );
  assert.equal(evaluateExpression('=Code.GetDistinct(Nothing)', context), null);
});

test('analysis advertises only the allowlisted native GetPercentLevel and GetDistinct mappings', () => {
  const mapped = fixture.toString()
    .replace(
      '<Value>=Parameters!Title.Value</Value>',
      '<Value>=Join(Code.GetDistinct(LookupSet("All", "All", Fields!Value.Value, "Choices")), ", ") &amp; " - " &amp; Code.GetPercentLevel(0.6)</Value>',
    );
  const analysis = analyzeRdl(mapped);
  assert.equal(analysis.compatible, true);
  assert.ok(analysis.capabilities.expressions.detected
    .some((entry) => entry.name === 'Code.GetPercentLevel' && entry.status === 'SUPPORTED'));
  assert.ok(analysis.capabilities.expressions.detected
    .some((entry) => entry.name === 'Code.GetDistinct' && entry.status === 'SUPPORTED'));
});

test('an RDL using RunningValue passes the expression capability gate', () => {
  const withRunningValue = fixture.toString().replace(
    '<Value>=Parameters!Title.Value</Value>',
    '<Value>=RunningValue(Fields!Amount.Value, Sum, Nothing)</Value>',
  );
  const analysis = analyzeRdl(withRunningValue);
  assert.equal(analysis.compatible, true);
  assert.ok(analysis.capabilities.expressions.detected.some((entry) => entry.name === 'RunningValue' && entry.status === 'SUPPORTED'));
});

test('dropCoveredPlaceholders removes only the empty cells a preceding span covers', () => {
  const cells = [
    { colSpan: 2, rowSpan: 1, items: [{ type: 'Textbox' }] },
    { colSpan: 1, rowSpan: 1, items: [] }, // covered by the span above -> dropped
    { colSpan: 1, rowSpan: 1, items: [{ type: 'Textbox' }] },
    { colSpan: 1, rowSpan: 1, items: [] }, // a genuinely blank standalone cell -> kept
  ];
  const result = dropCoveredPlaceholders(cells);
  assert.equal(result.length, 3);
  assert.deepEqual(result.map((cell) => cell.colSpan), [2, 1, 1]);
  assert.deepEqual(result.map((cell) => cell.items.length), [1, 1, 0]);
});

test('removes a literal-hidden tablix body column from the physical grid', () => {
  const hiddenColumnRdl = `<?xml version="1.0"?>
    <Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
      <ReportSections><ReportSection><Body><ReportItems><Tablix Name="HiddenBodyColumn">
        <TablixBody><TablixColumns><TablixColumn><Width>0.03125in</Width></TablixColumn></TablixColumns>
          <TablixRows><TablixRow><Height>0.2in</Height><TablixCells><TablixCell><CellContents>
            <Textbox Name="HiddenCell"><Value>not visible</Value></Textbox>
          </CellContents></TablixCell></TablixCells></TablixRow></TablixRows>
        </TablixBody>
        <TablixColumnHierarchy><TablixMembers><TablixMember><Visibility><Hidden>true</Hidden></Visibility></TablixMember></TablixMembers></TablixColumnHierarchy>
        <TablixRowHierarchy><TablixMembers><TablixMember/></TablixMembers></TablixRowHierarchy>
        <Width>0.03125in</Width><Height>0.2in</Height>
      </Tablix></ReportItems><Height>1in</Height></Body>
      <Width>3in</Width><Page><PageHeight>11in</PageHeight><PageWidth>8.5in</PageWidth>
        <LeftMargin>0.5in</LeftMargin><RightMargin>0.5in</RightMargin><TopMargin>0.5in</TopMargin><BottomMargin>0.5in</BottomMargin>
      </Page></ReportSection></ReportSections>
    </Report>`;
  const tablix = parseRdl(hiddenColumnRdl).body.items[0];
  assert.deepEqual(tablix.bodyColumns, []);
  assert.deepEqual(tablix.columns, []);
  assert.equal(tablix.rows[0].cells.length, 0);
});

test('the Combined Assurance Table 2 keeps every visible leaf header on a distinct grid column', {
  skip: hasSamples(COMBINED_ASSURANCE) ? false : MISSING_SAMPLES,
}, () => {
  const model = parseRdl(fs.readFileSync(samplePath(COMBINED_ASSURANCE)));
  assert.deepEqual(model.unsupported, []);
  const tablix = model.body.items.find((item) => item.name === 'Tablix2');
  assert.equal(tablix.rowHeaderColumns.length, 7);
  assert.equal(tablix.bodyColumns.length, 13);
  assert.equal(tablix.columns.length, 20);

  // Every declared static column-header row tiles the complete 13-column body (placeholders removed).
  for (const row of tablix.rows.slice(0, 2)) {
    assert.equal(row.cells.reduce((sum, cell) => sum + (cell.colSpan || 1), 0), 13);
  }

  const rows = materializeTablixRows(tablix, [], {}, {}, {});
  const placements = computeCellPlacements(rows, tablix.columns.length);
  const labelRow = rows.find((row) => row.cells.some((cell) => (cell.values || []).join('').includes('4th Line')));
  const columnOf = (needle) => {
    const index = labelRow.cells.findIndex((cell) => (cell.values || []).join('').includes(needle));
    return placements[rows.indexOf(labelRow)][index];
  };
  // The final assurance/activity/action headers must remain on their own columns rather than collapsing
  // onto the last grid position.
  assert.equal(columnOf('4th Line'), 13);
  assert.equal(columnOf('Combined Assurance Level'), 15);
  assert.equal(columnOf('Comment'), 19);
});
