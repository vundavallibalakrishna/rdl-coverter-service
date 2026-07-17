import { RDL_2016_SCHEMA_ATTRIBUTE_NAMES, RDL_2016_SCHEMA_ELEMENT_NAMES } from './rdl2016SchemaNames.js';
import { FUNCTION_NAMES } from './functions/index.js';

export const CapabilityStatus = Object.freeze({
  SUPPORTED: 'SUPPORTED',
  METADATA_ONLY: 'METADATA_ONLY',
  REJECTED: 'REJECTED',
});

const SCHEMA_SOURCES = Object.freeze({
  'http://schemas.microsoft.com/sqlserver/reporting/2008/01/reportdefinition': 'https://schemas.microsoft.com/sqlserver/reporting/2008/01/reportdefinition/ReportDefinition.xsd',
  'http://schemas.microsoft.com/sqlserver/reporting/2010/01/reportdefinition': 'https://schemas.microsoft.com/sqlserver/reporting/2010/01/reportdefinition/ReportDefinition.xsd',
  'http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition': 'https://learn.microsoft.com/en-us/openspecs/sql_server_protocols/ms-rdl/52ce3983-2bfc-4e72-9359-42aaf5fe4509',
});

const SUPPORTED_ELEMENTS = new Set([
  'Report', 'ReportSections', 'ReportSection', 'Body', 'Page', 'PageHeader', 'PageFooter',
  'PageWidth', 'PageHeight', 'TopMargin', 'RightMargin', 'BottomMargin', 'LeftMargin',
  'PrintOnFirstPage', 'PrintOnLastPage', 'Height', 'Width', 'Top', 'Left', 'ReportItems',
  'Textbox', 'Paragraphs', 'Paragraph', 'TextRuns', 'TextRun', 'Value', 'CanGrow', 'CanShrink',
  'Style', 'Color', 'BackgroundColor', 'FontFamily', 'FontSize', 'FontWeight', 'FontStyle',
  'TextAlign', 'VerticalAlign', 'TextDecoration', 'Format', 'PaddingLeft', 'PaddingRight', 'PaddingTop', 'PaddingBottom',
  'Border', 'TopBorder', 'RightBorder', 'BottomBorder', 'LeftBorder',
  'Visibility', 'Hidden', 'PageBreak', 'BreakLocation', 'Disabled',
  'KeepTogether', 'ZIndex', 'HideDuplicates', 'MarkupType',
  'Tablix', 'TablixBody', 'TablixColumns', 'TablixColumn', 'TablixRows', 'TablixRow',
  'TablixCells', 'TablixCell', 'CellContents', 'ColSpan', 'RowSpan', 'DataSetName',
  'TablixRowHierarchy', 'TablixColumnHierarchy', 'TablixMembers', 'TablixMember', 'TablixHeader', 'Size',
  'TablixCorner', 'TablixCornerRows', 'TablixCornerRow', 'TablixCornerCell',
  'RepeatColumnHeaders', 'RepeatRowHeaders', 'RepeatOnNewPage', 'KeepWithGroup', 'FixedData',
  'HideIfNoRows', 'NoRowsMessage',
  // Group/Parent drives supported recursive (parent/child) row grouping, rendered expanded with
  // indentation. Group/Variables are consumed (resolved as Variables!Name.Value in the current scope).
  'Group', 'Parent', 'GroupExpressions', 'GroupExpression', 'SortExpressions', 'SortExpression', 'Direction',
  'Variables', 'Variable',
  'Filters', 'Filter', 'FilterExpression', 'Operator', 'FilterValues', 'FilterValue',
  'Rectangle', 'Line', 'Image', 'Source', 'Sizing', 'EmbeddedImages', 'EmbeddedImage', 'MIMEType', 'ImageData',
  'DataSets', 'DataSet', 'Fields', 'Field', 'DataField', 'TypeName',
  'ReportParameters', 'ReportParameter', 'DataType', 'Nullable', 'AllowBlank', 'MultiValue', 'Prompt',
  'DefaultValue', 'Values', 'ValidValues', 'DataSetReference', 'NumberOfColumns', 'ColumnSpacing',
  // Chart family — structural and consumed elements for the supported Bar/Column/Pie renderer.
  'Chart', 'ChartCategoryHierarchy', 'ChartSeriesHierarchy', 'ChartMembers', 'ChartMember',
  'ChartData', 'ChartSeriesCollection', 'ChartSeries', 'ChartDataPoints', 'ChartDataPoint',
  'ChartDataPointValues', 'Y', 'ChartDataLabel', 'UseValueAsLabel', 'Visible',
  'ChartAreas', 'ChartArea', 'ChartCategoryAxes', 'ChartValueAxes', 'ChartAxis',
  'ChartLegends', 'ChartLegend', 'Position', 'ChartTitles', 'ChartTitle', 'Caption',
  'Type', 'Subtype', 'Minimum', 'Maximum', 'Interval', 'ChartNoDataMessage',
]);

const METADATA_ELEMENTS = new Set([
  'Description', 'Author', 'AutoRefresh', 'Label', 'LabelField', 'ValueField',
  // Advisory grouping hints acknowledged but not acted on (the group renders its natural instances).
  'DomainScope', 'NaturalGroup',
  'DataElementName', 'DataElementOutput', 'DataElementStyle',
  'DataSource', 'DataSourceName', 'DataSourceReference', 'ConnectionProperties', 'ConnectString',
  'DataProvider', 'IntegratedSecurity', 'Transaction', 'Query', 'CommandText', 'CommandType',
  'QueryParameters', 'QueryParameter', 'ParameterName', 'ParameterValue', 'Timeout',
  // Custom code is acknowledged but NEVER executed (no eval/Function/VM). Code.* invocations and
  // unknown expression functions remain fail-closed in the expression evaluator.
  'Code',
  // Chart cosmetic/positioning elements we accept but do not act on (never block rendering).
  'AutoFitTextDisabled', 'BackgroundGradientType', 'CalloutLineColor', 'CategoryAxisName',
  'ChartAxisScaleBreak', 'ChartAxisTitle', 'ChartBorderSkin', 'ChartEmptyPoints', 'ChartLegendTitle',
  'ChartMajorGridLines', 'ChartMajorTickMarks', 'ChartMarker', 'ChartMinorGridLines',
  'ChartMinorTickMarks', 'ChartSmartLabel', 'ColumnSeparatorColor', 'CrossAt', 'DockOutsideChartArea',
  'Enabled', 'HeaderSeparatorColor', 'Length', 'Location', 'MinMovingDistance', 'Palette', 'ValueAxisName',
]);

const METADATA_SUBTREES = new Set([
  'DataSources', 'Query', 'ReportParametersLayout', 'ParameterValues',
  // Interactive-only features (hyperlinks, drill-through, click-to-sort, drill-down toggles) are
  // irrelevant to a static PDF/DOCX. Acknowledge them but never act on them, so a report that merely
  // declares them still renders as static content (drill-down renders expanded) instead of failing closed.
  'ActionInfo', 'Drillthrough', 'InteractiveSort', 'ToggleItem', 'ToggleImage',
]);
const REJECTED_SUBTREES = new Set([
  'Subreport', 'Map', 'GaugePanel', 'CustomReportItem',
]);

export const EXPRESSION_FUNCTION_CAPABILITIES = Object.freeze({
  IIF: CapabilityStatus.SUPPORTED,
  First: CapabilityStatus.SUPPORTED,
  Format: CapabilityStatus.SUPPORTED,
  Sum: CapabilityStatus.SUPPORTED,
  Count: CapabilityStatus.SUPPORTED,
  CountDistinct: CapabilityStatus.SUPPORTED,
  Avg: CapabilityStatus.SUPPORTED,
  Min: CapabilityStatus.SUPPORTED,
  Max: CapabilityStatus.SUPPORTED,
  Switch: CapabilityStatus.SUPPORTED,
  Choose: CapabilityStatus.SUPPORTED,
  IsNothing: CapabilityStatus.SUPPORTED,
  Lookup: CapabilityStatus.SUPPORTED,
  LookupSet: CapabilityStatus.SUPPORTED,
  MultiLookup: CapabilityStatus.SUPPORTED,
  RunningValue: CapabilityStatus.SUPPORTED,
  RowNumber: CapabilityStatus.SUPPORTED,
  Previous: CapabilityStatus.SUPPORTED,
  Last: CapabilityStatus.SUPPORTED,
  CountRows: CapabilityStatus.SUPPORTED,
  StDev: CapabilityStatus.SUPPORTED,
  StDevP: CapabilityStatus.SUPPORTED,
  Var: CapabilityStatus.SUPPORTED,
  VarP: CapabilityStatus.SUPPORTED,
  DateAdd: CapabilityStatus.SUPPORTED,
  DateDiff: CapabilityStatus.SUPPORTED,
  DatePart: CapabilityStatus.SUPPORTED,
  CDate: CapabilityStatus.SUPPORTED,
  Join: CapabilityStatus.SUPPORTED,
  // Registry-backed functions (string / math / conversion / date / formatting / inspection) are added by
  // their category modules and classified SUPPORTED here automatically.
  ...Object.fromEntries(FUNCTION_NAMES.map((name) => [name, CapabilityStatus.SUPPORTED])),
});

function localName(name) {
  return String(name).replace(/^@_/, '').split(':').at(-1);
}

function inheritedStatus(path, parentStatus) {
  if (parentStatus === CapabilityStatus.REJECTED) return CapabilityStatus.REJECTED;
  if (parentStatus === CapabilityStatus.METADATA_ONLY) return CapabilityStatus.METADATA_ONLY;
  const names = path.split('.').map(localName);
  if (names.some((name) => REJECTED_SUBTREES.has(name))) return CapabilityStatus.REJECTED;
  if (names.some((name) => METADATA_SUBTREES.has(name))) return CapabilityStatus.METADATA_ONLY;
  return null;
}

function classifyElement(path, name, parentStatus) {
  const inherited = inheritedStatus(path, parentStatus);
  if (inherited) return inherited;
  if (name === 'rd:TypeName') return CapabilityStatus.SUPPORTED;
  if (name.startsWith('rd:') || name.startsWith('am:')) return CapabilityStatus.METADATA_ONLY;
  if (name === 'df:DefaultFontFamily') return CapabilityStatus.SUPPORTED;
  if (name.startsWith('df:')) return CapabilityStatus.REJECTED;
  const local = localName(name);
  if (METADATA_SUBTREES.has(local)) return CapabilityStatus.METADATA_ONLY;
  if (REJECTED_SUBTREES.has(local)) return CapabilityStatus.REJECTED;
  if (METADATA_ELEMENTS.has(local)) return CapabilityStatus.METADATA_ONLY;
  if (SUPPORTED_ELEMENTS.has(local)) return CapabilityStatus.SUPPORTED;
  return CapabilityStatus.REJECTED;
}

function classifyDeclaredElement(name) {
  if (METADATA_SUBTREES.has(name) || METADATA_ELEMENTS.has(name)) return CapabilityStatus.METADATA_ONLY;
  if (SUPPORTED_ELEMENTS.has(name)) return CapabilityStatus.SUPPORTED;
  return CapabilityStatus.REJECTED;
}

function classifyDeclaredAttribute(name) {
  if (name === 'Name' || name === 'MustUnderstand') return CapabilityStatus.SUPPORTED;
  return CapabilityStatus.REJECTED;
}

export function rdl2016SchemaCatalogue() {
  const entries = [
    ...RDL_2016_SCHEMA_ELEMENT_NAMES.map((name) => ({ name, kind: 'ELEMENT', status: classifyDeclaredElement(name) })),
    ...RDL_2016_SCHEMA_ATTRIBUTE_NAMES.map((name) => ({ name, kind: 'ATTRIBUTE', status: classifyDeclaredAttribute(name) })),
  ];
  return {
    namespace: 'http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition',
    source: SCHEMA_SOURCES['http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition'],
    policyVersion: 1,
    defaultStatus: CapabilityStatus.REJECTED,
    summary: Object.fromEntries(Object.values(CapabilityStatus).map((status) => [
      status,
      entries.filter((entry) => entry.status === status).length,
    ])),
    entries,
  };
}

function classifyAttribute(path, name, parentStatus) {
  const inherited = inheritedStatus(path, parentStatus);
  if (inherited) return inherited;
  const local = localName(name);
  if (local === 'Name') return CapabilityStatus.SUPPORTED;
  if (local === 'MustUnderstand') return CapabilityStatus.SUPPORTED;
  if (local === 'TypeName' && name.includes('rd:')) return CapabilityStatus.SUPPORTED;
  if (local === 'xmlns' || name.includes('xmlns:') || local === 'schemaLocation') return CapabilityStatus.METADATA_ONLY;
  if (name.includes('xml:space')) return CapabilityStatus.METADATA_ONLY;
  if (name.includes('rd:') || name.includes('am:')) return CapabilityStatus.METADATA_ONLY;
  return CapabilityStatus.REJECTED;
}

function expressionFunctions(value) {
  if (typeof value !== 'string' || !value.trim().startsWith('=')) return [];
  const code = value.replace(/"(?:[^"]|"")*"|'[^']*'/g, '');
  return [...code.matchAll(/\b([A-Za-z_][A-Za-z0-9_]*)\s*\(/g)].map((match) => match[1]);
}

export function inspectRdlCapabilities(parsed, namespace) {
  const entries = new Map();
  const detectedFunctions = new Map();
  let customCodeDetected = false;

  function record(path, kind, status) {
    const key = `${kind}:${path}`;
    if (!entries.has(key)) entries.set(key, { path, kind, status });
  }

  function visit(value, path = '', parentStatus = null) {
    if (Array.isArray(value)) {
      for (const entry of value) visit(entry, path, parentStatus);
      return;
    }
    if (!value || typeof value !== 'object') {
      for (const name of expressionFunctions(value)) {
        const configured = Object.entries(EXPRESSION_FUNCTION_CAPABILITIES)
          .find(([known]) => known.toLowerCase() === name.toLowerCase());
        detectedFunctions.set(name.toLowerCase(), {
          name: configured?.[0] || name,
          status: configured?.[1] || CapabilityStatus.REJECTED,
        });
      }
      if (typeof value === 'string' && value.trim().startsWith('=') && /\bCode\./i.test(value)) customCodeDetected = true;
      return;
    }

    for (const [name, child] of Object.entries(value)) {
      if (name === '?xml') continue;
      if (name === '#text') {
        visit(child, path, parentStatus);
        continue;
      }
      if (name.startsWith('@_')) {
        const attributePath = `${path}.@${name.slice(2)}`;
        record(attributePath, 'ATTRIBUTE', classifyAttribute(attributePath, name, parentStatus));
        continue;
      }
      const childPath = path ? `${path}.${name}` : name;
      const status = classifyElement(childPath, name, parentStatus);
      record(childPath, 'ELEMENT', status);
      visit(child, childPath, status);
    }
  }

  visit(parsed);
  if (customCodeDetected) detectedFunctions.set('code.*', { name: 'Code.*', status: CapabilityStatus.REJECTED });
  const allEntries = [...entries.values()].sort((left, right) => left.path.localeCompare(right.path));
  const functions = [...detectedFunctions.values()].sort((left, right) => left.name.localeCompare(right.name));
  const statusByPath = new Map(allEntries.map((entry) => [entry.path, entry.status]));
  const rejected = allEntries.filter((entry) => {
    if (entry.status !== CapabilityStatus.REJECTED) return false;
    const separator = entry.path.lastIndexOf('.');
    const parentPath = separator < 0 ? '' : entry.path.slice(0, separator);
    return statusByPath.get(parentPath) !== CapabilityStatus.REJECTED;
  });
  const summary = Object.fromEntries(Object.values(CapabilityStatus).map((status) => [
    status,
    allEntries.filter((entry) => entry.status === status).length,
  ]));
  return {
    schema: {
      namespace,
      source: SCHEMA_SOURCES[namespace] || null,
      policyVersion: 1,
      defaultStatus: CapabilityStatus.REJECTED,
      declaredCatalogue: namespace.endsWith('/2016/01/reportdefinition') ? {
        elements: RDL_2016_SCHEMA_ELEMENT_NAMES.length,
        attributes: RDL_2016_SCHEMA_ATTRIBUTE_NAMES.length,
        summary: rdl2016SchemaCatalogue().summary,
      } : null,
    },
    summary,
    entries: allEntries,
    rejected,
    expressions: {
      catalogue: Object.entries(EXPRESSION_FUNCTION_CAPABILITIES).map(([name, status]) => ({ name, status })),
      detected: functions,
      rejected: functions.filter((entry) => entry.status === CapabilityStatus.REJECTED),
    },
  };
}
