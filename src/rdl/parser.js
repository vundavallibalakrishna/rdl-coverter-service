import { createHash } from 'node:crypto';
import { XMLParser } from 'fast-xml-parser';
import { ServiceError } from '../errors.js';
import { toPoints } from '../units.js';
import { inspectRdlCapabilities } from './capabilities.js';
import { canonicalizeCulture } from './culture.js';
import { FUNCTION_NAMES } from './functions/index.js';
import { CUSTOM_CODE_FUNCTION_NAMES } from './functions/customCode.js';
import { asArray, childEntries, firstDefined, flattenCellItems, parseBoolean, RENDERABLE_CELL_ITEMS, styleTextValue, textValue } from './helpers.js';

const SUPPORTED_NAMESPACES = new Set([
  'http://schemas.microsoft.com/sqlserver/reporting/2008/01/reportdefinition',
  'http://schemas.microsoft.com/sqlserver/reporting/2010/01/reportdefinition',
  'http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition',
]);
const SUPPORTED_REQUIRED_NAMESPACES = new Set([
  'http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition/defaultfontfamily',
]);
const ITEM_NAMES = ['Textbox', 'Tablix', 'Rectangle', 'Line', 'Image', 'Subreport', 'Chart', 'Map', 'GaugePanel', 'CustomReportItem'];

// A size that may be a conditional expression (e.g. a border width of
// =IIF(Fields!Row_Number_1.Value = 1, "1pt", "0pt")). Such a value depends on row data and cannot be
// converted to points until render time, so keep the raw expression string; literal sizes convert now.
function sizeOrExpression(value, fallback) {
  if (typeof value === 'number') return value;
  const text = textValue(value, '');
  if (text.startsWith('=')) return text;
  return toPoints(text || `${fallback}pt`, fallback);
}

// An empty element (`<Style />`) is not a declared value — see styleTextValue. Drop blanks BEFORE the
// fallback chain runs, otherwise `<TopBorder><Style /></TopBorder>` would shadow the general Border it is
// supposed to inherit from. Numbers are already-normalized fallback values and pass through untouched.
function declaredValue(value) {
  if (value === undefined || value === null) return undefined;
  if (typeof value === 'number') return value;
  return textValue(value, '').trim() === '' ? undefined : value;
}

function borderOf(border, fallback = {}) {
  const width = firstDefined(declaredValue(border?.Width), declaredValue(fallback.width), '1pt');
  return {
    style: styleTextValue(firstDefined(declaredValue(border?.Style), declaredValue(border?.['rd:Style']), declaredValue(fallback.style)), 'None'),
    color: styleTextValue(firstDefined(declaredValue(border?.Color), declaredValue(fallback.color)), '#000000'),
    width: sizeOrExpression(width, 1),
  };
}

function styleOf(style = {}, defaultFontFamily = 'Arial') {
  const border = borderOf(style.Border);
  return {
    color: styleTextValue(style.Color, '#000000'),
    backgroundColor: styleTextValue(style.BackgroundColor, null),
    fontFamily: styleTextValue(style.FontFamily, defaultFontFamily),
    // FontSize / Padding / LineHeight are RdlSize ExpressionType — a literal OR an =expression evaluated per
    // row. Keep expressions raw (sizeOrExpression); the renderers resolve them via styleSize at render time.
    fontSize: sizeOrExpression(style.FontSize, 10),
    fontWeight: styleTextValue(style.FontWeight, 'Normal'),
    fontStyle: styleTextValue(style.FontStyle, 'Normal'),
    textDecoration: styleTextValue(style.TextDecoration, 'None'),
    textAlign: styleTextValue(style.TextAlign, 'Left'),
    verticalAlign: styleTextValue(style.VerticalAlign, 'Top'),
    writingMode: styleTextValue(style.WritingMode, 'Default'),
    format: styleTextValue(style.Format, null),
    paddingLeft: sizeOrExpression(style.PaddingLeft, 2),
    paddingRight: sizeOrExpression(style.PaddingRight, 2),
    paddingTop: sizeOrExpression(style.PaddingTop, 2),
    paddingBottom: sizeOrExpression(style.PaddingBottom, 2),
    lineHeight: style.LineHeight === undefined ? null : sizeOrExpression(style.LineHeight, 0),
    border,
    borders: {
      top: borderOf(style.TopBorder, border),
      right: borderOf(style.RightBorder, border),
      bottom: borderOf(style.BottomBorder, border),
      left: borderOf(style.LeftBorder, border),
    },
  };
}

function evaluationModeOf(value) {
  const mode = textValue(value?.['@_EvaluationMode'], 'Auto');
  if (!/^(Auto|Constant)$/i.test(mode)) {
    throw new ServiceError('RDL_INVALID', `Unsupported TextRun Value EvaluationMode: ${mode}`);
  }
  return /^constant$/i.test(mode) ? 'Constant' : 'Auto';
}

function textboxParagraphs(textbox, defaultFontFamily) {
  const paragraphs = asArray(textbox.Paragraphs?.Paragraph);
  if (paragraphs.length === 0) return [[{
    value: textValue(textbox.Value, ''),
    markupType: 'None',
    style: styleOf(textbox.Style, defaultFontFamily),
  }]];
  return paragraphs.map((paragraph) => asArray(paragraph.TextRuns?.TextRun).map((run) => ({
    value: textValue(run.Value, ''),
    evaluationMode: evaluationModeOf(run.Value),
    markupType: textValue(run.MarkupType, 'None'),
    // TextRun styles are independent within a textbox. Preserve their effective inherited style instead
    // of flattening every run to the first run's font. Renderers can then switch weight, size, family,
    // colour, and decoration at the exact run boundary declared by the RDL.
    style: styleOf({
      ...(textbox.Style || {}),
      ...(paragraph.Style || {}),
      ...(run.Style || {}),
      Border: textbox.Style?.Border,
    }, defaultFontFamily),
  })));
}

function textboxParagraphStyles(textbox, defaultFontFamily) {
  const paragraphs = asArray(textbox.Paragraphs?.Paragraph);
  if (paragraphs.length === 0) return [{ ...styleOf(textbox.Style, defaultFontFamily), spaceBefore: 0, spaceAfter: 0 }];
  return paragraphs.map((paragraph) => ({
    ...styleOf({
      ...(textbox.Style || {}),
      ...(paragraph.Style || {}),
      Border: textbox.Style?.Border,
    }, defaultFontFamily),
    spaceBefore: sizeOrExpression(paragraph.SpaceBefore, 0),
    spaceAfter: sizeOrExpression(paragraph.SpaceAfter, 0),
  }));
}

function textboxStyle(textbox, defaultFontFamily) {
  const paragraph = asArray(textbox.Paragraphs?.Paragraph)[0];
  const run = asArray(paragraph?.TextRuns?.TextRun)[0];
  return styleOf({
    ...(textbox.Style || {}),
    ...(paragraph?.Style || {}),
    ...(run?.Style || {}),
    Border: textbox.Style?.Border,
  }, defaultFontFamily);
}

function baseItem(type, value, defaultFontFamily) {
  return {
    type,
    name: value?.['@_Name'] || null,
    top: toPoints(textValue(value?.Top, '0pt')),
    left: toPoints(textValue(value?.Left, '0pt')),
    width: toPoints(textValue(value?.Width, '0pt')),
    height: toPoints(textValue(value?.Height, '0pt')),
    zIndex: Number.parseInt(textValue(value?.ZIndex, '0'), 10) || 0,
    keepTogether: parseBoolean(value?.KeepTogether),
    hidden: textValue(value?.Visibility?.Hidden, 'false'),
    style: styleOf(value?.Style, defaultFontFamily),
    pageBreak: value?.PageBreak ? {
      location: textValue(value.PageBreak.BreakLocation, 'None'),
      disabled: textValue(value.PageBreak.Disabled, 'false'),
    } : null,
    pageName: textValue(value?.PageName, null),
  };
}

function parseCell(cell = {}, defaultFontFamily) {
  const contents = cell.CellContents || {};
  const entries = childEntries(contents, ITEM_NAMES);
  return {
    colSpan: Number(contents.ColSpan || 1),
    rowSpan: Number(contents.RowSpan || 1),
    items: entries.map(([type, item]) => parseItem(type, item, defaultFontFamily)),
  };
}

// In RDL a TablixBody row lists one <TablixCell> per column: a cell with ColSpan=C is followed
// by C-1 empty placeholder cells representing the columns it covers. If those placeholders are
// kept they are miscounted as real grid columns, which shifts every later cell right and
// collapses trailing header labels (e.g. "4th Line", "Combined Assurance Level") onto the last
// column. Drop only the empty placeholders immediately covered by a preceding span; a covered
// slot that unexpectedly carries content is kept, so malformed input degrades safely.
export function dropCoveredPlaceholders(cells) {
  const result = [];
  let covered = 0;
  for (const cell of cells) {
    if (covered > 0 && (cell.colSpan || 1) === 1 && (cell.rowSpan || 1) === 1 && cell.items.length === 0) {
      covered -= 1;
      continue;
    }
    covered = 0;
    result.push(cell);
    if ((cell.colSpan || 1) > 1) covered = (cell.colSpan || 1) - 1;
  }
  return result;
}

function parseFilters(container) {
  return asArray(container?.Filter).map((filter) => ({
    expression: textValue(filter.FilterExpression),
    operator: textValue(filter.Operator, 'Equal'),
    values: asArray(filter.FilterValues?.FilterValue).map(textValue),
  }));
}

// Report- or group-scoped variables: a named value expression referenced elsewhere as Variables!Name.Value.
function parseVariables(container) {
  return asArray(container?.Variable)
    .map((variable) => ({ name: variable['@_Name'] || null, value: textValue(variable.Value, '') }))
    .filter((entry) => entry.name);
}

// Flatten every group's variable definitions into one name→expression map. They resolve lazily in the
// current row scope at evaluation time, which reproduces SSRS's per-scope value for the usual case where a
// variable is a deterministic function of its scope's fields.
function collectGroupVariables(members, target) {
  for (const member of members || []) {
    for (const variable of member.group?.variables || []) target[variable.name] = variable.value;
    collectGroupVariables(member.children, target);
  }
  return target;
}

function parseMember(member = {}, defaultFontFamily) {
  return {
    group: member.Group ? {
      name: member.Group['@_Name'] || null,
      expressions: asArray(member.Group.GroupExpressions?.GroupExpression).map((entry) => textValue(entry)),
      pageBreak: member.Group.PageBreak ? textValue(member.Group.PageBreak.BreakLocation, 'None') : 'None',
      filters: parseFilters(member.Group.Filters),
      // Recursive (parent/child) grouping: a Group/Parent expression turns a flat group into a
      // self-referencing hierarchy that renders expanded with indentation by recursion depth.
      parent: member.Group.Parent !== undefined && member.Group.Parent !== null ? textValue(member.Group.Parent, null) : null,
      variables: parseVariables(member.Group.Variables),
      // DomainScope / NaturalGroup are advisory grouping hints; acknowledged but not acted on (the group
      // renders its natural instances), so a report that declares them still renders instead of failing closed.
      domainScope: textValue(member.Group.DomainScope, null),
      naturalGroup: member.Group.NaturalGroup !== undefined,
    } : null,
    repeatOnNewPage: parseBoolean(member.RepeatOnNewPage),
    keepTogether: parseBoolean(member.KeepTogether),
    keepWithGroup: textValue(member.KeepWithGroup, 'None'),
    fixedData: parseBoolean(member.FixedData),
    hideIfNoRows: parseBoolean(member.HideIfNoRows),
    hidden: textValue(member.Visibility?.Hidden, 'false'),
    sortExpressions: asArray(member.SortExpressions?.SortExpression).map((sort) => ({ value: textValue(sort.Value), direction: textValue(sort.Direction, 'Ascending') })),
    header: member.TablixHeader ? {
      size: toPoints(textValue(member.TablixHeader.Size, '0pt')),
      cell: parseCell({ CellContents: member.TablixHeader.CellContents }, defaultFontFamily),
    } : null,
    children: asArray(member.TablixMembers?.TablixMember).map((child) => parseMember(child, defaultFontFamily)),
  };
}

function leafMemberPaths(members, parent = [], target = []) {
  for (const member of members || []) {
    const path = [...parent, member];
    if (member.children.length) leafMemberPaths(member.children, path, target);
    else target.push(path);
  }
  return target;
}

function rowHeaderColumns(paths) {
  const candidates = paths.map((path) => path.filter((member) => member.header).map((member) => member.header.size));
  candidates.sort((left, right) => right.length - left.length || right.reduce((sum, width) => sum + width, 0) - left.reduce((sum, width) => sum + width, 0));
  return candidates[0] || [];
}

function parseChartMember(member) {
  if (!member) return null;
  return {
    group: member.Group ? {
      name: member.Group['@_Name'] || null,
      expressions: asArray(member.Group.GroupExpressions?.GroupExpression).map((entry) => textValue(entry)),
    } : null,
    sortExpressions: asArray(member.SortExpressions?.SortExpression).map((sort) => ({ value: textValue(sort.Value), direction: textValue(sort.Direction, 'Ascending') })),
    filters: asArray(member.Group?.Filters?.Filter).map((filter) => ({
      expression: textValue(filter.FilterExpression),
      operator: textValue(filter.Operator, 'Equal'),
      values: asArray(filter.FilterValues?.FilterValue).map(textValue),
    })),
    label: textValue(member.Label, null),
  };
}

// Maps an SSRS series (Type, Subtype) to a supported chart type. Shape charts split into pie/doughnut
// by subtype (exploded variants render the same); Point renders as Scatter; an omitted Type defaults
// to a column chart.
function chartTypeFor(rawType, rawSubtype) {
  const type = String(rawType).toLowerCase();
  const subtype = String(rawSubtype || '').toLowerCase();
  if (type === 'bar') return 'bar';
  if (type === 'column' || type === '') return 'column';
  if (type === 'line') return 'line';
  if (type === 'area') return 'area';
  if (type === 'scatter' || type === 'point') return 'scatter';
  if (type === 'shape') return /doughnut/.test(subtype) ? 'doughnut' : 'pie';
  return null;
}

// Stacking (Stacked / PercentStacked subtypes) applies only to bar/column/area charts.
function chartStackMode(rawType, rawSubtype) {
  const type = String(rawType).toLowerCase();
  if (!['bar', 'column', 'area', ''].includes(type)) return 'none';
  const subtype = String(rawSubtype || '').toLowerCase();
  if (/percent/.test(subtype)) return 'percent';
  if (/stacked/.test(subtype)) return 'stacked';
  return 'none';
}

function chartExploded(rawType, rawSubtype) {
  return String(rawType).toLowerCase() === 'shape' && /exploded/.test(String(rawSubtype || '').toLowerCase());
}

function parseChart(value, defaultFontFamily) {
  const item = baseItem('Chart', value, defaultFontFamily);
  const seriesCollection = asArray(value.ChartData?.ChartSeriesCollection?.ChartSeries);
  // SSRS defaults an omitted series <Type> to a vertical column chart. Unknown/unsupported types are
  // recorded (fail-closed via collectUnsupported) rather than thrown, so /v1/analyze can report them.
  const rawType = textValue(seriesCollection[0]?.Type, 'Column');
  const rawSubtype = textValue(seriesCollection[0]?.Subtype, '');
  const chartType = chartTypeFor(rawType, rawSubtype);
  const stacked = chartStackMode(rawType, rawSubtype);
  const exploded = chartExploded(rawType, rawSubtype);
  const seriesMember = parseChartMember(asArray(value.ChartSeriesHierarchy?.ChartMembers?.ChartMember)[0]);
  const area = asArray(value.ChartAreas?.ChartArea)[0] || {};
  const valueAxis = asArray(area.ChartValueAxes?.ChartAxis)[0] || {};
  const categoryAxis = asArray(area.ChartCategoryAxes?.ChartAxis)[0] || {};
  const legend = asArray(value.ChartLegends?.ChartLegend)[0];
  const title = asArray(value.ChartTitles?.ChartTitle)[0];
  const legendStyle = styleOf(legend?.Style, defaultFontFamily);
  if (legend?.Style?.FontSize === undefined) legendStyle.fontSize = 8;
  const titleStyle = styleOf(title?.Style, defaultFontFamily);
  if (title?.Style?.FontSize === undefined) titleStyle.fontSize = 9;
  if (title?.Style?.FontWeight === undefined) titleStyle.fontWeight = 'Bold';
  if (title?.Style?.TextAlign === undefined) titleStyle.textAlign = 'Center';
  const parseAxis = (axis) => ({
    interval: textValue(axis.Interval, null),
    labelsAutoFitDisabled: textValue(axis.LabelsAutoFitDisabled, 'false'),
    allowLabelRotation: textValue(axis.AllowLabelRotation, 'None'),
    style: styleOf(axis.Style, defaultFontFamily),
  });
  const customProperties = (series) => Object.fromEntries(asArray(series.CustomProperties?.CustomProperty)
    .map((property) => [textValue(property.Name, ''), textValue(property.Value, '')])
    .filter(([name]) => name));
  const propertyAllowed = (name) => {
    if (/^PointWidth$/i.test(name)) return chartType === 'bar' || chartType === 'column';
    if (/^(?:PieLineColor|PieStartAngle)$/i.test(name)) return chartType === 'pie' || chartType === 'doughnut';
    return false;
  };
  return {
    ...item,
    datasetName: textValue(value.DataSetName),
    chartType,
    stacked,
    exploded,
    palette: textValue(value.Palette, 'Default'),
    customPaletteColors: asArray(value.ChartCustomPaletteColors?.ChartCustomPaletteColor).map((entry) => textValue(entry)),
    unsupportedType: chartType ? null : rawType,
    category: parseChartMember(asArray(value.ChartCategoryHierarchy?.ChartMembers?.ChartMember)[0]),
    series: seriesMember?.group ? seriesMember : null,
    seriesDefs: seriesCollection.map((series) => {
      const point = asArray(series.ChartDataPoints?.ChartDataPoint)[0] || {};
      const dataLabel = point.ChartDataLabel || {};
      return {
        name: series['@_Name'] || null,
        y: textValue(point.ChartDataPointValues?.Y, null),
        color: textValue(point.Style?.Color, null),
        dataLabel: {
          expression: textValue(dataLabel.Label, null),
          useValueAsLabel: parseBoolean(dataLabel.UseValueAsLabel),
          visible: parseBoolean(dataLabel.Visible),
          position: textValue(dataLabel.Position, 'Auto'),
          style: styleOf(dataLabel.Style, defaultFontFamily),
        },
        customProperties: customProperties(series),
        unsupportedCustomProperties: Object.keys(customProperties(series)).filter((name) => !propertyAllowed(name)),
      };
    }),
    valueAxis: parseAxis(valueAxis),
    categoryAxis: parseAxis(categoryAxis),
    // Draw the built-in legend only when the RDL declares one. Charts without <ChartLegends> (Chart1/
    // Chart2 here) rely on a manual legend built from separate report items, so a synthetic legend
    // would duplicate it.
    chartArea: { style: styleOf(area.Style, defaultFontFamily) },
    legend: {
      visible: Boolean(legend),
      hidden: textValue(legend?.Hidden, 'false'),
      position: textValue(legend?.Position, 'RightTop'),
      layout: textValue(legend?.Layout, 'AutoTable'),
      dockOutsideChartArea: textValue(legend?.DockOutsideChartArea, 'false'),
      style: legendStyle,
    },
    title: title ? {
      caption: textValue(title.Caption, null),
      hidden: textValue(title.Hidden, 'false'),
      position: textValue(title.Position, 'TopCenter'),
      style: titleStyle,
    } : null,
    noDataMessage: textValue(value.ChartNoDataMessage?.Caption, 'No Data Available'),
  };
}

function parseItem(type, value, defaultFontFamily) {
  if (type === 'Chart') return parseChart(value, defaultFontFamily);
  const item = baseItem(type, value, defaultFontFamily);
  if (type === 'Textbox') {
    const paragraphs = textboxParagraphs(value, defaultFontFamily);
    return {
      ...item,
      style: textboxStyle(value, defaultFontFamily),
      value: paragraphs.map((runs) => runs.map((run) => run.value).join('')).join('\n'),
      paragraphs,
      paragraphStyles: textboxParagraphStyles(value, defaultFontFamily),
      hideDuplicates: textValue(value.HideDuplicates, null),
      canGrow: parseBoolean(value.CanGrow, true),
      canShrink: parseBoolean(value.CanShrink),
    };
  }
  if (type === 'Rectangle') {
    return { ...item, items: childEntries(value.ReportItems, ITEM_NAMES).map(([childType, child]) => parseItem(childType, child, defaultFontFamily)) };
  }
  if (type === 'Line') return item;
  if (type === 'Image') {
    return { ...item, source: textValue(value.Source), value: textValue(value.Value), sizing: textValue(value.Sizing, 'FitProportional') };
  }
  if (type === 'Subreport') {
    return {
      ...item,
      reportName: textValue(value.ReportName, null),
      parameters: asArray(value.Parameters?.Parameter).map((parameter) => ({
        name: parameter['@_Name'] || null,
        value: textValue(parameter.Value, ''),
      })),
      mergeTransactions: parseBoolean(value.MergeTransactions),
      omitBorderOnPageBreak: parseBoolean(value.OmitBorderOnPageBreak),
    };
  }
  if (type === 'Tablix') {
    const rawRows = asArray(value.TablixBody?.TablixRows?.TablixRow);
    const rawBodyColumns = asArray(value.TablixBody?.TablixColumns?.TablixColumn)
      .map((column) => toPoints(textValue(column.Width, '72pt'), 72));
    const rowMembers = asArray(value.TablixRowHierarchy?.TablixMembers?.TablixMember).map((member) => parseMember(member, defaultFontFamily));
    const columnMembers = asArray(value.TablixColumnHierarchy?.TablixMembers?.TablixMember).map((member) => parseMember(member, defaultFontFamily));
    const rowMemberPaths = leafMemberPaths(rowMembers);
    const columnMemberPaths = leafMemberPaths(columnMembers);
    // A static hidden column is removed from the physical grid, as SSRS does. Keeping its tiny declared
    // width creates a phantom right edge (often a doubled border) and prevents merged cells from closing
    // at the visible table boundary. Expression-backed visibility remains in the model for render-time
    // handling; only literal Hidden=true is safe to eliminate during normalization.
    const hiddenColumn = (index) => {
      const path = columnMemberPaths.length === rawBodyColumns.length
        ? columnMemberPaths[index]
        : (columnMemberPaths.length === 1 ? columnMemberPaths[0] : null);
      return Boolean(path?.length && path.every((member) => String(member.hidden).trim().toLowerCase() === 'true'));
    };
    const visibleColumnIndexes = rawBodyColumns.map((_, index) => index).filter((index) => !hiddenColumn(index));
    const bodyColumns = visibleColumnIndexes.map((index) => rawBodyColumns[index]);
    const rows = rawRows.map((row) => ({
      height: toPoints(textValue(row.Height, '18pt'), 18),
      hidden: textValue(row.Visibility?.Hidden, 'false'),
      cells: dropCoveredPlaceholders(visibleColumnIndexes
        .map((index) => asArray(row.TablixCells?.TablixCell)[index])
        .filter(Boolean)
        .map((cell) => parseCell(cell, defaultFontFamily))),
    }));
    const headerColumns = rowHeaderColumns(rowMemberPaths);
    // Dynamic column groups (matrix / cross-tab) are detected here so the whole downstream pipeline
    // (column expansion, intersection scoping, renderers) can gate on it and leave static-column
    // tablixes byte-identical.
    const hasColumnGroups = columnMemberPaths.some((path) => path.some((member) => member.group?.expressions?.length));
    // TablixCorner: the static top-left block of a matrix. Parsed only for structure; consumed when
    // dynamic column groups are present.
    const tablixCorner = asArray(value.TablixCorner?.TablixCornerRows?.TablixCornerRow)
      .map((cornerRow) => asArray(cornerRow.TablixCornerCell).map((cell) => parseCell({ CellContents: cell.CellContents }, defaultFontFamily)));
    return {
      ...item,
      datasetName: textValue(value.DataSetName),
      bodyColumns,
      rowHeaderColumns: headerColumns,
      columns: [...headerColumns, ...bodyColumns],
      rows,
      rowMembers,
      rowMemberPaths,
      columnMembers,
      columnMemberPaths,
      hasColumnGroups,
      tablixCorner,
      repeatColumnHeaders: parseBoolean(value.RepeatColumnHeaders),
      repeatRowHeaders: parseBoolean(value.RepeatRowHeaders),
      fixedColumnHeaders: parseBoolean(value.FixedColumnHeaders),
      fixedRowHeaders: parseBoolean(value.FixedRowHeaders),
      noRowsMessage: textValue(value.NoRowsMessage, null),
      filters: parseFilters(value.Filters),
      sortExpressions: asArray(value.SortExpressions?.SortExpression).map((sort) => ({ value: textValue(sort.Value), direction: textValue(sort.Direction, 'Ascending') })),
    };
  }
  return item;
}

function parseParameters(report) {
  return asArray(report.ReportParameters?.ReportParameter).map((parameter) => {
    const validValuesReference = parameter.ValidValues?.DataSetReference;
    const validReference = firstDefined(
      validValuesReference,
      parameter.DefaultValue?.DataSetReference,
    );
    return {
      name: parameter['@_Name'],
      dataType: textValue(parameter.DataType, 'String'),
      nullable: parseBoolean(parameter.Nullable),
      allowBlank: parseBoolean(parameter.AllowBlank),
      multiValue: parseBoolean(parameter.MultiValue),
      prompt: textValue(parameter.Prompt, parameter['@_Name']),
      hidden: parseBoolean(parameter.Hidden),
      defaultValues: asArray(parameter.DefaultValue?.Values?.Value).map(textValue),
      lookupDataset: validReference ? textValue(validReference.DataSetName) : null,
      lookupValueField: validValuesReference ? textValue(validValuesReference.ValueField) : null,
      lookupLabelField: validValuesReference ? textValue(validValuesReference.LabelField) : null,
      staticValidValues: asArray(parameter.ValidValues?.ParameterValues?.ParameterValue).map((entry) => ({
        value: textValue(entry.Value),
        label: textValue(entry.Label, textValue(entry.Value)),
      })),
    };
  });
}

function parseDatasets(report) {
  return asArray(report.DataSets?.DataSet).map((dataset) => ({
    name: dataset['@_Name'],
    // A Field carries either a DataField (bound to a query result column) or a Value (computed from other
    // fields). Only a bound one is a key in the posted rows, so the two must stay distinguishable — see
    // rdl/fields.js. Without this a calculated field inherits its Name as a DataField and is then demanded
    // of data that can never contain it.
    fields: asArray(dataset.Fields?.Field).map((field) => {
      const dataField = field.DataField === undefined || field.DataField === null ? '' : textValue(field.DataField, '');
      const expression = field.Value === undefined || field.Value === null ? '' : textValue(field.Value, '');
      const calculated = !dataField && expression !== '';
      return {
        name: field['@_Name'],
        dataField: dataField || field['@_Name'],
        calculated,
        value: calculated ? expression : undefined,
        typeName: textValue(field['rd:TypeName'] || field.TypeName, 'System.String'),
      };
    }),
    parameterNames: asArray(dataset.Query?.QueryParameters?.QueryParameter).map((parameter) => parameter['@_Name']),
    hasQuery: Boolean(dataset.Query?.CommandText || dataset.Query?.CommandType),
  }));
}

function collectItemDatasets(items, target = new Set()) {
  for (const item of items) {
    if (item.datasetName) target.add(item.datasetName);
    if (item.items) collectItemDatasets(item.items, target);
    if (item.rows) for (const row of item.rows) for (const cell of row.cells) collectItemDatasets(cell.items, target);
  }
  return target;
}

// A bare subreport cannot render because its definition and invocation-specific datasets live outside the
// parent RDL. Preserve complete caller-facing dependency metadata so /v1/analyze can identify what a
// render-time bundle must supply without resolving a report-server path.
function collectSubreports(items, target = []) {
  const visitMembers = (members) => {
    for (const member of members || []) {
      if (member.header) collectSubreports(member.header.cell.items, target);
      visitMembers(member.children);
    }
  };

  for (const item of items || []) {
    if (item.type === 'Subreport') {
      target.push({
        name: item.name,
        reportName: item.reportName,
        parameters: item.parameters,
        keepTogether: item.keepTogether,
        mergeTransactions: item.mergeTransactions,
        omitBorderOnPageBreak: item.omitBorderOnPageBreak,
      });
    }
    if (item.type === 'Tablix') {
      visitMembers(item.rowMembers);
      visitMembers(item.columnMembers);
      for (const cornerRow of item.tablixCorner || []) {
        for (const cornerCell of cornerRow) collectSubreports(cornerCell.items, target);
      }
    }
    if (item.items) collectSubreports(item.items, target);
    if (item.rows) {
      for (const row of item.rows) {
        for (const cell of row.cells) collectSubreports(cell.items, target);
      }
    }
  }
  return target;
}

// A tablix cell is a distinct rendering context: only the types in RENDERABLE_CELL_ITEMS are actually drawn
// there, regardless of what the catalogue says about the element name elsewhere in the report. Report
// anything else so the report fails closed at analyze time instead of rendering a silently empty cell.
function collectUnsupportedCell(cellItems, target) {
  for (const item of flattenCellItems(cellItems)) {
    if (!RENDERABLE_CELL_ITEMS.has(item.type)) target.add(`TablixCellContent:${item.type}`);
  }
  return collectUnsupported(cellItems, target);
}

function collectUnsupported(items, target = new Set()) {
  for (const item of items) {
    if (['Subreport', 'Map', 'GaugePanel', 'CustomReportItem'].includes(item.type)) target.add(item.type);
    if (item.type === 'Chart' && item.unsupportedType) target.add(`ChartType:${item.unsupportedType}`);
    if (item.type === 'Chart') {
      for (const series of item.seriesDefs || []) {
        for (const property of series.unsupportedCustomProperties || []) target.add(`ChartProperty:${property}`);
      }
    }
    if (item.type === 'Image' && item.source !== 'Embedded') target.add(`ImageSource:${item.source || 'missing'}`);
    if (item.type === 'Textbox') {
      const writingMode = String(item.style?.writingMode || 'Default');
      if (!writingMode.startsWith('=') && !/^(Default|Horizontal|Vertical|Rotate270)$/i.test(writingMode)) {
        target.add(`WritingMode:${writingMode}`);
      }
      for (const paragraph of item.paragraphs || []) for (const run of paragraph) {
        const markupType = typeof run === 'object' ? run.markupType : 'None';
        if (!/^(None|HTML)$/i.test(String(markupType || 'None'))) target.add(`MarkupType:${markupType}`);
        const evaluationMode = typeof run === 'object' ? run.evaluationMode : 'Auto';
        if (!/^(Auto|Constant)$/i.test(String(evaluationMode || 'Auto'))) target.add(`EvaluationMode:${evaluationMode}`);
      }
    }
    if (item.type === 'Tablix') {
      const visitMembers = (members) => {
        for (const member of members || []) {
          if (member.header) collectUnsupportedCell(member.header.cell.items, target);
          visitMembers(member.children);
        }
      };
      visitMembers(item.rowMembers);
      visitMembers(item.columnMembers);
      for (const cornerRow of item.tablixCorner || []) for (const cornerCell of cornerRow) collectUnsupportedCell(cornerCell.items, target);
    }
    if (item.items) collectUnsupported(item.items, target);
    if (item.rows) for (const row of item.rows) for (const cell of row.cells) collectUnsupportedCell(cell.items, target);
  }
  return target;
}

function collectUnsupportedExpressions(value, target = new Set()) {
  if (typeof value === 'string' && value.trim().startsWith('=')) {
    const code = value.replace(/"(?:[^"]|"")*"|'[^']*'/g, '');
    const supportedFunctions = new Set(['iif', 'first', 'format', 'sum', 'count', 'countdistinct', 'avg', 'min', 'max', 'runningvalue',
      'switch', 'choose', 'isnothing', 'rownumber', 'previous', 'last', 'countrows', 'stdev', 'stdevp', 'var', 'varp', 'join',
      'lookup', 'lookupset', 'multilookup', 'dateadd', 'datediff', 'datepart', 'cdate',
      ...FUNCTION_NAMES.map((name) => name.toLowerCase()),
      ...CUSTOM_CODE_FUNCTION_NAMES.map((name) => name.toLowerCase())]);
    for (const match of code.matchAll(/\b((?:Code\.)?[A-Za-z_][A-Za-z0-9_]*)\s*\(/gi)) {
      const name = match[1];
      if (!supportedFunctions.has(name.toLowerCase())) {
        if (/^Code\./i.test(name)) target.add('CustomCode');
        else target.add(`ExpressionFunction:${name}`);
      }
    }
  } else if (Array.isArray(value)) {
    for (const entry of value) collectUnsupportedExpressions(entry, target);
  } else if (value && typeof value === 'object') {
    for (const [key, entry] of Object.entries(value)) {
      if (
        key === 'value' &&
        (
          /^constant$/i.test(String(value.evaluationMode || '')) ||
          Array.isArray(value.paragraphs)
        )
      ) continue;
      collectUnsupportedExpressions(entry, target);
    }
  }
  return target;
}

function attachDatasetFields(items, datasetsByName) {
  for (const item of items) {
    if (item.datasetName) item.datasetFields = datasetsByName.get(item.datasetName)?.fields || [];
    if (item.items) attachDatasetFields(item.items, datasetsByName);
    if (item.rows) for (const row of item.rows) for (const cell of row.cells) attachDatasetFields(cell.items, datasetsByName);
  }
}

export function parseRdl(input, limits = {}) {
  const xml = Buffer.isBuffer(input) ? input.toString('utf8') : String(input || '');
  if (!xml.trim()) throw new ServiceError('RDL_INVALID', 'RDL file is empty');
  if (Buffer.byteLength(xml) > (limits.maxRdlBytes || 10 * 1024 * 1024)) throw new ServiceError('RDL_INVALID', 'RDL file exceeds the configured size limit', 413);
  if (/<!DOCTYPE|<!ENTITY/i.test(xml)) throw new ServiceError('RDL_INVALID', 'DTD and entity declarations are not allowed');
  const rootAttributes = xml.match(/<Report\b([^>]*)>/i)?.[1] || '';
  const namespace = rootAttributes.match(/\bxmlns=["']([^"']+)["']/i)?.[1];
  if (!SUPPORTED_NAMESPACES.has(namespace)) throw new ServiceError('RDL_INVALID', `Unsupported RDL namespace: ${namespace || 'missing'}`);
  const declaredNamespaces = new Map([...rootAttributes.matchAll(/\bxmlns:([A-Za-z_][\w.-]*)=["']([^"']+)["']/g)].map((match) => [match[1], match[2]]));
  const requiredPrefixes = (rootAttributes.match(/\bMustUnderstand=["']([^"']*)["']/i)?.[1] || '').trim().split(/\s+/).filter(Boolean);
  for (const prefix of requiredPrefixes) {
    const requiredNamespace = declaredNamespaces.get(prefix);
    if (!requiredNamespace || !SUPPORTED_REQUIRED_NAMESPACES.has(requiredNamespace)) {
      throw new ServiceError('RDL_INVALID', `Unsupported required RDL namespace: ${prefix}`);
    }
  }
  const nodeCount = (xml.match(/<(?!!|\?|\/)[A-Za-z_][^>]*>/g) || []).length;
  if (nodeCount > (limits.maxXmlNodes || 250_000)) throw new ServiceError('RDL_INVALID', 'RDL contains too many XML nodes', 413);
  let depth = 0;
  let maximumDepth = 0;
  for (const match of xml.matchAll(/<\/?[A-Za-z_][^>]*>/g)) {
    const tag = match[0];
    if (tag.startsWith('</')) depth = Math.max(0, depth - 1);
    else if (!tag.endsWith('/>')) {
      depth += 1;
      maximumDepth = Math.max(maximumDepth, depth);
    }
  }
  if (maximumDepth > (limits.maxXmlDepth || 256)) throw new ServiceError('RDL_INVALID', 'RDL XML nesting is too deep', 413);

  let parsed;
  try {
    parsed = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', removeNSPrefix: false, parseTagValue: false, trimValues: false, processEntities: false }).parse(xml);
  } catch {
    throw new ServiceError('RDL_INVALID', 'RDL XML is malformed');
  }
  const report = parsed.Report;
  if (!report) throw new ServiceError('RDL_INVALID', 'RDL root Report element is missing');
  const capabilities = inspectRdlCapabilities(parsed, namespace);
  const section = firstDefined(asArray(report.ReportSections?.ReportSection)[0], report);
  const defaultFontFamily = styleTextValue(report['df:DefaultFontFamily'], 'Arial');
  const page = section.Page || report.Page || {};
  const body = section.Body || report.Body;
  if (!body) throw new ServiceError('RDL_INVALID', 'RDL Body element is missing');
  const items = childEntries(body.ReportItems, ITEM_NAMES).map(([type, item]) => parseItem(type, item, defaultFontFamily));
  const datasets = parseDatasets(report);
  attachDatasetFields(items, new Map(datasets.map((dataset) => [dataset.name, dataset])));
  const parameters = parseParameters(report);
  const renderingDatasets = collectItemDatasets(items);
  const parameterDatasets = new Set(parameters.map((parameter) => parameter.lookupDataset).filter(Boolean));
  const embeddedImages = Object.fromEntries(asArray(report.EmbeddedImages?.EmbeddedImage).map((image) => [image['@_Name'], {
    mimeType: textValue(image.MIMEType, 'image/png'),
    data: textValue(image.ImageData),
  }]));
  const pageHeader = page.PageHeader ? {
    height: toPoints(textValue(page.PageHeader.Height, '0pt')),
    printOnFirstPage: parseBoolean(page.PageHeader.PrintOnFirstPage, true),
    printOnLastPage: parseBoolean(page.PageHeader.PrintOnLastPage, true),
    items: childEntries(page.PageHeader.ReportItems, ITEM_NAMES).map(([type, item]) => parseItem(type, item, defaultFontFamily)),
  } : null;
  const pageFooter = page.PageFooter ? {
    height: toPoints(textValue(page.PageFooter.Height, '0pt')),
    printOnFirstPage: parseBoolean(page.PageFooter.PrintOnFirstPage, true),
    printOnLastPage: parseBoolean(page.PageFooter.PrintOnLastPage, true),
    items: childEntries(page.PageFooter.ReportItems, ITEM_NAMES).map(([type, item]) => parseItem(type, item, defaultFontFamily)),
  } : null;
  const allItems = [...items, ...(pageHeader?.items || []), ...(pageFooter?.items || [])];
  // Report-level variables first, then group variables (group wins on a name clash), producing one
  // name→expression map threaded to the evaluator via globals.variables.
  const variables = {};
  for (const variable of parseVariables(report.Variables)) variables[variable.name] = variable.value;
  const collectTablixVariables = (list) => {
    for (const item of list || []) {
      if (item.type === 'Tablix') { collectGroupVariables(item.rowMembers, variables); collectGroupVariables(item.columnMembers, variables); }
      if (item.items) collectTablixVariables(item.items);
    }
  };
  collectTablixVariables(allItems);
  const subreports = collectSubreports(allItems);
  const unsupported = [...new Set([
    ...collectUnsupportedExpressions(allItems, collectUnsupported(allItems)),
    ...capabilities.rejected.map((entry) => `RdlPath:${entry.path}`),
    ...capabilities.expressions.rejected.map((entry) => entry.name === 'Code.*' ? 'CustomCode' : `ExpressionFunction:${entry.name}`),
    ...(Number.parseInt(textValue(page.NumberOfColumns, '1'), 10) > 1 ? ['PageColumns'] : []),
  ])];

  const reportLanguage = textValue(report.Language, '').trim();
  const model = {
    namespace,
    name: textValue(report.Description, 'RDL Report'),
    // Report `Language` (culture). A literal is canonicalized now; an `=expression` is resolved at render
    // time from `language`. Null culture means no declared Language — the formatter keeps its legacy defaults.
    language: reportLanguage || null,
    culture: canonicalizeCulture(reportLanguage),
    identity: {
      name: textValue(report.Description, 'RDL Report'),
      definitionSha256: createHash('sha256').update(xml).digest('hex'),
    },
    page: {
      width: toPoints(textValue(page.PageWidth, '8.27in')),
      height: toPoints(textValue(page.PageHeight, '11.69in')),
      marginTop: toPoints(textValue(page.TopMargin, '0in')),
      marginRight: toPoints(textValue(page.RightMargin, '0in')),
      marginBottom: toPoints(textValue(page.BottomMargin, '0in')),
      marginLeft: toPoints(textValue(page.LeftMargin, '0in')),
      columns: Number.parseInt(textValue(page.NumberOfColumns, '1'), 10) || 1,
      columnSpacing: toPoints(textValue(page.ColumnSpacing, '0pt')),
      header: pageHeader,
      footer: pageFooter,
    },
    body: { height: toPoints(textValue(body.Height, '0pt')), width: toPoints(textValue(firstDefined(body.Width, section.Width), page.PageWidth || '8.27in')), items },
    datasets,
    parameters,
    variables,
    subreports,
    renderingDatasets: [...renderingDatasets],
    parameterDatasets: [...parameterDatasets],
    embeddedImages,
    defaultFontFamily,
    fonts: [...new Set([defaultFontFamily, ...[...xml.matchAll(/<FontFamily>([^<]+)<\/FontFamily>/g)].map((match) => match[1].trim())])],
    // Field names the report actually reads, from every `Fields!Name.Value` in every expression. SSRS does
    // not require a declared field to exist in the result set — it only matters where the report reads one
    // — so a dataset routinely declares columns its query never returns (a query edited after the fields
    // were generated, a shared dataset trimmed for one report). Demanding all of them rejects reports SSRS
    // renders. Dynamic access, `Fields(name).Value`, cannot be resolved statically and is not captured; a
    // field reached only that way behaves as SSRS does and yields nothing rather than failing the request.
    referencedFields: [...new Set([...xml.matchAll(/Fields!([A-Za-z_][A-Za-z0-9_]*)/g)].map((match) => match[1]))],
    features: {
      textboxes: (xml.match(/<Textbox\b/g) || []).length,
      tablixes: (xml.match(/<Tablix\b/g) || []).length,
      rectangles: (xml.match(/<Rectangle\b/g) || []).length,
      lines: (xml.match(/<Line\b/g) || []).length,
      images: (xml.match(/<Image\b/g) || []).length,
      groups: (xml.match(/<Group\b/g) || []).length,
      pageBreaks: (xml.match(/<PageBreak\b/g) || []).length,
      subreports: subreports.length,
    },
    unsupported,
    capabilities,
  };
  return model;
}

export function analyzeParsedRdl(model) {
  return {
    namespace: model.namespace,
    identity: model.identity,
    page: {
      width: model.page.width,
      height: model.page.height,
      marginTop: model.page.marginTop,
      marginRight: model.page.marginRight,
      marginBottom: model.page.marginBottom,
      marginLeft: model.page.marginLeft,
      header: model.page.header ? {
        height: model.page.header.height,
        printOnFirstPage: model.page.header.printOnFirstPage,
        printOnLastPage: model.page.header.printOnLastPage,
      } : null,
      footer: model.page.footer ? {
        height: model.page.footer.height,
        printOnFirstPage: model.page.footer.printOnFirstPage,
        printOnLastPage: model.page.footer.printOnLastPage,
      } : null,
    },
    parameters: model.parameters,
    datasets: model.datasets.map((dataset) => ({ ...dataset, requiredForRendering: model.renderingDatasets.includes(dataset.name), parameterOnly: model.parameterDatasets.includes(dataset.name) && !model.renderingDatasets.includes(dataset.name) })),
    fonts: model.fonts,
    features: model.features,
    subreports: model.subreports,
    compatible: model.unsupported.length === 0,
    blockingErrors: model.unsupported.map((type) => ({ code: 'UNSUPPORTED_FEATURE', feature: type })),
    capabilities: model.capabilities,
  };
}

export function analyzeRdl(input, limits) {
  return analyzeParsedRdl(parseRdl(input, limits));
}
