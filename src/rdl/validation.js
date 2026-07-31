import { ServiceError } from '../errors.js';
import { evaluateExpression, formatValue } from './expression.js';
import { isBoundField, normalizeRowFields, rowKeyFor } from './fields.js';
import { flattenCellItems, RENDERABLE_CELL_ITEMS } from './helpers.js';
import { normalizeDisplayText, renderMarkupText } from './text.js';

function hasValue(value) {
  return value !== undefined && value !== null && value !== '' && (!Array.isArray(value) || value.length > 0);
}

function validateParameterType(parameter, value) {
  const values = parameter.multiValue ? value : [value];
  if (parameter.multiValue && !Array.isArray(value)) throw new ServiceError('PARAMETER_INVALID', `Parameter ${parameter.name} must be an array`);
  for (const item of values) {
    if (item === null && parameter.nullable) continue;
    if (parameter.dataType === 'Integer' && !Number.isInteger(Number(item))) throw new ServiceError('PARAMETER_INVALID', `Parameter ${parameter.name} must be an integer`);
    if (parameter.dataType === 'Float' && !Number.isFinite(Number(item))) throw new ServiceError('PARAMETER_INVALID', `Parameter ${parameter.name} must be numeric`);
    if (parameter.dataType === 'Boolean' && !['true', 'false'].includes(String(item).toLowerCase())) throw new ServiceError('PARAMETER_INVALID', `Parameter ${parameter.name} must be boolean`);
    if (parameter.dataType === 'DateTime' && Number.isNaN(new Date(item).getTime())) throw new ServiceError('PARAMETER_INVALID', `Parameter ${parameter.name} must be an ISO date`);
  }
}

export function resolveParameterValues(definitions, parameters = {}) {
  const resolved = { ...parameters };
  for (const parameter of definitions || []) {
    const defaultValue = parameter.multiValue ? parameter.defaultValues : parameter.defaultValues[0];
    const value = hasValue(parameters[parameter.name]) ? parameters[parameter.name] : defaultValue;
    if (!hasValue(value) && !parameter.nullable && !parameter.allowBlank) {
      throw new ServiceError('PARAMETER_INVALID', `Required parameter is missing: ${parameter.name}`);
    }
    if (hasValue(value)) validateParameterType(parameter, value);
    resolved[parameter.name] = value;
  }
  return resolved;
}

function canonicalParameterValue(parameter, value) {
  if (parameter.multiValue) return (value || []).map((entry) => canonicalParameterValue({ ...parameter, multiValue: false }, entry));
  if (value === null || value === undefined) return null;
  if (parameter.dataType === 'Integer' || parameter.dataType === 'Float') return Number(value);
  if (parameter.dataType === 'Boolean') return String(value).toLowerCase() === 'true';
  if (parameter.dataType === 'DateTime') return new Date(value).toISOString();
  return String(value);
}

export function parameterSignature(definitions, parameters = {}) {
  return JSON.stringify((definitions || []).map((parameter) => [
    parameter.name,
    canonicalParameterValue(parameter, parameters[parameter.name]),
  ]));
}

export function validateRenderInput(model, request, limits) {
  if (model.unsupported.length > 0) throw new ServiceError(
    'UNSUPPORTED_FEATURE',
    'RDL contains unsupported elements, properties, or expressions',
    400,
    { features: model.unsupported.slice(0, 100) },
  );
  const resolvedParameters = resolveParameterValues(model.parameters, request.parameters || {});

  const datasets = request.datasets || {};
  let totalRows = 0;
  for (const dataset of model.datasets) {
    if (!model.renderingDatasets.includes(dataset.name)) continue;
    const rows = datasets[dataset.name];
    if (!Array.isArray(rows)) throw new ServiceError('DATASET_MISSING', `Required dataset is missing: ${dataset.name}`);
    totalRows += rows.length;
    if (totalRows > limits.maxRows) throw new ServiceError('RDL_INVALID', `Dataset rows exceed the ${limits.maxRows} row limit`, 413);
    for (const [index, row] of rows.entries()) {
      if (!row || typeof row !== 'object' || Array.isArray(row)) throw new ServiceError('FIELD_MISSING', `Dataset ${dataset.name} row ${index} must be an object`);
      for (const field of dataset.fields) {
        // A calculated field is computed from its siblings, never supplied by the caller (see rdl/fields.js).
        if (!isBoundField(field)) continue;
        if (rowKeyFor(row, field.dataField) === undefined) {
          throw new ServiceError(
            'FIELD_MISSING',
            `Dataset ${dataset.name} row ${index} is missing field ${field.dataField}`,
            400,
            // Column names, not row values: without them this error cannot be diagnosed without the RDL
            // and the query side by side, which is exactly the position it used to leave callers in.
            { dataset: dataset.name, field: field.dataField, availableFields: Object.keys(row).slice(0, 100) },
          );
        }
      }
    }
  }
  return { totalRows, parameters: resolvedParameters };
}

// Kept as a thin alias so both call sites in this module read as before; the shared implementation lives
// in rdl/fields.js because the renderers normalize the same rows and must agree on what a field means.
function normalizeFields(row, definitions = [], context = {}) {
  return normalizeRowFields(row, definitions, context);
}

function evaluateItemText(item, context) {
  if (!item.paragraphs) {
    const value = evaluateExpression(item.value, context);
    return item.style?.format ? formatValue(value, item.style.format) : value;
  }
  return normalizeDisplayText(item.paragraphs.map((runs) => runs.map((run) => {
    const definition = run && typeof run === 'object' ? run : { value: run, markupType: 'None' };
    const value = /^constant$/i.test(String(definition.evaluationMode || 'Auto'))
      ? definition.value
      : evaluateExpression(definition.value, context);
    if (value === null || value === undefined) return '';
    const format = expressionValue(definition.style?.format ?? item.style?.format, context);
    const formatted = format ? String(formatValue(value, format)) : String(value);
    return renderMarkupText(formatted, definition.markupType);
  }).join('')).join('\n'));
}

function expressionValue(value, context) {
  return typeof value === 'string' && value.trim().startsWith('=') ? evaluateExpression(value, context) : value;
}

function evalHidden(expression, context) {
  const result = evaluateExpression(expression, context);
  return result === true || String(result).toLowerCase() === 'true';
}

function collectMemberGroupFilters(members, target = []) {
  for (const member of members || []) {
    if (member.group?.filters?.length) target.push(...member.group.filters);
    collectMemberGroupFilters(member.children, target);
  }
  return target;
}

export function filterMatches(filter, fields, parameters, globals, dataset, datasets) {
  const context = { fields, parameters, globals, dataset, datasets };
  const actual = expressionValue(filter.expression, context);
  const expected = filter.values.map((value) => expressionValue(value, context));
  const operator = String(filter.operator || 'Equal').toLowerCase();
  if (operator === 'equal') return actual === expected[0] || String(actual) === String(expected[0]);
  if (operator === 'notequal') return actual !== expected[0] && String(actual) !== String(expected[0]);
  if (operator === 'greaterthan') return actual > expected[0];
  if (operator === 'greaterthanorequal') return actual >= expected[0];
  if (operator === 'lessthan') return actual < expected[0];
  if (operator === 'lessthanorequal') return actual <= expected[0];
  if (operator === 'between') return actual >= expected[0] && actual <= expected[1];
  if (operator === 'in') return expected.some((value) => actual === value || String(actual) === String(value));
  if (operator === 'like') {
    const pattern = String(expected[0] ?? '').replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
    return new RegExp(`^${pattern}$`, 'i').test(String(actual ?? ''));
  }
  throw new ServiceError('UNSUPPORTED_FEATURE', `Unsupported tablix filter operator: ${filter.operator}`);
}

function collectMemberSorts(members, target = []) {
  for (const member of members || []) {
    target.push(...(member.sortExpressions || []));
    collectMemberSorts(member.children, target);
  }
  return target;
}

function memberGroupsByName(members, target = new Map()) {
  for (const member of members || []) {
    if (member.group?.name) target.set(member.group.name, member.group);
    memberGroupsByName(member.children, target);
  }
  return target;
}

function dataRegionScopes(tablix, rows) {
  const scopes = {};
  if (tablix.datasetName) scopes[tablix.datasetName] = rows;
  if (tablix.name) scopes[tablix.name] = rows;
  return scopes;
}

function groupValue(group, fields, parameters, globals, dataset, datasets, rowIndex) {
  if (!group) return null;
  if (!group.expressions?.length) return `row:${rowIndex}`;
  const context = { fields, parameters, globals, dataset, datasets };
  return group.expressions.map((expression) => expressionValue(expression, context));
}

function scopeKey(scope, tablix, groups, fields, parameters, globals, dataset, datasets, rowIndex, memberPath = []) {
  if (!scope || scope === tablix.datasetName) return `dataset:${tablix.datasetName}`;
  // HideDuplicates is scoped to a containing GROUP INSTANCE, not merely to the target group's
  // expression value. The same child value (especially Nothing/null) under two different parent
  // instances represents two distinct scopes in SSRS and must display once in each. Include every
  // containing group from the current hierarchy path through the named scope so nested instances do
  // not collapse into one global `<group>:<value>` bucket.
  const instancePath = [];
  for (const [depth, member] of memberPath.entries()) {
    const group = member.group;
    if (!group) continue;
    instancePath.push([
      group.name || `depth:${depth}`,
      groupValue(group, fields, parameters, globals, dataset, datasets, rowIndex),
    ]);
    if (group.name === scope) return `${scope}:${JSON.stringify(instancePath)}`;
  }
  const group = groups.get(scope);
  if (!group) return `scope:${scope}`;
  return `${scope}:${JSON.stringify(groupValue(group, fields, parameters, globals, dataset, datasets, rowIndex))}`;
}

function headerDescriptors(path = []) {
  const descriptors = [];
  const groups = [];
  for (const member of path) {
    if (member.group) groups.push(member.group);
    if (member.header) descriptors.push({ member, header: member.header, groups: [...groups] });
  }
  return descriptors;
}

function descriptorKey(descriptor, fields, parameters, globals, dataset, datasets, rowIndex) {
  if (descriptor.groups.length === 0) return `row:${rowIndex}`;
  return JSON.stringify(descriptor.groups.map((group) => groupValue(group, fields, parameters, globals, dataset, datasets, rowIndex)));
}

function headerSpans(descriptors, widths) {
  if (descriptors.length === 0 || widths.length === 0) return [];
  if (descriptors.length === 1) return [widths.length];
  // When the deepest row-header path has exactly one descriptor per header column, each maps 1:1.
  // Assign span 1 deterministically instead of relying on the width-tolerance heuristic, which is
  // brittle when adjacent header sizes are close.
  if (descriptors.length === widths.length) return descriptors.map(() => 1);
  const spans = [];
  let columnIndex = 0;
  for (const [index, descriptor] of descriptors.entries()) {
    const remainingHeaders = descriptors.length - index - 1;
    const maximumSpan = Math.max(1, widths.length - columnIndex - remainingHeaders);
    let span = 1;
    let width = widths[columnIndex] || 0;
    while (span < maximumSpan && width + (widths[columnIndex + span] || 0) <= descriptor.header.size + 0.5) {
      width += widths[columnIndex + span] || 0;
      span += 1;
    }
    spans.push(span);
    columnIndex += span;
  }
  return spans;
}

function assertMaterializedGrid(rows, columnCount, tablixName) {
  if (columnCount === 0) return;
  const occupied = new Array(columnCount).fill(0);
  for (const [rowIndex, row] of rows.entries()) {
    const covered = occupied.map((remaining) => remaining > 0);
    let columnIndex = 0;
    for (const cell of row.cells || []) {
      while (columnIndex < columnCount && covered[columnIndex]) columnIndex += 1;
      const colSpan = Math.max(1, cell.colSpan || 1);
      if (columnIndex + colSpan > columnCount) {
        throw new ServiceError('RDL_INVALID', `Tablix ${tablixName || 'unnamed'} row ${rowIndex + 1} exceeds its declared column grid`);
      }
      for (let offset = 0; offset < colSpan; offset += 1) {
        if (covered[columnIndex + offset]) {
          throw new ServiceError('RDL_INVALID', `Tablix ${tablixName || 'unnamed'} row ${rowIndex + 1} contains overlapping cells`);
        }
        covered[columnIndex + offset] = true;
        if ((cell.rowSpan || 1) > 1) occupied[columnIndex + offset] = Math.max(occupied[columnIndex + offset], cell.rowSpan || 1);
      }
      columnIndex += colSpan;
    }
    if (covered.some((value) => !value)) {
      throw new ServiceError('RDL_INVALID', `Tablix ${tablixName || 'unnamed'} row ${rowIndex + 1} does not cover its declared column grid`);
    }
    for (let index = 0; index < occupied.length; index += 1) occupied[index] = Math.max(0, occupied[index] - 1);
  }
}

function itemValue(item, context) {
  return item.type === 'Textbox' ? evaluateItemText(item, context) : '';
}

function materializedCell(rawCell, context, duplicateState, duplicatePrefix, scopeResolver) {
  // A cell whose content is only container Rectangles has no border-bearing item of its own, so its edges
  // keep resolving from the tablix style (as they did before the Rectangle was flattened away). Without
  // this the inner textbox's own Border=None would silently drop the cell's grid lines.
  const containerWrapped = (rawCell.items || []).length > 0 && rawCell.items.every((item) => item.type === 'Rectangle');
  const cell = { ...rawCell, items: flattenCellItems(rawCell.items), containerWrapped };
  // Fail closed rather than silently drawing an empty cell. The capability catalogue classifies element
  // NAMES, so an element can be SUPPORTED at body level yet unhandled in a cell — exactly how a
  // Rectangle-wrapped cell rendered blank with compatible:true. analyze() reports this first; this is the
  // defence-in-depth if a caller reaches materialization without it.
  const unrenderable = cell.items.find((item) => !RENDERABLE_CELL_ITEMS.has(item.type));
  if (unrenderable) {
    throw new ServiceError('UNSUPPORTED_FEATURE', `Tablix cell content is not supported: ${unrenderable.type}`);
  }
  const values = cell.items.map((item, itemIndex) => {
    if (item.type === 'Tablix' || item.type === 'Subreport') return '';
    const value = itemValue(item, context);
    if (item.type !== 'Textbox' || !item.hideDuplicates) return value;
    const key = `${duplicatePrefix}:${item.name || itemIndex}`;
    const current = { scope: scopeResolver(item.hideDuplicates), value: String(value ?? '') };
    const previous = duplicateState.get(key);
    duplicateState.set(key, current);
    return previous && previous.scope === current.scope && previous.value === current.value ? '' : value;
  });
  const nestedTablixes = cell.items.filter((item) => item.type === 'Tablix').map((item) => {
    const depth = (context.nestedTablixDepth || 0) + 1;
    if (depth > 8) {
      throw new ServiceError('UNSUPPORTED_FEATURE', 'Nested tablix depth exceeds the supported limit of 8');
    }
    const parentDatasetName = context.tablixDatasetName;
    if (item.datasetName && parentDatasetName && item.datasetName !== parentDatasetName) {
      throw new ServiceError(
        'UNSUPPORTED_FEATURE',
        `Nested tablix ${item.name || 'unnamed'} must use the containing tablix dataset`,
      );
    }
    const scopedRows = context.nestedDataset || context.dataset || [];
    const nestedGlobals = { ...context.globals, __nestedTablixDepth: depth };
    return {
      item,
      rows: materializeTablixRows(item, scopedRows, context.parameters, nestedGlobals, context.datasets),
      columns: materializeTablixColumns(item, scopedRows, context.parameters, nestedGlobals, context.datasets),
      parameters: context.parameters,
      datasets: context.datasets,
      globals: nestedGlobals,
    };
  });
  for (const item of cell.items.filter((candidate) => candidate.type === 'Subreport')) {
    if (evalHidden(item.hidden, context)) continue;
    const resolved = item.resolvedSubreport;
    if (!resolved?.model || !resolved.instancesBySignature) {
      throw new ServiceError('UNSUPPORTED_FEATURE', `Subreport is not bundled: ${item.reportName || item.name || 'unnamed'}`);
    }
    const suppliedParameters = Object.fromEntries((item.parameters || []).map((parameter) => [
      parameter.name,
      evaluateExpression(parameter.value, context),
    ]));
    const childParameters = resolveParameterValues(resolved.model.parameters, suppliedParameters);
    const signature = parameterSignature(resolved.model.parameters, childParameters);
    const instance = resolved.instancesBySignature.get(signature);
    if (!instance) {
      throw new ServiceError(
        'DATASET_MISSING',
        `Subreport invocation data is missing: ${item.reportName || item.name || 'unnamed'}`,
      );
    }
    const childDatasets = Object.fromEntries(resolved.model.datasets.map((dataset) => [
      dataset.name,
      (instance.datasets[dataset.name] || []).map((row) => normalizeFields(row, dataset.fields)),
    ]));
    const childGlobals = {
      ...context.globals,
      ReportName: resolved.reportName,
      variables: resolved.model.variables || {},
      __nestedTablixDepth: (context.nestedTablixDepth || 0) + 1,
    };
    for (const childItem of resolved.model.body.items) {
      const nestedItem = {
        ...childItem,
        left: (item.left || 0) + (childItem.left || 0),
        top: (item.top || 0) + (childItem.top || 0),
      };
      const rawRows = instance.datasets[childItem.datasetName] || [];
      nestedTablixes.push({
        item: nestedItem,
        rows: materializeTablixRows(nestedItem, rawRows, instance.parameters, childGlobals, childDatasets),
        columns: materializeTablixColumns(nestedItem, rawRows, instance.parameters, childGlobals, childDatasets),
        parameters: instance.parameters,
        datasets: childDatasets,
        globals: childGlobals,
        subreport: true,
      });
    }
  }
  return {
    ...cell,
    values,
    nestedTablixes,
    // Preserve the exact materialization scope for expression-backed cell styles. Matrix values are
    // evaluated at the row∩column intersection; evaluating their background, borders, or font later
    // with only the row representative would incorrectly give every cell in that row the same style.
    fields: context.fields || {},
    scopeDataset: context.dataset || [],
    scopes: context.scopes || {},
    hidden: cell.items.some((item) => {
      const result = evaluateExpression(item.hidden, context);
      return result === true || String(result).toLowerCase() === 'true';
    }),
  };
}

export function prepareTablixData(tablix, rows, parameters, globals = {}, datasets = {}) {
  const normalized = (Array.isArray(rows) ? rows : []).map((row) => normalizeFields(row, tablix.datasetFields, { parameters, globals, datasets }));
  // Apply both tablix-level filters and row-group member filters (a group filter removes rows that fall
  // outside the group, e.g. a series/category exclusion).
  const activeFilters = [...(tablix.filters || []), ...collectMemberGroupFilters(tablix.rowMembers)];
  const filtered = normalized.filter((fields) => activeFilters.every((filter) => filterMatches(filter, fields, parameters, globals, normalized, datasets)));
  const sorts = [...(tablix.sortExpressions || []), ...collectMemberSorts(tablix.rowMembers)];
  if (sorts.length === 0) return filtered;
  return filtered.map((fields, index) => ({ fields, index })).sort((left, right) => {
    for (const sort of sorts) {
      const leftValue = expressionValue(sort.value, { fields: left.fields, parameters, globals, dataset: filtered, datasets });
      const rightValue = expressionValue(sort.value, { fields: right.fields, parameters, globals, dataset: filtered, datasets });
      const compared = typeof leftValue === 'string' || typeof rightValue === 'string'
        ? String(leftValue ?? '').localeCompare(String(rightValue ?? ''))
        : Number(leftValue ?? 0) - Number(rightValue ?? 0);
      if (compared !== 0) return /^desc/i.test(sort.direction) ? -compared : compared;
    }
    return left.index - right.index;
  }).map(({ fields }) => fields);
}

// Precomputes, for every named/expression row group, a map from group-key -> the rows in that group
// instance. Reused to give each cell the correct aggregate scope: a Sum() in a group scope aggregates
// only that group's rows, not the whole dataset.
function buildGroupInstances(members, sourceRows, parameters, globals, datasets, target = new Map()) {
  for (const member of members || []) {
    const group = member.group;
    if (group?.expressions?.length) {
      const byKey = new Map();
      sourceRows.forEach((fields, rowIndex) => {
        const key = JSON.stringify(groupValue(group, fields, parameters, globals, sourceRows, datasets, rowIndex));
        if (!byKey.has(key)) byKey.set(key, []);
        byKey.get(key).push(fields);
      });
      target.set(group, byKey);
    }
    buildGroupInstances(member.children, sourceRows, parameters, globals, datasets, target);
  }
  return target;
}

// Advanced tablix features (group header/footer subtotal rows, recursive/parent groups, and dynamic
// column groups / matrix) all require a per-instance recursive walk rather than the flat per-template
// loop below. Detection is deliberately narrow so that every static/detail/merged-header tablix that
// does NOT use these constructs keeps taking the unchanged flat path and materializes byte-identically.
export function needsAdvancedMaterialization(tablix) {
  if (tablix.hasColumnGroups) return true;
  const hasParent = (members) => (members || []).some((member) => member.group?.parent || hasParent(member.children));
  if (hasParent(tablix.rowMembers)) return true;
  // A group header/footer row is a STATIC leaf member (no Group element) nested inside an ancestor
  // group. This also covers SSRS's group-instance row-header chains: the recursive materializer must
  // emit those once per group instance rather than once per source row.
  return (tablix.rowMemberPaths || []).some((path) => {
    const leaf = path[path.length - 1];
    if (leaf?.group) return false;
    return path.some((member) => member.group?.expressions?.length);
  });
}

export function materializeTablixRows(tablix, rows, parameters, globals = {}, datasets = {}) {
  const sourceRows = prepareTablixData(tablix, rows, parameters, globals, datasets);
  if (needsAdvancedMaterialization(tablix)) return materializeAdvancedRows(tablix, sourceRows, parameters, globals, datasets);
  const memberPaths = tablix.rowMemberPaths || [];
  const groups = memberGroupsByName(tablix.rowMembers);
  const groupInstances = buildGroupInstances(tablix.rowMembers, sourceRows, parameters, globals, datasets);
  // Scope map for a cell: named group instances containing this row + the innermost group's rows (the
  // SSRS default no-scope aggregate scope). Falls back to the whole dataset when there is no row group.
  const scopeContext = (path, fields, rowIndex) => {
    const scopes = dataRegionScopes(tablix, sourceRows);
    let innermost = sourceRows;
    let nestedDataset = sourceRows;
    for (const member of path || []) {
      const group = member.group;
      if (!group) continue;
      if (!group.expressions?.length) {
        // A Group with no expressions is the RDL detail scope: one group instance per current row.
        // A nested data region sees that one-row intersection, while ordinary no-scope aggregates retain
        // the nearest named parent group scope (matching SSRS aggregate semantics).
        nestedDataset = [fields];
        if (group.name) scopes[group.name] = nestedDataset;
        continue;
      }
      const key = JSON.stringify(groupValue(group, fields, parameters, globals, sourceRows, datasets, rowIndex));
      const instance = groupInstances.get(group)?.get(key) || sourceRows;
      if (group.name) scopes[group.name] = instance;
      innermost = instance;
      nestedDataset = instance;
    }
    return { scopes, dataset: innermost, nestedDataset };
  };
  // Row groups that page-break between instances (Start/End/Between/StartAndEnd all read as "start each
  // new instance on a fresh page" for a top-down table). The first row of every new instance is tagged.
  const breakGroups = [];
  const collectBreakGroups = (members) => {
    for (const member of members || []) {
      if (member.group?.expressions?.length && /^(Start|End|Between|StartAndEnd)$/i.test(String(member.group.pageBreak || 'None'))) breakGroups.push(member.group);
      collectBreakGroups(member.children);
    }
  };
  collectBreakGroups(tablix.rowMembers);
  const lastBreakKey = new Map();
  // Static versus dynamic is declared by the TablixMember hierarchy, not inferred from expressions in
  // its cells. Static members may legitimately reference Fields (for example, to provide an IIF fallback
  // when the dataset is empty). A Group element, including a detail group without expressions, makes the
  // corresponding member path dynamic.
  const memberPathIsDynamic = (path) => (path || []).some((member) => member.group);
  const firstDynamic = tablix.rows.findIndex((row, index) => memberPathIsDynamic(memberPaths[index]));
  // Hard rule: a tablix's static column-header rows (everything above the first dynamic/detail row) repeat on
  // every page, regardless of whether the RDL sets RepeatColumnHeaders / a member's RepeatOnNewPage. SSRS
  // leaves this opt-in, but a multi-page table whose header does not repeat is unreadable, so the service
  // always repeats it. Both renderers key page-repeated headers off `isHeader`, so this drives PDF and DOCX
  // (w:tblHeader) alike. When there is no dynamic row (firstDynamic < 0) there is nothing to repeat.
  const repeatingHeaderCount = firstDynamic > 0 ? firstDynamic : 0;
  const output = [];
  const duplicateState = new Map();
  for (const [index, row] of tablix.rows.entries()) {
    const path = memberPaths[index] || [];
    const descriptors = headerDescriptors(path);
    const spans = headerSpans(descriptors, tablix.rowHeaderColumns || []);
    const dynamic = memberPathIsDynamic(path);
    // Static members remain visible on an empty data region unless the RDL explicitly opts into
    // HideIfNoRows. Their Fields references evaluate against the empty context as Nothing.
    if (!dynamic && sourceRows.length === 0 && path.some((member) => member.hideIfNoRows)) continue;
    const contexts = dynamic ? sourceRows : [sourceRows[0] || {}];
    for (const [rowIndex, fields] of contexts.entries()) {
      const { scopes, dataset, nestedDataset } = dynamic
        ? scopeContext(path, fields, rowIndex)
        : { scopes: dataRegionScopes(tablix, sourceRows), dataset: sourceRows, nestedDataset: sourceRows };
      const context = {
        fields,
        parameters,
        globals,
        dataset,
        outermostDataset: sourceRows,
        nestedDataset,
        datasets,
        scopes,
        tablixDatasetName: tablix.datasetName,
        tablixName: tablix.name,
        nestedTablixDepth: globals.__nestedTablixDepth || 0,
      };
      // Honor static Visibility.Hidden on the row and on any member in its path (a hidden group/row is
      // not rendered). Interactive drill-down toggles are ignored — everything renders expanded.
      if (evalHidden(row.hidden, context) || path.some((member) => evalHidden(member.hidden, context))) continue;
      const resolveScope = (scope) => scopeKey(scope, tablix, groups, fields, parameters, globals, sourceRows, datasets, rowIndex, path);
      const rowHeaders = [];
      for (const [descriptorIndex, descriptor] of descriptors.entries()) {
        const key = descriptorKey(descriptor, fields, parameters, globals, sourceRows, datasets, rowIndex);
        const previousFields = rowIndex > 0 ? contexts[rowIndex - 1] : null;
        const previousKey = previousFields
          ? descriptorKey(descriptor, previousFields, parameters, globals, sourceRows, datasets, rowIndex - 1)
          : null;
        if (dynamic && previousKey === key) continue;
        let rowSpan = 1;
        if (dynamic) {
          while (rowIndex + rowSpan < contexts.length && descriptorKey(
            descriptor, contexts[rowIndex + rowSpan], parameters, globals, sourceRows, datasets, rowIndex + rowSpan,
          ) === key) rowSpan += 1;
        }
        rowHeaders.push({
          ...materializedCell(descriptor.header.cell, context, duplicateState, `header:${index}:${descriptorIndex}`, resolveScope),
          colSpan: spans[descriptorIndex] || 1,
          rowSpan,
          isRowHeader: true,
        });
      }
      let pageBreakBefore = false;
      if (dynamic) {
        for (const group of breakGroups) {
          const groupKey = JSON.stringify(groupValue(group, fields, parameters, globals, sourceRows, datasets, rowIndex));
          const previous = lastBreakKey.get(group);
          if (previous !== undefined && previous !== groupKey) pageBreakBefore = true;
          lastBreakKey.set(group, groupKey);
        }
      }
      output.push({
        sourceIndex: index,
        isHeader: index < repeatingHeaderCount,
        isStatic: !dynamic,
        role: dynamic ? 'detail' : 'static',
        indentLevel: 0,
        pageBreakBefore,
        fields,
        scopeDataset: dataset,
        height: row.height,
        keepTogether: path.some((member) => member.keepTogether) || row.cells.some((cell) => cell.items.some((item) => item.keepTogether)),
        cells: [
          ...rowHeaders,
          ...row.cells.map((cell, cellIndex) => materializedCell(cell, context, duplicateState, `body:${index}:${cellIndex}`, resolveScope)),
        ],
      });
    }
  }
  return output;
}

// Rows in `a` that are also present (by object identity) in `b`. Both are always slices of the same
// prepared sourceRows array, so identity intersection gives a matrix body cell its row∩column scope.
function intersectRows(a, b) {
  const set = new Set(b);
  return a.filter((row) => set.has(row));
}

// Ordered distinct group instances for a member's group over `rows`, preserving first-appearance
// order. Rows are SLICED (never cloned) so positional expression functions keep working.
function partitionByGroup(group, rows, parameters, globals, sourceRows, datasets) {
  const byKey = new Map();
  const order = [];
  rows.forEach((fields, index) => {
    const key = JSON.stringify(groupValue(group, fields, parameters, globals, sourceRows, datasets, index));
    if (!byKey.has(key)) { byKey.set(key, { key, rows: [] }); order.push(key); }
    byKey.get(key).rows.push(fields);
  });
  return order.map((key) => byKey.get(key));
}

// Column axis of a matrix: ordered leaf column instances (each = the rows for one distinct column-group
// value) plus, per hierarchy depth, the header cells to render above the body. Static-column tablixes
// never reach here (gated by hasColumnGroups).
function matrixColumnData(tablix, sourceRows, parameters, globals, datasets) {
  const leaves = [];
  const headerRowsByLevel = [];
  // Each leaf in the design-time column hierarchy owns exactly one TablixBody column. A dynamic
  // leaf repeats that column for every group instance; a neighbouring static leaf remains a single
  // column. Do not repeat the complete body-column template for every dynamic instance: mixed
  // static/dynamic matrices (for example, a row-label column followed by a dynamic value column)
  // would otherwise multiply both columns and destroy the declared grid geometry.
  const templateIndexByMember = new Map(
    (tablix.columnMemberPaths || []).map((path, index) => [path[path.length - 1], index]),
  );
  const walk = (members, incomingRows, scopes, depth) => {
    let leafCount = 0;
    for (const member of members || []) {
      const group = member.group;
      if (group?.expressions?.length) {
        for (const instance of partitionByGroup(group, incomingRows, parameters, globals, sourceRows, datasets)) {
          const childScopes = group.name ? { ...scopes, [group.name]: instance.rows } : scopes;
          const childLeafCount = member.children?.length
            ? walk(member.children, instance.rows, childScopes, depth + 1)
            : (leaves.push({
              rows: instance.rows,
              scopes: childScopes,
              templateIndex: templateIndexByMember.get(member) ?? 0,
            }), 1);
          if (member.header) {
            (headerRowsByLevel[depth] ||= []).push({
              cell: member.header.cell,
              colSpan: childLeafCount,
              fields: instance.rows[0] || {},
              scopes: childScopes,
              rows: instance.rows,
            });
          }
          leafCount += childLeafCount;
        }
      } else if (member.children?.length) {
        leafCount += walk(member.children, incomingRows, scopes, depth);
      } else {
        leaves.push({
          rows: incomingRows,
          scopes,
          templateIndex: templateIndexByMember.get(member) ?? 0,
        });
        leafCount += 1;
      }
    }
    return leafCount;
  };
  walk(tablix.columnMembers, sourceRows, dataRegionScopes(tablix, sourceRows), 0);
  return { leaves, headerRowsByLevel };
}

export function materializeTablixColumns(tablix, rows, parameters, globals = {}, datasets = {}) {
  if (!tablix.hasColumnGroups) return tablix.columns;
  const sourceRows = prepareTablixData(tablix, rows, parameters, globals, datasets);
  const { leaves } = matrixColumnData(tablix, sourceRows, parameters, globals, datasets);
  return [
    ...(tablix.rowHeaderColumns || []),
    ...leaves.map((leaf) => tablix.bodyColumns[leaf.templateIndex]),
  ];
}

function staticLeafRole(leaf, ancestorPath) {
  if (!ancestorPath.some((member) => member.group?.expressions?.length)) return 'static';
  const parent = ancestorPath[ancestorPath.length - 1];
  const siblings = parent?.children || [];
  const dynamicIndex = siblings.findIndex((sibling) => sibling.group);
  return dynamicIndex === -1 || siblings.indexOf(leaf) < dynamicIndex ? 'header' : 'footer';
}

// Recursive per-instance walk that emits ordered "row units" for the advanced features. Each unit
// records the leaf template it renders, its representative fields, its innermost aggregate scope
// (dataset), the named scope map, and per-unit role/indent/page-break metadata.
function materializeAdvancedRows(tablix, sourceRows, parameters, globals, datasets) {
  const groups = new Map();
  memberGroupsByName(tablix.rowMembers, groups);
  memberGroupsByName(tablix.columnMembers, groups);
  const columnData = tablix.hasColumnGroups ? matrixColumnData(tablix, sourceRows, parameters, globals, datasets) : null;

  const leafOrder = [];
  const collectLeaves = (members) => {
    for (const member of members || []) {
      if (member.children?.length) collectLeaves(member.children);
      else leafOrder.push(member);
    }
  };
  collectLeaves(tablix.rowMembers);
  const templateIndexOf = new Map(leafOrder.map((member, index) => [member, index]));

  const units = [];
  const lastBreakKey = new Map();
  let pendingPageBreak = false;
  const emit = (leaf, path, fields, dataset, scopes, role, indentLevel, nestedDataset = dataset) => {
    units.push({ templateIndex: templateIndexOf.get(leaf) ?? 0, path, fields, dataset, nestedDataset, scopes, role, indentLevel, pageBreakBefore: pendingPageBreak, emitIndex: units.length });
    pendingPageBreak = false;
  };

  const walkRecursive = (member, path, incomingRows, scopes, baseIndent) => {
    const group = member.group;
    const byKey = new Map();
    const order = [];
    incomingRows.forEach((fields, index) => {
      const key = JSON.stringify(groupValue(group, fields, parameters, globals, sourceRows, datasets, index));
      const parentRaw = expressionValue(group.parent, { fields, parameters, globals, dataset: sourceRows, datasets });
      const parentKey = parentRaw === null || parentRaw === undefined || parentRaw === '' ? null : JSON.stringify([parentRaw]);
      if (!byKey.has(key)) { byKey.set(key, { key, parentKey, rows: [] }); order.push(key); }
      byKey.get(key).rows.push(fields);
    });
    const childrenByParent = new Map();
    for (const key of order) {
      const node = byKey.get(key);
      const parent = node.parentKey && byKey.has(node.parentKey) ? node.parentKey : null;
      if (!childrenByParent.has(parent)) childrenByParent.set(parent, []);
      childrenByParent.get(parent).push(key);
    }
    const visit = (key, depth) => {
      const node = byKey.get(key);
      const childScopes = group.name ? { ...scopes, [group.name]: node.rows } : scopes;
      if (member.children?.length) walkMembers(member.children, node.rows, childScopes, path, depth);
      else emit(member, path, node.rows[0] || {}, node.rows, childScopes, 'group', depth);
      for (const childKey of childrenByParent.get(key) || []) visit(childKey, depth + 1);
    };
    for (const rootKey of childrenByParent.get(null) || []) visit(rootKey, baseIndent);
  };

  function walkMembers(members, incomingRows, scopes, ancestorPath, indentLevel) {
    for (const member of members || []) {
      // Apply HideIfNoRows before descending so a static container and its descendants follow the same
      // empty-region visibility rule as a static leaf.
      if (incomingRows.length === 0 && !member.group && member.hideIfNoRows) continue;
      const path = [...ancestorPath, member];
      const group = member.group;
      if (group?.parent) { walkRecursive(member, path, incomingRows, scopes, indentLevel); continue; }
      if (group?.expressions?.length) {
        const breaks = /^(Start|End|Between|StartAndEnd)$/i.test(String(group.pageBreak || 'None'));
        for (const instance of partitionByGroup(group, incomingRows, parameters, globals, sourceRows, datasets)) {
          if (member.hideIfNoRows && instance.rows.length === 0) continue;
          if (breaks) {
            const previous = lastBreakKey.get(group);
            if (previous !== undefined && previous !== instance.key) pendingPageBreak = true;
            lastBreakKey.set(group, instance.key);
          }
          const childScopes = group.name ? { ...scopes, [group.name]: instance.rows } : scopes;
          if (member.children?.length) walkMembers(member.children, instance.rows, childScopes, path, indentLevel);
          else emit(member, path, instance.rows[0] || {}, instance.rows, childScopes, 'group', indentLevel);
        }
      } else if (group) {
        for (const row of incomingRows) {
          if (member.children?.length) walkMembers(member.children, [row], scopes, path, indentLevel);
          else emit(member, path, row, incomingRows, scopes, 'detail', indentLevel, [row]);
        }
      } else if (member.children?.length) {
        walkMembers(member.children, incomingRows, scopes, path, indentLevel);
      } else {
        emit(member, path, incomingRows[0] || {}, incomingRows, scopes, staticLeafRole(member, ancestorPath), indentLevel);
      }
    }
  }

  walkMembers(tablix.rowMembers, sourceRows, dataRegionScopes(tablix, sourceRows), [], 0);

  // Visibility changes the physical row sequence, so remove hidden units before calculating vertical
  // spans. Counting a hidden unit and then dropping it later makes an ancestor span one row too far and
  // overlap the first row of the next group instance.
  const visibleUnits = units.filter((unit) => {
    const template = tablix.rows[unit.templateIndex];
    if (!template) return false;
    const context = {
      fields: unit.fields,
      parameters,
      globals,
      dataset: unit.dataset,
      outermostDataset: sourceRows,
      nestedDataset: unit.nestedDataset,
      datasets,
      scopes: unit.scopes,
      tablixDatasetName: tablix.datasetName,
      tablixName: tablix.name,
      nestedTablixDepth: globals.__nestedTablixDepth || 0,
    };
    return !evalHidden(template.hidden, context) && !unit.path.some((member) => evalHidden(member.hidden, context));
  });

  // Advanced tablixes can still begin with ordinary static column-header members before their first
  // grouped row. Preserve the same repeatable-header contract as the flat materializer: only the leading
  // contiguous static units repeat, while group headers and subtotal rows remain ordinary table rows.
  const firstDynamicUnit = visibleUnits.findIndex((unit) => unit.role !== 'static');
  const repeatingHeaderCount = firstDynamicUnit > 0 ? firstDynamicUnit : 0;

  const rowHeaderColumns = tablix.rowHeaderColumns || [];
  const duplicateState = new Map();
  const unitDescriptors = visibleUnits.map((unit) => headerDescriptors(unit.path));
  const descriptorValue = (unit, descriptors, descriptor) => descriptorKey(descriptor, unit.fields, parameters, globals, sourceRows, datasets, unit.emitIndex);

  const output = [];
  for (const [index, unit] of visibleUnits.entries()) {
    const template = tablix.rows[unit.templateIndex];
    const context = {
      fields: unit.fields,
      parameters,
      globals,
      dataset: unit.dataset,
      outermostDataset: sourceRows,
      datasets,
      scopes: unit.scopes,
      tablixDatasetName: tablix.datasetName,
      tablixName: tablix.name,
    };
    const resolveScope = (scope) => scopeKey(scope, tablix, groups, unit.fields, parameters, globals, sourceRows, datasets, unit.emitIndex, unit.path);

    const descriptors = unitDescriptors[index];
    const rowHeaders = [];
    const staticSpans = unit.role === 'static' ? headerSpans(descriptors, rowHeaderColumns) : [];
    descriptors.forEach((descriptor, descriptorIndex) => {
      const key = descriptorValue(unit, descriptors, descriptor);
      const previous = index > 0 ? unitDescriptors[index - 1].find((entry) => entry.member === descriptor.member) : null;
      if (previous && descriptorValue(visibleUnits[index - 1], unitDescriptors[index - 1], previous) === key) return;
      let rowSpan = 1;
      for (let ahead = index + 1; ahead < visibleUnits.length; ahead += 1) {
        const match = unitDescriptors[ahead].find((entry) => entry.member === descriptor.member);
        if (match && descriptorValue(visibleUnits[ahead], unitDescriptors[ahead], match) === key) rowSpan += 1;
        else break;
      }
      rowHeaders.push({
        ...materializedCell(descriptor.header.cell, context, duplicateState, `header:${descriptorIndex}`, resolveScope),
        // Dynamic group owners keep a stable single hierarchy column because they may also span later
        // rows with deeper descendants. A standalone static band has no such descendants, so preserve
        // its declared header size as a horizontal span across the row-header grid.
        colSpan: staticSpans[descriptorIndex] || 1,
        rowSpan,
        isRowHeader: true,
      });
    });

    // Advanced group header/footer rows often have fewer header descriptors than the deepest detail path.
    // Materialize the unowned tail as a real blank grid cell so body cells always begin after the complete
    // row-header grid. Active ancestor row spans occupy the earlier columns and are skipped by placement.
    const rowHeaderTail = Math.max(0, rowHeaderColumns.length - descriptors.length);
    if (rowHeaderTail > 0 && unit.role !== 'static') {
      rowHeaders.push({ colSpan: rowHeaderTail, rowSpan: 1, items: [], values: [], hidden: false, isRowHeader: true });
    }

    let bodyCells;
    if (columnData) {
      bodyCells = [];
      for (const column of columnData.leaves) {
        const intersection = intersectRows(unit.dataset, column.rows);
        const cellContext = {
          fields: intersection[0] || unit.fields,
          parameters,
          globals,
          dataset: intersection,
          outermostDataset: sourceRows,
          nestedDataset: intersection,
          datasets,
          scopes: { ...unit.scopes, ...column.scopes },
          tablixDatasetName: tablix.datasetName,
          tablixName: tablix.name,
          nestedTablixDepth: globals.__nestedTablixDepth || 0,
        };
        const cellResolve = (scope) => scopeKey(scope, tablix, groups, cellContext.fields, parameters, globals, sourceRows, datasets, unit.emitIndex, unit.path);
        const cell = template.cells[column.templateIndex];
        if (cell) {
          bodyCells.push(materializedCell(
            cell,
            cellContext,
            duplicateState,
            `body:${unit.templateIndex}:${column.rows.length}:${column.templateIndex}`,
            cellResolve,
          ));
        }
      }
    } else {
      bodyCells = template.cells.map((cell, cellIndex) => materializedCell(cell, context, duplicateState, `body:${unit.templateIndex}:${cellIndex}`, resolveScope));
    }

    output.push({
      sourceIndex: unit.templateIndex,
      isHeader: index < repeatingHeaderCount,
      isStatic: unit.role === 'static',
      role: unit.role,
      indentLevel: unit.indentLevel || 0,
      pageBreakBefore: unit.pageBreakBefore || false,
      fields: unit.fields,
      scopeDataset: unit.dataset,
      height: template.height,
      keepTogether: unit.path.some((member) => member.keepTogether) || template.cells.some((cell) => cell.items.some((item) => item.keepTogether)),
      cells: [...rowHeaders, ...bodyCells],
    });
  }

  if (columnData) {
    const headerRows = [];
    const cornerRowSpan = columnData.headerRowsByLevel.filter(Boolean).length || 1;
    columnData.headerRowsByLevel.forEach((level, depth) => {
      if (!level) return;
      const cells = [];
      if (depth === 0 && rowHeaderColumns.length) {
        const cornerCell = tablix.tablixCorner?.[0]?.[0];
        const cornerContext = {
          fields: {},
          parameters,
          globals,
          dataset: sourceRows,
          outermostDataset: sourceRows,
          datasets,
          scopes: dataRegionScopes(tablix, sourceRows),
          tablixDatasetName: tablix.datasetName,
          tablixName: tablix.name,
          nestedTablixDepth: globals.__nestedTablixDepth || 0,
        };
        cells.push({
          ...(cornerCell
            ? materializedCell(cornerCell, cornerContext, duplicateState, `corner`, () => `corner`)
            : { items: [], values: [], hidden: false }),
          colSpan: rowHeaderColumns.length,
          rowSpan: cornerRowSpan,
          isRowHeader: true,
        });
      }
      for (const header of level) {
        const headerContext = {
          fields: header.fields,
          parameters,
          globals,
          dataset: header.rows,
          outermostDataset: sourceRows,
          datasets,
          scopes: header.scopes,
          tablixDatasetName: tablix.datasetName,
          tablixName: tablix.name,
          nestedTablixDepth: globals.__nestedTablixDepth || 0,
        };
        const headerResolve = (scope) => scopeKey(scope, tablix, groups, header.fields, parameters, globals, sourceRows, datasets, 0);
        cells.push({
          ...materializedCell(header.cell, headerContext, duplicateState, `colheader:${depth}`, headerResolve),
          colSpan: header.colSpan,
          rowSpan: 1,
        });
      }
      headerRows.push({
        sourceIndex: -1,
        isHeader: true,
        isStatic: true,
        role: 'columnHeader',
        indentLevel: 0,
        pageBreakBefore: false,
        fields: {},
        scopeDataset: sourceRows,
        height: tablix.rows[0]?.height || 18,
        keepTogether: false,
        cells,
      });
    });
    const result = [...headerRows, ...output];
    assertMaterializedGrid(result, rowHeaderColumns.length + columnData.leaves.length, tablix.name);
    return result;
  }

  assertMaterializedGrid(output, rowHeaderColumns.length + tablix.bodyColumns.length, tablix.name);
  return output;
}
