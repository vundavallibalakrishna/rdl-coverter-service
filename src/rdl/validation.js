import { ServiceError } from '../errors.js';
import { evaluateExpression, formatValue } from './expression.js';
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

export function validateRenderInput(model, request, limits) {
  if (model.unsupported.length > 0) throw new ServiceError(
    'UNSUPPORTED_FEATURE',
    'RDL contains unsupported elements, properties, or expressions',
    400,
    { features: model.unsupported.slice(0, 100) },
  );
  const parameters = request.parameters || {};
  const resolvedParameters = { ...parameters };
  for (const parameter of model.parameters) {
    const defaultValue = parameter.multiValue ? parameter.defaultValues : parameter.defaultValues[0];
    const value = hasValue(parameters[parameter.name]) ? parameters[parameter.name] : defaultValue;
    if (!hasValue(value) && !parameter.nullable && !parameter.allowBlank) throw new ServiceError('PARAMETER_INVALID', `Required parameter is missing: ${parameter.name}`);
    if (hasValue(value)) validateParameterType(parameter, value);
    resolvedParameters[parameter.name] = value;
  }

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
        if (!(field.dataField in row)) throw new ServiceError('FIELD_MISSING', `Dataset ${dataset.name} row ${index} is missing field ${field.dataField}`);
      }
    }
  }
  return { totalRows, parameters: resolvedParameters };
}

function expressionUsesFields(value) {
  return typeof value === 'string' && /Fields!/i.test(value);
}

function itemUsesFields(item) {
  if (expressionUsesFields(item.value) || expressionUsesFields(item.hidden) || expressionUsesFields(JSON.stringify(item.paragraphs || []))) return true;
  return (item.items || []).some(itemUsesFields);
}

function normalizeFields(row, definitions = []) {
  const normalized = { ...row };
  for (const field of definitions) normalized[field.name] = row[field.dataField];
  return normalized;
}

function evaluateItemText(item, context) {
  if (!item.paragraphs) {
    const value = evaluateExpression(item.value, context);
    return item.style?.format ? formatValue(value, item.style.format) : value;
  }
  return normalizeDisplayText(item.paragraphs.map((runs) => runs.map((run) => {
    const definition = run && typeof run === 'object' ? run : { value: run, markupType: 'None' };
    const value = evaluateExpression(definition.value, context);
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

function groupValue(group, fields, parameters, globals, dataset, datasets, rowIndex) {
  if (!group) return null;
  if (!group.expressions?.length) return `row:${rowIndex}`;
  const context = { fields, parameters, globals, dataset, datasets };
  return group.expressions.map((expression) => expressionValue(expression, context));
}

function scopeKey(scope, tablix, groups, fields, parameters, globals, dataset, datasets, rowIndex) {
  if (!scope || scope === tablix.datasetName) return `dataset:${tablix.datasetName}`;
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
    const value = itemValue(item, context);
    if (item.type !== 'Textbox' || !item.hideDuplicates) return value;
    const key = `${duplicatePrefix}:${item.name || itemIndex}`;
    const current = { scope: scopeResolver(item.hideDuplicates), value: String(value ?? '') };
    const previous = duplicateState.get(key);
    duplicateState.set(key, current);
    return previous && previous.scope === current.scope && previous.value === current.value ? '' : value;
  });
  return {
    ...cell,
    values,
    hidden: cell.items.some((item) => {
      const result = evaluateExpression(item.hidden, context);
      return result === true || String(result).toLowerCase() === 'true';
    }),
  };
}

export function prepareTablixData(tablix, rows, parameters, globals = {}, datasets = {}) {
  const normalized = (Array.isArray(rows) ? rows : []).map((row) => normalizeFields(row, tablix.datasetFields));
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
  // A group header/footer subtotal row is a STATIC leaf member (no Group element at all) nested inside
  // an ancestor group that has grouping expressions. A detail leaf (the `Details` group, expressions
  // empty) still has a Group element and is intentionally NOT matched, so ordinary grouped tables stay
  // on the flat path.
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
    const scopes = { [tablix.datasetName]: sourceRows };
    let innermost = sourceRows;
    for (const member of path || []) {
      const group = member.group;
      if (!group?.expressions?.length) continue;
      const key = JSON.stringify(groupValue(group, fields, parameters, globals, sourceRows, datasets, rowIndex));
      const instance = groupInstances.get(group)?.get(key) || sourceRows;
      if (group.name) scopes[group.name] = instance;
      innermost = instance;
    }
    return { scopes, dataset: innermost };
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
  const firstDynamic = tablix.rows.findIndex((row, index) => (
    row.cells.some((cell) => cell.items.some(itemUsesFields)) || memberPaths[index]?.some((member) => member.group)
  ));
  let repeatingHeaderCount = 0;
  if (firstDynamic > 0 && tablix.repeatColumnHeaders) repeatingHeaderCount = firstDynamic;
  else if (firstDynamic > 0) {
    while (repeatingHeaderCount < firstDynamic && memberPaths[repeatingHeaderCount]?.some((member) => member.repeatOnNewPage)) repeatingHeaderCount += 1;
  }
  const output = [];
  const duplicateState = new Map();
  for (const [index, row] of tablix.rows.entries()) {
    const path = memberPaths[index] || [];
    const descriptors = headerDescriptors(path);
    const spans = headerSpans(descriptors, tablix.rowHeaderColumns || []);
    const dynamic = row.cells.some((cell) => cell.items.some(itemUsesFields)) || path.some((member) => member.group);
    const contexts = dynamic ? sourceRows : [sourceRows[0] || {}];
    for (const [rowIndex, fields] of contexts.entries()) {
      const { scopes, dataset } = dynamic ? scopeContext(path, fields, rowIndex) : { scopes: { [tablix.datasetName]: sourceRows }, dataset: sourceRows };
      const context = { fields, parameters, globals, dataset, datasets, scopes };
      // Honor static Visibility.Hidden on the row and on any member in its path (a hidden group/row is
      // not rendered). Interactive drill-down toggles are ignored — everything renders expanded.
      if (evalHidden(row.hidden, context) || path.some((member) => evalHidden(member.hidden, context))) continue;
      const resolveScope = (scope) => scopeKey(scope, tablix, groups, fields, parameters, globals, sourceRows, datasets, rowIndex);
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
  const bodyColumnCount = tablix.bodyColumns.length;
  const leaves = [];
  const headerRowsByLevel = [];
  const walk = (members, incomingRows, scopes, depth) => {
    let leafCount = 0;
    for (const member of members || []) {
      const group = member.group;
      if (group?.expressions?.length) {
        for (const instance of partitionByGroup(group, incomingRows, parameters, globals, sourceRows, datasets)) {
          const childScopes = group.name ? { ...scopes, [group.name]: instance.rows } : scopes;
          const childLeafCount = member.children?.length
            ? walk(member.children, instance.rows, childScopes, depth + 1)
            : (leaves.push({ rows: instance.rows, scopes: childScopes }), 1);
          if (member.header) {
            (headerRowsByLevel[depth] ||= []).push({
              cell: member.header.cell,
              colSpan: childLeafCount * bodyColumnCount,
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
        leaves.push({ rows: incomingRows, scopes });
        leafCount += 1;
      }
    }
    return leafCount;
  };
  walk(tablix.columnMembers, sourceRows, { [tablix.datasetName]: sourceRows }, 0);
  return { leaves, headerRowsByLevel };
}

export function materializeTablixColumns(tablix, rows, parameters, globals = {}, datasets = {}) {
  if (!tablix.hasColumnGroups) return tablix.columns;
  const sourceRows = prepareTablixData(tablix, rows, parameters, globals, datasets);
  const { leaves } = matrixColumnData(tablix, sourceRows, parameters, globals, datasets);
  return [...(tablix.rowHeaderColumns || []), ...leaves.flatMap(() => tablix.bodyColumns)];
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
  const emit = (leaf, path, fields, dataset, scopes, role, indentLevel) => {
    units.push({ templateIndex: templateIndexOf.get(leaf) ?? 0, path, fields, dataset, scopes, role, indentLevel, pageBreakBefore: pendingPageBreak, emitIndex: units.length });
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
          else emit(member, path, row, incomingRows, scopes, 'detail', indentLevel);
        }
      } else if (member.children?.length) {
        walkMembers(member.children, incomingRows, scopes, path, indentLevel);
      } else {
        emit(member, path, incomingRows[0] || {}, incomingRows, scopes, staticLeafRole(member, ancestorPath), indentLevel);
      }
    }
  }

  walkMembers(tablix.rowMembers, sourceRows, { [tablix.datasetName]: sourceRows }, [], 0);

  const rowHeaderColumns = tablix.rowHeaderColumns || [];
  const duplicateState = new Map();
  const unitDescriptors = units.map((unit) => headerDescriptors(unit.path));
  const descriptorValue = (unit, descriptors, descriptor) => descriptorKey(descriptor, unit.fields, parameters, globals, sourceRows, datasets, unit.emitIndex);

  const output = [];
  for (const [index, unit] of units.entries()) {
    const template = tablix.rows[unit.templateIndex];
    if (!template) continue;
    const context = { fields: unit.fields, parameters, globals, dataset: unit.dataset, datasets, scopes: unit.scopes };
    if (evalHidden(template.hidden, context) || unit.path.some((member) => evalHidden(member.hidden, context))) continue;
    const resolveScope = (scope) => scopeKey(scope, tablix, groups, unit.fields, parameters, globals, sourceRows, datasets, unit.emitIndex);

    const descriptors = unitDescriptors[index];
    const spans = headerSpans(descriptors, rowHeaderColumns);
    const rowHeaders = [];
    descriptors.forEach((descriptor, descriptorIndex) => {
      const key = descriptorValue(unit, descriptors, descriptor);
      const previous = index > 0 ? unitDescriptors[index - 1].find((entry) => entry.member === descriptor.member) : null;
      if (previous && descriptorValue(units[index - 1], unitDescriptors[index - 1], previous) === key) return;
      let rowSpan = 1;
      for (let ahead = index + 1; ahead < units.length; ahead += 1) {
        const match = unitDescriptors[ahead].find((entry) => entry.member === descriptor.member);
        if (match && descriptorValue(units[ahead], unitDescriptors[ahead], match) === key) rowSpan += 1;
        else break;
      }
      rowHeaders.push({
        ...materializedCell(descriptor.header.cell, context, duplicateState, `header:${descriptorIndex}`, resolveScope),
        colSpan: spans[descriptorIndex] || 1,
        rowSpan,
        isRowHeader: true,
      });
    });

    let bodyCells;
    if (columnData) {
      bodyCells = [];
      for (const column of columnData.leaves) {
        const intersection = intersectRows(unit.dataset, column.rows);
        const cellContext = { fields: intersection[0] || unit.fields, parameters, globals, dataset: intersection, datasets, scopes: { ...unit.scopes, ...column.scopes } };
        const cellResolve = (scope) => scopeKey(scope, tablix, groups, cellContext.fields, parameters, globals, sourceRows, datasets, unit.emitIndex);
        template.cells.forEach((cell, cellIndex) => {
          bodyCells.push(materializedCell(cell, cellContext, duplicateState, `body:${unit.templateIndex}:${column.rows.length}:${cellIndex}`, cellResolve));
        });
      }
    } else {
      bodyCells = template.cells.map((cell, cellIndex) => materializedCell(cell, context, duplicateState, `body:${unit.templateIndex}:${cellIndex}`, resolveScope));
    }

    output.push({
      sourceIndex: unit.templateIndex,
      isHeader: false,
      isStatic: false,
      role: unit.role,
      indentLevel: unit.indentLevel || 0,
      pageBreakBefore: unit.pageBreakBefore || false,
      fields: unit.fields,
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
        const cornerContext = { fields: {}, parameters, globals, dataset: sourceRows, datasets, scopes: { [tablix.datasetName]: sourceRows } };
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
        const headerContext = { fields: header.fields, parameters, globals, dataset: header.rows, datasets, scopes: header.scopes };
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
        height: tablix.rows[0]?.height || 18,
        keepTogether: false,
        cells,
      });
    });
    return [...headerRows, ...output];
  }

  return output;
}
