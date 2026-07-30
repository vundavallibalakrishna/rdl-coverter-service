import { evaluateExpression } from './expression.js';

/**
 * How an RDL dataset field maps onto a posted row.
 *
 * An RDL `<Field>` is one of two things, and conflating them fails exports that cannot succeed:
 *
 *  - **Bound**: `<Field Name="Owner"><DataField>OwnerName</DataField></Field>` — the value comes from the
 *    query result column named by DataField, so it must be a key in every posted row.
 *  - **Calculated**: `<Field Name="Age"><Value>=DateDiff("d", Fields!Raised.Value, Today())</Value></Field>`
 *    — the value is computed from other fields. It is not a query column and can never be a key in a
 *    posted row, so requiring it rejects a perfectly valid report.
 *
 * Both live in this one module so the request validator and the renderers agree on what a field means.
 */

/**
 * The key in `row` that supplies `dataField`, or undefined if none does.
 *
 * Resolution is case-insensitive on the fallback path: SQL Server column names are case-insensitive under
 * the usual collations and SSRS binds them that way, so an RDL saying `ACTIONID` against a result column
 * `ActionId` is the same column, not a missing one. An exact match always wins, so a row that genuinely
 * carries both spellings is unaffected.
 */
export function rowKeyFor(row, dataField) {
  if (!row || typeof row !== 'object' || dataField === undefined || dataField === null) return undefined;
  if (Object.hasOwn(row, dataField)) return dataField;
  const wanted = String(dataField).toLowerCase();
  return Object.keys(row).find((key) => key.toLowerCase() === wanted);
}

// Only a bound field can be expected in the posted rows.
export function isBoundField(field) {
  return !field?.calculated;
}

/**
 * Rekeys a posted row from query column names to RDL field names, and computes the calculated fields.
 *
 * Bound fields are resolved first because a calculated field's expression reads its siblings by field
 * name. A calculated expression that throws yields undefined rather than failing the render: a blank cell
 * is recoverable, an aborted export is not.
 */
export function normalizeRowFields(row, definitions = [], context = {}) {
  const normalized = { ...row };

  for (const field of definitions) {
    if (!isBoundField(field)) continue;
    const key = rowKeyFor(row, field.dataField);
    // Nested data regions receive the already-normalized rows from their containing group scope. Preserve
    // that internal field value when the original DataField key is no longer present; this also makes
    // normalization idempotent without weakening the request validator's DataField requirement.
    if (key !== undefined) normalized[field.name] = row[key];
    else if (!Object.hasOwn(normalized, field.name)) normalized[field.name] = undefined;
  }

  for (const field of definitions) {
    if (isBoundField(field)) continue;
    try {
      normalized[field.name] = evaluateExpression(field.value, { ...context, fields: normalized });
    } catch {
      normalized[field.name] = undefined;
    }
  }

  return normalized;
}
