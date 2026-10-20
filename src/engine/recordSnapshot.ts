// recordSnapshot — read column values out of a raw Web API row.
//
// Rows fetched with `GET /<set>(<id>)` (Merge's Target / Subordinate,
// Update's "before" row) never carry a lookup under its logical name.
// Dataverse returns the GUID at `_<col>_value`, and because PPTB's
// queryData always sends `Prefer: odata.include-annotations="*"`, the same
// key also carries:
//
//   _<col>_value@Microsoft.Dynamics.CRM.lookuplogicalname   → target entity
//   _<col>_value@OData.Community.Display.V1.FormattedValue  → display name
//
// The lookuplogicalname annotation is the only reliable way to know which
// entity a polymorphic lookup (customer / owner / regarding) points at —
// `targets[0]` is just the first allowed type.

import { isLookupLike, type ColumnMeta } from '../mock/metadata';
import { attrRef } from './odataAttr';
import type { LookupFieldValue } from '../state/writeState';

const LOOKUP_LOGICAL_NAME = '@Microsoft.Dynamics.CRM.lookuplogicalname';
const FORMATTED_VALUE = '@OData.Community.Display.V1.FormattedValue';

/** A lookup read from a row: id + resolved target entity + display name. */
export interface RowLookupValue extends LookupFieldValue {
  name?: string;
}

/**
 * Read a lookup-like column from a raw row. Returns null when the lookup is
 * empty. `targetEntity` comes from the lookuplogicalname annotation; when
 * that's missing it falls back to the column's only target, and is `''`
 * for a polymorphic column (callers must not bind to a guessed entity).
 */
export function readRowLookup(
  row: Record<string, unknown>,
  col: ColumnMeta,
): RowLookupValue | null {
  if (!isLookupLike(col)) return null;
  const key = attrRef(col, col.logicalName);
  const id = row[key];
  if (typeof id !== 'string' || !id) return null;
  const annotated = row[`${key}${LOOKUP_LOGICAL_NAME}`];
  const targetEntity =
    typeof annotated === 'string' && annotated
      ? annotated
      : col.targets.length === 1
        ? col.targets[0]
        : '';
  const name = row[`${key}${FORMATTED_VALUE}`];
  return { id, targetEntity, name: typeof name === 'string' && name ? name : undefined };
}

/** Read any column from a raw row — lookups via `readRowLookup`. */
export function readRowValue(row: Record<string, unknown>, col: ColumnMeta): unknown {
  return isLookupLike(col) ? readRowLookup(row, col) : row[col.logicalName];
}

/** True when two lookup values point at the same row (GUIDs compare case-insensitively). */
export function sameLookup(
  a: LookupFieldValue | null | undefined,
  b: LookupFieldValue | null | undefined,
): boolean {
  if (!a?.id || !b?.id) return !a?.id && !b?.id;
  if (a.id.toLowerCase() !== b.id.toLowerCase()) return false;
  // An unresolved target (polymorphic, no annotation) can't disagree.
  return !a.targetEntity || !b.targetEntity || a.targetEntity === b.targetEntity;
}

/** Display text for a lookup value: its name, else `entity(guid…)`. */
export function formatLookup(lk: RowLookupValue | LookupFieldValue): string {
  const name = (lk as RowLookupValue).name;
  if (name) return lk.targetEntity ? `${name} (${lk.targetEntity})` : name;
  return `${lk.targetEntity || '?'}(${lk.id.slice(0, 8)}…)`;
}
