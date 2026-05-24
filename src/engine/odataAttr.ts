// odataAttr — translate a studio-shape column logical name into the
// attribute reference Dataverse OData v9.2 actually expects.
//
// Reference: Dataverse REST Builder (https://github.com/GuidoPreite/DRB)
// uses this pattern in every $select / $filter / $orderby that touches a
// lookup column. The OData schema exposes:
//
//   • the lookup VALUE as `_<logical>_value` (a primitive GUID property)
//   • the navigation PROPERTY as `<logical>` (only usable in $expand or
//     inside lambda predicates, e.g. `primarycontactid/contactid`)
//
// Trying to put `customerid` in a $select against
// `Microsoft.Dynamics.CRM.incident` produces the dreaded 0x80060888:
//   "Could not find a property named 'customerid' on type
//    'Microsoft.Dynamics.CRM.incident'."
// because `customerid` is a nav property, not an addressable attribute.
//
// Customer columns are polymorphic (account|contact) so the value column
// is `_customerid_value` too — same encoding.
//
// EntityName / partylist / certain virtual types could be added here if
// they need different handling — for now we only special-case lookups.

import type { ColumnMeta, TableMeta } from '../mock/metadata';

/** True if the column's value lives at `_<logical>_value` in OData. */
export function isLookupLike(col: ColumnMeta | undefined): boolean {
  if (!col) return false;
  return (
    col.attributeType === 'Lookup' ||
    col.attributeType === 'Customer' ||
    col.attributeType === 'Owner'
  );
}

/** Encode a single column's OData attribute reference. */
export function attrRef(col: ColumnMeta | undefined, logicalName: string): string {
  // Prefer the pre-computed oDataName the metadata provider stamped onto
  // the column — it's the single source of truth for "what name does this
  // attribute have in $select / $filter / $expand($select)?". Fall back
  // to deriving it from the attributeType for callers (mocks, tests) that
  // haven't gone through the live provider.
  if (col?.oDataName) return col.oDataName;
  if (isLookupLike(col)) return `_${logicalName}_value`;
  return logicalName;
}

/**
 * Resolve `logicalName` against the table and return the OData attribute
 * reference. Convenience wrapper used by encoders that don't already have
 * a resolved ColumnMeta on hand.
 */
export function attrRefByName(table: TableMeta | undefined, logicalName: string): string {
  if (!table) return logicalName;
  const col = table.columns.find(c => c.logicalName === logicalName);
  return attrRef(col, logicalName);
}

/**
 * Encode a possibly-aliased attribute reference (e.g. lambda paths like
 * `c/primarycontactid`). The leading alias path passes through unchanged;
 * only the trailing segment gets the lookup treatment.
 *
 * Example: `c/primarycontactid` with primarycontactid as Lookup →
 *          `c/_primarycontactid_value`
 */
export function attrRefPath(
  table: TableMeta | undefined,
  expr: string,
): string {
  if (!expr.includes('/')) return attrRefByName(table, expr);
  const parts = expr.split('/');
  const last = parts.pop() ?? expr;
  return [...parts, attrRefByName(table, last)].join('/');
}
