// Utilities for extracting OData annotation values + filtering annotation
// keys out of a row's column list.
//
// The annotation contract is the same regardless of how the query was
// authored (FetchXML vs OData), so the helpers are query-shape agnostic.
//
// OData annotation suffixes Dataverse emits:
//   • Regular columns:  <col>@OData.Community.Display.V1.FormattedValue
//   • Lookup columns:   _<col>_value@OData.Community.Display.V1.FormattedValue
//   • Aliased columns:  <alias>@OData.Community.Display.V1.FormattedValue
//   • Attribute name:   <col>@OData.Community.Display.V1.AttributeName
//   • Polymorphic nav:  _<col>_value@Microsoft.Dynamics.CRM.associatednavigationproperty
//   • Lookup target:    _<col>_value@Microsoft.Dynamics.CRM.lookuplogicalname
//
// See https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query-data-web-api#retrieve-formatted-values

const FORMATTED_VALUE_SUFFIX = '@OData.Community.Display.V1.FormattedValue';
const ATTRIBUTE_NAME_SUFFIX = '@OData.Community.Display.V1.AttributeName';
const ASSOC_NAV_SUFFIX = '@Microsoft.Dynamics.CRM.associatednavigationproperty';
const LOOKUP_LOGICAL_SUFFIX = '@Microsoft.Dynamics.CRM.lookuplogicalname';

/**
 * Returns the formatted-value annotation for a column, or the raw value if
 * no annotation exists. Handles regular, lookup-value, and aliased columns.
 *
 * @example
 *   getFormattedValue({ statuscode: 1, 'statuscode@…FormattedValue': 'Active' }, 'statuscode')
 *     → 'Active'
 *   getFormattedValue({ _ownerid_value: 'guid', '_ownerid_value@…FormattedValue': 'Jane Doe' }, '_ownerid_value')
 *     → 'Jane Doe'
 */
export function getFormattedValue(record: Record<string, unknown>, column: string): unknown {
  if (!record || !column) return null;

  // Direct hit: exact-name annotation
  const direct = `${column}${FORMATTED_VALUE_SUFFIX}`;
  if (direct in record) return record[direct];

  // Already in lookup form (`_xxx_value`) — direct hit was tried above, fall through to raw.
  if (column.startsWith('_') && column.endsWith('_value')) return record[column];

  // Caller passed the bare attribute name for a lookup — try `_<col>_value@...FV`
  const lookupForm = `_${column}_value${FORMATTED_VALUE_SUFFIX}`;
  if (lookupForm in record) return record[lookupForm];

  return record[column];
}

/**
 * For lookup-value columns, returns the polymorphic target's logical name
 * (e.g. `account` or `contact` for `_customerid_value`). Returns `undefined`
 * for non-lookup columns or when the annotation is absent.
 */
export function getLookupTargetEntity(record: Record<string, unknown>, column: string): string | undefined {
  if (!record || !column) return undefined;
  const direct = record[`${column}${LOOKUP_LOGICAL_SUFFIX}`];
  if (typeof direct === 'string') return direct;
  if (!column.startsWith('_') || !column.endsWith('_value')) {
    const wrapped = record[`_${column}_value${LOOKUP_LOGICAL_SUFFIX}`];
    if (typeof wrapped === 'string') return wrapped;
  }
  return undefined;
}

/**
 * For lookup-value columns, returns the navigation-property name Dataverse
 * uses for this row's specific target (e.g. `customerid_account` when the
 * Customer happens to point at an account row).
 */
export function getAssociatedNavProperty(record: Record<string, unknown>, column: string): string | undefined {
  if (!record || !column) return undefined;
  const direct = record[`${column}${ASSOC_NAV_SUFFIX}`];
  if (typeof direct === 'string') return direct;
  if (!column.startsWith('_') || !column.endsWith('_value')) {
    const wrapped = record[`_${column}_value${ASSOC_NAV_SUFFIX}`];
    if (typeof wrapped === 'string') return wrapped;
  }
  return undefined;
}

/**
 * For aliased columns, returns the underlying attribute logical name from
 * the `AttributeName` annotation. Used by the column header resolver to
 * pull the right `DisplayName` from metadata.
 */
export function getOriginalAttributeName(record: Record<string, unknown>, column: string): string | undefined {
  if (!record || !column) return undefined;
  const v = record[`${column}${ATTRIBUTE_NAME_SUFFIX}`];
  return typeof v === 'string' ? v : undefined;
}

/**
 * True if a row key is an OData / Dynamics CRM annotation that should not
 * appear as a grid column. Matches all known annotation namespaces (lower-
 * and upper-case forms Dataverse uses interchangeably). Also matches the
 * internal `__rowKey` synthetic column injected by the flattener for
 * stable row IDs after collection-expand multiplication.
 */
export function isAnnotationColumn(columnName: string): boolean {
  if (!columnName) return false;
  if (columnName.startsWith('__')) return true; // synthetic columns (e.g. __rowKey)
  return (
    columnName.includes('@OData.') ||
    columnName.includes('@Microsoft.Dynamics.CRM.') ||
    columnName.includes('@odata.') ||
    columnName.includes('@microsoft.dynamics.crm.')
  );
}

/** Drop annotation keys from a column list — leaves only displayable cols. */
export function filterDisplayableColumns(columns: string[]): string[] {
  if (!columns) return [];
  return columns.filter(c => !isAnnotationColumn(c));
}

/** Fast check: does this row carry any formatted-value annotation at all? */
export function hasFormattedValues(record: Record<string, unknown>): boolean {
  if (!record) return false;
  for (const k of Object.keys(record)) {
    if (k.endsWith(FORMATTED_VALUE_SUFFIX)) return true;
  }
  return false;
}
