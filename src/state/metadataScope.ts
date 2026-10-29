// Which tables, columns and relationships the pickers offer (#44), matching
// FetchXML Studio (its #119) so both tools behave the same.
//
// "Advanced Find only" (Settings → Query Scope, off by default) keeps what
// Advanced Find shows: `IsValidForAdvancedFind`, "whether the item is shown
// in Advanced Find" (Microsoft Learn). It isn't a query restriction: items
// Advanced Find hides query fine, which is what lets DRS build requests
// Advanced Find can't. So the metadata loads once, unfiltered, and only the
// pickers are scoped. Lookups of what a request already uses (encoding a
// pasted URL, validation, the results grid) always see everything.
//
// Image and File columns are never Advanced Find-eligible (`entityimage`
// reports false), so they stay in scope: Manage Image / File Data would have
// nothing to offer otherwise (#305). Columns with IsValidForRead = false are
// left to each read picker, as before; a write picker can still offer them.

import { usePersistedSettings } from '../host/usePersistedSettings';

interface Scoped {
  isValidForAdvancedFind?: boolean;
}

export const isEntityInScope = (e: Scoped, advancedFindOnly: boolean): boolean =>
  !advancedFindOnly || e.isValidForAdvancedFind !== false;

export const isColumnInScope = (
  c: Scoped & { attributeType?: string },
  advancedFindOnly: boolean,
): boolean =>
  !advancedFindOnly ||
  c.isValidForAdvancedFind !== false ||
  c.attributeType === 'Image' ||
  c.attributeType === 'File';

export const isNavInScope = (n: Scoped, advancedFindOnly: boolean): boolean =>
  !advancedFindOnly || n.isValidForAdvancedFind !== false;

/** The columns a picker offers. */
export const scopedColumns = <C extends Scoped & { attributeType?: string }>(
  columns: readonly C[],
  advancedFindOnly: boolean,
): C[] => columns.filter((c) => isColumnInScope(c, advancedFindOnly));

/** The navigation properties a picker offers. */
export const scopedNavs = <N extends Scoped>(navs: readonly N[], advancedFindOnly: boolean): N[] =>
  navs.filter((n) => isNavInScope(n, advancedFindOnly));

/** The "Advanced Find only" setting; pickers re-render when it changes. */
export function useAdvancedFindOnly(): boolean {
  return usePersistedSettings()[0].advancedFindOnly;
}
