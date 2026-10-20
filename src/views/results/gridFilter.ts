// Client-side row filtering for the results grid (#17).
//
// Two filters, both case-insensitive "contains" over the rows already
// loaded (the OData request is never changed):
//   • the toolbar search — any visible column matches;
//   • the filter row — one text per grid column, all must match.
// A cell matches on its formatted value or its raw value; a "(Raw)" twin
// column (`<key>__raw`) matches on the raw value only.
//
// The search runs over a per-row index (each row's visible cell text,
// lower-cased and joined once), so each keystroke is one substring scan per
// row instead of re-reading and re-formatting every cell.

import { getFormattedValue } from './FormattedValueUtils';

/** Grid column id → filter text. Empty / whitespace texts are ignored. */
export type ColumnFilters = Record<string, string>;

/** Separator between cells in the index — can't appear in a typed query. */
const SEP = '\u0001';
const RAW_SUFFIX = '__raw';

const toText = (value: unknown): string => {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object') return JSON.stringify(value);
  return String(value);
};

/** A cell's searchable text (formatted + raw, lower-cased) for a grid column id. */
export function cellSearchText(row: Record<string, unknown>, columnId: string): string {
  if (columnId.endsWith(RAW_SUFFIX)) {
    return toText(row[columnId.slice(0, -RAW_SUFFIX.length)]).toLowerCase();
  }
  const raw = toText(row[columnId]);
  const formatted = toText(getFormattedValue(row, columnId));
  return (formatted === raw ? raw : `${formatted}${SEP}${raw}`).toLowerCase();
}

/** One lower-cased haystack per row over the given column keys. */
export function buildSearchIndex(rows: Record<string, unknown>[], columnKeys: string[]): string[] {
  return rows.map((r) => columnKeys.map((k) => cellSearchText(r, k)).join(SEP));
}

/** The column filters that actually filter (non-blank), normalised. */
export function activeColumnFilters(filters: ColumnFilters): [string, string][] {
  return Object.entries(filters)
    .map(([id, text]): [string, string] => [id, text.trim().toLowerCase()])
    .filter(([, text]) => text.length > 0);
}

/** True when the search or any column filter would narrow the rows. */
export function isFiltering(query: string, filters: ColumnFilters): boolean {
  return query.trim().length > 0 || activeColumnFilters(filters).length > 0;
}

/**
 * The rows matching the search `query` (any column, via `index`) and every
 * column filter. Returns `rows` itself when nothing filters.
 */
export function filterRows<T extends Record<string, unknown>>(
  rows: T[],
  index: string[],
  query: string,
  filters: ColumnFilters,
): T[] {
  const q = query.trim().toLowerCase();
  const cols = activeColumnFilters(filters);
  if (!q && cols.length === 0) return rows;
  return rows.filter(
    (r, i) =>
      (!q || (index[i] ?? '').includes(q)) &&
      cols.every(([id, text]) => cellSearchText(r, id).includes(text)),
  );
}
