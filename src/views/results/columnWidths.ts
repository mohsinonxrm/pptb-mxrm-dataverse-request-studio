// Persisted results-grid column widths (#12).
//
// Widths the user drags are stored per (table, grid column id) in the
// display settings, so a width sticks for that table's column across runs
// and across different $select lists. Column ids are the grid's own ids —
// response keys such as `name`, `_ownerid_value`, `primarycontactid.fullname`
// or a `(Raw)` twin's `<key>__raw` — so each grid column has its own width.
//
// The store keeps the most recently resized MAX_WIDTH_TABLES tables so the
// settings blob can't grow without bound.

/** table key → (grid column id → width in px). Insertion order = least → most recent. */
export type ColumnWidthStore = Record<string, Record<string, number>>;

export const MAX_WIDTH_TABLES = 50;
const MIN_STORED_WIDTH = 20;

/** Store key for a table: its logical name, lower-cased; `_` when there's none. */
export function widthTableKey(table: string | undefined): string {
  return table?.trim().toLowerCase() || '_';
}

/** The stored widths for a table (empty when none). */
export function widthsForTable(
  store: ColumnWidthStore | undefined,
  table: string | undefined,
): Record<string, number> {
  const own = store?.[widthTableKey(table)];
  return own && typeof own === 'object' ? own : {};
}

/**
 * `store` with `widths` merged into the table's widths (rounded to whole
 * pixels; nonsense values dropped). The table moves to the most-recent end,
 * and the oldest tables beyond MAX_WIDTH_TABLES are dropped.
 */
export function withColumnWidths(
  store: ColumnWidthStore | undefined,
  table: string | undefined,
  widths: Record<string, number>,
): ColumnWidthStore {
  const key = widthTableKey(table);
  const merged: Record<string, number> = { ...widthsForTable(store, table) };
  for (const [col, w] of Object.entries(widths)) {
    if (Number.isFinite(w) && w >= MIN_STORED_WIDTH) merged[col] = Math.round(w);
  }
  const rest = Object.entries(store ?? {}).filter(([k]) => k !== key);
  const kept = rest.slice(Math.max(0, rest.length - (MAX_WIDTH_TABLES - 1)));
  return Object.fromEntries([...kept, [key, merged]]);
}

/** `store` without the table's widths ("Reset column widths"). */
export function withoutTableWidths(
  store: ColumnWidthStore | undefined,
  table: string | undefined,
): ColumnWidthStore {
  const key = widthTableKey(table);
  return Object.fromEntries(Object.entries(store ?? {}).filter(([k]) => k !== key));
}
