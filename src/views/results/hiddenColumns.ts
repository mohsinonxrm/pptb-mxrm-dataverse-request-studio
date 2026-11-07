// Results-grid columns the user hid with the Columns menu (#292).
//
// Stored per table in the display settings as the hidden column keys (the
// grid's response keys, such as `name` or `_ownerid_value`), so a column the
// next run adds shows by default. Hiding a column also hides its "(Raw)"
// twin. Like the column widths, only the most recently changed
// MAX_HIDDEN_TABLES tables are kept.

import { widthTableKey } from './columnWidths';
import type { ColumnChoice } from './exportResults';

/** table key → hidden column keys. Insertion order = least → most recent. */
export type HiddenColumnStore = Record<string, string[]>;

export const MAX_HIDDEN_TABLES = 50;

/** The hidden column keys for a table (empty when none). */
export function hiddenForTable(
  store: HiddenColumnStore | undefined,
  table: string | undefined,
): string[] {
  const own = store?.[widthTableKey(table)];
  return Array.isArray(own) ? own.filter((k) => typeof k === 'string') : [];
}

/**
 * `store` with the table's hidden columns set to `hidden`. An empty list
 * removes the table ("Show all columns"); otherwise the table moves to the
 * most-recent end and the oldest tables beyond MAX_HIDDEN_TABLES are dropped.
 */
export function withHiddenColumns(
  store: HiddenColumnStore | undefined,
  table: string | undefined,
  hidden: string[],
): HiddenColumnStore {
  const key = widthTableKey(table);
  const rest = Object.entries(store ?? {}).filter(([k]) => k !== key);
  if (hidden.length === 0) return Object.fromEntries(rest);
  const kept = rest.slice(Math.max(0, rest.length - (MAX_HIDDEN_TABLES - 1)));
  return Object.fromEntries([...kept, [key, [...new Set(hidden)]]]);
}

/**
 * The column keys to show: `keys` without the hidden ones. If that would
 * leave nothing (every column the response has is hidden), all show.
 */
export function visibleColumnKeys(keys: string[], hidden: string[]): string[] {
  if (hidden.length === 0) return keys;
  const set = new Set(hidden);
  const shown = keys.filter((k) => !set.has(k));
  return shown.length > 0 ? shown : keys;
}

/** "Columns" or, with some hidden, "Columns · 4 of 23". */
export function columnsButtonLabel(choices: ColumnChoice[]): string {
  const shown = choices.filter((c) => c.shown).length;
  return shown === choices.length ? 'Columns' : `Columns · ${shown} of ${choices.length}`;
}
