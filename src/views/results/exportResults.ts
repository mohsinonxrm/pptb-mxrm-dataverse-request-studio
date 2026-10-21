// Exporting Retrieve-Multiple results to CSV / JSON (#15).
//
// The CSV mirrors what the grid shows: the same columns in the same order,
// with the same headers, over the rows the grid currently lists (so an
// active row search narrows the export too). "Shown" values match the grid
// (formatted values where Dataverse sent them, raw for "(Raw)" twins and
// when the grid is in raw mode); "raw" exports the underlying values
// (GUIDs, numbers, ISO dates). JSON exports the OData envelope as-is.
//
// Cells that a spreadsheet would run as a formula are neutralised for the
// CSV file and the "Copy as TSV" clipboard text alike (`spreadsheetSafe`,
// #239): Dataverse text often comes from outside the organisation.
//
// Pure functions only — the toolbar wires them to `saveFileSafe`.

import { getFormattedValue } from './FormattedValueUtils';

/** A grid column as exported: its header and where its value comes from. */
export interface ExportColumn {
  /** Grid column id (a "(Raw)" twin column ends in `__raw`). */
  id: string;
  /** Header as the grid shows it. */
  header: string;
  /** Row key the value is read from. */
  sourceCol: string;
  /** Whether the grid shows this column's raw value. */
  showsRaw: boolean;
}

/** What the grid currently lists — reported by ResultsGrid for export. */
export interface GridExportView {
  columns: ExportColumn[];
  rows: Record<string, unknown>[];
}

/** "shown": values as the grid shows them; "raw": underlying values. */
export type ExportValueMode = 'shown' | 'raw';

const toText = (value: unknown): string => {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object') return JSON.stringify(value);
  return String(value);
};

/** A cell's text for export. */
export function exportCellText(
  record: Record<string, unknown>,
  column: ExportColumn,
  mode: ExportValueMode,
): string {
  if (mode === 'raw' || column.showsRaw) return toText(record[column.sourceCol]);
  return toText(getFormattedValue(record, column.sourceCol) ?? record[column.sourceCol]);
}

/**
 * Columns to export for a mode. Raw mode leaves out the "(Raw)" twins:
 * their formatted column already exports as the raw value.
 */
export function exportColumnsFor(columns: ExportColumn[], mode: ExportValueMode): ExportColumn[] {
  return mode === 'raw' ? columns.filter((c) => !c.id.endsWith('__raw')) : columns;
}

/** RFC 4180 field: quoted when it holds a comma, quote, line break or edge whitespace. */
export function csvField(text: string): string {
  return /[",\r\n]|^\s|\s$/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

// A spreadsheet runs a cell as a formula when it starts with `=`, `+`, `-` or
// `@` (leading spaces are skipped), or with a tab or carriage return.
const FORMULA_START = /^(?:[\t\r]|\s*[=+\-@])/;
// For text derived from a number only `=` and `@` count: Dataverse formats
// negatives as "-5", "-$1,234.00" or "-1.234,50 €".
const FORMULA_START_NUMERIC = /^(?:[\t\r]|\s*[=@])/;
// Plain numeric text ("-5", "+1.5e3") is a number to a spreadsheet, not a formula.
const PLAIN_NUMBER = /^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?\s*$/;

/**
 * `text` as a spreadsheet cell that can't run as a formula (CSV / formula
 * injection, #239). Text that starts like a formula gets a leading `'`, as
 * OWASP recommends; the spreadsheet then shows the stored text with that
 * mark in front. `source` is the underlying Dataverse value: when it's a
 * number or boolean, the text is Dataverse's own formatting of it, so a
 * leading sign is left alone (a choice label starting with `=` or `@` is
 * still neutralised).
 */
export function spreadsheetSafe(text: string, source?: unknown): string {
  const numeric = typeof source === 'number' || typeof source === 'boolean';
  const start = numeric ? FORMULA_START_NUMERIC : FORMULA_START;
  if (!start.test(text) || PLAIN_NUMBER.test(text)) return text;
  return `'${text}`;
}

/**
 * One cell of the "Copy as TSV" text: `spreadsheetSafe`, and also neutralised
 * when it starts with a double quote — pasted TSV isn't escaped, and a
 * spreadsheet takes a leading `"` as a text qualifier and parses what's
 * inside (`"=HYPERLINK(…)"`). Tabs and line breaks become spaces.
 */
export function tsvCell(text: string, source?: unknown): string {
  const safe = spreadsheetSafe(text, source);
  return (/^\s*"/.test(safe) ? `'${safe}` : safe).replace(/[\t\n\r]+/g, ' ');
}

/**
 * CSV text for the given rows and columns: a header row, CRLF line breaks,
 * and a UTF-8 byte-order mark so Excel reads non-ASCII text correctly.
 * Headers and cells go through `spreadsheetSafe` before quoting.
 */
export function toCsv(
  rows: Record<string, unknown>[],
  columns: ExportColumn[],
  mode: ExportValueMode,
): string {
  const cols = exportColumnsFor(columns, mode);
  const lines = [cols.map((c) => csvField(spreadsheetSafe(c.header))).join(',')];
  for (const r of rows) {
    lines.push(
      cols
        .map((c) => csvField(spreadsheetSafe(exportCellText(r, c, mode), r[c.sourceCol])))
        .join(','),
    );
  }
  return '﻿' + lines.join('\r\n') + '\r\n';
}

/** Default file name: `<table>-<yyyyMMdd-HHmmss>.<ext>` in local time. */
export function exportFileName(table: string | undefined, ext: string, at = new Date()): string {
  const p = (n: number) => String(n).padStart(2, '0');
  const stamp =
    `${at.getFullYear()}${p(at.getMonth() + 1)}${p(at.getDate())}-` +
    `${p(at.getHours())}${p(at.getMinutes())}${p(at.getSeconds())}`;
  const base = (table || 'results').replace(/[^\w.-]+/g, '_');
  return `${base}-${stamp}.${ext}`;
}
