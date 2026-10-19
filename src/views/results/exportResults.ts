// Exporting Retrieve-Multiple results to CSV / JSON (#15).
//
// The CSV mirrors what the grid shows: the same columns in the same order,
// with the same headers, over the rows the grid currently lists (so an
// active row search narrows the export too). "Shown" values match the grid
// (formatted values where Dataverse sent them, raw for "(Raw)" twins and
// when the grid is in raw mode); "raw" exports the underlying values
// (GUIDs, numbers, ISO dates). JSON exports the OData envelope as-is.
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

/**
 * CSV text for the given rows and columns: a header row, CRLF line breaks,
 * and a UTF-8 byte-order mark so Excel reads non-ASCII text correctly.
 */
export function toCsv(
  rows: Record<string, unknown>[],
  columns: ExportColumn[],
  mode: ExportValueMode,
): string {
  const cols = exportColumnsFor(columns, mode);
  const lines = [cols.map((c) => csvField(c.header)).join(',')];
  for (const r of rows) lines.push(cols.map((c) => csvField(exportCellText(r, c, mode))).join(','));
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
