// How a run reads in Recent runs and the History dialog (#226).

import type { RecentRun } from '../state/readState';

/** "5 min ago", "3 hr ago", "2d ago". */
export function timeAgo(ts: number): string {
  const s = Math.floor((Date.now() - ts) / 1000);
  if (s < 60) return 'just now';
  const m = Math.floor(s / 60);
  if (m < 60) return `${m} min ago`;
  const h = Math.floor(m / 60);
  if (h < 24) return `${h} hr ago`;
  const d = Math.floor(h / 24);
  return `${d}d ago`;
}

/** Modes whose runs return rows. */
const ROW_MODES = new Set([
  'retrieve-multiple',
  'retrieve-single',
  'retrieve-nextlink',
  'predefined-query',
]);

/**
 * The tail of a Recent runs line: the mode's own summary, or a row count for
 * reads. A row count means nothing for a write ("204 · 1 rows") or for a
 * function / download, so those show the status alone unless the mode gave
 * a summary (#309).
 */
export function runDetail(r: Pick<RecentRun, 'modeId' | 'summary' | 'rowCount'>): string {
  if (r.summary) return ` · ${r.summary}`;
  if (!ROW_MODES.has(r.modeId)) return '';
  const n = r.rowCount;
  if (n === undefined) return ' · — rows';
  return ` · ${n.toLocaleString()} ${n === 1 ? 'row' : 'rows'}`;
}

/**
 * The accessible name of a restorable run: the line as shown (method, time,
 * status, detail), then the URL, so it starts with the visible text that
 * speech users say (WCAG 2.5.3, #387). The History dialog shows the URL
 * too, which this then contains as well.
 */
export function runAccessibleName(
  r: Pick<RecentRun, 'method' | 'url' | 'ts' | 'status' | 'modeId' | 'summary' | 'rowCount'>,
): string {
  return `${r.method} ${timeAgo(r.ts)} · ${r.status || 'error'}${runDetail(r)} ${r.url}, restore run`;
}
