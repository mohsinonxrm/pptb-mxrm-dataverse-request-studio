// Merging `@odata.nextLink` pages into a result, and the Retrieve-all loop
// (#223). Kept free of React so the budget, Stop and stale-run rules are
// testable on their own.
import type { ExecResult } from '../../engine/dataverseExecutor';

type Row = Record<string, unknown>;

/** Rows fetched since the last merge, and the nextLink that follows them. */
export interface PageBatch {
  rows: Row[];
  /** Bytes received for these pages. */
  bytes: number;
  /** Absent once the last page is in. */
  nextLink?: string;
}

/** A Retrieve all pauses after this many rows, so it can't fill memory unasked. */
export const RETRIEVE_ALL_ROW_BUDGET = 50_000;
/** …or this many pages, for very small page sizes. */
export const RETRIEVE_ALL_PAGE_BUDGET = 1_000;
/** Pages reach the grid at most this often, so merging stays close to linear. */
export const RETRIEVE_ALL_COMMIT_MS = 500;

export function pageOf(res: ExecResult): PageBatch {
  const body = res.body as { value?: Row[]; '@odata.nextLink'?: string } | null;
  return { rows: body?.value ?? [], bytes: res.bytes, nextLink: body?.['@odata.nextLink'] };
}

/**
 * Appends a batch to the result: the envelope is kept, `value` grows and
 * the nextLink is replaced (or dropped once exhausted). The size adds the
 * batch's bytes rather than re-serializing everything loaded so far.
 */
export function mergePage(r: ExecResult, batch: PageBatch): ExecResult {
  const prevBody = (r.body as Record<string, unknown> | null) ?? {};
  const prevRows = (prevBody.value as Row[] | undefined) ?? [];
  const body: Record<string, unknown> = { ...prevBody, value: prevRows.concat(batch.rows) };
  if (batch.nextLink) body['@odata.nextLink'] = batch.nextLink;
  else delete body['@odata.nextLink'];
  return { ...r, body, bytes: r.bytes + batch.bytes };
}

/**
 * 'stopped': the user pressed Stop, or a newer scroll / Retrieve all took
 * over. 'superseded': a newer Execute or a reset.
 */
export type PagingStatus = 'current' | 'stopped' | 'superseded';

export type RetrieveAllOutcome =
  | { kind: 'done'; rows: number }
  /** The row or page budget was reached; more pages remain. */
  | { kind: 'paused'; rows: number }
  | { kind: 'stopped'; rows: number }
  | { kind: 'failed'; rows: number; error: ExecResult }
  /** Nothing was merged after the newer run started. */
  | { kind: 'superseded' };

/**
 * Follows nextLinks from `nextLink` until the last page, a failure, Stop,
 * or the budget. Pages are collected and handed to `commit` in batches,
 * at most every `commitMs`, and before returning, except after Stop or a
 * newer run, when nothing more is merged. The nextLink in each batch is the
 * one after its rows, so a later Retrieve all resumes exactly where this one
 * left off. `rows` counts the rows this call merged.
 */
export async function retrieveAll(opts: {
  nextLink: string;
  fetchPage: (url: string) => Promise<ExecResult>;
  commit: (batch: PageBatch) => void;
  status: () => PagingStatus;
  rowBudget?: number;
  pageBudget?: number;
  commitMs?: number;
  now?: () => number;
}): Promise<RetrieveAllOutcome> {
  const {
    fetchPage,
    commit,
    status,
    rowBudget = RETRIEVE_ALL_ROW_BUDGET,
    pageBudget = RETRIEVE_ALL_PAGE_BUDGET,
    commitMs = RETRIEVE_ALL_COMMIT_MS,
    now = () => performance.now(),
  } = opts;
  let next: string | undefined = opts.nextLink;
  let pending: Row[] = [];
  let pendingBytes = 0;
  let pendingPages = 0;
  let loaded = 0;
  let pages = 0;
  let lastCommit = now();

  const flush = () => {
    if (pendingPages === 0) return;
    commit({ rows: pending, bytes: pendingBytes, nextLink: next });
    pending = [];
    pendingBytes = 0;
    pendingPages = 0;
    lastCommit = now();
  };

  while (next) {
    const res = await fetchPage(next);
    const st = status();
    if (st === 'superseded') return { kind: 'superseded' };
    // After Stop, the held pages are dropped as well. Stop re-enables
    // Retrieve all and scrolling, and those start again from the merged
    // nextLink; merging held pages now would add them a second time.
    if (st === 'stopped') return { kind: 'stopped', rows: loaded - pending.length };
    if (!res.ok) {
      flush();
      return { kind: 'failed', rows: loaded, error: res };
    }
    const page = pageOf(res);
    for (const row of page.rows) pending.push(row);
    pendingBytes += page.bytes;
    pendingPages += 1;
    loaded += page.rows.length;
    pages += 1;
    next = page.nextLink;
    if (!next) break;
    if (loaded >= rowBudget || pages >= pageBudget) {
      flush();
      return { kind: 'paused', rows: loaded };
    }
    if (now() - lastCommit >= commitMs) flush();
  }
  flush();
  return { kind: 'done', rows: loaded };
}
