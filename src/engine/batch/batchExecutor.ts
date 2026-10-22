// $batch executor (#9) — sends a compiled batch through PPTB 1.2.7's
// `executeBatch` / `executeTransaction` and maps the host's per-operation
// results back onto the composer's items.
//
// Host behaviour this relies on (desktop-app v1.2.7,
// src/main/utilities/dataverseBatch.ts + managers/dataverseManager.ts):
//   • the host writes the multipart body itself; we hand it
//     { method, url (relative to /api/data/v9.2), headers, body, contentId }
//   • results come back parsed, correlated by Content-ID, in request order
//   • `executeTransaction` returns only the failing response when a change
//     set rolls back (it often has no Content-ID)
//   • `executeBatch` without `Prefer: odata.continue-on-error` throws when
//     Dataverse stops early (fewer results than requests) — so the partial
//     results are lost; we say so and point at Continue on error
//   • outer `headers` and per-request headers go through the host's
//     additional-headers consent prompt; a decline rejects the call before
//     anything is sent

import { dvHost, hostSupportsBatch, isEmbedded } from '../../host/pptbBridge';
import type { CompiledItem, DispatchSegment } from './batchModel';

export interface HostBatchResult {
  contentId?: string;
  status: number;
  statusText?: string;
  headers: Record<string, string>;
  body?: unknown;
}

export type ItemOutcome =
  /** Dataverse answered this request. */
  | 'done'
  /** The change set failed, so this request was rolled back (or never applied). */
  | 'rolled-back'
  /** An earlier failure stopped the batch before this request ran. */
  | 'not-run'
  /** The host call failed as a whole, so there's no per-request result. */
  | 'unknown';

export interface BatchItemResult {
  itemId: string;
  contentId: string;
  outcome: ItemOutcome;
  status?: number;
  statusText?: string;
  headers?: Record<string, string>;
  body?: unknown;
}

export interface SegmentResult {
  transaction: boolean;
  contentIds: string[];
  ms: number;
  /** False when any request failed, the change set rolled back, or the call threw. */
  ok: boolean;
  /** Host / transport error — nothing in this segment has a per-request result. */
  error?: string;
  /** A change-set error response the host couldn't tie to one request. */
  changesetError?: HostBatchResult;
  /** True when this segment was never sent because an earlier one failed. */
  skipped?: boolean;
  /**
   * The host call threw, but its message shows Dataverse answered with a
   * failed request (the host couldn't parse the partial response) — so at
   * least one request in the segment failed.
   */
  requestFailed?: boolean;
}

export interface BatchRunResult {
  ms: number;
  ok: boolean;
  items: BatchItemResult[];
  segments: SegmentResult[];
}

const isFailure = (r: HostBatchResult) => r.status >= 400;

/**
 * Map one host call's results onto its items. Pure — unit tested.
 */
export function mapSegmentResults(
  segment: DispatchSegment,
  results: HostBatchResult[],
): { items: BatchItemResult[]; segment: SegmentResult } {
  const byId = new Map<string, HostBatchResult>();
  const orphans: HostBatchResult[] = [];
  for (const r of results) {
    if (r.contentId && !byId.has(r.contentId)) byId.set(r.contentId, r);
    else orphans.push(r);
  }
  // No Content-IDs at all but one result per request: the host already
  // ordered them, so match by position.
  if (byId.size === 0 && orphans.length === segment.items.length) {
    segment.items.forEach((c, i) => byId.set(c.contentId, orphans[i]));
    orphans.length = 0;
  }

  const anyFailure = results.some(isFailure);
  const rolledBack = segment.transaction && anyFailure;
  const items = segment.items.map((c): BatchItemResult => {
    const r = byId.get(c.contentId);
    if (!r) {
      return {
        itemId: c.item.id,
        contentId: c.contentId,
        outcome: rolledBack ? 'rolled-back' : 'not-run',
      };
    }
    return {
      itemId: c.item.id,
      contentId: c.contentId,
      // A 2xx inside a failed change set was undone with the rest.
      outcome: rolledBack && !isFailure(r) ? 'rolled-back' : 'done',
      status: r.status,
      statusText: r.statusText,
      headers: r.headers,
      body: r.body,
    };
  });
  const changesetError = orphans.find(isFailure);
  return {
    items,
    segment: {
      transaction: segment.transaction,
      contentIds: segment.items.map((c) => c.contentId),
      ms: 0,
      ok: !anyFailure && items.every((i) => i.outcome === 'done'),
      changesetError,
    },
  };
}

/**
 * What Recent Runs shows for a batch run (#204). `status` is the outer
 * `$batch` response — 200 whenever Dataverse answered at least one call,
 * 0 when none got an answer (host / transport error, consent declined) —
 * never a single part's status. `summary` counts the parts, e.g.
 * "3 requests · 1 failed". Pure — unit tested.
 */
export function summarizeBatchRun(res: BatchRunResult): {
  status: number;
  summary: string;
  succeeded: number;
} {
  const answered = res.segments.some((s) => !s.skipped && (!s.error || s.requestFailed));
  const n = res.items.length;
  let succeeded = 0;
  let failed = 0;
  let rolledBack = 0;
  let notRun = 0;
  let unknown = 0;
  for (const i of res.items) {
    if (i.outcome === 'done') {
      if ((i.status ?? 0) >= 400) failed++;
      else succeeded++;
    } else if (i.outcome === 'rolled-back') rolledBack++;
    else if (i.outcome === 'not-run') notRun++;
    else unknown++;
  }
  const parts = [`${n} request${n === 1 ? '' : 's'}`];
  if (failed) parts.push(`${failed} failed`);
  if (rolledBack) parts.push(`${rolledBack} rolled back`);
  if (notRun) parts.push(`${notRun} not run`);
  if (unknown) parts.push(`${unknown} no result`);
  if (n > 0 && succeeded === n) parts.push('all succeeded');
  return { status: answered ? 200 : 0, summary: parts.join(' · '), succeeded };
}

const CONSENT_DECLINED_RE = /additional headers were rejected|header consent/i;
const PARTIAL_RE = /invalid content id correlation|missing operation results/i;
const CORRELATION_RE = /invalid content id correlation/i;

/**
 * Drop the Electron IPC wrapper the host adds to a rejection —
 * `Error invoking remote method 'dataverse.executeBatch': Error: …` — so the
 * message starts with the host's own text (#198).
 */
export function stripIpcPrefix(message: string): string {
  return message
    .replace(/^\s*Error invoking remote method '[^']*':\s*/i, '')
    .replace(/^(?:\w*Error:\s*)+/, '')
    .trim();
}

/** Plain-language message for a host rejection. */
export function describeHostError(
  rawMessage: string,
  ctx: { transaction: boolean; continueOnError: boolean; sentHeaders: boolean },
): string {
  const message = stripIpcPrefix(rawMessage);
  if (CONSENT_DECLINED_RE.test(message)) {
    return (
      'Power Platform ToolBox didn’t send this batch because its request headers weren’t ' +
      'approved, so nothing reached Dataverse. Send again and allow the headers when ToolBox ' +
      `asks, or remove them. Host message: ${message}`
    );
  }
  // #198 — PPTB 1.2.8's batch parser throws on a continue-on-error response
  // in which a failed request is followed by more requests (`[bad, good]`
  // fails; `[bad]` and `[good, bad]` work). Dataverse ran the batch; the
  // host just can't read the answer.
  if (!ctx.transaction && ctx.continueOnError && CORRELATION_RE.test(message)) {
    return (
      'A request in this batch failed, and Power Platform ToolBox couldn’t read the ' +
      'continue-on-error response because more requests followed the failed one (a known ' +
      'ToolBox batch-parser bug). Dataverse still ran every request: the ones that succeeded ' +
      'were applied. To see each result, put requests that might fail last in the batch, or ' +
      `send them on their own. Host message: ${message}`
    );
  }
  if (!ctx.transaction && !ctx.continueOnError && PARTIAL_RE.test(message)) {
    return (
      'A request failed and Dataverse stopped the batch there. Requests before it were ' +
      'applied; the rest didn’t run. This version of Power Platform ToolBox can’t return ' +
      'partial results without continue-on-error — turn on Continue on error to see each ' +
      `request’s result. Host message: ${message}`
    );
  }
  return message;
}

/**
 * Run the plan: one host call per segment, in order. Without
 * continue-on-error the run stops after the first failed segment, which is
 * what a single native $batch does.
 */
export async function executeBatchPlan(
  plan: DispatchSegment[],
  opts: { continueOnError: boolean },
): Promise<BatchRunResult> {
  const t0 = performance.now();
  const items: BatchItemResult[] = [];
  const segments: SegmentResult[] = [];
  let stopped = false;

  for (const seg of plan) {
    const contentIds = seg.items.map((c) => c.contentId);
    if (stopped) {
      segments.push({ transaction: seg.transaction, contentIds, ms: 0, ok: false, skipped: true });
      for (const c of seg.items) {
        items.push({ itemId: c.item.id, contentId: c.contentId, outcome: 'not-run' });
      }
      continue;
    }
    const requests = seg.items.map(toHostRequest);
    // The continue-on-error preference only matters to a call holding
    // independent requests; a change set is all-or-nothing anyway, and
    // leaving it off spares a consent prompt.
    const outer =
      !seg.transaction && opts.continueOnError ? { Prefer: 'odata.continue-on-error' } : undefined;
    const s0 = performance.now();
    try {
      const results = seg.transaction
        ? await dvHost.executeTransaction(requests, outer)
        : await dvHost.executeBatch(requests, outer);
      const mapped = mapSegmentResults(seg, results as HostBatchResult[]);
      mapped.segment.ms = Math.round(performance.now() - s0);
      segments.push(mapped.segment);
      items.push(...mapped.items);
      if (!mapped.segment.ok && !opts.continueOnError) stopped = true;
    } catch (e) {
      const raw = e instanceof Error ? e.message : String(e);
      const sentHeaders =
        !!outer || requests.some((r) => r.headers && Object.keys(r.headers).length);
      segments.push({
        transaction: seg.transaction,
        contentIds,
        ms: Math.round(performance.now() - s0),
        ok: false,
        requestFailed: PARTIAL_RE.test(raw),
        error: describeHostError(raw, {
          transaction: seg.transaction,
          continueOnError: opts.continueOnError,
          sentHeaders,
        }),
      });
      for (const c of seg.items) {
        items.push({
          itemId: c.item.id,
          contentId: c.contentId,
          outcome: seg.transaction ? 'rolled-back' : 'unknown',
        });
      }
      if (!opts.continueOnError) stopped = true;
    }
  }

  return {
    ms: Math.round(performance.now() - t0),
    ok: segments.every((s) => s.ok),
    items,
    segments,
  };
}

export function toHostRequest(c: CompiledItem): DataverseAPI.BatchRequest {
  const w = c.wire!;
  const req: DataverseAPI.BatchRequest = { method: w.method, url: w.path, contentId: c.contentId };
  if (Object.keys(w.headers).length > 0) req.headers = { ...w.headers };
  if (w.body !== undefined) req.body = w.body;
  return req;
}

/** Why Execute is unavailable in this environment, or null when it can run. */
export function batchUnavailableReason(): string | null {
  if (!isEmbedded())
    return 'Run Dataverse Request Studio inside Power Platform ToolBox to send a batch.';
  if (!hostSupportsBatch())
    return 'Sending a $batch requires Power Platform ToolBox 1.2.7 or later.';
  return null;
}
