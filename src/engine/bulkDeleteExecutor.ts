// Live calls for the Bulk Delete mode (#10): the pre-confirmation count,
// the BulkDelete action, the current user (for "notify me"), and job
// status polling. Kept apart from dataverseExecutor.ts so the mode stays a
// self-contained (lazy-loaded) addition.

import { dvHost } from '../host/pptbBridge';
import { mapError, type ExecResult } from './dataverseExecutor';
import {
  COUNT_CAP,
  buildBulkDeleteBody,
  bulkDeleteOperationPath,
  countPath,
  jobStatusPath,
  type BulkDeleteState,
} from './bulkDelete';

interface HostError {
  message?: string;
  code?: string;
  statusCode?: number;
  status?: number;
  innererror?: { message?: string };
}

const errMessage = (e: unknown) => {
  const h: HostError = typeof e === 'object' && e !== null ? (e as HostError) : {};
  return h.message ?? h.innererror?.message ?? String(e);
};

export type MatchCount =
  | { ok: true; count: number; capped: boolean }
  | { ok: false; error: string };

/** Rows the job would delete right now, via `$count` on the same filter. */
export async function countMatching(s: BulkDeleteState): Promise<MatchCount> {
  const path = countPath(s);
  if (!path) return { ok: false, error: 'Table metadata not loaded.' };
  try {
    const res = (await dvHost.queryData(path)) as { '@odata.count'?: number; value?: unknown[] };
    const n = res['@odata.count'];
    if (typeof n !== 'number') {
      return { ok: false, error: 'The response had no @odata.count.' };
    }
    return { ok: true, count: n, capped: n >= COUNT_CAP };
  } catch (e) {
    return { ok: false, error: errMessage(e) };
  }
}

export async function executeBulkDelete(s: BulkDeleteState): Promise<ExecResult> {
  const { body, problems } = buildBulkDeleteBody(s, new Date());
  if (!body || problems.length) {
    return mapError({ statusCode: 400, message: problems[0] ?? 'Request incomplete.' }, 0);
  }
  const t0 = performance.now();
  try {
    const value = await dvHost.execute({
      operationName: 'BulkDelete',
      operationType: 'action',
      parameters: body,
    });
    const ms = Math.round(performance.now() - t0);
    return {
      status: 200,
      statusText: 'OK',
      ms,
      bytes: JSON.stringify(value ?? {}).length,
      ok: true,
      outcome: 'ok',
      body: value ?? {},
      headers: { 'OData-Version': '4.0' },
    };
  } catch (e) {
    return mapError(e, Math.round(performance.now() - t0));
  }
}

export function jobIdFromResult(body: unknown): string | null {
  if (!body || typeof body !== 'object') return null;
  const id = (body as { JobId?: unknown }).JobId;
  return typeof id === 'string' ? id : null;
}

export async function currentUserId(): Promise<string> {
  const r = await dvHost.execute({ operationName: 'WhoAmI', operationType: 'function' });
  const id = (r as { UserId?: unknown }).UserId;
  if (typeof id !== 'string') throw new Error('WhoAmI returned no UserId.');
  return id.toLowerCase();
}

export interface JobStatus {
  job: Record<string, unknown> | null;
  operation: Record<string, unknown> | null;
  error?: string;
}

export async function fetchJobStatus(jobId: string): Promise<JobStatus> {
  try {
    const jobRes = (await dvHost.queryData(jobStatusPath(jobId))) as unknown;
    // queryData may hand back the entity itself or wrap it in { value: [...] }.
    const job = Array.isArray((jobRes as { value?: unknown[] })?.value)
      ? (((jobRes as { value: Record<string, unknown>[] }).value[0] ?? null) as Record<
          string,
          unknown
        > | null)
      : ((jobRes as Record<string, unknown>) ?? null);
    let operation: Record<string, unknown> | null = null;
    try {
      const op = await dvHost.queryData(bulkDeleteOperationPath(jobId));
      operation = op?.value?.[0] ?? null;
    } catch {
      // bulkdeleteoperation needs its own read privilege; the job row is enough.
    }
    return { job, operation };
  } catch (e) {
    return { job: null, operation: null, error: errMessage(e) };
  }
}
