// Live execution for the Bulk Write mode (#8): the availability check and
// the collection-bound action call. Kept apart from dataverseExecutor.ts so
// the mode stays a self-contained (lazy-loaded) addition.

import { dvHost } from '../host/pptbBridge';
import { mapError, sendable, type ExecResult } from './dataverseExecutor';
import {
  buildBulkWriteBody,
  supportMessagesFor,
  supportQueryPath,
  unsupportedMessage,
  type BulkOperation,
  type BulkSupport,
  type BulkWriteState,
} from './bulkWrite';

/**
 * Check `sdkmessagefilters` for the operation's message(s) on this table.
 * A failed lookup (no read access to sdkmessagefilter, offline) reports
 * `unknown` so the user can still try the request.
 */
export async function checkBulkSupport(table: string, op: BulkOperation): Promise<BulkSupport> {
  try {
    const missing: string[] = [];
    for (const message of supportMessagesFor(op)) {
      const res = await dvHost.queryData(supportQueryPath(message, table));
      if (!res?.value?.length) missing.push(message);
    }
    if (missing.length === 0) return { state: 'supported' };
    return { state: 'unsupported', message: unsupportedMessage(op, table, missing) };
  } catch (e) {
    return {
      state: 'unknown',
      message: `Couldn't check ${op} availability: ${e instanceof Error ? e.message : String(e)}`,
    };
  }
}

/**
 * POST /<set>/Microsoft.Dynamics.CRM.<op> with { Targets }. PPTB's execute()
 * posts to the entity set when given entityName without entityId (#55).
 *
 * `headers` is the mode's composed header map (user rows + bypass), the same
 * one the Code tab shows. PPTB 1.2.7+ sends it as `additionalHeaders` after
 * the user approves it (#257); older hosts block bypass via the advisory.
 */
export async function executeBulkWrite(
  s: BulkWriteState,
  headers?: Record<string, string>,
): Promise<ExecResult> {
  const { body, issues } = buildBulkWriteBody(s);
  if (issues.length > 0) {
    return mapError(
      { statusCode: 400, message: `Row ${issues[0].rowIndex + 1}: ${issues[0].message}` },
      0,
    );
  }
  const wire = sendable(headers);
  const t0 = performance.now();
  try {
    const value = await dvHost.execute(
      {
        operationName: s.operation,
        operationType: 'action',
        entityName: s.table,
        parameters: body,
      },
      wire,
    );
    const ms = Math.round(performance.now() - t0);
    const out = value && Object.keys(value).length > 0 ? value : null;
    return {
      status: out ? 200 : 204,
      statusText: out ? 'OK' : 'No Content',
      ms,
      bytes: out ? JSON.stringify(out).length : 0,
      ok: true,
      outcome: 'ok',
      body: out,
      headers: { 'OData-Version': '4.0' },
    };
  } catch (e) {
    return mapError(e, Math.round(performance.now() - t0), wire);
  }
}

/** `Ids` from a CreateMultipleResponse, or [] for 204 operations. */
export function idsFromResult(body: unknown): string[] {
  if (!body || typeof body !== 'object') return [];
  const ids = (body as { Ids?: unknown }).Ids;
  return Array.isArray(ids) ? ids.filter((x): x is string => typeof x === 'string') : [];
}
