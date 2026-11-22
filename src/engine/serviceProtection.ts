// Dataverse service protection limits (#269): which limit a 429 hit, and
// what to do about it.
//
// Dataverse throttles a user (across every app signed in as them) on three
// limits, each over a sliding 5-minute window, and answers with 429 and one
// of these codes. The host doesn't pass response headers through, so the
// Retry-After duration isn't known here (phase 2 of #269, after #225).
// https://learn.microsoft.com/power-apps/developer/data-platform/api-limits

import type { ExecResult } from './dataverseExecutor';

export interface ServiceProtectionLimit {
  /** The Dataverse code, when the error had one. */
  code?: string;
  /** Short name of the limit, for a title. */
  limit: string;
  /** What happened and what to do, in a sentence or two. */
  advice: string;
}

const LIMITS: Record<string, Omit<ServiceProtectionLimit, 'code'>> = {
  '0x80072322': {
    limit: 'number of requests',
    advice:
      'Dataverse allows about 6,000 requests per user in 5 minutes. Wait a few minutes before sending more. Fewer, larger requests help: a bigger page size, or modest Bulk Write and $batch sizes (a batch is one request, but its work counts toward the execution-time limit).',
  },
  '0x80072321': {
    limit: 'execution time',
    advice:
      'Requests used about 20 minutes of server time in 5 minutes. Wait a few minutes, then make each request lighter: fewer columns ($select), simpler filters, smaller pages or smaller batches.',
  },
  '0x80072326': {
    limit: 'concurrent requests',
    advice:
      'Too many requests were running at once (the limit is 52 or more). Try again in a moment, once the others finish.',
  },
};

const GENERIC: Omit<ServiceProtectionLimit, 'code'> = {
  limit: 'too many requests',
  advice: 'Dataverse is throttling requests from your user. Wait a few minutes, then try again.',
};

/** Not just this tool: everything signed in as the same user counts. */
export const SERVICE_PROTECTION_SCOPE =
  'Dataverse counts these limits per user on each server, so other apps and flows signed in as you count too.';

export const SERVICE_PROTECTION_DOCS =
  'https://learn.microsoft.com/power-apps/developer/data-platform/api-limits';

/** The limit a failed result hit, or null when it isn't throttling. */
export function serviceProtectionLimit(
  result: Pick<ExecResult, 'status' | 'body'> | null | undefined,
): ServiceProtectionLimit | null {
  if (!result) return null;
  const err = (result.body as { error?: { code?: string; message?: string } } | null)?.error;
  const code = (err?.code ?? /\b0x8007232[126]\b/i.exec(err?.message ?? '')?.[0])?.toLowerCase();
  if (code && LIMITS[code]) return { code, ...LIMITS[code] };
  return result.status === 429 ? GENERIC : null;
}
