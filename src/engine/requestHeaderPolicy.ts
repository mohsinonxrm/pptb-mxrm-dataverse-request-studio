// Which request headers Execute sends through PPTB (#109).
//
// The write group, Bulk Write, Execute Action and Execute Workflow send every
// header row the user enabled (minus the host-owned ones). The read group,
// Execute Function and Associate / Disassociate send an allow-list: the
// headers that change what those requests do (impersonation, conditional
// retrieve, consistency, bypass). Anything else stays in the Code tab, and
// the Headers pane says so. Each header set goes through the host's consent
// prompt once, so a plain read that only uses the defaults sends none and
// never prompts.

import { isCredentialHeader } from './sensitiveHeaders';

// Headers PPTB's Dataverse client always sets itself. Its client spreads
// `additionalHeaders` over its own, so sending these would override them
// (and prompt for consent for nothing). `Prefer` is host-owned too: the host
// always sends `return=representation`, and a custom Prefer replaces the
// rest of the host's own (on reads, `odata.include-annotations="*"`). Reads
// send one through `readPrefer`, which puts the annotations back.
const HOST_OWNED_HEADERS = new Set([
  'authorization',
  'accept',
  'content-type',
  'content-length',
  'odata-version',
  'odata-maxversion',
  'prefer',
]);

/** True for a header the host builds itself (never sent as an additional header). */
export const isHostOwnedHeader = (name: string): boolean =>
  HOST_OWNED_HEADERS.has(name.trim().toLowerCase());

/**
 * The part of a mode's composed headers (the Code tab map) that we hand to
 * the host as `additionalHeaders`: host-owned headers removed, names
 * de-duplicated case-insensitively (the later row wins, as the composers
 * append precondition/bypass rows after the user's). `undefined` when
 * nothing is left — the call then stays header-free and never prompts.
 */
export function toWireHeaders(
  headers?: Record<string, string>,
): Record<string, string> | undefined {
  if (!headers) return undefined;
  const byName = new Map<string, [string, string]>();
  for (const [rawName, value] of Object.entries(headers)) {
    const name = rawName.trim();
    const key = name.toLowerCase();
    // Credential headers (Cookie, Proxy-Authorization, …) are never sent (#217).
    if (!key || HOST_OWNED_HEADERS.has(key) || isCredentialHeader(key)) continue;
    byName.delete(key);
    byName.set(key, [name, value]);
  }
  return byName.size > 0 ? Object.fromEntries(byName.values()) : undefined;
}

/**
 * The modes that send an allow-list rather than every header: the reads,
 * Associate / Disassociate, and `impersonation` for Execute Function and
 * the Bulk Delete Job.
 */
export type HeaderPolicy = 'read' | 'relate' | 'impersonation';

/** Impersonation: by systemuserid, or by Microsoft Entra object id. */
const IMPERSONATION = ['MSCRMCallerID', 'CallerObjectId'];
/** Bypass custom logic / flows (the Bypass pane's headers). */
const BYPASS = [
  'MSCRM.BypassCustomPluginExecution',
  'MSCRM.BypassBusinessLogicExecution',
  'MSCRM.SuppressCallbackRegistrationExpanderJob',
];

/** The headers each policy sends, by name as the Headers pane writes them. */
export const SENT_HEADERS: Record<HeaderPolicy, readonly string[]> = {
  read: [...IMPERSONATION, 'If-None-Match', 'Consistency'],
  relate: [...IMPERSONATION, ...BYPASS],
  impersonation: IMPERSONATION,
};

const allowed = (policy: HeaderPolicy) => new Set(SENT_HEADERS[policy].map((n) => n.toLowerCase()));

const same = (a: string, b: string) => a.trim().toLowerCase() === b.toLowerCase();

export const isImpersonationHeader = (name: string): boolean =>
  IMPERSONATION.some((n) => same(name, n));

export const isBypassHeader = (name: string): boolean => BYPASS.some((n) => same(name, n));

/**
 * `If-None-Match: null` was Retrieve Single's default before 1.5.0. It never
 * matches an etag, so it changes nothing; it isn't sent, so old saved
 * requests don't raise a consent prompt for it.
 */
const isNoOp = (name: string, value: string) =>
  same(name, 'if-none-match') && value.trim().toLowerCase() === 'null';

/** The allow-listed part of a mode's composed headers; `undefined` when none. */
export function allowedHeaders(
  policy: HeaderPolicy,
  headers?: Record<string, string>,
): Record<string, string> | undefined {
  const wire = toWireHeaders(headers);
  if (!wire) return undefined;
  const ok = allowed(policy);
  const out = Object.entries(wire).filter(
    ([name, value]) => ok.has(name.toLowerCase()) && !isNoOp(name, value),
  );
  return out.length > 0 ? Object.fromEntries(out) : undefined;
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const ZERO_GUID = '00000000-0000-0000-0000-000000000000';

/**
 * Why the impersonation headers can't be sent as they are, or null:
 * both kinds at once (Dataverse documents one or the other), or a value that
 * isn't a user's id — the presets' all-zero placeholder included.
 */
export function impersonationProblem(headers?: Record<string, string>): string | null {
  const rows = Object.entries(toWireHeaders(headers) ?? {}).filter(([n]) =>
    isImpersonationHeader(n),
  );
  if (rows.length > 1) {
    return `Turn on one impersonation header, not both (${rows.map(([n]) => n).join(' and ')}).`;
  }
  const [row] = rows;
  if (row && (!GUID_RE.test(row[1].trim()) || row[1].trim() === ZERO_GUID)) {
    return `${row[0]} needs the id of the user to run as (a GUID), not ${JSON.stringify(row[1])}.`;
  }
  return null;
}

/**
 * The `Prefer` a read sends for its page size (#219), or `undefined`. The
 * host's own is `odata.include-annotations="*"`, which a custom Prefer
 * replaces, so it's sent again first: the grid needs the annotations. The
 * page size goes last: Dataverse rejects it ahead of `include-annotations`
 * (0x80060888 "Invalid prefer header", confirmed live). Other preferences
 * stay in the Code tab, like the rest of the Prefer pane.
 */
export function readPrefer(maxpagesize: number | null | undefined): string | undefined {
  return maxpagesize && maxpagesize > 0
    ? `odata.include-annotations="*",odata.maxpagesize=${maxpagesize}`
    : undefined;
}

/** What a read sends: the allow-listed rows, plus `Prefer` for a page size. */
export function readRequestHeaders(
  headers: Record<string, string> | undefined,
  maxpagesize: number | null | undefined,
): Record<string, string> | undefined {
  const prefer = readPrefer(maxpagesize);
  const sent = { ...allowedHeaders('read', headers), ...(prefer ? { Prefer: prefer } : {}) };
  return Object.keys(sent).length > 0 ? sent : undefined;
}
