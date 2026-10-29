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

/** The modes that send an allow-list rather than every header. */
export type HeaderPolicy = 'read' | 'relate' | 'function';

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
  function: IMPERSONATION,
};

const allowed = (policy: HeaderPolicy) => new Set(SENT_HEADERS[policy].map((n) => n.toLowerCase()));

export const isImpersonationHeader = (name: string): boolean =>
  IMPERSONATION.some((n) => n.toLowerCase() === name.trim().toLowerCase());

/**
 * `If-None-Match: null` was Retrieve Single's default before 1.5.0. It never
 * matches an etag, so it changes nothing; it isn't sent, so old saved
 * requests don't raise a consent prompt for it.
 */
const isNoOp = (name: string, value: string) =>
  name.trim().toLowerCase() === 'if-none-match' && value.trim().toLowerCase() === 'null';

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

/**
 * Enabled header rows Execute doesn't send under this policy (Code tab only).
 * Host-owned rows aren't listed: the host sends its own.
 */
export function unsentHeaderNames(
  policy: HeaderPolicy,
  headers?: Record<string, string>,
): string[] {
  const wire = toWireHeaders(headers);
  if (!wire) return [];
  const ok = allowed(policy);
  return Object.keys(wire).filter((name) => !ok.has(name.toLowerCase()));
}

/** Split a Prefer value on commas outside double quotes. */
function preferTokens(value: string): string[] {
  const out: string[] = [];
  let cur = '';
  let quoted = false;
  for (const ch of value) {
    if (ch === '"') quoted = !quoted;
    if (ch === ',' && !quoted) {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out.map((t) => t.trim()).filter(Boolean);
}

/**
 * The `Prefer` a read sends, or `undefined` when the host's own will do.
 * The host's is `odata.include-annotations="*"`, which a custom Prefer
 * replaces, so it always comes first here: the grid needs the annotations.
 * The user's other preferences follow, and the page size goes last: Dataverse
 * rejects it ahead of `include-annotations` (0x80060888 "Invalid prefer
 * header", confirmed live). Annotation and `return=representation` tokens
 * from a Prefer header row are dropped (`*` already covers the first; the
 * second is for writes).
 */
export function readPrefer(
  maxpagesize: number | null | undefined,
  preferRow?: string,
): string | undefined {
  const extra = preferTokens(preferRow ?? '').filter((t) => {
    const k = t.toLowerCase();
    return (
      !k.startsWith('odata.include-annotations') &&
      !k.startsWith('odata.maxpagesize') &&
      k !== 'return=representation'
    );
  });
  const pageSize = maxpagesize && maxpagesize > 0 ? [`odata.maxpagesize=${maxpagesize}`] : [];
  if (extra.length === 0 && pageSize.length === 0) return undefined;
  return ['odata.include-annotations="*"', ...extra, ...pageSize].join(',');
}

/**
 * What a read sends: the allow-listed rows, plus `Prefer` when the page size
 * or a Prefer header row asks for more than the host's default.
 */
export function readRequestHeaders(
  headers: Record<string, string> | undefined,
  maxpagesize: number | null | undefined,
): Record<string, string> | undefined {
  const preferRow = Object.entries(headers ?? {}).find(
    ([name]) => name.trim().toLowerCase() === 'prefer',
  )?.[1];
  const prefer = readPrefer(maxpagesize, preferRow);
  const sent = { ...allowedHeaders('read', headers), ...(prefer ? { Prefer: prefer } : {}) };
  return Object.keys(sent).length > 0 ? sent : undefined;
}
