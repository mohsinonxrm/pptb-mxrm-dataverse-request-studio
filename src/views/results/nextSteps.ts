// Next steps for failed and empty results (#291). Pure helpers: ErrorState
// and the grid's empty state turn these into buttons and hints.

import type { EntityListItem } from '../../host/metadataProvider';
import { statusTextFor, type ExecResult } from '../../engine/dataverseExecutor';

/** What a 403 message says is missing. Every field is optional: Dataverse
 *  names a privilege (`prvReadAccount`) in some messages and only the access
 *  right (`AccessRights: WriteAccess`) in others. */
export interface MissingPrivilege {
  /** e.g. `prvReadAccount`. */
  privilege?: string;
  /** e.g. `Read`, `Append To`. */
  access?: string;
  /** Table logical name, when the message gives one. */
  entity?: string;
}

const ACCESS_LABEL: Record<string, string> = {
  Create: 'Create',
  Read: 'Read',
  Write: 'Write',
  Delete: 'Delete',
  Append: 'Append',
  AppendTo: 'Append To',
  Assign: 'Assign',
  Share: 'Share',
};

/** The privilege a 403 message names, or null when it names none. */
export function missingPrivilege(message: string | undefined): MissingPrivilege | null {
  if (!message) return null;
  const out: MissingPrivilege = {};
  const prv = /\b(prv[A-Z][A-Za-z0-9_]*)/.exec(message)?.[1];
  if (prv) {
    out.privilege = prv;
    // AppendTo before Append, so `prvAppendToAccount` isn't read as Append.
    const m = /^prv(AppendTo|Append|Assign|Create|Delete|Read|Share|Write)(.*)$/.exec(prv);
    if (m) out.access = ACCESS_LABEL[m[1]];
  }
  const rights = /AccessRights:\s*([A-Za-z]+?)Access\b/.exec(message)?.[1];
  if (!out.access && rights && ACCESS_LABEL[rights]) out.access = ACCESS_LABEL[rights];
  const entity = /\bentity '([a-z0-9_]+)'/i.exec(message)?.[1];
  if (entity) out.entity = entity.toLowerCase();
  return out.privilege || out.access ? out : null;
}

/** One sentence naming the missing privilege, for the 403 card. */
export function describeMissingPrivilege(p: MissingPrivilege): string {
  const what = p.access
    ? `${p.access} access${p.entity ? ` on ${p.entity}` : ''}`
    : p.entity
      ? `access on ${p.entity}`
      : '';
  const name = p.privilege ? `the ${p.privilege} privilege` : '';
  const missing = name && what ? `${name} (${what})` : name || what;
  return `Your user is missing ${missing}. A Dataverse administrator can add it to one of your security roles.`;
}

/** The path segment Dataverse couldn't resolve ("Resource not found for the
 *  segment 'acounts'"), or null. */
export function notFoundSegment(message: string | undefined): string | null {
  return /Resource not found for the segment '([^']+)'/i.exec(message ?? '')?.[1] ?? null;
}

function editDistance(a: string, b: string): number {
  // Levenshtein with one row; strings are short entity-set names.
  const row = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    let prev = row[0];
    row[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const cur = row[j];
      row[j] = Math.min(row[j] + 1, row[j - 1] + 1, prev + (a[i - 1] === b[j - 1] ? 0 : 1));
      prev = cur;
    }
  }
  return row[b.length];
}

/**
 * The table whose entity set the user most likely meant by `segment`, from
 * `entities` (pass the scoped list, so nothing outside the query scope is
 * offered). Matches, in order: the entity set or logical name ignoring
 * case, then the closest entity set within a small edit distance. Null when
 * nothing is close, or two tables are equally close.
 */
export function suggestEntitySet(
  segment: string,
  entities: readonly EntityListItem[],
): EntityListItem | null {
  const seg = segment.toLowerCase();
  if (!seg) return null;
  const exact = entities.find(
    (e) => e.entitySetName.toLowerCase() === seg || e.logicalName.toLowerCase() === seg,
  );
  // An exact entity-set match means the segment wasn't the problem.
  if (exact) return exact.entitySetName === segment ? null : exact;
  const limit = Math.max(1, Math.floor(seg.length / 4));
  let best: EntityListItem | null = null;
  let bestDistance = Infinity;
  let tie = false;
  for (const e of entities) {
    const d = editDistance(seg, e.entitySetName.toLowerCase());
    if (d > limit) continue;
    if (d < bestDistance) {
      best = e;
      bestDistance = d;
      tie = false;
    } else if (d === bestDistance) {
      tie = true;
    }
  }
  return tie ? null : best;
}

/** `url` with its first path segment equal to `from` replaced by `to`. The
 *  query string is left alone. */
export function replacePathSegment(url: string, from: string, to: string): string {
  const q = url.indexOf('?');
  const path = q === -1 ? url : url.slice(0, q);
  const rest = q === -1 ? '' : url.slice(q);
  const parts = path.split('/');
  // Match the segment itself, or the segment of a key lookup: `acounts(…)`.
  const i = parts.findIndex((p) => p === from || p.startsWith(`${from}(`));
  if (i === -1) return url;
  parts[i] = to + parts[i].slice(from.length);
  return parts.join('/') + rest;
}

/**
 * Plain-text report of a failed run, for pasting into a ticket or chat.
 * Holds the method, URL, status, error code and message, and nothing from
 * the request headers, so no credentials or impersonation ids. Bearer
 * tokens are masked in case a host message ever quotes one.
 */
export function errorReport({
  result,
  method,
  requestUrl,
  operation,
  at = new Date(),
}: {
  result: ExecResult;
  method?: string;
  requestUrl?: string;
  /** What ran, for writes without a single request URL: "update on account". */
  operation?: string;
  at?: Date;
}): string {
  const err = (result.body as { error?: { code?: string; message?: string } } | null)?.error;
  const status = result.status
    ? `${result.status} ${result.statusText || statusTextFor(result.status) || ''}`.trim()
    : `none (${result.statusText || 'the host reported no HTTP status'})`;
  const mask = (s: string) => s.replace(/\bBearer\s+[A-Za-z0-9._~+/=-]+/gi, 'Bearer [removed]');
  const lines = [
    'Dataverse Request Studio: error report',
    `Time: ${at.toISOString()}`,
    ...(operation ? [`Operation: ${operation}`] : []),
    ...(requestUrl ? [`Request: ${method ? `${method} ` : ''}${mask(requestUrl)}`] : []),
    `Status: ${status}`,
    ...(err?.code ? [`Code: ${err.code}`] : []),
    ...(err?.message ? [`Message: ${mask(err.message)}`] : []),
    `Duration: ${result.ms} ms`,
  ];
  return lines.join('\n');
}
