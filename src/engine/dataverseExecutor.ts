// Live executor — routes each studio mode's "Execute" through the PPTB
// host's `window.dataverseAPI` surface (via the `dvHost` wrapper in
// `pptbBridge.ts`).
//
// Each `execute*` takes a mode's state (plus the headers it composed) and
// resolves to an `ExecResult` with the host's real status, never a thrown
// error; `runtime.ts` dispatches to them.
//
// Request headers: the write group (Create / Update / Upsert / Delete /
// Merge), Execute Action and Execute Workflow (#266) pass the headers their
// mode composed for the Code tab. On PPTB
// 1.2.7+ the precondition (`If-Match` / `If-None-Match`), bypass (`MSCRM.*`)
// and user-added headers go out as `additionalHeaders` once the user
// approves them in the host. Headers the host builds itself (see
// requestHeaderPolicy) stay with the host. Older hosts drop the map; there
// `hostHeaderAdvisories` blocks the requests whose meaning depends on it.
// The read group, Execute Function and Associate / Disassociate send an
// allow-list of their headers instead (impersonation, conditional retrieve,
// consistency, bypass; #109), and reads a `Prefer` that keeps the host's
// annotations (`readHeaders`).
//
// What's expressly NOT implemented:
//   - Binary (file/image) chunked upload — host doesn't expose this API.
//     ManageFile / ManageImage modes return `host-not-supported` from the
//     live executor so the user can BUILD the request and copy out, but
//     Execute is gated with a clear error.
//
// All other modes use the host API surface directly. Each function:
//   1. Times the call (start clock).
//   2. Routes to the correct `dvHost.*` method.
//   3. Wraps the result in the studio's `ExecResult` envelope.
//   4. Translates host errors to the same shape the mock produces
//      so `ResultsView` rendering doesn't need to know which executor
//      ran.

import {
  dvHost,
  HostNotAvailableError,
  hostSupportsAdditionalHeaders,
  isEmbedded,
} from '../host/pptbBridge';
import { metadata } from '../host/metadataProvider';
import { cleanHostMessage } from './hostErrorText';
import { jsonBytes } from './byteSize';
import { hasCollectionInvolvedNestedExpand } from '../editors/expand/expandTree';
import {
  buildRetrieveMultiple,
  COUNT_ONLY_PREFER,
  buildRetrieveSingle,
  buildCreateBody,
  buildUpdateBody,
  buildUpsertBody,
  buildMergeBody,
  buildAssociateRequests,
  buildDisassociateRequests,
  PATCH_EXISTING_HEADERS,
} from './urlBuilder';
import {
  buildExecuteActionBody,
  buildExecuteFunctionWirePath,
  invalidFunctionParamReason,
} from './executeBuilders';
import type {
  RetrieveMultipleState,
  RetrieveSingleState,
  PredefinedQueryState,
  RetrieveNextLinkState,
} from '../state/readState';
import type {
  CreateState,
  UpdateState,
  UpsertState,
  DeleteState,
  MergeState,
} from '../state/writeState';
import type { AssociateState, DisassociateState } from '../state/relateState';
import type {
  ExecuteActionState,
  ExecuteFunctionState,
  ExecuteWorkflowState,
} from '../state/executeState';
import type {
  ManageFileState,
  ManageImageState,
  ManageAttachmentState,
} from '../state/binaryState';
// Action lookup goes through the live CSDL provider. PPTB-only.
import { actions } from '../host/csdlProvider';
import {
  allowedHeaders,
  impersonationProblem,
  isBypassHeader,
  isImpersonationHeader,
  readRequestHeaders,
  toWireHeaders,
} from './requestHeaderPolicy';

export { toWireHeaders };
import { invalidUrlPartReason } from '../primitives/urlParts';

// ── Public types ─────────────────────────────────────────────────────────

/** Coarse outcome bucket on every `ExecResult`. ResultsView inspects this
 *  to render the "no rows" empty-state distinctly from a network success
 *  that happens to return a populated body. Derived from `status` by
 *  `outcomeFor`; `'error'` covers every other failure (4xx, 5xx, 501, and
 *  host errors whose status is unknown). */
export type ExecOutcome = 'ok' | '404' | '401' | '403' | 'error' | 'empty' | 'partial';

/** Normalized execution result handed back to every mode. */
export interface ExecResult {
  /** HTTP status. `0` when the host rejected without reporting one. */
  status: number;
  statusText: string;
  ms: number;
  bytes: number;
  ok: boolean;
  outcome: ExecOutcome;
  body: unknown;
  /** Response headers. The PPTB host doesn't expose them, so this is `{}`
   *  for every live result; the Headers tab says so instead of showing
   *  made-up values. */
  headers: Record<string, string>;
  /** Associate / Disassociate on a collection-valued nav property: one
   *  entry per target, in request order (#99). Absent for every other op. */
  targets?: RelateTargetResult[];
  /** Count only (#272): Dataverse stopped counting at its limit (5,000, or
   *  500 for elastic tables). Absent when the response didn't say. */
  countLimitExceeded?: boolean;
}

/**
 * What happened to one target of a multi-target Associate / Disassociate.
 *   • `done`    — linked / unlinked by this run.
 *   • `skipped` — Associate only: already linked, so no request was sent.
 *                 Disassociate never skips (see executeDisassociate).
 *   • `failed`  — the request failed; `status` / `error` come from the same
 *                 mapping as every other failure.
 */
export interface RelateTargetResult {
  targetId: string;
  state: 'done' | 'skipped' | 'failed';
  status?: number;
  statusText?: string;
  error?: { code?: string; message: string };
}

/** Per-state counts for a relate result's targets (all zero when absent). */
export function relateTargetCounts(r: ExecResult | null | undefined): {
  done: number;
  skipped: number;
  failed: number;
} {
  const c = { done: 0, skipped: 0, failed: 0 };
  for (const t of r?.targets ?? []) c[t.state]++;
  return c;
}

/** Requests that went through: the `done` targets of a per-target result,
 *  otherwise `requested` when the run succeeded. `null` when nothing ran or
 *  the whole run failed. */
export function relateDoneCount(r: ExecResult | null, requested: number): number | null {
  if (!r) return null;
  const done = r.targets ? relateTargetCounts(r).done : r.ok ? requested : 0;
  return r.ok || done > 0 ? done : null;
}

// ── Helpers ──────────────────────────────────────────────────────────────

const STATUS_TEXT: Record<number, string> = {
  200: 'OK',
  201: 'Created',
  204: 'No Content',
  304: 'Not Modified',
  400: 'Bad Request',
  401: 'Unauthorized',
  403: 'Forbidden',
  404: 'Not Found',
  409: 'Conflict',
  412: 'Precondition Failed',
  429: 'Too Many Requests',
  500: 'Internal Server Error',
  501: 'Not Implemented',
};

/** Outcome bucket for a status. 2xx is 'ok'; anything else is a failure. */
function outcomeFor(status: number): ExecOutcome {
  if (status >= 200 && status < 300) return 'ok';
  if (status === 401) return '401';
  if (status === 403) return '403';
  if (status === 404) return '404';
  return 'error';
}

/** Reason phrase for a status from the map above, or undefined when it
 *  isn't listed. The status badges use it so a create reads "201 Created"
 *  rather than "201 OK" (#123). */
export function statusTextFor(status: number): string | undefined {
  return STATUS_TEXT[status];
}

interface HostError {
  message?: string;
  code?: string;
  statusCode?: number;
  status?: number;
  // PPTB error envelopes also carry these in some cases
  innererror?: { message?: string };
}

function isHostError(e: unknown): e is HostError {
  return typeof e === 'object' && e !== null;
}

/** `error.code` on the ExecResult when PPTB didn't send a request because
 *  the user declined (or didn't answer) its additional-headers prompt. */
export const HEADER_CONSENT_DECLINED_CODE = 'DRS.HeaderConsentDeclined';
/** `error.code` on the ExecResult when PPTB refused one of the
 *  `additionalHeaders` (invalid or conflicting name) before sending anything. */
export const HEADER_REJECTED_CODE = 'DRS.HeaderRejected';

// PPTB 1.2.7 checks consent and validates `additionalHeaders` BEFORE it sends
// anything, and rejects the call with these messages (desktop-app
// dataverseHeaderConsentManager.ts / dataverseBatch.ts). Nothing reached
// Dataverse, so we say so rather than report a Dataverse failure.
const CONSENT_DECLINED_RE = /additional headers were rejected|header consent/i;
const HEADER_INVALID_RE =
  /invalid dataverse header|conflicting duplicate dataverse header|additional dataverse headers/i;

// PPTB rejects failed Dataverse calls with a plain `Error` whose message is
// `<code>: <message>` (or `HTTP <status>: <body>` when the response has no
// error envelope), so `statusCode` / `status` are usually absent. Recover
// the status from the message where it's unambiguous: an `HTTP <nnn>`
// prefix, or a Dataverse error code that always maps to one status. The
// precondition codes let WriteResultCard render its 412 / 404 copy.
const STATUS_BY_DATAVERSE_CODE: Record<string, number> = {
  '0x80040217': 404, // ObjectDoesNotExist (also If-Match: * on a missing row)
  '0x80040220': 403, // Caller is missing a privilege
  '0x80040237': 412, // DuplicateRecord (also If-None-Match: * on an existing row)
  '0x80060882': 412, // ConcurrencyVersionMismatch: stale If-Match ETag
  '0x80060888': 400, // OData: could not find a property
  // Service protection limits (#269): Dataverse answers these with 429.
  '0x80072322': 429, // Number of requests
  '0x80072321': 429, // Combined execution time
  '0x80072326': 429, // Concurrent requests
};

/** The first Dataverse error code (`0x8xxxxxxx`) in a host message, as written. */
function dataverseCodeIn(message: string): string | undefined {
  return /\b0x8[0-9a-fA-F]{7}\b/.exec(message)?.[0];
}

/** Status recoverable from a host message: an explicit `HTTP <nnn>` wins,
 *  then a Dataverse code with a fixed status. */
function statusFromMessage(message: string): number | undefined {
  // "HTTP 404", or axios' "Request failed with status code 404" (the PPTB
  // VS Code extension passes that text through).
  const http = /\b(?:HTTP|status code) (\d{3})\b/.exec(message);
  if (http) return Number(http[1]);
  const code = dataverseCodeIn(message);
  return code ? STATUS_BY_DATAVERSE_CODE[code.toLowerCase()] : undefined;
}

/** The one shape every failed ExecResult takes. `statusText` defaults from
 *  the status; status `0` means the host didn't report one. `rawMessage`
 *  keeps the host's unedited text (IPC wrapper included) for the JSON view. */
function failure(
  status: number,
  error: { code?: string; message: string; rawMessage?: string },
  ms: number,
  statusText?: string,
): ExecResult {
  const { code: _code, ...withoutCode } = error;
  return {
    status,
    statusText:
      statusText ?? STATUS_TEXT[status] ?? (status ? 'Error' : 'Status not reported by host'),
    ms,
    bytes: 0,
    ok: false,
    outcome: outcomeFor(status),
    body: { error: error.code ? error : withoutCode },
    headers: {},
  };
}

/**
 * Map anything a dvHost call threw to an ExecResult. `sentHeaders` is the
 * `additionalHeaders` map the call carried, if any, so a host-side consent or
 * header-validation rejection reads as "not sent" instead of a Dataverse error.
 */
export function mapError(
  err: unknown,
  ms: number,
  sentHeaders?: Record<string, string>,
): ExecResult {
  // Outside PPTB there is no host to send the request to. Report it as
  // 501 Not Implemented (README → "Standalone vs PPTB-embedded").
  if (err instanceof HostNotAvailableError) {
    return failure(
      501,
      {
        code: 'HostNotAvailable',
        message:
          'Running standalone: Execute needs the PPTB host. ' +
          'Build the request here and copy it from the Code tab, or open DRS inside Power Platform ToolBox.',
      },
      ms,
    );
  }
  const e: HostError = isHostError(err) ? err : {};
  const rawMessage = e.message ?? e.innererror?.message ?? String(err);
  // Drop Electron's "Error invoking remote method '…': Error: " (#188); the
  // raw text stays on the result body when it differs.
  const message = cleanHostMessage(rawMessage);
  const raw = rawMessage !== message ? { rawMessage } : {};
  if (sentHeaders && CONSENT_DECLINED_RE.test(message)) {
    return failure(
      403,
      {
        code: HEADER_CONSENT_DECLINED_CODE,
        message:
          `Power Platform ToolBox didn't send this request because its headers ` +
          `(${Object.keys(sentHeaders).join(', ')}) weren't approved, so nothing reached Dataverse. ` +
          `Execute again and allow the headers when ToolBox asks. Host message: ${message}`,
        ...raw,
      },
      ms,
      'Headers not approved',
    );
  }
  if (sentHeaders && HEADER_INVALID_RE.test(message)) {
    return failure(
      400,
      {
        code: HEADER_REJECTED_CODE,
        message:
          `Power Platform ToolBox rejected a request header, so nothing was sent to Dataverse. ` +
          `Fix the header in the Headers or Prefer pane. Host message: ${message}`,
        ...raw,
      },
      ms,
    );
  }
  // 0 = the host didn't say. Better than a made-up 500 for what is usually
  // a 400-class rejection.
  const status = e.statusCode ?? e.status ?? statusFromMessage(message) ?? 0;
  const code = e.code ?? dataverseCodeIn(message);
  return failure(status, { code, message, ...raw }, ms);
}

function ok(status: number, body: unknown, ms: number): ExecResult {
  const bytes = jsonBytes(body);
  return {
    status,
    statusText: STATUS_TEXT[status] ?? 'OK',
    ms,
    bytes,
    ok: status >= 200 && status < 300,
    outcome: outcomeFor(status),
    body,
    // The host returns parsed bodies only — no response headers to show.
    headers: {},
  };
}

/** Returned by the binary modes — host doesn't expose chunked upload. */
function hostNotSupported(feature: string, ms: number): ExecResult {
  return failure(
    501,
    {
      code: '0x80048345',
      message:
        `${feature} isn't supported through the PPTB Dataverse API today. ` +
        `Build the request here and run it from an external Web API client.`,
    },
    ms,
  );
}

/** Wrap an async dvHost call with timing + error mapping. Pass the headers
 *  the call sends so a declined consent prompt maps to a clear result. */
async function timed<T>(
  fn: () => Promise<T>,
  sentHeaders?: Record<string, string>,
): Promise<{ ms: number; value: T } | ExecResult> {
  const t0 = performance.now();
  try {
    const value = await fn();
    return { ms: Math.round(performance.now() - t0), value };
  } catch (e) {
    return mapError(e, Math.round(performance.now() - t0), sentHeaders);
  }
}

/**
 * Headers that can't be sent as they are, so the request isn't (#109):
 *   • impersonation that isn't one user's id (both headers on, or the
 *     presets' all-zero placeholder);
 *   • on a host that can't send headers (before PPTB 1.2.7), an enabled
 *     impersonation header — the request would run as the signed-in user —
 *     and, with `bypass`, a bypass header — plug-ins would still run.
 */
export function refuseHeaderProblems(
  headers?: Record<string, string>,
  opts: { bypass?: boolean } = {},
): ExecResult | null {
  const refuse = (message: string) =>
    mapError({ statusCode: 400, message: `Not sent: ${message}` }, 0);
  const problem = impersonationProblem(headers);
  if (problem) return refuse(problem);
  if (!isEmbedded() || hostSupportsAdditionalHeaders() || !headers) return null;
  const update = 'update ToolBox (1.2.7 or later), or copy the request from the Code tab.';
  const caller = Object.keys(toWireHeaders(headers) ?? {}).find(isImpersonationHeader);
  if (caller) {
    return refuse(
      `this version of Power Platform ToolBox can't send ${caller}, so the request would run ` +
        `as you, not as the user it names. Turn the header off to run it as yourself, or ${update}`,
    );
  }
  const bypass = opts.bypass && Object.keys(toWireHeaders(headers) ?? {}).find(isBypassHeader);
  if (bypass) {
    return refuse(
      `this version of Power Platform ToolBox can't send ${bypass}, so plug-ins and flows would ` +
        `still run. Turn the header off to run them, or ${update}`,
    );
  }
  return null;
}

/** `toWireHeaders`, but only when the host can actually send them — so the
 *  consent-declined mapping never fires on a host that dropped them. */
export function sendable(headers?: Record<string, string>): Record<string, string> | undefined {
  return hostSupportsAdditionalHeaders() ? toWireHeaders(headers) : undefined;
}

function isExecResult(x: unknown): x is ExecResult {
  return typeof x === 'object' && x !== null && 'ok' in x && 'status' in x;
}

// ── Read group ───────────────────────────────────────────────────────────

/**
 * What a read sends (#109): the allow-listed header rows (impersonation,
 * `If-None-Match`, `Consistency`) and a `Prefer` when the page size (#219)
 * or a Prefer row asks for more than the host's default. Only when the host
 * can send headers (1.2.7+); a plain read sends none and never prompts.
 * Pages after the first must carry the same headers: the same page size
 * (Dataverse "Page results"), and the same caller, or a later page would
 * run as the signed-in user.
 */
export function readHeaders(
  headers: Record<string, string> | undefined,
  maxpagesize: number | null | undefined,
): Record<string, string> | undefined {
  return hostSupportsAdditionalHeaders() ? readRequestHeaders(headers, maxpagesize) : undefined;
}

/** `readHeaders` with only a page size — kept for the paging callers and tests (#219). */
export const readPreferHeaders = (maxpagesize: number | null | undefined) =>
  readHeaders(undefined, maxpagesize);

/** True when the read's headers carry a page size (then `$top` is dropped, #219). */
const sendsPageSize = (h: Record<string, string> | undefined) =>
  /odata\.maxpagesize=/i.test(h?.Prefer ?? '');

export async function executeRetrieveMultiple(
  s: RetrieveMultipleState,
  composed?: Record<string, string>,
): Promise<ExecResult> {
  const dropped = refuseHeaderProblems(composed);
  if (dropped) return dropped;
  if (s.countOnly) return executeCount(s, composed);
  const headers = readHeaders(composed, s.prefer.maxpagesize);
  // Without the page size on the wire, dropping `$top` (as the Code tab
  // does) would leave the query uncapped: up to 5,000 rows (#219). Except
  // where Dataverse can reject top-level `$top` (a nested `$expand` that
  // involves a collection, 0x80060888): there it stays dropped, as before.
  const keepTop = !sendsPageSize(headers) && !hasCollectionInvolvedNestedExpand(s.expand, s.table);
  const built = buildRetrieveMultiple(s, { keepTop });
  const res = await timed(() => dvHost.queryData(built.relativeNoBase, headers), headers);
  if (isExecResult(res)) return res;
  return ok(200, res.value, res.ms);
}

/** Code on a count-only response that carried no `@odata.count`. */
export const COUNT_UNREADABLE_CODE = 'DRS.CountUnreadable';

/** `@odata.count` and whether Dataverse stopped counting at its limit. */
export function readCount(body: unknown): { count: number; limitExceeded?: boolean } | null {
  if (!body || typeof body !== 'object') return null;
  const b = body as Record<string, unknown>;
  const count = b['@odata.count'];
  if (typeof count !== 'number' || !Number.isInteger(count) || count < 0) return null;
  const exceeded = b['@Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded'];
  return typeof exceeded === 'boolean' ? { count, limitExceeded: exceeded } : { count };
}

/**
 * Count only (#272): one row with `$count=true` (see buildRetrieveMultiple).
 * The request's own headers go with it (impersonation, say), but never a
 * page size. The result's body is the count itself.
 */
async function executeCount(
  s: RetrieveMultipleState,
  composed?: Record<string, string>,
): Promise<ExecResult> {
  const sent = readHeaders(composed, null);
  // Sent headers replace the host's own Prefer, which asks for every
  // annotation; ask again for the one that says the count is capped.
  const headers = sent ? { ...sent, Prefer: COUNT_ONLY_PREFER } : undefined;
  const built = buildRetrieveMultiple(s);
  const res = await timed(() => dvHost.queryData(built.relativeNoBase, headers), headers);
  if (isExecResult(res)) return res;
  const read = readCount(res.value);
  if (!read) {
    // Dataverse answered 200, so say that, but there's no count to show.
    return {
      ...failure(
        200,
        {
          code: COUNT_UNREADABLE_CODE,
          message: "Dataverse answered without @odata.count, so there's no count to show.",
        },
        res.ms,
      ),
      outcome: 'error',
    };
  }
  return {
    ...ok(200, read.count, res.ms),
    ...(read.limitExceeded !== undefined ? { countLimitExceeded: read.limitExceeded } : {}),
  };
}

/**
 * A request whose ids or names would change the URL's shape isn't sent
 * (#221). The modes block Execute for these already; this is the backstop
 * for state that reached the executor some other way (an imported request).
 */
function refuseUrlParts(
  ids: Parameters<typeof invalidUrlPartReason>[0],
  names: Parameters<typeof invalidUrlPartReason>[1] = [],
): ExecResult | null {
  const reason = invalidUrlPartReason(ids, names);
  return reason ? mapError({ statusCode: 400, message: `Not sent: ${reason}` }, 0) : null;
}

export async function executeRetrieveSingle(
  s: RetrieveSingleState,
  composed?: Record<string, string>,
): Promise<ExecResult> {
  const dropped = refuseHeaderProblems(composed);
  if (dropped) return dropped;
  const refused = refuseUrlParts([['Record id', s.recordId]]);
  if (refused) return refused;
  const built = buildRetrieveSingle(s);
  // No page size on a single row; a Prefer row can still ask for more.
  const headers = readHeaders(composed, null);
  // RetrieveSingle returns a single object, not a {value:[…]} envelope. The
  // host's queryData wraps responses in `{ value: Record<string, unknown>[] }`
  // — for /accounts(id) it returns the single record as the first element.
  // Cast through unknown so TS lets us index a typed array shape that's
  // really just the OData response we know is a singleton in this code path.
  const res = await timed(() => dvHost.queryData(built.relativeNoBase, headers), headers);
  // The host rejects every non-2xx, so a matching If-None-Match arrives as a
  // failed 304. That's the answer the header asks for: the row hasn't
  // changed since that etag.
  if (isExecResult(res) && res.status === 304) {
    return { ...ok(304, null, res.ms), ok: true, outcome: 'ok' };
  }
  if (isExecResult(res)) return res;
  const v = res.value as unknown;
  const row = Array.isArray(v) ? (v[0] ?? null) : v && typeof v === 'object' ? v : null;
  return ok(row ? 200 : 404, row, res.ms);
}

export async function executeNextLink(
  s: RetrieveNextLinkState,
  composed?: Record<string, string>,
): Promise<ExecResult> {
  return executeAbsoluteUrl(s.url, s.prefer.maxpagesize, composed);
}

/**
 * Follow an absolute Dataverse Web API URL (typically an `@odata.nextLink`
 * pulled out of a previous response). The host's `queryData` takes a path
 * relative to `/api/data/v9.2/`, so we strip the scheme/host/api/version
 * prefix before calling. Used by ResultsGrid's infinite scroll +
 * Retrieve-All command. `maxpagesize` and `composed` are the page size and
 * headers the first page was requested with; every later page sends them
 * too (#219, #109).
 */
export async function executeAbsoluteUrl(
  url: string,
  maxpagesize?: number | null,
  composed?: Record<string, string>,
): Promise<ExecResult> {
  const dropped = refuseHeaderProblems(composed);
  if (dropped) return dropped;
  const headers = readHeaders(composed, maxpagesize);
  let path = url;
  const apiIdx = path.indexOf('/api/data/');
  if (apiIdx >= 0) {
    path = path.slice(apiIdx + '/api/data/'.length);
    const slash = path.indexOf('/');
    if (slash >= 0) path = path.slice(slash + 1);
  }
  const res = await timed(() => dvHost.queryData(path, headers), headers);
  if (isExecResult(res)) return res;
  return ok(200, res.value, res.ms);
}

export async function executePredefined(
  s: PredefinedQueryState,
  composed?: Record<string, string>,
): Promise<ExecResult> {
  const dropped = refuseHeaderProblems(composed);
  if (dropped) return dropped;
  if (!s.queryId) {
    return mapError({ statusCode: 400, message: 'Pick a saved query first.' }, 0);
  }
  const refused = refuseUrlParts(
    [
      ['Query id', s.queryId],
      ['Source row id', s.related?.recordId],
    ],
    [['Relationship', s.related?.navProperty]],
  );
  if (refused) return refused;
  // Predefined queries: GET /<entitySet>?savedQuery=<id>  or  ?userQuery=<id>
  // We need the entity set name from the table. The metadata provider gives
  // us TableMeta synchronously when cached.
  const tbl = metadata.peekTable(s.table) ?? (await metadata.getTable(s.table));
  if (!tbl) return mapError({ statusCode: 404, message: `Unknown table ${s.table}` }, 0);
  const param = s.queryType === 'savedQuery' ? 'savedQuery' : 'userQuery';
  const headers = readHeaders(composed, s.prefer.maxpagesize);
  // Dataverse ignores `$top` next to a page size, so it's only sent when the
  // page size isn't (#219); that matches the URL bar whenever it can.
  const top = s.top && !sendsPageSize(headers) ? `&$top=${s.top}` : '';
  // Related collection (#5): /<sourceSet>(<id>)/<nav>?savedQuery=<id>
  let collection = tbl.entitySetName;
  if (s.related) {
    const r = s.related;
    if (!r.sourceTable || !r.navProperty || !r.recordId) {
      return mapError(
        { statusCode: 400, message: 'Pick the related collection and the source row.' },
        0,
      );
    }
    const src = metadata.peekTable(r.sourceTable) ?? (await metadata.getTable(r.sourceTable));
    if (!src) return mapError({ statusCode: 404, message: `Unknown table ${r.sourceTable}` }, 0);
    collection = `${src.entitySetName}(${r.recordId})/${r.navProperty}`;
  }
  const path = `${collection}?${param}=${s.queryId}${top}`;
  const res = await timed(() => dvHost.queryData(path, headers), headers);
  if (isExecResult(res)) return res;
  return ok(200, res.value, res.ms);
}

// ── Write group ──────────────────────────────────────────────────────────

export async function executeCreate(
  s: CreateState,
  headers?: Record<string, string>,
): Promise<ExecResult> {
  const dropped = refuseHeaderProblems(headers);
  if (dropped) return dropped;
  const body = buildCreateBody(s);
  const wire = sendable(headers);
  const res = await timed(() => dvHost.create(s.table, body, wire), wire);
  if (isExecResult(res)) return res;
  return ok(201, res.value, res.ms);
}

/**
 * Update's "Require existing record" (`If-Match: *`) on a host that can't send
 * headers (PPTB < 1.2.7, #218). Its `update()` is a plain PATCH, which
 * Dataverse treats as an upsert, so a mistyped or stale GUID would create a
 * record. Read the record first and stop unless it exists. A record deleted
 * between this read and the PATCH would still be re-created: only the header
 * closes that window of milliseconds. Returns the result to show when the
 * update must not be sent, or `null` to go ahead.
 */
async function requireExistingRecord(table: string, id: string): Promise<ExecResult | null> {
  const meta = metadata.peekTable(table) ?? (await metadata.getTable(table));
  const res = await timed(() =>
    dvHost.retrieve(table, id, meta?.primaryKey ? [meta.primaryKey] : undefined),
  );
  if (!isExecResult(res)) return null;
  const why =
    "This host can't send If-Match: *, so Dataverse Request Studio " +
    'checks that the record exists before updating.';
  const err = (res.body as { error?: { code?: string; message?: string; rawMessage?: string } })
    ?.error;
  if (res.status === 404) {
    return failure(
      404,
      {
        code: err?.code,
        message: `No ${meta?.displayName ?? table} record has the id ${id}, so nothing was updated. ${why}`,
      },
      res.ms,
    );
  }
  return failure(
    res.status,
    {
      code: err?.code,
      message:
        `Nothing was updated: the check that the record exists failed (${err?.message ?? res.statusText}). ` +
        why,
      ...(err?.rawMessage ? { rawMessage: err.rawMessage } : {}),
    },
    res.ms,
    res.statusText,
  );
}

/**
 * PATCH a record that must already exist (#254): single-valued Associate /
 * Disassociate set or clear the lookup with a PATCH of the source record,
 * which Dataverse would otherwise treat as an upsert. Sends `If-Match: *`
 * on hosts that pass headers (atomic; the user approves it once), and reads
 * the record first on hosts that can't, as `executeUpdate` does (#218).
 */
async function patchExistingRecord(
  table: string,
  id: string,
  body: Record<string, unknown>,
  /** The mode's allow-listed headers (impersonation, bypass; #109). */
  extra?: Record<string, string>,
): Promise<{ ms: number; value: void } | ExecResult> {
  if (isEmbedded() && !hostSupportsAdditionalHeaders()) {
    const stop = await requireExistingRecord(table, id);
    if (stop) return stop;
  }
  const wire = sendable({ ...extra, ...PATCH_EXISTING_HEADERS });
  const res = await timed(() => dvHost.update(table, id, body, wire), wire);
  // Declining the header prompt shouldn't block setting a lookup that
  // worked without one before: fall back to reading the record first. Not
  // when the user's own headers were declined too: without them the PATCH
  // would run as the signed-in user, or with plug-ins the user bypassed.
  const declined =
    isExecResult(res) &&
    (res.body as { error?: { code?: string } } | null)?.error?.code ===
      HEADER_CONSENT_DECLINED_CODE;
  if (!declined || extra) return res;
  const stop = await requireExistingRecord(table, id);
  if (stop) return stop;
  return timed(() => dvHost.update(table, id, body));
}

export async function executeUpdate(
  s: UpdateState,
  headers?: Record<string, string>,
): Promise<ExecResult> {
  const dropped = refuseHeaderProblems(headers);
  if (dropped) return dropped;
  if (!s.recordId) return mapError({ statusCode: 400, message: 'Pick a record to update.' }, 0);
  const refused = refuseUrlParts([['Record id', s.recordId]]);
  if (refused) return refused;
  // CRITICAL: PUT single-column (PUT /<entityset>(<id>)/<column> with body
  // `{ value: <scalar> }`) is a real Web API pattern but PPTB's
  // dataverseAPI.update is PATCH-only — there's no raw-request hook for the
  // property-path URL. If we silently fell through to PATCH here we'd write
  // a different shape than DRS authored (and what the URL bar / Code tab
  // promise). Fail loud and direct the user to PATCH-one-field instead.
  if (s.method === 'PUT') {
    return mapError(
      {
        statusCode: 400,
        message:
          "PUT single-column (PUT /<entityset>(<id>)/<column>) isn't supported by the PPTB host. " +
          "Its dataverseAPI.update is PATCH-only; there's no raw-request hook for property-path URLs. " +
          'To set one column from inside PPTB, switch back to PATCH and put a single field in the body. ' +
          'To use the PUT pattern, copy the URL + body from the Code tab and run it outside PPTB.',
      },
      0,
    );
  }
  const body = buildUpdateBody(s);
  // The host's `update()` is PATCH semantics. For Upsert-via-Update (when
  // concurrency is `none`) the server still does the right thing.
  if (s.concurrency.kind === 'update-only' && isEmbedded() && !hostSupportsAdditionalHeaders()) {
    const stop = await requireExistingRecord(s.table, s.recordId);
    if (stop) return stop;
  }
  const wire = sendable(headers);
  const res = await timed(() => dvHost.update(s.table, s.recordId!, body, wire), wire);
  if (isExecResult(res)) return res;
  // dvHost.update resolves to void: the host returns no row, even when the
  // request asked for return=representation (#224). Report the 204 as is.
  return ok(204, null, res.ms);
}

export async function executeUpsert(
  s: UpsertState,
  headers?: Record<string, string>,
): Promise<ExecResult> {
  const dropped = refuseHeaderProblems(headers);
  if (dropped) return dropped;
  // The host's `update()` is PATCH semantics, addressed by logical name +
  // GUID. For alt-key Upsert the address is `<entitySet>(<keyCol>='value')`
  // and PPTB's dataverseAPI doesn't expose a raw-request hook for that
  // URL shape. Same family of limitation as Delete-single-property and
  // Update-PUT: we still author the correct URL + body, but Execute is
  // gated by an advisory in UpsertMode pointing users to copy the
  // request out for external execution.
  if (s.key.kind === 'alternate') {
    return mapError(
      {
        statusCode: 400,
        message:
          "Alternate-key Upsert (PATCH /<entityset>(<keyCol>='value')) isn't supported by the PPTB host — " +
          'its dataverseAPI.update only accepts GUID addressing. ' +
          'DRS authored the correct request — copy the URL + body from the Code tab and run it from outside PPTB ' +
          '(Postman / curl / JS SDK / Power Automate).',
      },
      0,
    );
  }
  if (s.key.kind !== 'guid' || !s.key.recordId) {
    return mapError({ statusCode: 400, message: 'Pick or enter a record GUID.' }, 0);
  }
  const recordId = s.key.recordId;
  const refused = refuseUrlParts([['Record id', recordId]]);
  if (refused) return refused;
  // CRITICAL: route the body through buildUpsertBody (== buildCreateBody)
  // so lookups become `<navProp>@odata.bind`, multi-select picklists become
  // comma-strings, and empty values are dropped. The previous version cast
  // s.fieldValues directly, which is the in-memory shape (Lookup is
  // `{id, targetEntity}` object) — not what Web API accepts.
  const body = buildUpsertBody(s);
  const wire = sendable(headers);
  const res = await timed(() => dvHost.update(s.table, recordId, body, wire), wire);
  if (isExecResult(res)) return res;
  // Same as Update: no row from the host, and no way to tell 201 from 204.
  return ok(204, null, res.ms);
}

export async function executeDelete(
  s: DeleteState,
  headers?: Record<string, string>,
): Promise<ExecResult> {
  const dropped = refuseHeaderProblems(headers);
  if (dropped) return dropped;
  if (!s.recordId) return mapError({ statusCode: 400, message: 'Pick a record to delete.' }, 0);
  const refused = refuseUrlParts([['Record id', s.recordId]]);
  if (refused) return refused;
  // CRITICAL: single-property delete (DELETE /<entityset>(<id>)/<column>) is
  // a real Web API pattern (per the docs) but PPTB's dataverseAPI surface
  // only exposes a whole-row `delete(entityLogicalName, id)`. There is no
  // generic raw-request method we can route a property-path DELETE through.
  // If we called `dvHost.delete(table, id)` here in single-property mode,
  // we'd silently whole-row-delete the record — which is exactly the bug
  // the user reported. Fail loud instead.
  if (s.scope.kind === 'single-property') {
    return mapError(
      {
        statusCode: 400,
        message:
          "Single-property DELETE (DELETE /<entityset>(<id>)/<column>) isn't supported by the PPTB host. " +
          "Its dataverseAPI only does whole-row deletes; there's no raw-request hook for property-path URLs. " +
          'DRS authored the correct request — copy the URL or the Code-tab fetch snippet and run it from outside PPTB (Postman, curl, the JS SDK, Power Automate, etc.).',
      },
      0,
    );
  }
  const wire = sendable(headers);
  const res = await timed(() => dvHost.delete(s.table, s.recordId!, wire), wire);
  if (isExecResult(res)) return res;
  return ok(204, null, res.ms);
}

export async function executeMerge(
  s: MergeState,
  headers?: Record<string, string>,
): Promise<ExecResult> {
  const dropped = refuseHeaderProblems(headers);
  if (dropped) return dropped;
  // Per docs: POST /Merge with { Target, Subordinate, UpdateContent, PerformParentingChecks }
  const refused = refuseUrlParts([
    ['Master record id', s.targetId],
    ['Subordinate record id', s.subordinateId],
  ]);
  if (refused) return refused;
  const body = buildMergeBody(s);
  const wire = sendable(headers);
  const res = await timed(
    () =>
      dvHost.execute(
        {
          operationName: 'Merge',
          operationType: 'action',
          parameters: body,
        },
        wire,
      ),
    wire,
  );
  if (isExecResult(res)) return res;
  return ok(204, res.value, res.ms);
}

// ── Relate group ─────────────────────────────────────────────────────────

// Collection-valued Associate / Disassociate is one host call per target
// (no batch), so a run can partly succeed. Each target gets its own result
// and a failure doesn't stop the rest (#99).

/** Max target GUIDs per "which are linked?" query, to keep the URL short. */
const LINK_CHECK_CHUNK = 25;

/**
 * Of `targetIds`, the ones currently linked to the source row through a
 * collection-valued nav property. `null` when it couldn't be determined
 * (metadata missing, no read privilege, query error): the caller then
 * sends every request as before. A missing host still throws, so standalone
 * runs report 501 like every other mode.
 */
async function linkedTargets(
  table: string,
  sourceId: string,
  navProperty: string,
  targetEntity: string,
  targetIds: string[],
): Promise<Set<string> | null> {
  try {
    const src = metadata.peekTable(table) ?? (await metadata.getTable(table));
    const tgt = metadata.peekTable(targetEntity) ?? (await metadata.getTable(targetEntity));
    if (!src?.entitySetName || !tgt?.primaryKey) return null;
    const linked = new Set<string>();
    for (let i = 0; i < targetIds.length; i += LINK_CHECK_CHUNK) {
      const chunk = targetIds.slice(i, i + LINK_CHECK_CHUNK);
      const filter = chunk.map((id) => `${tgt.primaryKey} eq ${id}`).join(' or ');
      const res = await dvHost.queryData(
        `/${src.entitySetName}(${sourceId})/${navProperty}?$select=${tgt.primaryKey}&$filter=${filter}`,
      );
      for (const row of (res.value ?? []) as Record<string, unknown>[]) {
        const id = row[tgt.primaryKey];
        if (typeof id === 'string') linked.add(id.toLowerCase());
      }
    }
    return linked;
  } catch (e) {
    if (e instanceof HostNotAvailableError) throw e;
    return null;
  }
}

/**
 * Run `send` for each target in order, skipping the ones `skip` says are
 * already in the requested state, and fold the per-target outcomes into one
 * ExecResult:
 *   • nothing failed       → 204, ok (skips included).
 *   • nothing succeeded    → the first failure's status and error (skipped
 *                            targets don't count as successes: nothing
 *                            changed).
 *   • some of each         → outcome 'partial', not ok. The status is the
 *                            first failure's real HTTP status: Dataverse
 *                            has no status for "some targets failed", so
 *                            DRS doesn't invent one (D2, #99).
 * `targets` and the body's `targets` list every target either way.
 */
async function runPerTarget(
  targetIds: string[],
  skip: (id: string) => boolean,
  send: (id: string) => Promise<unknown>,
  t0: number,
  /** The headers each request sends, to report a declined consent prompt. */
  sentHeaders?: Record<string, string>,
): Promise<ExecResult> {
  const targets: RelateTargetResult[] = [];
  let firstFailure: ExecResult | null = null;
  for (const targetId of targetIds) {
    if (skip(targetId)) {
      targets.push({ targetId, state: 'skipped' });
      continue;
    }
    try {
      await send(targetId);
      targets.push({ targetId, state: 'done' });
    } catch (e) {
      const f = mapError(e, 0, sentHeaders);
      // Standalone: nothing can be sent, so don't report it N times. Nor
      // ask N times when the user declined the headers: no target is sent.
      const code = (f.body as { error?: { code?: string } } | null)?.error?.code;
      if (e instanceof HostNotAvailableError || code === HEADER_CONSENT_DECLINED_CODE)
        return { ...f, ms: Math.round(performance.now() - t0) };
      firstFailure ??= f;
      targets.push({
        targetId,
        state: 'failed',
        status: f.status,
        statusText: f.statusText,
        error: (f.body as { error: { code?: string; message: string } }).error,
      });
    }
  }
  const ms = Math.round(performance.now() - t0);
  const succeeded = targets.filter((t) => t.state === 'done').length;
  if (!firstFailure) return { ...ok(204, { targets }, ms), targets };
  const error = (firstFailure.body as { error: unknown }).error;
  const body = { error, targets };
  if (succeeded === 0) {
    return { ...firstFailure, ms, body, bytes: jsonBytes(body), targets };
  }
  return {
    ...firstFailure,
    ms,
    bytes: jsonBytes(body),
    ok: false,
    outcome: 'partial',
    body,
    targets,
  };
}

export async function executeAssociate(
  s: AssociateState,
  composed?: Record<string, string>,
): Promise<ExecResult> {
  const dropped = refuseHeaderProblems(composed, { bypass: true });
  if (dropped) return dropped;
  const sourceId = s.sourceId;
  const navProperty = s.navProperty;
  if (!sourceId || !navProperty) {
    return mapError({ statusCode: 400, message: 'Pick a source record + navigation property.' }, 0);
  }
  const refused = refuseUrlParts(
    [['Source record id', sourceId], ...s.targets.map((id) => ['Target record id', id] as const)],
    [['Navigation property', navProperty]],
  );
  if (refused) return refused;
  const reqs = buildAssociateRequests(s);
  // Impersonation and bypass headers (#109).
  const relate = allowedHeaders('relate', composed);
  if (reqs.length === 0)
    return mapError({ statusCode: 400, message: 'Pick at least one target.' }, 0);

  // Resolve cardinality so we pick the right PPTB host method:
  //   • Single-valued (N:1) → PATCH @odata.bind via dvHost.update
  //                           PPTB doesn't expose PUT $ref and the docs-
  //                           preferred shape is PATCH anyway.
  //   • Collection-valued    → POST $ref per target via dvHost.associate
  const tbl = metadata.peekTable(s.table) ?? (await metadata.getTable(s.table));
  const nav = tbl?.navigationProperties.find((n) => n.name === navProperty);
  const singleValued = nav?.cardinality === 'ManyToOne';
  const relatedEntity = nav?.targetEntity ?? '';

  const t0 = performance.now();
  if (singleValued) {
    // reqs has exactly one entry per buildAssociateRequests's single-
    // valued branch. Body shape: `{ "<nav>@odata.bind": "<set>(<id>)" }`.
    // Must not create the source record when its id is wrong (#254).
    const r = reqs[0];
    const res = await patchExistingRecord(s.table, sourceId, r.body, relate);
    if (isExecResult(res)) return res;
    return ok(204, null, res.ms);
  }
  // Per-target POST $ref. PPTB's .associate takes the GUID and emits the
  // right URL + body internally. Targets that are already linked are
  // skipped, so re-running after a partial failure only sends the rest (an
  // N:N re-associate would otherwise fail on the duplicate link).
  const ids = reqs.map((r) => r.targetId);
  let linked: Set<string> | null;
  try {
    linked = await linkedTargets(s.table, sourceId, navProperty, relatedEntity, ids);
  } catch (e) {
    return mapError(e, Math.round(performance.now() - t0));
  }
  const wire = sendable(relate);
  return runPerTarget(
    ids,
    (id) => !!linked?.has(id.toLowerCase()),
    (id) => dvHost.associate(s.table, sourceId, navProperty, relatedEntity, id, wire),
    t0,
    wire,
  );
}

export async function executeDisassociate(
  s: DisassociateState,
  composed?: Record<string, string>,
): Promise<ExecResult> {
  const dropped = refuseHeaderProblems(composed, { bypass: true });
  if (dropped) return dropped;
  // Cache narrowed locals — TS doesn't carry the null-check through closure
  // scope into the loop body otherwise.
  const sourceId = s.sourceId;
  const navProperty = s.navProperty;
  if (!sourceId || !navProperty) {
    return mapError({ statusCode: 400, message: 'Pick a source record + navigation property.' }, 0);
  }
  const refused = refuseUrlParts(
    [['Source record id', sourceId], ...s.targetIds.map((id) => ['Target record id', id] as const)],
    [['Navigation property', navProperty]],
  );
  if (refused) return refused;
  const reqs = buildDisassociateRequests(s);
  // Impersonation and bypass headers (#109).
  const relate = allowedHeaders('relate', composed);
  if (reqs.length === 0)
    return mapError({ statusCode: 400, message: 'Pick at least one target.' }, 0);

  const t0 = performance.now();
  const patch = reqs.find((r) => r.method === 'PATCH');
  if (patch) {
    // Single-valued → PATCH /<source>(<id>) with { "<nav>@odata.bind": null }.
    // Routes through dvHost.update which accepts arbitrary body shapes
    // (including null values). Replaces the old "silently skip" branch
    // that did nothing because PPTB's .disassociate requires a target id.
    // Must not create the source record when its id is wrong (#254).
    const res = await patchExistingRecord(s.table, sourceId, patch.body ?? {}, relate);
    if (isExecResult(res)) return res;
    return ok(204, null, res.ms);
  }
  // Collection-valued → DELETE per target via PPTB's .disassociate. Every
  // target's DELETE is sent: unlike Associate there's no "already in that
  // state" pre-check, because the check's GET only returns related rows the
  // caller can read at their privilege depth — a linked row outside that
  // scope would look unlinked and its DELETE would never be sent. Dataverse
  // has no documented error code for "this row isn't linked", so a failed
  // DELETE is reported as the failure Dataverse returned.
  const ids = reqs.map((r) => r.targetId).filter((id): id is string => !!id);
  const wire = sendable(relate);
  return runPerTarget(
    ids,
    () => false,
    (id) => dvHost.disassociate(s.table, sourceId, navProperty, id, wire),
    t0,
    wire,
  );
}

// ── Execute group ────────────────────────────────────────────────────────

export async function executeAction(
  s: ExecuteActionState,
  headers?: Record<string, string>,
): Promise<ExecResult> {
  const dropped = refuseHeaderProblems(headers);
  if (dropped) return dropped;
  if (!s.actionName) return mapError({ statusCode: 400, message: 'Pick an action.' }, 0);
  const action = await actions.find(s.actionName);
  if (!action) return mapError({ statusCode: 404, message: `Unknown action ${s.actionName}` }, 0);

  const params = buildExecuteActionBody(s);
  // Resolve binding from the action CSDL — the source state only stores the
  // record id; the entity name lives on action.binding.entityType. The host
  // maps the three binding shapes as (#55):
  //   entityName + entityId → POST /<set>(<id>)/Microsoft.Dynamics.CRM.<op>
  //   entityName only       → POST /<set>/Microsoft.Dynamics.CRM.<op>  (collection)
  //   neither               → POST /<op>                               (unbound)
  // Collection-bound actions used to fall through to the unbound form.
  const boundEntity = action.binding.kind === 'unbound' ? undefined : action.binding.entityType;
  let boundEntityId: string | undefined;
  if (action.binding.kind === 'entity') {
    if (!s.boundRecordId) {
      return mapError(
        { statusCode: 400, message: `Action '${action.name}' is bound — pick a source record.` },
        0,
      );
    }
    const refused = refuseUrlParts([['Bound record id', s.boundRecordId]]);
    if (refused) return refused;
    boundEntityId = s.boundRecordId;
  }

  // The Headers pane's rows (impersonation, bypass, …) go out on 1.2.7+ (#266).
  const sent = sendable(headers);
  const res = await timed(
    () =>
      dvHost.execute(
        {
          operationName: s.actionName!,
          operationType: 'action',
          entityName: boundEntity,
          entityId: boundEntityId,
          parameters: params,
        },
        sent,
      ),
    sent,
  );
  if (isExecResult(res)) return res;
  return ok(200, res.value, res.ms);
}

export async function executeFunction(
  s: ExecuteFunctionState,
  composed?: Record<string, string>,
): Promise<ExecResult> {
  const dropped = refuseHeaderProblems(composed);
  if (dropped) return dropped;
  if (!s.functionName) return mapError({ statusCode: 400, message: 'Pick a function.' }, 0);
  const action = await actions.find(s.functionName);
  if (!action)
    return mapError({ statusCode: 404, message: `Unknown function ${s.functionName}` }, 0);
  if (action.binding.kind === 'entity' && !s.boundRecordId) {
    return mapError(
      { statusCode: 400, message: `Function '${action.name}' is bound — pick a source record.` },
      0,
    );
  }
  if (action.binding.kind === 'entity') {
    const refused = refuseUrlParts([['Bound record id', s.boundRecordId]]);
    if (refused) return refused;
  }
  const badParam = invalidFunctionParamReason(action.parameters, s.paramValues);
  if (badParam) return mapError({ statusCode: 400, message: `Not sent: ${badParam}` }, 0);
  // Bound functions need the table's real entity set name for the URL; make
  // sure its metadata is in the registry the builder reads from.
  if (action.binding.kind !== 'unbound') {
    await metadata.getTable(action.binding.entityType);
  }

  // Send the exact URL the builder shows (#55). The host's execute() would
  // re-encode the raw editor values itself — EntityReference objects and
  // numeric enums came out differently from the URL bar, it ignored the
  // parameter-alias toggle, and it guesses the entity set name by
  // pluralizing. A function call is a plain GET, so route the builder's
  // path through queryData instead: one source of truth for preview, Code
  // tab and execution.
  const path = buildExecuteFunctionWirePath(s);
  if (!path) {
    return mapError(
      {
        statusCode: 400,
        message:
          action.binding.kind === 'unbound'
            ? `Couldn't build the URL for function '${action.name}'.`
            : `Couldn't resolve the entity set for '${action.binding.entityType}' — ` +
              `table metadata isn't available.`,
      },
      0,
    );
  }
  // Impersonation headers only (#109).
  const headers = sendable(allowedHeaders('impersonation', composed));
  const res = await timed(() => dvHost.queryData(path, headers), headers);
  if (isExecResult(res)) return res;
  return ok(200, res.value, res.ms);
}

export async function executeWorkflow(
  s: ExecuteWorkflowState,
  headers?: Record<string, string>,
): Promise<ExecResult> {
  const dropped = refuseHeaderProblems(headers);
  if (dropped) return dropped;
  if (!s.workflowId || !s.entityId) {
    return mapError({ statusCode: 400, message: 'Pick a workflow and a record.' }, 0);
  }
  const refused = refuseUrlParts([
    ['Workflow id', s.workflowId],
    ['Record id', s.entityId],
  ]);
  if (refused) return refused;
  // POST /workflows(<wf-id>)/Microsoft.Dynamics.CRM.ExecuteWorkflow
  // body: { EntityId: '<target-guid>' }
  const sent = sendable(headers);
  const res = await timed(
    () =>
      dvHost.execute(
        {
          operationName: 'ExecuteWorkflow',
          operationType: 'action',
          entityName: 'workflow',
          entityId: s.workflowId!,
          parameters: { EntityId: s.entityId },
        },
        sent,
      ),
    sent,
  );
  if (isExecResult(res)) return res;
  return ok(204, res.value, res.ms);
}

// ── Binary group ─────────────────────────────────────────────────────────

export async function executeManageFile(_s: ManageFileState): Promise<ExecResult> {
  return hostNotSupported('Manage File (chunked upload/download)', 0);
}

export async function executeManageImage(_s: ManageImageState): Promise<ExecResult> {
  return hostNotSupported('Manage Image (chunked upload/download)', 0);
}

export async function executeManageAttachment(_s: ManageAttachmentState): Promise<ExecResult> {
  return hostNotSupported('Manage Attachment/Annotation (chunked upload/download)', 0);
}

// ── Unexpected failures ──────────────────────────────────────────────────

/**
 * Result for an exception thrown while running a request (outside `timed`,
 * e.g. a `$metadata` fetch failing inside `actions.find`). `runtime` routes
 * every throw through this so a mode's Execute always gets a result back
 * and its spinner clears (#102).
 */
export function execFailure(err: unknown, ms: number): ExecResult {
  return mapError(err, ms);
}
