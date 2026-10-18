// Live executor — routes each studio mode's "Execute" through the PPTB
// host's `window.dataverseAPI` surface (via the `dvHost` wrapper in
// `pptbBridge.ts`).
//
// This module implements the SAME function signatures as the mock
// `executor.ts` — `(state, sim?) => Promise<ExecResult>`. The `sim` arg is
// ignored in live mode (it's a mock-only affordance for synthesizing 4xx
// responses without round-trips); the host returns real status codes.
//
// Request headers: the write group (Create / Update / Upsert / Delete /
// Merge) passes the headers its mode composed for the Code tab. On PPTB
// 1.2.7+ the precondition (`If-Match` / `If-None-Match`), bypass (`MSCRM.*`)
// and user-added headers go out as `additionalHeaders` once the user
// approves them in the host. Headers the host builds itself (see
// HOST_OWNED_HEADERS) stay with the host. Older hosts drop the map; there
// `hostHeaderAdvisories` blocks the requests whose meaning depends on it.
//
// What's expressly NOT implemented:
//   - Headers on the read / relate / execute groups. Those still run with
//     the host's own headers, including its default `Prefer`.
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

import { dvHost, HostNotAvailableError, hostSupportsAdditionalHeaders } from '../host/pptbBridge';
import { metadata } from '../host/metadataProvider';
import { cleanHostMessage } from './hostErrorText';
import {
  buildRetrieveMultiple,
  buildRetrieveSingle,
  buildCreateBody,
  buildUpdate,
  buildUpdateBody,
  buildUpsertBody,
  buildDelete,
  buildMergeBody,
  buildAssociateRequests,
  buildDisassociateRequests,
} from './urlBuilder';
import { buildExecuteActionBody, buildExecuteFunction } from './executeBuilders';
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

// ── Public types ─────────────────────────────────────────────────────────

/** Coarse outcome bucket on every `ExecResult`. ResultsView inspects this
 *  to render the "no rows" empty-state distinctly from a network success
 *  that happens to return a populated body. Derived from `status` by
 *  `outcomeFor`; `'error'` covers every other failure (4xx, 5xx, 501, and
 *  host errors whose status is unknown). */
export type ExecOutcome = 'ok' | '404' | '401' | '403' | 'error' | 'empty';

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
};

/** The first Dataverse error code (`0x8xxxxxxx`) in a host message, as written. */
function dataverseCodeIn(message: string): string | undefined {
  return /\b0x[0-9a-fA-F]{8}\b/.exec(message)?.[0];
}

/** Status recoverable from a host message: an explicit `HTTP <nnn>` wins,
 *  then a Dataverse code with a fixed status. */
function statusFromMessage(message: string): number | undefined {
  const http = /\bHTTP (\d{3})\b/.exec(message);
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
function mapError(err: unknown, ms: number, sentHeaders?: Record<string, string>): ExecResult {
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
          `Fix the header in the Headers pane. Host message: ${message}`,
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
  const bytes = body ? JSON.stringify(body).length : 0;
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

// Headers PPTB's Dataverse client always sets itself. Its client spreads
// `additionalHeaders` over its own, so sending these would override them
// (and prompt for consent for nothing). `Prefer` is host-owned too: the host
// always sends `return=representation` on writes and would only append ours.
const HOST_OWNED_HEADERS = new Set([
  'authorization',
  'accept',
  'content-type',
  'content-length',
  'odata-version',
  'odata-maxversion',
  'prefer',
]);

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
    if (!key || HOST_OWNED_HEADERS.has(key)) continue;
    byName.delete(key);
    byName.set(key, [name, value]);
  }
  return byName.size > 0 ? Object.fromEntries(byName.values()) : undefined;
}

/** `toWireHeaders`, but only when the host can actually send them — so the
 *  consent-declined mapping never fires on a host that dropped them. */
function sendable(headers?: Record<string, string>): Record<string, string> | undefined {
  return hostSupportsAdditionalHeaders() ? toWireHeaders(headers) : undefined;
}

function isExecResult(x: unknown): x is ExecResult {
  return typeof x === 'object' && x !== null && 'ok' in x && 'status' in x;
}

// ── Read group ───────────────────────────────────────────────────────────

export async function executeRetrieveMultiple(s: RetrieveMultipleState): Promise<ExecResult> {
  const built = buildRetrieveMultiple(s);
  const res = await timed(() => dvHost.queryData(built.relativeNoBase));
  if (isExecResult(res)) return res;
  return ok(200, res.value, res.ms);
}

export async function executeRetrieveSingle(s: RetrieveSingleState): Promise<ExecResult> {
  const built = buildRetrieveSingle(s);
  // RetrieveSingle returns a single object, not a {value:[…]} envelope. The
  // host's queryData wraps responses in `{ value: Record<string, unknown>[] }`
  // — for /accounts(id) it returns the single record as the first element.
  // Cast through unknown so TS lets us index a typed array shape that's
  // really just the OData response we know is a singleton in this code path.
  const res = await timed(() => dvHost.queryData(built.relativeNoBase));
  if (isExecResult(res)) return res;
  const v = res.value as unknown;
  const row = Array.isArray(v) ? (v[0] ?? null) : v && typeof v === 'object' ? v : null;
  return ok(row ? 200 : 404, row, res.ms);
}

export async function executeNextLink(s: RetrieveNextLinkState): Promise<ExecResult> {
  return executeAbsoluteUrl(s.url);
}

/**
 * Follow an absolute Dataverse Web API URL (typically an `@odata.nextLink`
 * pulled out of a previous response). The host's `queryData` takes a path
 * relative to `/api/data/v9.2/`, so we strip the scheme/host/api/version
 * prefix before calling. Used by ResultsGrid's infinite scroll +
 * Retrieve-All command.
 */
export async function executeAbsoluteUrl(url: string): Promise<ExecResult> {
  let path = url;
  const apiIdx = path.indexOf('/api/data/');
  if (apiIdx >= 0) {
    path = path.slice(apiIdx + '/api/data/'.length);
    const slash = path.indexOf('/');
    if (slash >= 0) path = path.slice(slash + 1);
  }
  const res = await timed(() => dvHost.queryData(path));
  if (isExecResult(res)) return res;
  return ok(200, res.value, res.ms);
}

export async function executePredefined(s: PredefinedQueryState): Promise<ExecResult> {
  if (!s.queryId) {
    return mapError({ statusCode: 400, message: 'Pick a saved query first.' }, 0);
  }
  // Predefined queries: GET /<entitySet>?savedQuery=<id>  or  ?userQuery=<id>
  // We need the entity set name from the table. The metadata provider gives
  // us TableMeta synchronously when cached.
  const tbl = metadata.peekTable(s.table) ?? (await metadata.getTable(s.table));
  if (!tbl) return mapError({ statusCode: 404, message: `Unknown table ${s.table}` }, 0);
  const param = s.queryType === 'savedQuery' ? 'savedQuery' : 'userQuery';
  const top = s.top ? `&$top=${s.top}` : '';
  const path = `${tbl.entitySetName}?${param}=${s.queryId}${top}`;
  const res = await timed(() => dvHost.queryData(path));
  if (isExecResult(res)) return res;
  return ok(200, res.value, res.ms);
}

// ── Write group ──────────────────────────────────────────────────────────

export async function executeCreate(
  s: CreateState,
  headers?: Record<string, string>,
): Promise<ExecResult> {
  const body = buildCreateBody(s);
  const wire = sendable(headers);
  const res = await timed(() => dvHost.create(s.table, body, wire), wire);
  if (isExecResult(res)) return res;
  return ok(201, res.value, res.ms);
}

export async function executeUpdate(
  s: UpdateState,
  headers?: Record<string, string>,
): Promise<ExecResult> {
  if (!s.recordId) return mapError({ statusCode: 400, message: 'Pick a record to update.' }, 0);
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
  const wire = sendable(headers);
  const res = await timed(() => dvHost.update(s.table, s.recordId!, body, wire), wire);
  if (isExecResult(res)) return res;
  // dvHost.update resolves to void on 204 No Content. Best-effort GET to
  // hydrate the updated row for the Results pane, then fall back to a
  // minimal envelope.
  void buildUpdate; // keep import used
  return ok(204, { '@odata.context': 'update' }, res.ms);
}

export async function executeUpsert(
  s: UpsertState,
  headers?: Record<string, string>,
): Promise<ExecResult> {
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
  // CRITICAL: route the body through buildUpsertBody (== buildCreateBody)
  // so lookups become `<navProp>@odata.bind`, multi-select picklists become
  // comma-strings, and empty values are dropped. The previous version cast
  // s.fieldValues directly, which is the in-memory shape (Lookup is
  // `{id, targetEntity}` object) — not what Web API accepts.
  const body = buildUpsertBody(s);
  const wire = sendable(headers);
  const res = await timed(() => dvHost.update(s.table, recordId, body, wire), wire);
  if (isExecResult(res)) return res;
  return ok(204, { '@odata.context': 'upsert' }, res.ms);
}

export async function executeDelete(
  s: DeleteState,
  headers?: Record<string, string>,
): Promise<ExecResult> {
  if (!s.recordId) return mapError({ statusCode: 400, message: 'Pick a record to delete.' }, 0);
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
  void buildDelete;
  const wire = sendable(headers);
  const res = await timed(() => dvHost.delete(s.table, s.recordId!, wire), wire);
  if (isExecResult(res)) return res;
  return ok(204, null, res.ms);
}

export async function executeMerge(
  s: MergeState,
  headers?: Record<string, string>,
): Promise<ExecResult> {
  // Per docs: POST /Merge with { Target, Subordinate, UpdateContent, PerformParentingChecks }
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

export async function executeAssociate(s: AssociateState): Promise<ExecResult> {
  const sourceId = s.sourceId;
  const navProperty = s.navProperty;
  if (!sourceId || !navProperty) {
    return mapError({ statusCode: 400, message: 'Pick a source record + navigation property.' }, 0);
  }
  const reqs = buildAssociateRequests(s);
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
  try {
    if (singleValued) {
      // reqs has exactly one entry per buildAssociateRequests's single-
      // valued branch. Body shape: `{ "<nav>@odata.bind": "<set>(<id>)" }`.
      const r = reqs[0];
      await dvHost.update(s.table, sourceId, r.body);
    } else {
      // Per-target POST $ref. PPTB's .associate takes the GUID and emits
      // the right URL + body internally.
      for (const r of reqs) {
        await dvHost.associate(s.table, sourceId, navProperty, relatedEntity, r.targetId);
      }
    }
    return ok(204, null, Math.round(performance.now() - t0));
  } catch (e) {
    return mapError(e, Math.round(performance.now() - t0));
  }
}

export async function executeDisassociate(s: DisassociateState): Promise<ExecResult> {
  // Cache narrowed locals — TS doesn't carry the null-check through closure
  // scope into the loop body otherwise.
  const sourceId = s.sourceId;
  const navProperty = s.navProperty;
  if (!sourceId || !navProperty) {
    return mapError({ statusCode: 400, message: 'Pick a source record + navigation property.' }, 0);
  }
  const reqs = buildDisassociateRequests(s);
  if (reqs.length === 0)
    return mapError({ statusCode: 400, message: 'Pick at least one target.' }, 0);

  const t0 = performance.now();
  try {
    for (const r of reqs) {
      if (r.method === 'PATCH') {
        // Single-valued → PATCH /<source>(<id>) with { "<nav>@odata.bind": null }.
        // Routes through dvHost.update which accepts arbitrary body shapes
        // (including null values). Replaces the old "silently skip" branch
        // that did nothing because PPTB's .disassociate requires a target id.
        await dvHost.update(s.table, sourceId, r.body ?? {});
      } else {
        // Collection-valued → DELETE per target via PPTB's .disassociate.
        // r.targetId is non-null by construction in this branch.
        if (!r.targetId) continue;
        await dvHost.disassociate(s.table, sourceId, navProperty, r.targetId);
      }
    }
    return ok(204, null, Math.round(performance.now() - t0));
  } catch (e) {
    return mapError(e, Math.round(performance.now() - t0));
  }
}

// ── Execute group ────────────────────────────────────────────────────────

export async function executeAction(s: ExecuteActionState): Promise<ExecResult> {
  if (!s.actionName) return mapError({ statusCode: 400, message: 'Pick an action.' }, 0);
  const action = await actions.find(s.actionName);
  if (!action) return mapError({ statusCode: 404, message: `Unknown action ${s.actionName}` }, 0);

  const params = buildExecuteActionBody(s);
  // Resolve binding from the action CSDL — the source state only stores the
  // record id; the entity name lives on action.binding.entityType.
  const boundEntity = action.binding.kind === 'entity' ? action.binding.entityType : undefined;
  const boundEntityId = boundEntity ? (s.boundRecordId ?? undefined) : undefined;

  const res = await timed(() =>
    dvHost.execute({
      operationName: s.actionName!,
      operationType: 'action',
      entityName: boundEntity,
      entityId: boundEntityId,
      parameters: params,
    }),
  );
  if (isExecResult(res)) return res;
  return ok(200, res.value, res.ms);
}

export async function executeFunction(s: ExecuteFunctionState): Promise<ExecResult> {
  if (!s.functionName) return mapError({ statusCode: 400, message: 'Pick a function.' }, 0);
  const action = await actions.find(s.functionName);
  if (!action)
    return mapError({ statusCode: 404, message: `Unknown function ${s.functionName}` }, 0);
  void buildExecuteFunction;

  const boundEntity = action.binding.kind === 'entity' ? action.binding.entityType : undefined;
  const boundEntityId = boundEntity ? (s.boundRecordId ?? undefined) : undefined;

  // The host's execute() handles parameter encoding (URL-aliased for
  // functions, JSON-body for actions) so we just pass the raw param map.
  const res = await timed(() =>
    dvHost.execute({
      operationName: s.functionName!,
      operationType: 'function',
      entityName: boundEntity,
      entityId: boundEntityId,
      parameters: s.paramValues as Record<string, unknown>,
    }),
  );
  if (isExecResult(res)) return res;
  return ok(200, res.value, res.ms);
}

export async function executeWorkflow(s: ExecuteWorkflowState): Promise<ExecResult> {
  if (!s.workflowId || !s.entityId) {
    return mapError({ statusCode: 400, message: 'Pick a workflow and a record.' }, 0);
  }
  // POST /workflows(<wf-id>)/Microsoft.Dynamics.CRM.ExecuteWorkflow
  // body: { EntityId: '<target-guid>' }
  const res = await timed(() =>
    dvHost.execute({
      operationName: 'ExecuteWorkflow',
      operationType: 'action',
      entityName: 'workflow',
      entityId: s.workflowId!,
      parameters: { EntityId: s.entityId },
    }),
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

// ── Bridge connectivity check (used by an "are we online" indicator) ─────
export async function pingHost(): Promise<boolean> {
  try {
    await dvHost.execute({ operationName: 'WhoAmI', operationType: 'function' });
    return true;
  } catch (e) {
    if (e instanceof HostNotAvailableError) return false;
    // Other errors (network blip, 401) still mean "the host is here, just
    // can't reach Dataverse right now". Treat as "online" so we don't
    // hide the live executor.
    return true;
  }
}
