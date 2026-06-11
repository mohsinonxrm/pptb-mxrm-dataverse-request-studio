// Live executor — routes each studio mode's "Execute" through the PPTB
// host's `window.dataverseAPI` surface (via the `dvHost` wrapper in
// `pptbBridge.ts`).
//
// This module implements the SAME function signatures as the mock
// `executor.ts` — `(state, sim?) => Promise<ExecResult>`. The `sim` arg is
// ignored in live mode (it's a mock-only affordance for synthesizing 4xx
// responses without round-trips); the host returns real status codes.
//
// What's expressly NOT implemented:
//   - Custom `MSCRM.*` bypass headers — the host strips them. The studio
//     still builds + shows them in the UI and the Code tab, but the live
//     executor doesn't send them. Users can copy the snippet from the Code
//     tab and run it via an external Web API client.
//   - `Prefer: return=representation` / `Prefer: odata.maxpagesize` —
//     the host owns the Prefer header; we can't override it.
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

import { dvHost, HostNotAvailableError } from '../host/pptbBridge';
import { metadata } from '../host/metadataProvider';
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
 *  that happens to return a populated body. */
export type ExecOutcome = 'ok' | '404' | '401' | '403' | 'empty';

/** Normalized execution result handed back to every mode. */
export interface ExecResult {
  status: number;
  statusText: string;
  ms: number;
  bytes: number;
  ok: boolean;
  outcome: ExecOutcome;
  body: unknown;
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
  412: 'Precondition Failed',
  500: 'Internal Server Error',
};

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

function mapError(err: unknown, ms: number): ExecResult {
  const e: HostError = isHostError(err) ? err : {};
  const status = e.statusCode ?? e.status ?? 500;
  const message = e.message ?? e.innererror?.message ?? String(err);
  const code = e.code ?? '0x80040217';
  const outcome: ExecResult['outcome'] =
    status === 404 ? '404' : status === 401 ? '401' : status === 403 ? '403' : 'ok';
  return {
    status,
    statusText: STATUS_TEXT[status] ?? 'Error',
    ms,
    bytes: 0,
    ok: false,
    outcome: outcome === 'ok' ? '404' : outcome,
    body: { error: { code, message } },
    headers: {},
  };
}

function ok(status: number, body: unknown, ms: number): ExecResult {
  const bytes = body ? JSON.stringify(body).length : 0;
  return {
    status,
    statusText: STATUS_TEXT[status] ?? 'OK',
    ms,
    bytes,
    ok: status >= 200 && status < 300,
    outcome: 'ok',
    body,
    headers: { 'OData-Version': '4.0' },
  };
}

/** Returned by the binary modes — host doesn't expose chunked upload. */
function hostNotSupported(feature: string, ms: number): ExecResult {
  return {
    status: 501,
    statusText: 'Not Implemented',
    ms,
    bytes: 0,
    ok: false,
    outcome: '403',
    body: {
      error: {
        code: '0x80048345',
        message:
          `${feature} isn't supported through the PPTB Dataverse API today. ` +
          `Build the request here and run it from an external Web API client.`,
      },
    },
    headers: {},
  };
}

/** Wrap an async dvHost call with timing + error mapping. */
async function timed<T>(fn: () => Promise<T>): Promise<{ ms: number; value: T } | ExecResult> {
  const t0 = performance.now();
  try {
    const value = await fn();
    return { ms: Math.round(performance.now() - t0), value };
  } catch (e) {
    return mapError(e, Math.round(performance.now() - t0));
  }
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

export async function executeCreate(s: CreateState): Promise<ExecResult> {
  const body = buildCreateBody(s);
  const res = await timed(() => dvHost.create(s.table, body));
  if (isExecResult(res)) return res;
  return ok(201, res.value, res.ms);
}

export async function executeUpdate(s: UpdateState): Promise<ExecResult> {
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
  const res = await timed(() => dvHost.update(s.table, s.recordId!, body));
  if (isExecResult(res)) return res;
  // dvHost.update resolves to void on 204 No Content. Best-effort GET to
  // hydrate the updated row for the Results pane, then fall back to a
  // minimal envelope.
  void buildUpdate; // keep import used
  return ok(204, { '@odata.context': 'update' }, res.ms);
}

export async function executeUpsert(s: UpsertState): Promise<ExecResult> {
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
  const res = await timed(() => dvHost.update(s.table, recordId, body));
  if (isExecResult(res)) return res;
  return ok(204, { '@odata.context': 'upsert' }, res.ms);
}

export async function executeDelete(s: DeleteState): Promise<ExecResult> {
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
  const res = await timed(() => dvHost.delete(s.table, s.recordId!));
  if (isExecResult(res)) return res;
  return ok(204, null, res.ms);
}

export async function executeMerge(s: MergeState): Promise<ExecResult> {
  // Per docs: POST /Merge with { Target, Subordinate, UpdateContent, PerformParentingChecks }
  const body = buildMergeBody(s);
  const res = await timed(() =>
    dvHost.execute({
      operationName: 'Merge',
      operationType: 'action',
      parameters: body,
    }),
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
