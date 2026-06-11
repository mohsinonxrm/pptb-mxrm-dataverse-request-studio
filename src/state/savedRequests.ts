// Saved Requests — named, per-mode state snapshots persisted to localStorage.
//
// Persistence layers:
//   • In-memory recents      — last 8 runs per mode, lost on reload
//   • localStorage saves     — this module, up to 50 entries per org scope
//   • Export/Import to disk  — future, via the host's filesystem API
//
// One saved entry = a named snapshot of a mode's state shape. Each mode
// plugs in by:
//   1. Defining its own state interface in readState/writeState/etc.
//   2. Adding its modeId to `SavedModeId` below.
//   3. Registering a (serialize, deserialize, autoSuggest) triple in
//      MODE_HELPERS so the storage layer stays mode-agnostic.
//   4. The unified library + Save/Load UX is shared — the SaveButton just
//      hands `unknown` state through; only the mode itself knows the shape.
//
// `dirty: Set<string>` is a runtime-only field on every mode's state.
// JSON.stringify can't serialize a Set, AND it isn't semantically part
// of the request — it tracks "has the user edited since last execute".
// We strip it on serialize and re-init it as `new Set()` on deserialize.

import type { RetrieveMultipleState, RetrieveSingleState, PredefinedQueryState } from './readState';
import type { DeleteState, UpdateState, CreateState, UpsertState, MergeState } from './writeState';
import type { AssociateState, DisassociateState } from './relateState';
import type {
  ExecuteActionState,
  ExecuteFunctionState,
  ExecuteWorkflowState,
} from './executeState';
import type { ManageFileState, ManageImageState, ManageAttachmentState } from './binaryState';
import { getEnv } from '../mock/environment';

// ── Storage shape ─────────────────────────────────────────────────────
//
// Org-scoping: saved requests live under a per-org localStorage bucket so
// connections don't bleed into one another. Switching from
// `contoso.crm.dynamics.com` to `fabrikam.crm.dynamics.com` shows a
// different library, and a saved request from one org can't be loaded
// against the wrong schema by accident.
//
// We still tag each entry with `orgScope` inside the payload — that way
// if someone migrates a key by hand, or we later add a "show all orgs"
// view, the origin is recoverable. The current library always filters by
// the live env host via `getSavedRequestsKey()`.

export const SAVED_REQUESTS_KEY_PREFIX = 'drs:saved-requests:v1';
export const MAX_SAVED = 50;

/**
 * The localStorage key for the CURRENT connection. Recomputed at every
 * read/write — when the host switches orgs, the next call returns the
 * new key automatically. Falls back to `:default` if no env (SSR / tests).
 */
export function getSavedRequestsKey(): string {
  const scope = getOrgScope();
  return `${SAVED_REQUESTS_KEY_PREFIX}:${scope}`;
}

/** Current org scope = the Dataverse host (e.g. `contoso.crm.dynamics.com`). */
export function getOrgScope(): string {
  try {
    return getEnv().host || 'default';
  } catch {
    return 'default';
  }
}

/**
 * Back-compat alias used by the singleton store's cross-tab `storage`
 * event listener. We compare event.key against `getSavedRequestsKey()`
 * directly, so this constant is kept only so existing imports compile.
 * Prefer `getSavedRequestsKey()` for new code.
 */
export const SAVED_REQUESTS_KEY = SAVED_REQUESTS_KEY_PREFIX;

/** Union of every mode id that participates in the saved-request library. Add new modes here. */
export type SavedModeId =
  | 'retrieve-multiple'
  | 'retrieve-single'
  | 'predefined-query'
  | 'delete'
  | 'update'
  | 'create'
  | 'upsert'
  | 'merge'
  | 'associate'
  | 'disassociate'
  // Execute group — exec-action covers OOB / Custom API / Custom Action
  // via state.category; exec-function covers OOB / Custom API functions.
  | 'exec-action'
  | 'exec-function'
  | 'exec-workflow'
  // Binary group — file / image / attachment / annotation column or record
  // operations. Each saved under their own id so a saved manage-file can't
  // be loaded into manage-image and vice versa (different state shapes).
  | 'manage-file'
  | 'manage-image'
  | 'manage-attachment';

export interface SavedRequest {
  /** Stable random id — never reused across renames; survives content edits. */
  id: string;
  /** User-supplied name (auto-suggested on first save). */
  name: string;
  /** Mode the entry came from. Lets the loader route + filter the library. */
  modeId: SavedModeId;
  /** Dataverse host this entry was saved against. Used as a visual hint and
   *  as a safety net if the user ever moves a key between buckets manually. */
  orgScope: string;
  /**
   * JSON-serializable snapshot of the mode's state (sans `dirty`).
   * Opaque to the storage layer — each mode casts to its own
   * SerializedXxxState shape inside its loader.
   */
  state: unknown;
  /** Epoch ms when this entry was created OR last overwritten. */
  savedAt: number;
  /** Optional epoch ms — set when the user runs the loaded request. */
  lastRunAt?: number;
}

// ── Per-mode serialized shapes ────────────────────────────────────────

/** Same as RetrieveMultipleState minus the non-serializable `dirty` Set. */
export type SerializedRetrieveMultipleState = Omit<RetrieveMultipleState, 'dirty'>;
/** Same as RetrieveSingleState minus the non-serializable `dirty` Set. */
export type SerializedRetrieveSingleState = Omit<RetrieveSingleState, 'dirty'>;
/** Same as PredefinedQueryState minus the non-serializable `dirty` Set. */
export type SerializedPredefinedQueryState = Omit<PredefinedQueryState, 'dirty'>;
/**
 * Same as DeleteState minus the non-serializable `dirty` Set AND the
 * one-shot confirmation fields. We DON'T persist `confirmText` /
 * `acknowledged` — re-typing the record name is a deliberate safety
 * affordance every time the request is loaded. Persisting "true" past
 * the original session would defeat the purpose.
 */
export type SerializedDeleteState = Omit<DeleteState, 'dirty' | 'confirmText' | 'acknowledged'>;
/**
 * Same as UpdateState minus the non-serializable `dirty` Set. We DO
 * persist `fieldValues` — saved update requests are about preserving
 * "here's a payload shape I want to write" and the user can edit
 * before re-executing.
 */
export type SerializedUpdateState = Omit<UpdateState, 'dirty'>;
/** Same as CreateState minus the non-serializable `dirty` Set. */
export type SerializedCreateState = Omit<CreateState, 'dirty'>;
export type SerializedUpsertState = Omit<UpsertState, 'dirty'>;
/**
 * Same as MergeState minus the non-serializable `dirty` Set AND the
 * transient live-fetched row snapshots. Snapshots are repopulated on
 * load after MergeMode refires its fetch.
 */
export type SerializedMergeState = Omit<
  MergeState,
  'dirty' | 'targetSnapshot' | 'subordinateSnapshot'
>;
/** Same as AssociateState minus the non-serializable `dirty` Set. */
export type SerializedAssociateState = Omit<AssociateState, 'dirty'>;
/** Same as DisassociateState minus the non-serializable `dirty` Set. */
export type SerializedDisassociateState = Omit<DisassociateState, 'dirty'>;
/** Same as ExecuteActionState minus the non-serializable `dirty` Set. */
export type SerializedExecuteActionState = Omit<ExecuteActionState, 'dirty'>;
/** Same as ExecuteFunctionState minus the non-serializable `dirty` Set. */
export type SerializedExecuteFunctionState = Omit<ExecuteFunctionState, 'dirty'>;
/** Same as ExecuteWorkflowState minus the non-serializable `dirty` Set. */
export type SerializedExecuteWorkflowState = Omit<ExecuteWorkflowState, 'dirty'>;
/**
 * Same as ManageFileState minus the non-serializable `dirty` Set AND the
 * transient binary payload fields (`bodyBase64`). A multi-MB base64 blob
 * in localStorage blows past every browser's quota; the user re-picks the
 * file when they re-open a saved manage-file request.
 */
export type SerializedManageFileState = Omit<ManageFileState, 'dirty' | 'bodyBase64'>;
/** Same as ManageImageState minus dirty + bodyBase64 (same rationale). */
export type SerializedManageImageState = Omit<ManageImageState, 'dirty' | 'bodyBase64'>;
/** Same as ManageAttachmentState minus dirty + bodyBase64 (same rationale). */
export type SerializedManageAttachmentState = Omit<ManageAttachmentState, 'dirty' | 'bodyBase64'>;

// ── Serialize / Deserialize ───────────────────────────────────────────

export function serializeRetrieveMultiple(
  state: RetrieveMultipleState,
): SerializedRetrieveMultipleState {
  // Strip the dirty Set; everything else is plain-object-safe (filter
  // tree, expand tree, prefer, headers are all interface-typed).
  const { dirty: _dirty, ...rest } = state;
  return rest;
}

export function deserializeRetrieveMultiple(
  snap: SerializedRetrieveMultipleState,
): RetrieveMultipleState {
  return { ...snap, dirty: new Set<string>() };
}

export function serializeRetrieveSingle(state: RetrieveSingleState): SerializedRetrieveSingleState {
  const { dirty: _dirty, ...rest } = state;
  return rest;
}

export function deserializeRetrieveSingle(
  snap: SerializedRetrieveSingleState,
): RetrieveSingleState {
  return { ...snap, dirty: new Set<string>() };
}

export function serializePredefinedQuery(
  state: PredefinedQueryState,
): SerializedPredefinedQueryState {
  const { dirty: _dirty, ...rest } = state;
  return rest;
}

export function deserializePredefinedQuery(
  snap: SerializedPredefinedQueryState,
): PredefinedQueryState {
  return { ...snap, dirty: new Set<string>() };
}

export function serializeDelete(state: DeleteState): SerializedDeleteState {
  // Strip dirty + confirmation fields (see SerializedDeleteState comment).
  const { dirty: _dirty, confirmText: _ct, acknowledged: _ack, ...rest } = state;
  return rest;
}

export function deserializeDelete(snap: SerializedDeleteState): DeleteState {
  // Re-init the one-shot safety fields. The user MUST re-type the
  // record name + check the box every time, even after loading a save.
  return { ...snap, dirty: new Set<string>(), confirmText: '', acknowledged: false };
}

export function serializeUpdate(state: UpdateState): SerializedUpdateState {
  const { dirty: _dirty, ...rest } = state;
  return rest;
}

export function deserializeUpdate(snap: SerializedUpdateState): UpdateState {
  // `nullFields` backfill — entries saved before the field existed default
  // to []. The serializer always emits the field going forward.
  return { ...snap, nullFields: snap.nullFields ?? [], dirty: new Set<string>() };
}

export function serializeCreate(state: CreateState): SerializedCreateState {
  const { dirty: _dirty, ...rest } = state;
  return rest;
}

export function deserializeCreate(snap: SerializedCreateState): CreateState {
  return { ...snap, nullFields: snap.nullFields ?? [], dirty: new Set<string>() };
}

export function serializeUpsert(state: UpsertState): SerializedUpsertState {
  const { dirty: _dirty, ...rest } = state;
  return rest;
}

export function deserializeUpsert(snap: SerializedUpsertState): UpsertState {
  return { ...snap, nullFields: snap.nullFields ?? [], dirty: new Set<string>() };
}

export function serializeMerge(state: MergeState): SerializedMergeState {
  // Strip dirty + the live snapshots — the latter get refetched by
  // MergeMode after deserialization once targetId / subordinateId settle.
  const { dirty: _d, targetSnapshot: _t, subordinateSnapshot: _s, ...rest } = state;
  return rest;
}

export function deserializeMerge(snap: SerializedMergeState): MergeState {
  return { ...snap, dirty: new Set<string>(), targetSnapshot: null, subordinateSnapshot: null };
}

export function serializeAssociate(state: AssociateState): SerializedAssociateState {
  const { dirty: _dirty, ...rest } = state;
  return rest;
}

export function deserializeAssociate(snap: SerializedAssociateState): AssociateState {
  // `targetNames` may be missing on entries saved before the field existed
  // (one-time migration). Default to {} so the editor renders the GUIDs
  // without "(name not resolved)" noise until the user re-picks.
  return { ...snap, targetNames: snap.targetNames ?? {}, dirty: new Set<string>() };
}

export function serializeDisassociate(state: DisassociateState): SerializedDisassociateState {
  const { dirty: _dirty, ...rest } = state;
  return rest;
}

export function deserializeDisassociate(snap: SerializedDisassociateState): DisassociateState {
  return { ...snap, targetNames: snap.targetNames ?? {}, dirty: new Set<string>() };
}

export function serializeExecuteAction(state: ExecuteActionState): SerializedExecuteActionState {
  const { dirty: _d, ...rest } = state;
  return rest;
}
export function deserializeExecuteAction(snap: SerializedExecuteActionState): ExecuteActionState {
  return { ...snap, dirty: new Set<string>() };
}

export function serializeExecuteFunction(
  state: ExecuteFunctionState,
): SerializedExecuteFunctionState {
  const { dirty: _d, ...rest } = state;
  return rest;
}
export function deserializeExecuteFunction(
  snap: SerializedExecuteFunctionState,
): ExecuteFunctionState {
  return { ...snap, dirty: new Set<string>() };
}

export function serializeExecuteWorkflow(
  state: ExecuteWorkflowState,
): SerializedExecuteWorkflowState {
  const { dirty: _d, ...rest } = state;
  return rest;
}
export function deserializeExecuteWorkflow(
  snap: SerializedExecuteWorkflowState,
): ExecuteWorkflowState {
  return { ...snap, dirty: new Set<string>() };
}

export function serializeManageFile(state: ManageFileState): SerializedManageFileState {
  // Strip the in-memory base64 payload — re-picked from disk on load.
  const { dirty: _d, bodyBase64: _b, ...rest } = state;
  return rest;
}
export function deserializeManageFile(snap: SerializedManageFileState): ManageFileState {
  return { ...snap, bodyBase64: '', dirty: new Set<string>() };
}

export function serializeManageImage(state: ManageImageState): SerializedManageImageState {
  const { dirty: _d, bodyBase64: _b, ...rest } = state;
  return rest;
}
export function deserializeManageImage(snap: SerializedManageImageState): ManageImageState {
  return { ...snap, bodyBase64: '', dirty: new Set<string>() };
}

export function serializeManageAttachment(
  state: ManageAttachmentState,
): SerializedManageAttachmentState {
  const { dirty: _d, bodyBase64: _b, ...rest } = state;
  return rest;
}
export function deserializeManageAttachment(
  snap: SerializedManageAttachmentState,
): ManageAttachmentState {
  return { ...snap, bodyBase64: '', dirty: new Set<string>() };
}

// ── localStorage I/O ──────────────────────────────────────────────────

/**
 * Read the full saved-requests list. Returns [] on any failure — never
 * throws (corruption shouldn't break the app's load path).
 */
export function loadSaved(): SavedRequest[] {
  if (typeof window === 'undefined') return [];
  try {
    const raw = window.localStorage.getItem(getSavedRequestsKey());
    if (!raw) return [];
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) return [];
    const scope = getOrgScope();
    // Light shape-validate — drop entries missing required fields rather
    // than nuking the whole array. Backfill `orgScope` for entries saved
    // before the field existed (one-time migration; we tag them with the
    // current host since we have no way to know the original).
    return parsed
      .filter(
        (e): e is SavedRequest =>
          e != null &&
          typeof e === 'object' &&
          typeof (e as SavedRequest).id === 'string' &&
          typeof (e as SavedRequest).name === 'string' &&
          typeof (e as SavedRequest).modeId === 'string' &&
          typeof (e as SavedRequest).state === 'object' &&
          typeof (e as SavedRequest).savedAt === 'number',
      )
      .map((e) => ({ ...e, orgScope: e.orgScope || scope }));
  } catch {
    return [];
  }
}

/** Overwrite the entire list. Best-effort — silently no-ops if storage fails. */
export function persistSaved(list: SavedRequest[]): void {
  if (typeof window === 'undefined') return;
  try {
    window.localStorage.setItem(getSavedRequestsKey(), JSON.stringify(list));
  } catch {
    // QuotaExceededError, privacy-mode, etc. — surface via the caller's
    // ok/error return so the UI can show "couldn't save" toast.
    throw new Error('localStorage write failed (quota or privacy mode).');
  }
}

// ── Hash + dirty detection ────────────────────────────────────────────
//
// Save button enables only when the current state diverges from the
// last-saved snapshot. Implemented by hashing the serialized state and
// comparing strings.
//
// We use a stable JSON stringification — Object.keys() isn't sorted by
// JSON.stringify in all engines, but in practice modern V8/SpiderMonkey
// preserve insertion order, and our state is built via consistent
// React state updates, so the key order matches across renders.
// For extra safety against false-positives, we could `JSON.stringify`
// with a sorted-keys replacer; deferred until it becomes a real issue.

/** Mode-agnostic state hash — stringifies whatever the mode hands in. */
export function hashState(state: unknown): string {
  try {
    return JSON.stringify(state);
  } catch {
    return '';
  }
}

// ── Auto-suggested name ───────────────────────────────────────────────

/**
 * Build a default name for a fresh save. Format varies per mode but the
 * shape is consistent: leading entity/table, then 1–3 short clause stats,
 * trailing date. Empty fields are omitted. User can override before Save.
 *
 * Modes call this via the SaveButton's `autoSuggestName` prop; the
 * function dispatches by modeId so callers don't need to know which
 * helper to invoke. New modes plug in by adding a case.
 */
export function autoSuggestName(modeId: SavedModeId, state: unknown): string {
  switch (modeId) {
    case 'retrieve-multiple':
      return autoSuggestRetrieveMultiple(state as SerializedRetrieveMultipleState);
    case 'retrieve-single':
      return autoSuggestRetrieveSingle(state as SerializedRetrieveSingleState);
    case 'predefined-query':
      return autoSuggestPredefinedQuery(state as SerializedPredefinedQueryState);
    case 'delete':
      return autoSuggestDelete(state as SerializedDeleteState);
    case 'update':
      return autoSuggestUpdate(state as SerializedUpdateState);
    case 'create':
      return autoSuggestCreate(state as SerializedCreateState);
    case 'upsert':
      return autoSuggestUpsert(state as SerializedUpsertState);
    case 'merge':
      return autoSuggestMerge(state as SerializedMergeState);
    case 'associate':
      return autoSuggestAssociate(state as SerializedAssociateState);
    case 'disassociate':
      return autoSuggestDisassociate(state as SerializedDisassociateState);
    case 'exec-action':
      return autoSuggestExecuteAction(state as SerializedExecuteActionState);
    case 'exec-function':
      return autoSuggestExecuteFunction(state as SerializedExecuteFunctionState);
    case 'exec-workflow':
      return autoSuggestExecuteWorkflow(state as SerializedExecuteWorkflowState);
    case 'manage-file':
      return autoSuggestManageFile(state as SerializedManageFileState);
    case 'manage-image':
      return autoSuggestManageImage(state as SerializedManageImageState);
    case 'manage-attachment':
      return autoSuggestManageAttachment(state as SerializedManageAttachmentState);
  }
}

function autoSuggestManageAttachment(state: SerializedManageAttachmentState): string {
  // Format: "attachment · upload · activitymimeattachment.body · contract.pdf · 2026-05-23"
  const parts: string[] = [state.target, state.operation];
  if (state.target === 'attachment') parts.push('activitymimeattachment.body');
  else parts.push('annotation.documentbody');
  if (state.recordId) parts.push(`#${state.recordId.replace(/[{}-]/g, '').slice(0, 8)}`);
  if (state.fileName) parts.push(state.fileName);
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}

function autoSuggestManageFile(state: SerializedManageFileState): string {
  // Format: "file · upload · account.sample_contractfile · contract.pdf · 2026-05-19"
  const parts: string[] = ['file', state.operation];
  if (state.table && state.fileColumn) parts.push(`${state.table}.${state.fileColumn}`);
  else if (state.table) parts.push(state.table);
  if (state.recordId) parts.push(`#${state.recordId.replace(/[{}-]/g, '').slice(0, 8)}`);
  if (state.fileName) parts.push(state.fileName);
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}

function autoSuggestManageImage(state: SerializedManageImageState): string {
  const parts: string[] = ['image', state.operation];
  if (state.table && state.imageColumn) parts.push(`${state.table}.${state.imageColumn}`);
  else if (state.table) parts.push(state.table);
  if (state.recordId) parts.push(`#${state.recordId.replace(/[{}-]/g, '').slice(0, 8)}`);
  if (state.operation === 'download') parts.push(state.downloadSize);
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}

function autoSuggestExecuteAction(state: SerializedExecuteActionState): string {
  // Format: "action · WinOpportunity · 3 params · 2026-05-19"
  const parts: string[] = [state.category];
  parts.push(state.actionName || 'untitled');
  if (state.boundRecordId) parts.push(`#${state.boundRecordId.replace(/[{}-]/g, '').slice(0, 8)}`);
  const n = Object.values(state.paramValues || {}).filter((v) => v != null && v !== '').length;
  if (n) parts.push(`${n} param${n === 1 ? '' : 's'}`);
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}
function autoSuggestExecuteFunction(state: SerializedExecuteFunctionState): string {
  const parts: string[] = ['function'];
  parts.push(state.functionName || 'untitled');
  if (state.boundRecordId) parts.push(`#${state.boundRecordId.replace(/[{}-]/g, '').slice(0, 8)}`);
  const n = Object.values(state.paramValues || {}).filter((v) => v != null && v !== '').length;
  if (n) parts.push(`${n} param${n === 1 ? '' : 's'}`);
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}
function autoSuggestExecuteWorkflow(state: SerializedExecuteWorkflowState): string {
  const parts: string[] = ['workflow'];
  if (state.workflowId) parts.push(`wf:${state.workflowId.replace(/[{}-]/g, '').slice(0, 8)}`);
  if (state.entityId) parts.push(`#${state.entityId.replace(/[{}-]/g, '').slice(0, 8)}`);
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}

function autoSuggestDisassociate(state: SerializedDisassociateState): string {
  const parts: string[] = ['disassociate'];
  parts.push(state.table || 'untitled');
  if (state.navProperty) parts.push(`× ${state.navProperty}`);
  if (state.targetIds?.length)
    parts.push(`${state.targetIds.length} target${state.targetIds.length === 1 ? '' : 's'}`);
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}

function autoSuggestAssociate(state: SerializedAssociateState): string {
  // Format: "associate · account → contact_customer_accounts · 3 targets · 2026-05-19"
  const parts: string[] = ['associate'];
  parts.push(state.table || 'untitled');
  if (state.navProperty) parts.push(`→ ${state.navProperty}`);
  if (state.targets?.length)
    parts.push(`${state.targets.length} target${state.targets.length === 1 ? '' : 's'}`);
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}

function autoSuggestRetrieveMultiple(state: SerializedRetrieveMultipleState): string {
  const parts: string[] = [];
  parts.push(state.table || 'untitled');
  if (state.select.length) parts.push(`${state.select.length} sel`);
  const filterCount = countFilterRules(state.filter);
  if (filterCount) parts.push(`${filterCount} flt`);
  if (state.orderby.length) parts.push(`${state.orderby.length} sort`);
  if (state.expand.length) parts.push(`${state.expand.length} exp`);
  if (state.apply?.enabled) parts.push('apply');
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}

function autoSuggestMerge(state: SerializedMergeState): string {
  const parts: string[] = ['merge'];
  parts.push(state.table || 'untitled');
  if (state.targetId) parts.push(`tgt:${state.targetId.replace(/[{}-]/g, '').slice(0, 8)}`);
  if (state.subordinateId)
    parts.push(`sub:${state.subordinateId.replace(/[{}-]/g, '').slice(0, 8)}`);
  const overrides = Object.values(state.fieldChoices || {}).filter((c) => c !== 'target').length;
  if (overrides) parts.push(`${overrides} override${overrides === 1 ? '' : 's'}`);
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}

function autoSuggestUpsert(state: SerializedUpsertState): string {
  const parts: string[] = ['upsert'];
  parts.push(state.table || 'untitled');
  // Disposition mode — distinguishes the three Upsert flavors in the
  // library list. None = full upsert; create-only / update-only constrain
  // the operation.
  if (state.concurrency.kind === 'create-only') parts.push('create-only');
  else if (state.concurrency.kind === 'update-only') parts.push('update-only');
  else if (state.concurrency.kind === 'etag') parts.push('etag');
  // Key identification
  if (state.key.kind === 'guid' && state.key.recordId) {
    parts.push(`#${state.key.recordId.replace(/[{}-]/g, '').slice(0, 8)}`);
  } else if (state.key.kind === 'alternate') {
    parts.push(`altkey:${state.key.keyName || '?'}`);
  }
  const n = Object.keys(state.fieldValues || {}).length;
  if (n) parts.push(`${n} field${n === 1 ? '' : 's'}`);
  if (state.bypass.businessLogic !== 'none' || state.bypass.suppressFlows) parts.push('bypass');
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}

function autoSuggestCreate(state: SerializedCreateState): string {
  const parts: string[] = ['create'];
  parts.push(state.table || 'untitled');
  const n = Object.keys(state.fieldValues || {}).length;
  if (n) parts.push(`${n} field${n === 1 ? '' : 's'}`);
  if (state.prefer?.returnRepresentation) parts.push('return=rep');
  if (state.bypass.businessLogic !== 'none' || state.bypass.suppressFlows) parts.push('bypass');
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}

function autoSuggestUpdate(state: SerializedUpdateState): string {
  const parts: string[] = [state.method.toLowerCase()];
  parts.push(state.table || 'untitled');
  if (state.recordId) parts.push(`#${state.recordId.replace(/[{}-]/g, '').slice(0, 8)}`);
  if (state.method === 'PUT' && state.putColumn) {
    parts.push(state.putColumn);
  } else {
    const n = Object.keys(state.fieldValues || {}).length;
    if (n) parts.push(`${n} field${n === 1 ? '' : 's'}`);
  }
  if (state.concurrency.kind === 'etag') parts.push('etag');
  if (state.bypass.businessLogic !== 'none' || state.bypass.suppressFlows) parts.push('bypass');
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}

function autoSuggestDelete(state: SerializedDeleteState): string {
  const parts: string[] = ['delete'];
  parts.push(state.table || 'untitled');
  if (state.recordId) parts.push(`#${state.recordId.replace(/[{}-]/g, '').slice(0, 8)}`);
  if (state.scope.kind === 'single-property') parts.push(`clear ${state.scope.column || '?'}`);
  if (state.concurrency.kind !== 'none')
    parts.push(state.concurrency.kind === 'etag' ? 'etag' : 'if-match');
  if (state.bypass.businessLogic !== 'none' || state.bypass.suppressFlows) parts.push('bypass');
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}

function autoSuggestPredefinedQuery(state: SerializedPredefinedQueryState): string {
  // Format: "savedQuery · account · #q-1a2b3c4d · 2026-05-18"
  // Saved queries are identified by GUID; we abbreviate to first 8 chars so
  // two saves on the same table with different queries can be told apart in
  // the library.
  const parts: string[] = [state.queryType === 'userQuery' ? 'userQuery' : 'savedQuery'];
  parts.push(state.table || 'untitled');
  if (state.queryId) parts.push(`#${state.queryId.replace(/[{}-]/g, '').slice(0, 8)}`);
  if (state.top && state.top !== 50) parts.push(`top:${state.top}`);
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}

function autoSuggestRetrieveSingle(state: SerializedRetrieveSingleState): string {
  const parts: string[] = [];
  parts.push(state.table || 'untitled');
  // Single-record fetches: the recordId is the whole point. Show a short
  // prefix so the user can tell two entries on the same table apart.
  if (state.recordId) parts.push(`#${state.recordId.replace(/[{}-]/g, '').slice(0, 8)}`);
  if (state.select.length) parts.push(`${state.select.length} sel`);
  if (state.expand.length) parts.push(`${state.expand.length} exp`);
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}

function countFilterRules(group: SerializedRetrieveMultipleState['filter']): number {
  let n = 0;
  for (const node of group.rules) {
    if (node.type === 'rule' || node.type === 'function') n++;
    else if (node.type === 'group') n += countFilterRules(node);
    else if (node.type === 'lambda') n += countFilterRules(node.inner) || 1;
  }
  return n;
}

/** Best-effort `table` accessor for the library-row meta line. Returns
 *  whatever the entry's state declares as `table`, or '—' if unset. */
export function tableNameFromState(state: unknown): string {
  if (state && typeof state === 'object' && 'table' in state) {
    const t = (state as { table?: unknown }).table;
    if (typeof t === 'string' && t) return t;
  }
  return '—';
}

// ── id generation ─────────────────────────────────────────────────────

export function newSavedId(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return `sv_${crypto.randomUUID()}`;
  }
  return `sv_${Date.now().toString(36)}_${Math.floor(Math.random() * 1e6).toString(36)}`;
}
