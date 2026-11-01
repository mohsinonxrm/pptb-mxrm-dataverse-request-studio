// Saved Requests — named, per-mode state snapshots persisted to localStorage.
//
// Persistence layers:
//   • In-memory recents      — last 8 runs per mode, lost on reload
//   • localStorage saves     — this module, up to MAX_SAVED entries per org scope
//   • Export/Import to disk  — a versioned JSON file (buildSavedExport /
//                              parseSavedImport), via the host's file dialogs
//
// Durability (#105, checked against the PPTB v1.2.7 source): a tool runs in
// a BrowserView with no `partition` (toolWindowManager.ts:458), i.e. in
// Electron's default persistent session under userData, and loads from
// `pptb-webview://<toolId>/` (browserviewProtocolManager.ts:236), so the
// localStorage origin is keyed by the tool id alone — not by version or
// install path. Updates keep the same id and nothing in PPTB clears
// storage data on update or uninstall, so the library survives both.
// `toolboxAPI.settings` is keyed by the same tool id (settingsManager.ts:244),
// so moving there would not be more durable. What does change the id is the
// install source (registry id vs `npm-<pkg>` vs `local-<pkg>`), which is
// what Export/Import is for.
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

import type {
  RetrieveMultipleState,
  RetrieveSingleState,
  PredefinedQueryState,
  RecentRun,
} from './readState';
import type { DeleteState, UpdateState, CreateState, UpsertState, MergeState } from './writeState';
import type { AssociateState, DisassociateState } from './relateState';
import type {
  ExecuteActionState,
  ExecuteFunctionState,
  ExecuteWorkflowState,
} from './executeState';
import type { ManageFileState, ManageImageState, ManageAttachmentState } from './binaryState';
import { getEnv } from '../mock/environment';
import { stripCredentialRows } from '../engine/sensitiveHeaders';
import { invalidUrlPartReason } from '../primitives/urlParts';
import {
  collectInvalidNames,
  expandNameProblems,
  type FilterGroup,
} from '../editors/filter/filterTree';
import {
  autoSuggestBatch,
  autoSuggestExecuteMultiple,
  type SerializedBatchState,
} from '../engine/batch/batchModel';

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
/**
 * Version of the SavedRequest entry shape, stamped on every entry and on
 * export files. Entries written before the field existed read as 1. The
 * stored value stays a plain array so older builds can still read it.
 * Bump when an entry's shape changes in a way a reader must know about.
 */
export const SAVED_SCHEMA_VERSION = 1;
/**
 * Library size per org scope, shared by every mode. A bound, not a target:
 * an entry is a few KB of JSON, so 200 stays far below the browser's
 * localStorage quota (a quota failure is reported by `persistSaved`
 * anyway). At the cap, saving a NEW request is refused until the user
 * deletes or replaces one — nothing is evicted silently (#100).
 */
export const MAX_SAVED = 200;

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

/** Union of every mode id that participates in the saved-request library. Add new modes here. */
export type SavedModeId =
  | 'retrieve-multiple'
  | 'retrieve-single'
  | 'predefined-query'
  | 'delete'
  // Bulk Delete job (#10) — query + schedule; see engine/bulkDelete.
  | 'bulk-delete'
  | 'update'
  | 'create'
  | 'upsert'
  | 'merge'
  // Bulk Write (#8) — rows + operation; see engine/bulkWrite.
  | 'bulk-write'
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
  | 'manage-attachment'
  // $batch composer (#9) — a list of raw requests grouped into change sets.
  | 'batch'
  // ExecuteMultiple-style mode (#11) — same state shape, one independent group.
  | 'execute-multiple';

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
  /** Entry shape version (SAVED_SCHEMA_VERSION); missing on pre-#105 entries = 1. */
  schemaVersion?: number;
}

/**
 * Every SavedModeId, as a record so a mode added to the union but not here
 * is a compile error. The 1.2.0 modes ($batch, ExecuteMultiple, Bulk Write,
 * Bulk Delete Job) were once missing from this list, which made import and
 * the other-connections view drop their entries (#220).
 */
const SAVED_MODES: Record<SavedModeId, true> = {
  'retrieve-multiple': true,
  'retrieve-single': true,
  'predefined-query': true,
  delete: true,
  'bulk-delete': true,
  update: true,
  create: true,
  upsert: true,
  merge: true,
  'bulk-write': true,
  associate: true,
  disassociate: true,
  'exec-action': true,
  'exec-function': true,
  'exec-workflow': true,
  'manage-file': true,
  'manage-image': true,
  'manage-attachment': true,
  batch: true,
  'execute-multiple': true,
};

/** Runtime list of SavedModeId, for validating stored and imported entries. */
export const SAVED_MODE_IDS: ReadonlySet<string> = new Set(Object.keys(SAVED_MODES));

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
  return {
    ...snap,
    nullFields: snap.nullFields ?? [],
    rawBody: snap.rawBody ?? null,
    dirty: new Set<string>(),
  };
}

export function serializeCreate(state: CreateState): SerializedCreateState {
  const { dirty: _dirty, ...rest } = state;
  return rest;
}

export function deserializeCreate(snap: SerializedCreateState): CreateState {
  return {
    ...snap,
    nullFields: snap.nullFields ?? [],
    rawBody: snap.rawBody ?? null,
    dirty: new Set<string>(),
  };
}

export function serializeUpsert(state: UpsertState): SerializedUpsertState {
  const { dirty: _dirty, ...rest } = state;
  return rest;
}

export function deserializeUpsert(snap: SerializedUpsertState): UpsertState {
  return {
    ...snap,
    nullFields: snap.nullFields ?? [],
    rawBody: snap.rawBody ?? null,
    dirty: new Set<string>(),
  };
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
//
// Reading never loses data (#220). The stored value is classified once, by
// `readLibrary`, for the current library, the other-connections view and
// import alike:
//   • entries    — what this build can list and load;
//   • preserved  — array items it can't (damaged, an unknown mode, a newer
//                  schema). Kept as the raw values and written back after
//                  the entries on every write, so a newer DRS's entries
//                  survive a round trip through an older one;
//   • unreadable — the stored text isn't a JSON array at all. The store
//                  refuses to write over it until the user keeps a copy or
//                  starts a new library; a backup copy is kept in storage.

/** A connection's library as read from storage. */
export interface SavedLibrary {
  entries: SavedRequest[];
  preserved: unknown[];
  /** The raw stored text when it isn't a JSON array, else null. */
  unreadable: string | null;
}

/** True for an entry this build can list and load. */
function isUsableEntry(e: unknown): e is SavedRequest {
  return (
    isSavedRequestShape(e) &&
    SAVED_MODE_IDS.has(e.modeId) &&
    (e.schemaVersion ?? 1) <= SAVED_SCHEMA_VERSION
  );
}

/**
 * Read a library bucket (default: the current connection's). Never throws.
 * Entries saved before `orgScope` / `schemaVersion` existed are backfilled
 * (tagged with the bucket's scope, schema 1).
 */
export function readLibrary(key = getSavedRequestsKey(), scope = getOrgScope()): SavedLibrary {
  const empty: SavedLibrary = { entries: [], preserved: [], unreadable: null };
  if (typeof window === 'undefined') return empty;
  let raw: string | null;
  try {
    raw = window.localStorage.getItem(key);
  } catch {
    return empty;
  }
  if (!raw) return empty;
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return { ...empty, unreadable: raw };
  }
  if (!Array.isArray(parsed)) return { ...empty, unreadable: raw };
  const entries: SavedRequest[] = [];
  const preserved: unknown[] = [];
  for (const e of parsed) {
    if (isUsableEntry(e)) {
      entries.push({
        ...e,
        state: stripCredentialRows(e.state),
        orgScope: e.orgScope || scope,
        schemaVersion: e.schemaVersion ?? 1,
      });
    } else {
      preserved.push(e);
    }
  }
  return { entries, preserved, unreadable: null };
}

/** Prefix of the backup keys for unreadable libraries — outside the library prefix. */
export const SAVED_BACKUP_KEY_PREFIX = 'drs:saved-requests-corrupt:v1';

/**
 * Keep a copy of an unreadable library under its own key (outside the
 * library prefix, so no connection lists it). An identical copy is never
 * written twice. Returns the backup key, or null when storage refused it.
 */
export function backupUnreadable(scope: string, raw: string): string | null {
  if (typeof window === 'undefined') return null;
  const prefix = `${SAVED_BACKUP_KEY_PREFIX}:${scope}:`;
  try {
    const ls = window.localStorage;
    for (let i = 0; i < ls.length; i++) {
      const k = ls.key(i);
      if (k?.startsWith(prefix) && ls.getItem(k) === raw) return k;
    }
    const key = `${prefix}${Date.now()}`;
    ls.setItem(key, raw);
    return key;
  } catch {
    return null;
  }
}

function isSavedRequestShape(e: unknown): e is SavedRequest {
  return (
    e != null &&
    typeof e === 'object' &&
    typeof (e as SavedRequest).id === 'string' &&
    typeof (e as SavedRequest).name === 'string' &&
    typeof (e as SavedRequest).modeId === 'string' &&
    typeof (e as SavedRequest).state === 'object' &&
    (e as SavedRequest).state !== null &&
    typeof (e as SavedRequest).savedAt === 'number'
  );
}

/**
 * Overwrite the current connection's library with `list`, followed by the
 * `preserved` raw items `readLibrary` couldn't use. Throws when storage
 * can't be written.
 */
export function persistSaved(list: SavedRequest[], preserved: readonly unknown[] = []): void {
  if (typeof window === 'undefined') return;
  try {
    window.localStorage.setItem(getSavedRequestsKey(), JSON.stringify([...list, ...preserved]));
  } catch {
    // QuotaExceededError, privacy-mode, etc. — surface via the caller's
    // ok/error return so the UI can show "couldn't save" toast.
    throw new Error('localStorage write failed (quota or privacy mode).');
  }
}

/**
 * Remove credential header rows already stored by an earlier build (#217),
 * in every connection's library. Rewrites a bucket only when an entry's
 * state changed, and keeps every other item exactly as stored (damaged
 * items and entries of modes this build doesn't know included). Buckets
 * that aren't a JSON array are left alone. Best-effort; never throws.
 */
export function scrubStoredCredentials(): void {
  if (typeof window === 'undefined') return;
  try {
    const ls = window.localStorage;
    const keys: string[] = [];
    for (let i = 0; i < ls.length; i++) {
      const k = ls.key(i);
      if (k?.startsWith(`${SAVED_REQUESTS_KEY_PREFIX}:`)) keys.push(k);
    }
    for (const key of keys) {
      let items: unknown;
      try {
        items = JSON.parse(ls.getItem(key) ?? '');
      } catch {
        continue;
      }
      if (!Array.isArray(items)) continue;
      let changed = false;
      const next = items.map((e) => {
        if (!e || typeof e !== 'object' || !('state' in e)) return e;
        const state = stripCredentialRows((e as { state: unknown }).state);
        if (state === (e as { state: unknown }).state) return e;
        changed = true;
        return { ...e, state };
      });
      if (changed) ls.setItem(key, JSON.stringify(next));
    }
  } catch {
    // Storage unavailable or full: entries are still stripped when read.
  }
}

// ── Other connections' libraries (#6) ─────────────────────────────────
//
// Every connection this browser profile has saved requests for has its own
// `drs:saved-requests:v1:<host>` key. These read / write those buckets
// directly so the library can list another connection's entries and copy or
// move them into the current one.

/** One other connection's saved-request library. */
export interface OrgLibrary {
  /** Org scope (the Dataverse host, or `default` outside PPTB). */
  scope: string;
  entries: SavedRequest[];
}

function readBucket(key: string, scope: string): SavedRequest[] {
  return readLibrary(key, scope).entries;
}

/**
 * Libraries of every connection except the current one that hold at least
 * one valid entry, largest first. Never throws.
 */
export function listOtherOrgLibraries(): OrgLibrary[] {
  if (typeof window === 'undefined') return [];
  const current = getSavedRequestsKey();
  const prefix = `${SAVED_REQUESTS_KEY_PREFIX}:`;
  const out: OrgLibrary[] = [];
  try {
    for (let i = 0; i < window.localStorage.length; i++) {
      const key = window.localStorage.key(i);
      if (!key || !key.startsWith(prefix) || key === current) continue;
      const scope = key.slice(prefix.length);
      const entries = readBucket(key, scope);
      if (entries.length) out.push({ scope, entries });
    }
  } catch {
    return out;
  }
  return out.sort((a, b) => b.entries.length - a.entries.length || a.scope.localeCompare(b.scope));
}

/**
 * Drop entries from another connection's library (the second half of a
 * move). Throws when storage can't be written, like `persistSaved`.
 */
export function removeFromOrgLibrary(scope: string, ids: ReadonlySet<string>): void {
  const key = `${SAVED_REQUESTS_KEY_PREFIX}:${scope}`;
  if (key === getSavedRequestsKey()) return;
  try {
    const raw = window.localStorage.getItem(key);
    const parsed = raw ? (JSON.parse(raw) as unknown) : [];
    if (!Array.isArray(parsed)) return;
    // Filter the raw array so entries this build can't read are kept as-is.
    const kept = parsed.filter(
      (e) => !(e && typeof e === 'object' && ids.has((e as { id?: unknown }).id as string)),
    );
    if (kept.length) window.localStorage.setItem(key, JSON.stringify(kept));
    else window.localStorage.removeItem(key);
  } catch {
    throw new Error('localStorage write failed (quota or privacy mode).');
  }
}

// ── Export / Import file ──────────────────────────────────────────────
//
// A self-describing JSON envelope. Import also accepts a bare array (a raw
// copy of the localStorage value). Entries are re-tagged with the importing
// org scope by the store; their original scope stays in the file only.

export const SAVED_EXPORT_FORMAT = 'drs-saved-requests';

export interface SavedRequestsExport {
  format: typeof SAVED_EXPORT_FORMAT;
  schemaVersion: number;
  exportedAt: string;
  orgScope: string;
  entries: SavedRequest[];
}

/** Serialize a library (one org scope) to the export file text. */
export function buildSavedExport(entries: SavedRequest[], orgScope: string): string {
  const file: SavedRequestsExport = {
    format: SAVED_EXPORT_FORMAT,
    schemaVersion: SAVED_SCHEMA_VERSION,
    exportedAt: new Date().toISOString(),
    orgScope,
    entries: entries.map((e) => ({
      ...e,
      state: stripCredentialRows(e.state),
      schemaVersion: e.schemaVersion ?? 1,
    })),
  };
  return JSON.stringify(file, null, 2);
}

type UrlPart = readonly [label: string, value: unknown];

/** What an entry puts into a URL: ids, names, and problems found in its query trees. */
interface EntryUrlParts {
  ids: UrlPart[];
  names: UrlPart[];
  /** From the filter / $expand trees (column, nav and alias names). */
  problems?: () => string[];
}

const isGroup = (v: unknown): v is FilterGroup =>
  !!v && typeof v === 'object' && Array.isArray((v as FilterGroup).rules);
const filterNames = (v: unknown) => (isGroup(v) ? collectInvalidNames(v) : []);
const expandNames = (v: unknown) =>
  Array.isArray(v) ? expandNameProblems(v as Parameters<typeof expandNameProblems>[0]) : [];

/**
 * The values each mode puts into a URL as-is: record ids (GUIDs), names
 * (tables, navigation properties, columns) and its query trees. Batch and
 * ExecuteMultiple aren't listed: their items also hold alternate keys and
 * `$n` references, and their builder checks those per item.
 */
const IMPORT_URL_PARTS: Partial<
  Record<SavedModeId, (s: Record<string, unknown>) => EntryUrlParts>
> = {
  'retrieve-multiple': (s) => ({
    ids: [],
    names: [
      ...listOf(s.select, 'Selected column'),
      ...(Array.isArray(s.orderby)
        ? s.orderby.map((o) => ['Sort column', (o as { col?: unknown })?.col] as const)
        : []),
    ],
    problems: () => [
      ...filterNames(s.filter),
      ...filterNames((s.apply as { prefilter?: unknown } | undefined)?.prefilter),
      ...expandNames(s.expand),
    ],
  }),
  'retrieve-single': (s) => ({
    ids: [['Record id', s.recordId]],
    names: listOf(s.select, 'Selected column'),
    problems: () => expandNames(s.expand),
  }),
  'predefined-query': (s) => {
    const r = (s.related ?? {}) as Record<string, unknown>;
    return {
      ids: [
        ['Query id', s.queryId],
        ['Source row id', r.recordId],
      ],
      names: [
        ['Source table', r.sourceTable],
        ['Relationship', r.navProperty],
      ],
    };
  },
  'bulk-delete': (s) => ({ ids: [], names: [], problems: () => filterNames(s.filter) }),
  delete: (s) => ({
    ids: [['Record id', s.recordId]],
    names: [['Column', (s.scope as { column?: unknown } | undefined)?.column]],
  }),
  update: (s) => ({ ids: [['Record id', s.recordId]], names: [['Column', s.putColumn]] }),
  upsert: (s) => {
    const k = (s.key ?? {}) as Record<string, unknown>;
    return { ids: k.kind === 'guid' ? [['Record id', k.recordId]] : [], names: [] };
  },
  merge: (s) => ({
    ids: [
      ['Master record id', s.targetId],
      ['Subordinate record id', s.subordinateId],
    ],
    names: [],
  }),
  associate: (s) => ({
    ids: [['Source record id', s.sourceId], ...listOf(s.targets, 'Target record id')],
    names: [['Navigation property', s.navProperty]],
  }),
  disassociate: (s) => ({
    ids: [['Source record id', s.sourceId], ...listOf(s.targetIds, 'Target record id')],
    names: [['Navigation property', s.navProperty]],
  }),
  'exec-action': (s) => ({ ids: [['Bound record id', s.boundRecordId]], names: [] }),
  'exec-function': (s) => ({ ids: [['Bound record id', s.boundRecordId]], names: [] }),
  'exec-workflow': (s) => ({
    ids: [
      ['Workflow id', s.workflowId],
      ['Record id', s.entityId],
    ],
    names: [],
  }),
  'manage-file': (s) => ({
    ids: [['Record id', s.recordId]],
    names: [['File column', s.fileColumn]],
  }),
  'manage-image': (s) => ({
    ids: [['Record id', s.recordId]],
    names: [['Image column', s.imageColumn]],
  }),
  'manage-attachment': (s) => ({
    ids: [
      ['Record id', s.recordId],
      ['Parent activity id', s.parentActivityId],
      ['Parent record id', s.parentEntityId],
    ],
    names: [['Parent table', s.parentEntityLogical]],
  }),
};

/** A list field as parts; a value that isn't an array is reported as one part. */
function listOf(v: unknown, label: string): UrlPart[] {
  return Array.isArray(v) ? v.map((x) => [label, x] as const) : v == null ? [] : [[label, v]];
}

/**
 * Why an entry from outside this library (an import file, another
 * connection's library) can't be loaded safely, or null (#221). Its ids must
 * be GUIDs and its names plain names: such a file is untrusted, and a value
 * like `x)/other_nav?$filter=…` would send a different request from the one
 * its name describes. Empty values are fine (the mode asks for them).
 */
export function importedEntryProblem(e: SavedRequest): string | null {
  const s = e.state as Record<string, unknown>;
  const parts = IMPORT_URL_PARTS[e.modeId]?.(s) ?? { ids: [], names: [] };
  const names: UrlPart[] = [['Table', s.table], ...parts.names];
  for (const [label, value] of [...parts.ids, ...names]) {
    if (value != null && typeof value !== 'string') return `${label} isn't text.`;
  }
  const text = (ps: UrlPart[]) => ps as ReadonlyArray<readonly [string, string | null | undefined]>;
  const reason = invalidUrlPartReason(text(parts.ids), text(names));
  if (reason) return reason;
  try {
    return parts.problems?.()[0] ?? null;
  } catch {
    // A tree that isn't one (nodes that aren't objects, …).
    return "Its query can't be read.";
  }
}

/**
 * Parse an export file. Throws an Error with a user-facing message when the
 * file isn't a saved-requests export at all; otherwise returns the valid
 * entries and how many were dropped (bad shape, unknown mode, newer schema,
 * or an id or name that isn't one; `problems` says why for those).
 */
export function parseSavedImport(text: string): {
  entries: SavedRequest[];
  invalid: number;
  problems: string[];
} {
  let parsed: unknown;
  try {
    parsed = JSON.parse(text.charCodeAt(0) === 0xfeff ? text.slice(1) : text);
  } catch {
    throw new Error('The file is not valid JSON.');
  }
  let raw: unknown[];
  if (Array.isArray(parsed)) {
    raw = parsed;
  } else if (
    parsed &&
    typeof parsed === 'object' &&
    (parsed as SavedRequestsExport).format === SAVED_EXPORT_FORMAT &&
    Array.isArray((parsed as SavedRequestsExport).entries)
  ) {
    const v = (parsed as SavedRequestsExport).schemaVersion;
    if (typeof v === 'number' && v > SAVED_SCHEMA_VERSION) {
      throw new Error(
        `The file was exported by a newer version of Dataverse Request Studio (schema ${v}). Update the tool and try again.`,
      );
    }
    raw = (parsed as SavedRequestsExport).entries;
  } else {
    throw new Error('The file is not a Dataverse Request Studio saved-requests export.');
  }
  const entries: SavedRequest[] = [];
  const problems: string[] = [];
  for (const e of raw) {
    if (!isUsableEntry(e)) continue;
    const problem = importedEntryProblem(e);
    if (problem) problems.push(`"${e.name}": ${problem}`);
    else entries.push(e);
  }
  return { entries, invalid: raw.length - entries.length, problems };
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
    case 'bulk-delete': {
      // Format: "bulk delete · account · 2026-10-17"
      const b = state as { table?: string };
      return ['bulk delete', b.table || 'untitled', new Date().toISOString().slice(0, 10)].join(
        ' · ',
      );
    }
    case 'update':
      return autoSuggestUpdate(state as SerializedUpdateState);
    case 'create':
      return autoSuggestCreate(state as SerializedCreateState);
    case 'upsert':
      return autoSuggestUpsert(state as SerializedUpsertState);
    case 'merge':
      return autoSuggestMerge(state as SerializedMergeState);
    case 'bulk-write': {
      // Format: "CreateMultiple · account · 12 rows · 2026-10-17"
      const b = state as { operation?: string; table?: string; rows?: unknown[] };
      const n = b.rows?.length ?? 0;
      return [
        b.operation || 'bulk',
        b.table || 'untitled',
        `${n} row${n === 1 ? '' : 's'}`,
        new Date().toISOString().slice(0, 10),
      ].join(' · ');
    }
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
    case 'batch':
      return autoSuggestBatch(state as SerializedBatchState);
    case 'execute-multiple':
      return autoSuggestExecuteMultiple(state as SerializedBatchState);
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
  // Related collection (#5): name the source table + nav.
  if (state.related?.navProperty) {
    parts.push(`via ${state.related.sourceTable}/${state.related.navProperty}`);
  }
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

// ── Same-target check (silent-overwrite guard) ────────────────────────
//
// State keys that pick WHAT a request targets. Changing one of them
// (switching the table, the action, the workflow, …) starts a different
// request, so a save made before the switch must not be silently
// overwritten by the new draft. Keys absent from both states are ignored,
// so each mode is compared on whichever of these it carries.
const TARGET_KEYS = ['table', 'actionName', 'functionName', 'workflowId'] as const;

function targetKeyValue(state: unknown, key: string): string {
  if (!state || typeof state !== 'object') return '';
  const v = (state as Record<string, unknown>)[key];
  return typeof v === 'string' ? v : '';
}

/**
 * True when two serialized snapshots point at the same target (same table /
 * action / function / workflow). The SaveButton uses this to decide whether
 * a remembered `lastSavedId` still belongs to the current draft before it
 * pre-fills that entry's name or overwrites it without asking (#58).
 */
export function sameSaveTarget(a: unknown, b: unknown): boolean {
  return TARGET_KEYS.every((k) => targetKeyValue(a, k) === targetKeyValue(b, k));
}

// ── id generation ─────────────────────────────────────────────────────

/**
 * Wrap a Recent-runs snapshot as a transient SavedRequest so a mode can
 * restore it through its existing `onLoadSaved` validator + hydrator (#2).
 * Never persisted. Returns null when the run was recorded without a
 * snapshot (nothing to restore).
 */
export function recentRunAsSavedEntry(run: RecentRun, modeId: SavedModeId): SavedRequest | null {
  if (run.snapshot === undefined || run.snapshot === null) return null;
  return {
    id: run.id,
    name: `Recent run (${run.method} · ${run.status})`,
    modeId,
    orgScope: getOrgScope(),
    state: run.snapshot,
    savedAt: run.ts,
  };
}

export function newSavedId(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return `sv_${crypto.randomUUID()}`;
  }
  return `sv_${Date.now().toString(36)}_${Math.floor(Math.random() * 1e6).toString(36)}`;
}
