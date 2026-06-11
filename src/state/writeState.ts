// Shared state shape for the Write group (Create / Update / Upsert / Delete / Merge).
//
// Each mode picks the subset it needs. We start with Create — Update/Upsert
// reuse the same field-value shape, and Delete only needs Target + a record id.

import type { PreferSpec } from '../editors/PreferEditor';
import type { HeaderItem } from '../editors/HeadersEditor';

// ──────────────────────────────────────────────────────────────
// Bypass options — shared by every write mode.
//
// Three discrete server-side behaviors, one combined toggle. Per
// https://learn.microsoft.com/en-us/power-apps/developer/data-platform/bypass-custom-business-logic
// and
// https://learn.microsoft.com/en-us/power-apps/developer/data-platform/bypass-power-automate-flows
// we model these as a single options object so the central header composer
// (and the BypassEditor pane) can render and emit them uniformly.
//
//   businessLogic — emits one of:
//     'none'  → no header
//     'sync'  → MSCRM.BypassBusinessLogicExecution: CustomSync
//     'async' → MSCRM.BypassBusinessLogicExecution: CustomAsync
//     'both'  → MSCRM.BypassBusinessLogicExecution: CustomSync,CustomAsync
//     'steps' → MSCRM.BypassBusinessLogicExecutionStepIds: <stepIds joined>
//
//   stepIds        — only meaningful when businessLogic = 'steps'.
//   useLegacyHeader — when true and businessLogic='sync', emit the legacy
//                     MSCRM.BypassCustomPluginExecution: true (requires the
//                     older prvBypassCustomPlugins privilege). Kept as an
//                     escape hatch — defaults false.
//   suppressFlows   — emits MSCRM.SuppressCallbackRegistrationExpanderJob: true.
//                     No privilege required, but the doc has a strong "don't
//                     unless you know" warning we surface as an advisory.
// ──────────────────────────────────────────────────────────────

export type BypassBusinessLogicMode = 'none' | 'sync' | 'async' | 'both' | 'steps';

export interface BypassOptions {
  businessLogic: BypassBusinessLogicMode;
  /** Plugin step GUIDs — only meaningful when mode='steps'. Server default cap 3, max 10. */
  stepIds: string[];
  /** Legacy header escape — uses MSCRM.BypassCustomPluginExecution (sync-only). */
  useLegacyHeader: boolean;
  /** MSCRM.SuppressCallbackRegistrationExpanderJob — Power Automate flow trigger bypass. */
  suppressFlows: boolean;
}

export const defaultBypassOptions = (): BypassOptions => ({
  businessLogic: 'none',
  stepIds: [],
  useLegacyHeader: false,
  suppressFlows: false,
});

// ──────────────────────────────────────────────────────────────
// Field-value shapes
// ──────────────────────────────────────────────────────────────
//
// Each entry in `fieldValues` corresponds to one writable column. The presence
// of a key means "include this field in the POST body" — removing the entry
// drops the field entirely from the request. This mirrors how the Dataverse
// Web API treats the body: only the keys you send are evaluated.
//
// The serializer (engine/urlBuilder.buildCreateBody) reads the column metadata
// to decide the on-wire shape:
//
//   • Lookup/Customer/Owner → `<col>@odata.bind` = `/<entitySet>(<guid>)`
//   • MultiSelectPicklist   → string of comma-separated integers (e.g. "1,2,3")
//   • DateTime              → ISO 8601 string (the runtime accepts both forms)
//   • DateOnly              → "yyyy-mm-dd"
//   • Money / Decimal / Double / Integer / BigInt → JSON number
//   • Boolean               → JSON boolean
//   • Picklist / State / Status / EntityName → JSON integer
//   • String / Memo / Uniqueidentifier → JSON string
//
// We don't tag the union (no `kind` discriminator). The serializer dispatches
// on AttributeTypeCode, so the in-memory value stays simple — drives the
// FieldSetEditor without needing a wrapper around every primitive input.

/** Stored shape for a lookup-like field (Lookup / Customer / Owner). */
export interface LookupFieldValue {
  /** GUID of the related row. */
  id: string;
  /** Entity logical name of the related table (must be one of the column's targets). */
  targetEntity: string;
}

/** Any value the FieldSetEditor can produce for a single column. */
export type CreateFieldValue = string | number | boolean | number[] | LookupFieldValue | null;

// ──────────────────────────────────────────────────────────────
// Create
// ──────────────────────────────────────────────────────────────
export interface CreateState {
  /** Logical name of the target table. */
  table: string;

  /**
   * Map of fieldLogicalName -> value. Presence of a key = include in body.
   * Deleting a key removes it from the request entirely.
   */
  fieldValues: Record<string, CreateFieldValue>;

  /**
   * Columns the user wants to send as EXPLICIT NULL in the body. Separate
   * from `fieldValues` to keep the "key absent = not in body" rule intact
   * (legacy null inside CreateFieldValue is still treated as skip).
   *
   * Body shape per docs:
   *   • Regular columns → `"<col>": null`
   *   • Lookup columns  → `"<nav>@odata.bind": null` (clears the lookup)
   *
   * Less common on Create (the server usually defaults missing columns) but
   * heavily used on Update / Upsert / Disassociate-single-valued where the
   * user wants to explicitly clear a value.
   *
   * Reference:
   *   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/associate-disassociate-entities-using-web-api#disassociate-by-using-a-single-valued-navigation-property
   */
  nullFields: string[];

  /** Prefer header — return=representation lives here and controls 201 vs 204. */
  prefer: PreferSpec;

  /** HTTP headers (Content-Type, OData-Version, OData-MaxVersion, Accept, etc.). */
  headers: HeaderItem[];

  /** $select on the URL — only used when Prefer: return=representation is set. */
  returnSelect: string[];

  /**
   * Send `MSCRM.SuppressDuplicateDetection: false` when true.
   * Per docs, duplicate detection is suppressed by default on create.
   */
  duplicateDetection: boolean;

  /** Custom-business-logic + Power-Automate bypass options. See BypassOptions. */
  bypass: BypassOptions;

  /** Tracks which clauses were edited since the last execute. */
  dirty: Set<string>;
}

// ──────────────────────────────────────────────────────────────
// Update
// ──────────────────────────────────────────────────────────────
//
// PATCH /<entitySet>(<id>) — the body shape is identical to Create, so we
// reuse CreateFieldValue and the buildCreateBody serializer. The differences
// live in:
//
//   • the URL (`(id)` segment)
//   • the response codes — 200 OK on `Prefer: return=representation` (not 201)
//     and 204 No Content otherwise
//   • the optional concurrency headers (`If-Match: *`, `If-Match: <etag>`,
//     `If-None-Match: *`) which gate the operation
//
// The docs (see /webapi/update-delete-entities-using-web-api) recommend
// `If-Match: *` as the default to prevent accidental upsert — we adopt that
// stance for the Update mode (Upsert mode will land separately and flip the
// default).

/**
 * Concurrency control for an Update / Upsert request.
 *
 *   • `none`        — no precondition header. Behaves as upsert (create if
 *                     missing) — only useful in Upsert mode.
 *   • `update-only` — `If-Match: *`. Prevents accidental create. Default for
 *                     the Update mode.
 *   • `etag`        — `If-Match: W/"<etag>"`. Optimistic concurrency: server
 *                     returns 412 Precondition Failed if the etag doesn't
 *                     match the row's current ETag.
 *   • `create-only` — `If-None-Match: *`. Upsert "create only" — server
 *                     returns 412 if the row already exists.
 */
export type ConcurrencyMode =
  | { kind: 'none' }
  | { kind: 'update-only' }
  | { kind: 'etag'; etag: string }
  | { kind: 'create-only' };

/**
 * Method selector for Update — controls the verb and URL shape.
 *
 *   • 'PATCH' (default) — multi-field update. URL: /<entitySet>(<id>).
 *     Body is the partial diff: `{ field1: value1, ... }`.
 *   • 'PUT' (single column) — set ONE column's value via the property URL.
 *     URL: /<entitySet>(<id>)/<putColumn>. Body: `{ "value": <scalar> }`.
 *     Per docs (update-delete-entities-using-web-api#update-a-single-property-value):
 *     "To update a single property value, use a PUT request and add the
 *      property name to the entity's Uri."
 *
 * Single-column UX is the v2.2 "drill into a column" pattern — clicking a row
 * in the Diff pane (or picking from the Field set) sets putColumn and flips
 * the method, surfacing the change in the URL bar.
 */
export type UpdateMethod = 'PATCH' | 'PUT';

export interface UpdateState {
  /** Logical name of the target table. */
  table: string;

  /** GUID of the record to update. */
  recordId: string | null;

  /** HTTP method — switches URL shape + body shape. */
  method: UpdateMethod;

  /** When method === 'PUT', the single column to update (URL segment). */
  putColumn: string | null;

  /**
   * Map of fieldLogicalName -> new value. Same encoding rules as CreateState:
   * presence = include in PATCH body. Per docs:
   *   "Only include the properties you are changing in the request body."
   * In PUT mode, only fieldValues[putColumn] is read.
   */
  fieldValues: Record<string, CreateFieldValue>;

  /**
   * Columns the user wants to send as EXPLICIT NULL — clears the column.
   * For lookups emits `"<nav>@odata.bind": null`; for regular columns emits
   * `"<col>": null`. Ignored in PUT mode (PUT single-property has its own
   * shape — clearing via PUT isn't a supported pattern in the Web API).
   */
  nullFields: string[];

  /** Prefer header — return=representation flips 204 ↔ 200 + body. */
  prefer: PreferSpec;

  /** HTTP headers (Content-Type, OData-*, MSCRMCallerID, etc.). */
  headers: HeaderItem[];

  /** $select on the URL — only emitted when Prefer: return=representation is on. */
  returnSelect: string[];

  /** Optimistic concurrency control. */
  concurrency: ConcurrencyMode;

  /** Custom-business-logic + Power-Automate bypass options. */
  bypass: BypassOptions;

  /** Tracks which clauses were edited since the last execute. */
  dirty: Set<string>;
}

// ──────────────────────────────────────────────────────────────
// Upsert
// ──────────────────────────────────────────────────────────────
//
// Same PATCH verb as Update, but addressed either by GUID OR by an alternate
// key. The concurrency mode flips the semantic:
//
//   • none          → upsert (create if missing, update if exists) — DEFAULT
//   • update-only   → If-Match: *      → 404 if missing
//   • create-only   → If-None-Match: * → 412 if exists
//
// Per docs (use-upsert-insert-update-record):
//   "When using alternate keys, don't include the alternate key values in the
//    body of the request" — the URL carries them, the body carries everything
//    else that needs to change (or get inserted).

/** How to address the upsert target — by primary key or by alternate key. */
export type UpsertKeyMode =
  | { kind: 'guid'; recordId: string | null }
  | {
      kind: 'alternate';
      /** Name of the AlternateKeyDef on the target table. */
      keyName: string;
      /** Map of keyColumnLogicalName -> value the user typed. */
      keyValues: Record<string, string>;
    };

export interface UpsertState {
  table: string;
  key: UpsertKeyMode;

  /** Body — same shape as Create/Update. */
  fieldValues: Record<string, CreateFieldValue>;

  /**
   * Columns the user wants to send as EXPLICIT NULL — see CreateState.nullFields.
   * On Upsert this is occasionally useful when "update an existing row's column
   * to null" is the intent; on the create-path the server defaults missing
   * columns anyway, so explicit-null vs absent is mostly equivalent there.
   */
  nullFields: string[];

  prefer: PreferSpec;
  headers: HeaderItem[];
  returnSelect: string[];
  concurrency: ConcurrencyMode;

  /** Custom-business-logic + Power-Automate bypass options. */
  bypass: BypassOptions;

  dirty: Set<string>;
}

// ──────────────────────────────────────────────────────────────
// Delete
// ──────────────────────────────────────────────────────────────
//
// DELETE /<entitySet>(<id>)
//   • 204 No Content    — happy path
//   • 404 Not Found     — row doesn't exist
//   • 412 Precondition Failed — when If-Match etag mismatches
//
// Optional scope:
//   • whole-row        — DELETE /<entitySet>(<id>) (default)
//   • single-property  — DELETE /<entitySet>(<id>)/<columnLogicalName>
//                        clears one column's value (not supported on
//                        single-valued navigation properties)

export type DeleteScope = { kind: 'whole-row' } | { kind: 'single-property'; column: string };

export interface DeleteState {
  table: string;
  recordId: string | null;
  scope: DeleteScope;
  concurrency: ConcurrencyMode; // Update-only / etag / none — same union
  headers: HeaderItem[];
  /**
   * When true, send `MSCRM.BypassCustomPluginExecution: true` (LEGACY header).
   * Kept for backward compat — new code paths should set `bypass.businessLogic = 'sync'`
   * which emits the modern `MSCRM.BypassBusinessLogicExecution: CustomSync`.
   * @deprecated Use `bypass.businessLogic` (+ `bypass.useLegacyHeader=true` if
   *   the legacy wire format is specifically required).
   */
  bypassCustomPlugins: boolean;
  /** Custom-business-logic + Power-Automate bypass options. Supersedes bypassCustomPlugins. */
  bypass: BypassOptions;
  /** Text the user must type to confirm the delete (matches the record's primary name). Gates Execute. */
  confirmText: string;
  /** Explicit "I understand this can't be undone" acknowledgement. Gates Execute. */
  acknowledged: boolean;
  dirty: Set<string>;
}

// ──────────────────────────────────────────────────────────────
// Merge
// ──────────────────────────────────────────────────────────────
//
// POST /Merge — action that merges Subordinate INTO Target. The merge action
// only supports `account`, `contact`, and `incident`. The body:
//
//   {
//     "Target":      { "@odata.type": "Microsoft.Dynamics.CRM.<table>", "<id>": "<guid>" },
//     "Subordinate": { "@odata.type": "Microsoft.Dynamics.CRM.<table>", "<id>": "<guid>" },
//     "UpdateContent": { "@odata.type": "Microsoft.Dynamics.CRM.<table>", ...changes },
//     "PerformParentingChecks": false
//   }
//
// Reference: https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/merge

/**
 * Per-field choice in the Merge dialog. Mirrors the model-driven app's
 * "Merge duplicate records" UX (https://learn.microsoft.com/en-us/power-apps/user/merge-duplicate-records):
 *
 *   • 'target'      → keep the Target's value (default). Field is NOT emitted
 *                     in UpdateContent — server retains Target's existing data.
 *   • 'subordinate' → use the Subordinate's value. Field is emitted in
 *                     UpdateContent with Subordinate's value, overwriting
 *                     whatever Target had.
 *   • 'custom'      → user-typed override. Field is emitted in UpdateContent
 *                     with customValues[field].
 *
 * Per docs (merge-entity-using-web-api):
 *   "Merging moves any useful data from the Subordinate record to the Target
 *    record. Any existing data in the Target record aren't overwritten."
 *
 * So the server already implicitly fills in empty Target fields from
 * Subordinate. UpdateContent only matters when the user wants to OVERRIDE
 * existing Target data — that's exactly when 'subordinate' or 'custom' fire.
 */
export type MergeFieldChoice = 'target' | 'subordinate' | 'custom';

export interface MergeState {
  /** Target / Subordinate share the same table type. */
  table: string;
  /** "Winner" record — keeps its id, absorbs the Subordinate. */
  targetId: string | null;
  /** "Loser" record — gets deactivated and re-parented. */
  subordinateId: string | null;

  /**
   * Per-field user choice. Keyed by column logicalName. Fields not in this
   * map default to 'target' (keep Target's value).
   */
  fieldChoices: Record<string, MergeFieldChoice>;

  /**
   * Custom-override values when fieldChoices[col] === 'custom'. The
   * MergeFieldDiff editor only emits into this map for the 'custom' choice;
   * 'subordinate' choices read directly from the Subordinate's row at body-
   * build time (see subordinateSnapshot below).
   */
  customValues: Record<string, CreateFieldValue>;

  /**
   * Live snapshots of the picked records. Populated by MergeMode after the
   * one-shot row-fetch resolves. NOT persisted in saved requests (transient
   * cache only — stale on reload, refetched after table+id rehydration).
   *
   *   • `targetSnapshot`      drives the field-diff "before" column for the Target
   *   • `subordinateSnapshot` drives both the diff display AND the
   *     `'subordinate'` choice path in `buildMergeBody` — the encoder needs
   *     access to the Subordinate's actual values to copy them into
   *     `UpdateContent`.
   */
  targetSnapshot?: Record<string, unknown> | null;
  subordinateSnapshot?: Record<string, unknown> | null;

  /** Enforces matching parent records on Target/Subordinate when true. */
  performParentingChecks: boolean;

  /** When true, sends `MSCRM.SuppressDuplicateDetection: true` so duplicate detection rules don't fire on the UpdateContent overwrite. */
  suppressDuplicateDetection: boolean;

  /** Custom-business-logic + Power-Automate bypass options. */
  bypass: BypassOptions;

  headers: HeaderItem[];
  dirty: Set<string>;
}
