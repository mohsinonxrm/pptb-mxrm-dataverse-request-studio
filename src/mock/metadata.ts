// Mock Dataverse metadata — modeled after the real Web API
// AttributeMetadata response shapes.
//
// Each column carries:
//   - logicalName / displayName / required / description
//   - attributeType (AttributeTypeCode enum string)
//   - per-type properties (Format, MaxLength, Precision, Targets, DateTimeBehavior,
//     OptionSet, MinValue/MaxValue) that drive the FilterEditor's operator
//     allowlist and value-input control.
//   - performance-relevant flags (isLogical, sourceType) that drive the
//     antipattern-advisory system.
//
// Field shapes are a discriminated union keyed on `attributeType`.
//
// ── Reference: when the real Web API is wired up ────────────────────────────
// This mock maps 1:1 to the Dataverse Web API metadata endpoint at
//   GET /api/data/v9.2/EntityDefinitions(LogicalName='<table>')
//     ?$expand=Attributes(...) — cast to the typed metadata classes
//
// Key MS Learn references for the mapping below:
//   • Use the Web API with metadata
//     https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-web-api-metadata
//   • Query metadata with the Web API
//     https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query-metadata-web-api
//   • Retrieve metadata by name or MetadataId
//     https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/retrieve-metadata-name-metadataid
//   • Create/update entity definitions
//     https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/create-update-entity-definitions-using-web-api
//   • Create/update column definitions
//     https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/create-update-column-definitions-using-web-api
//   • Create/update entity relationships
//     https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/create-update-entity-relationships-using-web-api
//   • AttributeMetadata.SourceType (SDK reference for the SourceType values)
//     https://learn.microsoft.com/en-us/dotnet/api/microsoft.xrm.sdk.metadata.attributemetadata.sourcetype

// ============================================================
// Attribute type code (subset that's relevant to filter UX)
// ============================================================
export type AttributeTypeCode =
  | 'BigInt'
  | 'Boolean'
  | 'Customer'
  | 'DateTime'
  | 'Decimal'
  | 'Double'
  | 'EntityName'
  | 'File'
  | 'Image'
  | 'Integer'
  | 'Lookup'
  | 'Memo'
  | 'Money'
  | 'Owner'
  | 'Picklist'
  | 'State'
  | 'Status'
  | 'String'
  | 'Uniqueidentifier'
  | 'MultiSelectPicklist';

// ============================================================
// Format enums (per the AttributeMetadata classes)
// ============================================================
export type StringFormat =
  | 'Email'
  | 'Text'
  | 'TextArea'
  | 'Url'
  | 'Phone'
  | 'TickerSymbol'
  | 'VersionNumber'
  | 'PhoneticGuide';
export type MemoFormat = 'Email' | 'Text' | 'TextArea' | 'Url' | 'Phone';
export type IntegerFormat = 'None' | 'Duration' | 'Language' | 'Locale' | 'TimeZone';
export type DateTimeFormat = 'DateAndTime' | 'DateOnly';
export type DateTimeBehavior = 'UserLocal' | 'DateOnly' | 'TimeZoneIndependent';

// ============================================================
// SourceType — for specialized columns (calculated / rollup / formula).
// Mirrors AttributeMetadata.SourceType from the SDK:
//   null → simple (most columns; field omitted)
//   0    → simple attribute
//   1    → calculated  (legacy classic formula)
//   2    → rollup      (aggregate over related records)
//   3    → formula     (Power Fx)
//   4    → prompt      (AI-generated)
// Per query-antipatterns docs, filtering on calculated/rollup/formula columns
// is throttled with the FilteringOnCalculatedColumns anti-pattern flag.
// Reference: https://learn.microsoft.com/en-us/dotnet/api/microsoft.xrm.sdk.metadata.attributemetadata.sourcetype
// ============================================================
export type SourceType = 0 | 1 | 2 | 3 | 4;

// ============================================================
// Discriminated column meta — one shape per AttributeTypeCode
// ============================================================
export interface BaseColumnMeta {
  logicalName: string;
  displayName: string;
  /**
   * Backwards-compat boolean — true iff `requiredLevel` is
   * `SystemRequired` or `ApplicationRequired`. New code should read
   * `requiredLevel` directly so it can distinguish system-locked from
   * business-required (which the maker can change).
   */
  required?: boolean;
  /**
   * Granular required-level from AttributeMetadata.RequiredLevel.Value.
   * Drives the Create mode's required-field gate.
   */
  requiredLevel?: 'None' | 'Recommended' | 'ApplicationRequired' | 'SystemRequired';
  /** True iff this attribute can appear in a POST body (Create). False
   *  for system-managed columns, formula/rollup outputs, etc. Drives the
   *  FieldSetEditor's column filter in Create mode. */
  isValidForCreate?: boolean;
  /** True iff this attribute can appear in a PATCH body (Update). Often
   *  a subset of isValidForCreate (write-once columns). Drives the
   *  FieldSetEditor's column filter in Update mode. */
  isValidForUpdate?: boolean;
  /** True iff this attribute can be read. Always true on standard
   *  attributes; surfaced for defensive $select filtering in Read modes. */
  isValidForRead?: boolean;
  /**
   * True iff the attribute is exposed via the OData / Web API. Hard wall —
   * if false, the attribute can't be selected, filtered, ordered, or
   * written to. We filter server-side AND drop the column at
   * mapAttribute time, so consumers should never see `false` here. The
   * field is carried for diagnostics + future "show all attributes"
   * affordances.
   */
  isValidOData?: boolean;
  description?: string;

  /**
   * Pre-computed OData attribute reference — the string DRS emits when
   * the column appears in `$select`, on the LHS of a `$filter` rule, or
   * inside an inner `$expand($select=...)`.
   *
   * For Lookup / Customer / Owner this is `_<logicalName>_value` (the
   * primitive GUID property in the OData schema). For every other
   * AttributeType it equals `logicalName`. Computed once in
   * `metadataProvider.mapAttribute` so encoders don't have to re-derive
   * it at every emit site. Optional during the type-only transition so
   * mock fixtures and tests don't break — encoders fall back to
   * `logicalName` when it's absent.
   */
  oDataName?: string;

  // ── Performance-relevant flags (drive antipattern advisories) ──
  // Mapped from AttributeMetadata.IsLogical (Web API):
  //   IsLogical = true → column lives in a different physical table; included
  //   in $select forces a join. Per query-antipatterns LargeAmountOfLogicalAttributes.
  isLogical?: boolean;

  // Mapped from AttributeMetadata.SourceType (Web API):
  //   non-null → specialized (calculated/rollup/formula/prompt). When set on
  //   a column used in $filter, Dataverse throws FilteringOnCalculatedColumns
  //   anti-pattern errors and may throttle the query.
  sourceType?: SourceType;
}

export interface StringColumnMeta extends BaseColumnMeta {
  attributeType: 'String';
  format: StringFormat;
  maxLength: number;
}
export interface MemoColumnMeta extends BaseColumnMeta {
  attributeType: 'Memo';
  format: MemoFormat;
  maxLength: number;
}
export interface IntegerColumnMeta extends BaseColumnMeta {
  attributeType: 'Integer';
  format: IntegerFormat;
  minValue?: number;
  maxValue?: number;
}
export interface BigIntColumnMeta extends BaseColumnMeta {
  attributeType: 'BigInt';
  minValue?: number;
  maxValue?: number;
}
export interface DecimalColumnMeta extends BaseColumnMeta {
  attributeType: 'Decimal';
  precision: number;
  minValue?: number;
  maxValue?: number;
}
export interface DoubleColumnMeta extends BaseColumnMeta {
  attributeType: 'Double';
  precision: number;
  minValue?: number;
  maxValue?: number;
}
export interface MoneyColumnMeta extends BaseColumnMeta {
  attributeType: 'Money';
  /** 0=Specific 1=Currency Decimal Precision 2=Pricing Decimal Precision */
  precisionSource: 0 | 1 | 2;
  precision: number;
  minValue?: number;
  maxValue?: number;
}
export interface BooleanColumnMeta extends BaseColumnMeta {
  attributeType: 'Boolean';
  trueOption: { value: 1; label: string };
  falseOption: { value: 0; label: string };
  defaultValue?: boolean;
}
export interface DateTimeColumnMeta extends BaseColumnMeta {
  attributeType: 'DateTime';
  format: DateTimeFormat;
  dateTimeBehavior: DateTimeBehavior;
}
export interface PicklistColumnMeta extends BaseColumnMeta {
  attributeType: 'Picklist';
  options: { value: number; label: string }[];
  isGlobal?: boolean;
  defaultFormValue?: number;
}
export interface MultiSelectPicklistColumnMeta extends BaseColumnMeta {
  attributeType: 'MultiSelectPicklist';
  options: { value: number; label: string }[];
}
export interface StateColumnMeta extends BaseColumnMeta {
  attributeType: 'State';
  options: { value: number; label: string }[];
}
export interface StatusColumnMeta extends BaseColumnMeta {
  attributeType: 'Status';
  /** statuscode entries with a back-reference to the state value */
  options: { value: number; label: string; state: number }[];
}
export interface LookupColumnMeta extends BaseColumnMeta {
  attributeType: 'Lookup';
  targets: string[]; // exactly 1 entity in a Lookup
}
export interface CustomerColumnMeta extends BaseColumnMeta {
  attributeType: 'Customer';
  targets: string[]; // typically ['account','contact']
}
export interface OwnerColumnMeta extends BaseColumnMeta {
  attributeType: 'Owner';
  targets: string[]; // typically ['systemuser','team']
}
export interface UniqueIdColumnMeta extends BaseColumnMeta {
  attributeType: 'Uniqueidentifier';
  isPrimaryKey?: boolean;
}
export interface EntityNameColumnMeta extends BaseColumnMeta {
  attributeType: 'EntityName';
  options: { value: number; label: string }[];
}
export interface FileColumnMeta extends BaseColumnMeta {
  attributeType: 'File';
  /** Max file size in KB. Defaults to 32768 (32 MB) per file column doc. Configurable up to 10 GB. */
  maxSizeInKB?: number;
  /** Special target table — when set, identifies the underlying body column on attachment/annotation. */
  binaryTarget?: 'file' | 'annotation' | 'attachment';
}
export interface ImageColumnMeta extends BaseColumnMeta {
  attributeType: 'Image';
  /**
   * Max file size in KB — image columns max out at 30 MB regardless. Per
   * image-column-data docs.
   */
  maxSizeInKB?: number;
  /** True for the table's primary image (entityimage). Settable on Create; thumbnail-only. */
  isPrimaryImage?: boolean;
  /**
   * Per [ImageAttributeMetadata.CanStoreFullImage]. When false, only the
   * 144×144 thumbnail is stored. When true, the full-size original is
   * downloadable via `?size=full`.
   */
  canStoreFullImage?: boolean;
}

export type ColumnMeta =
  | StringColumnMeta
  | MemoColumnMeta
  | IntegerColumnMeta
  | BigIntColumnMeta
  | DecimalColumnMeta
  | DoubleColumnMeta
  | MoneyColumnMeta
  | BooleanColumnMeta
  | DateTimeColumnMeta
  | PicklistColumnMeta
  | MultiSelectPicklistColumnMeta
  | StateColumnMeta
  | StatusColumnMeta
  | EntityNameColumnMeta
  | LookupColumnMeta
  | CustomerColumnMeta
  | OwnerColumnMeta
  | UniqueIdColumnMeta
  | FileColumnMeta
  | ImageColumnMeta;

// ============================================================
// Helpers — type predicates and "options" extractor
// ============================================================
export const isLookupLike = (
  c: ColumnMeta,
): c is LookupColumnMeta | CustomerColumnMeta | OwnerColumnMeta =>
  c.attributeType === 'Lookup' || c.attributeType === 'Customer' || c.attributeType === 'Owner';

export const isNumericLike = (
  c: ColumnMeta,
): c is
  | IntegerColumnMeta
  | BigIntColumnMeta
  | DecimalColumnMeta
  | DoubleColumnMeta
  | MoneyColumnMeta =>
  c.attributeType === 'Integer' ||
  c.attributeType === 'BigInt' ||
  c.attributeType === 'Decimal' ||
  c.attributeType === 'Double' ||
  c.attributeType === 'Money';

export const isDateLike = (c: ColumnMeta): c is DateTimeColumnMeta =>
  c.attributeType === 'DateTime';

export const isTextLike = (c: ColumnMeta): c is StringColumnMeta | MemoColumnMeta =>
  c.attributeType === 'String' || c.attributeType === 'Memo';

export const isOptionSetLike = (
  c: ColumnMeta,
): c is
  | PicklistColumnMeta
  | MultiSelectPicklistColumnMeta
  | StateColumnMeta
  | StatusColumnMeta
  | EntityNameColumnMeta
  | BooleanColumnMeta =>
  c.attributeType === 'Picklist' ||
  c.attributeType === 'MultiSelectPicklist' ||
  c.attributeType === 'State' ||
  c.attributeType === 'Status' ||
  c.attributeType === 'EntityName' ||
  c.attributeType === 'Boolean';

/** Returns the option list for any choice/state/status/boolean column, else undefined. */
export function columnOptions(c: ColumnMeta): { value: number; label: string }[] | undefined {
  if (
    c.attributeType === 'Picklist' ||
    c.attributeType === 'MultiSelectPicklist' ||
    c.attributeType === 'State' ||
    c.attributeType === 'Status' ||
    c.attributeType === 'EntityName'
  ) {
    return c.options;
  }
  if (c.attributeType === 'Boolean') {
    return [c.falseOption, c.trueOption];
  }
  return undefined;
}

export const isHiddenInFilter = (c: ColumnMeta): boolean =>
  c.attributeType === 'File' || c.attributeType === 'Image';

// ── Antipattern predicates (sourced from query-antipatterns docs) ───────────
// Reference: https://learn.microsoft.com/en-us/power-apps/developer/data-platform/query-antipatterns

/**
 * Per query-antipatterns "Avoid using conditions on large text columns":
 *   - Any Memo column qualifies as "large text".
 *   - String columns are large when MaxLength > 850.
 */
export const isLargeText = (c: ColumnMeta): boolean => {
  if (c.attributeType === 'Memo') return true;
  if (c.attributeType === 'String' && c.maxLength > 850) return true;
  return false;
};

/**
 * Per query-antipatterns "Avoid using formula or calculated columns":
 *   sourceType 1 (Calculated), 2 (Rollup), 3 (Formula), 4 (Prompt) are all
 *   computed at retrieval time and trigger FilteringOnCalculatedColumns.
 */
export const isComputedColumn = (c: ColumnMeta): boolean =>
  c.sourceType != null && c.sourceType > 0;

/**
 * Human-readable label for SourceType. Used in advisory copy.
 */
export const sourceTypeLabel = (s: SourceType | undefined): string | undefined => {
  if (s == null || s === 0) return undefined;
  switch (s) {
    case 1:
      return 'calculated';
    case 2:
      return 'rollup';
    case 3:
      return 'formula';
    case 4:
      return 'prompt';
  }
};

/**
 * Per query-antipatterns "Minimize the number of selected logical columns":
 *   IsLogical=true means the value is stored in a different physical table
 *   and Dataverse must join to retrieve it.
 */
export const isLogicalColumn = (c: ColumnMeta): boolean => c.isLogical === true;

/** Underlying Edm type — used for OData literal rendering. */
export function edmTypeOf(
  c: ColumnMeta,
): 'string' | 'number' | 'integer' | 'boolean' | 'date' | 'datetime' | 'guid' {
  switch (c.attributeType) {
    case 'String':
    case 'Memo':
      return 'string';
    case 'Integer':
    case 'BigInt':
      return 'integer';
    case 'Decimal':
    case 'Double':
    case 'Money':
      return 'number';
    case 'Boolean':
      return 'boolean';
    case 'DateTime':
      return c.format === 'DateOnly' ? 'date' : 'datetime';
    case 'Lookup':
    case 'Customer':
    case 'Owner':
    case 'Uniqueidentifier':
      return 'guid';
    case 'Picklist':
    case 'MultiSelectPicklist':
    case 'State':
    case 'Status':
    case 'EntityName':
      return 'integer';
    default:
      return 'string';
  }
}

// ============================================================
// Navigation properties (relationships) — used by $expand + lambda
// ============================================================
export interface NavProperty {
  name: string;
  relationshipName: string;
  cardinality: 'OneToMany' | 'ManyToOne' | 'ManyToMany';
  targetEntity: string;
  /**
   * For ManyToOne navs: the source-side attribute logical name (e.g.
   * `regardingobjectid`, `customerid`, `primarycontactid`). The FK column.
   *
   * Critical for resolving the correct `@odata.bind` property name on
   * polymorphic lookups. Customer/Owner-like and multi-target Lookups have
   * MULTIPLE NavProperty entries with the same `referencingAttribute` but
   * different `targetEntity` and different `name` (the target-disambiguated
   * nav-property name, e.g. `customerid_account` vs `customerid_contact`,
   * or `regardingobjectid_account_task` vs `regardingobjectid_contact_task`).
   *
   * When the encoder builds a write body and a user picks a target for a
   * polymorphic lookup, we match by `(referencingAttribute === col.logicalName
   * && targetEntity === picked.targetEntity)` and emit `<navName>@odata.bind`.
   *
   * Always undefined for OneToMany / ManyToMany navs (which don't have a
   * single source attribute).
   */
  referencingAttribute?: string;
}

/**
 * Alternate key on a table. Used by Upsert to address rows by a business
 * identifier (e.g. accountnumber) instead of the GUID primary key. Multiple
 * columns form a composite key — emit as `(col1='v1',col2='v2')` in the URL.
 *
 * See https://learn.microsoft.com/en-us/power-apps/maker/data-platform/define-alternate-keys-reference-records
 */
export interface AlternateKeyDef {
  /** Internal name of the key definition (e.g. `akey_accountnumber`). */
  name: string;
  /** UI label — e.g. "By Account Number". */
  displayName: string;
  /** Column logical names that make up the key (order matters in the URL). */
  columns: string[];
}

export interface TableMeta {
  logicalName: string;
  entitySetName: string;
  displayName: string;
  primaryKey: string;
  primaryName: string;
  columns: ColumnMeta[];
  navigationProperties: NavProperty[];
  /** Alternate keys defined on the table (optional — used by Upsert). */
  alternateKeys?: AlternateKeyDef[];
  /** True if this table is one of the Merge-supported types (account / contact / incident). */
  supportsMerge?: boolean;
}

// ════════════════════════════════════════════════════════════════════════
// Live-metadata registry — populated by host/metadataProvider.ts
// ════════════════════════════════════════════════════════════════════════
//
// `findTable()` and `TABLES` read from this registry. The provider transforms
// raw Dataverse metadata (EntityMetadata + AttributeMetadata[] +
// RelationshipMetadata) into the studio's `TableMeta` shape and calls
// `__registerLiveTable` once per fetched entity.
//
// In standalone mode (no PPTB host), the registry stays empty. Editors that
// depend on `findTable()` return undefined and render their "pick a table"
// state. There is NO mock-data fallback — DRS is now a live-Dataverse-only
// tool, matching the user's directive.

const __liveTables = new Map<string, TableMeta>();
const __liveListeners = new Set<() => void>();

export function __registerLiveTable(t: TableMeta): void {
  __liveTables.set(t.logicalName, t);
  for (const l of __liveListeners) l();
}
export function __registerLiveTables(list: TableMeta[]): void {
  for (const t of list) __liveTables.set(t.logicalName, t);
  for (const l of __liveListeners) l();
}
export function __subscribeLiveTables(cb: () => void): () => void {
  __liveListeners.add(cb);
  return () => __liveListeners.delete(cb);
}
export function __clearLiveTables(): void {
  __liveTables.clear();
  for (const l of __liveListeners) l();
}
export function __hasLiveTables(): boolean {
  return __liveTables.size > 0;
}

/**
 * Live-only entity list. Empty until the metadata provider has fetched and
 * registered at least one table. Reactive components should use the
 * `useLiveEntities()` hook or `useLazyMetadata().loadEntities()` instead
 * (those give you the raw EntityMetadata[] which is what the entity picker
 * actually wants).
 */
export function getLiveTables(): TableMeta[] {
  return Array.from(__liveTables.values());
}

/**
 * Backward-compat shim: editors that import `TABLES` see a snapshot of the
 * current live registry. The array is regenerated on every property access
 * via a Proxy so it stays fresh as new tables are registered.
 */
export const TABLES: TableMeta[] = new Proxy([] as TableMeta[], {
  get(_t, key: string | symbol) {
    const arr = Array.from(__liveTables.values());
    return arr[key as keyof typeof arr];
  },
  has(_t, key) {
    return key in Array.from(__liveTables.values());
  },
  ownKeys() {
    return Reflect.ownKeys(Array.from(__liveTables.values()));
  },
  getOwnPropertyDescriptor(_t, key) {
    return Object.getOwnPropertyDescriptor(Array.from(__liveTables.values()), key);
  },
});

/**
 * Synchronous TableMeta lookup. Returns the studio-shape TableMeta when the
 * entity is in the live registry, otherwise undefined. The mode-level
 * `useLiveTable(logical)` hook subscribes to registry updates so child
 * editors re-render when a fetch resolves.
 */
export const findTable = (logicalName: string): TableMeta | undefined =>
  __liveTables.get(logicalName);

export const findColumn = (table: TableMeta, logicalName: string): ColumnMeta | undefined =>
  table.columns.find((c) => c.logicalName === logicalName);

// ══════════════════════════════════════════════════════════════════════
// Canonical nav-path resolver — SINGLE SOURCE OF TRUTH.
// ══════════════════════════════════════════════════════════════════════
//
// A "nav-path" is a slash-delimited column reference that walks one or more
// single-valued (N:1) navigation properties to reach a column on a RELATED
// entity — e.g. `msdyn_opportunityid/abc_salesstage` or the chained form
// `primarycontactid/createdby/fullname`. The convention is: every segment
// EXCEPT the last is a nav property; the last segment is the leaf column.
//
// Before this resolver existed, the same walk was reimplemented in ~5 places
// (FilterEditor, NavPathColumnPicker, antipatterns, columnDisplayMap, the
// lambda encoder) and — critically — was MISSING from the $filter encoder's
// `colLookup`, which resolved the leaf against the ROOT table only. That made
// the encoder serialise a related-entity custom OptionSet (Edm.Int32) as a
// quoted string, which Dataverse rejects with 0x80060888 (issue #33). Routing
// every committed-path resolution through this one function makes the leaf's
// AttributeType drive quoting everywhere, consistently.
//
// Resolution is metadata-driven and lazy-aware: each hop is looked up in the
// live registry via `findTable`. If a related entity hasn't been fetched yet,
// `pendingTarget` names the first missing entity so callers can trigger a load
// (`useWarmReferencedTables` / `useLiveTable`) and re-resolve on the next
// render. This is the same self-healing contract the lambda encoder relies on.

export interface ResolvedNavPath {
  /** The resolved leaf column, on the deepest entity. Undefined when the
   *  path can't be fully resolved (bad segment, or a hop not yet loaded). */
  leaf?: ColumnMeta;
  /** The entity the leaf lives on (the deepest entity walked). For a bare
   *  column this is the root table. */
  ownerTable?: TableMeta;
  /** Entities visited, starting with the root: `[root, …related]`. */
  chain: TableMeta[];
  /** Logical name of the first nav target NOT yet in the live registry, if
   *  any. Drives lazy-loading; undefined once the whole chain is resolvable. */
  pendingTarget?: string;
}

/**
 * Resolve a (possibly aliased) nav-path against `rootTable`, walking N:1
 * navigation properties through the live registry.
 *
 * @param rootTable the entity the path is anchored on (root, or a lambda's
 *   target entity when resolving inside a lambda predicate).
 * @param path      slash-delimited path: `col` | `nav/col` | `nav/nav/col`.
 * @param opts.alias a lambda alias to strip from the head (e.g. `c` strips
 *   `c/jobtitle` → `jobtitle`, `c/primarycontactid/name` →
 *   `primarycontactid/name`).
 */
export function resolveNavPath(
  rootTable: TableMeta | undefined,
  path: string,
  opts?: { alias?: string },
): ResolvedNavPath {
  if (!rootTable) return { chain: [] };
  let p = path;
  const alias = opts?.alias;
  if (alias && p.startsWith(alias + '/')) p = p.slice(alias.length + 1);
  const segs = p.split('/').filter(Boolean);
  const chain: TableMeta[] = [rootTable];
  if (segs.length === 0) return { ownerTable: rootTable, chain };

  let current: TableMeta = rootTable;
  // Walk every segment except the leaf as an N:1 nav hop.
  for (let i = 0; i < segs.length - 1; i++) {
    const nav = current.navigationProperties.find(
      (n) => n.name === segs[i] && n.cardinality === 'ManyToOne',
    );
    // Unknown nav segment — can't go deeper; report what we have.
    if (!nav) return { ownerTable: current, chain };
    const target = findTable(nav.targetEntity);
    // Target not loaded yet — surface it for lazy-loading + re-resolve later.
    if (!target) return { ownerTable: current, chain, pendingTarget: nav.targetEntity };
    chain.push(target);
    current = target;
  }
  const leafSeg = segs[segs.length - 1];
  return {
    leaf: current.columns.find((c) => c.logicalName === leafSeg || c.oDataName === leafSeg),
    ownerTable: current,
    chain,
  };
}

/**
 * Predicate that hides "companion" read-only logical columns from picker
 * UIs. These are the denormalized `*name` / `*yominame` / `*codename` text
 * fields that Dataverse auto-populates from picklists, lookups, state/status
 * codes, etc. — they're not user-selectable in the modern UX and the
 * resolved name is better surfaced via FormattedValue annotations.
 *
 * Empirically validated against 8 standard entities (account, activity-
 * pointer, task, email, case, opportunity, lead, contact): this single
 * rule cleanly catches every `*name` companion (44+ per entity on
 * picklist-heavy tables) while preserving every genuinely-useful column
 * (subject on activities, address composites on customer tables, etc.).
 */
export function isCompanionLogicalReadOnly(c: ColumnMeta): boolean {
  return c.isLogical === true && c.isValidForCreate === false && c.isValidForUpdate === false;
}

// ══════════════════════════════════════════════════════════════════════
// Saved queries — TYPE-ONLY exports.
// ══════════════════════════════════════════════════════════════════════
// Pre-mock-strip we shipped a hardcoded fixture of common system + user
// queries for the Predefined-Query mode's picker. Now that DRS is live-only,
// the picker will fetch real savedquery + userquery rows from Dataverse via
// the host. For now we export an empty array so existing imports compile.
// The Predefined-Query mode editor will be wired to a live query loader in
// a follow-up pass.

export interface SavedQuery {
  id: string;
  name: string;
  type: 'savedQuery' | 'userQuery';
  entity: string;
  description: string;
  filterSummary: string;
  columns: string[];
}

export const SAVED_QUERIES: SavedQuery[] = [];
