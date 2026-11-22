// Metadata provider — transforms raw Dataverse metadata into the studio's
// `TableMeta` / `ColumnMeta` shape, then publishes it into the synchronous
// live registry in `metadata/metadata.ts` so every editor's `findTable(logical)`
// call returns live data.
//
// Architecture:
//
//   pptbClient.ts  →  dataverseMetadata.ts  →  metadataProvider.ts
//   (raw queries)     (3-tier cache loaders)   (raw → TableMeta + register)
//
// This module no longer carries its own cache — the `cache.ts` singleton
// is the single source of truth. Repeat reads are free because the
// loaders short-circuit on cache hit.

import * as loader from './dataverseMetadata';
import {
  __registerLiveTable,
  __clearLiveTables,
  __unregisterLiveTable,
  findTable,
  getLiveTables,
} from '../metadata/metadata';
import type {
  TableMeta,
  ColumnMeta,
  NavProperty,
  AlternateKeyDef,
  AttributeTypeCode,
  SourceType,
} from '../metadata/metadata';
import type { EntityMetadata, AttributeMetadata, RelationshipMetadata } from './pptbClient';
import { metadataCache } from './cache';
import { activeOrg } from './activeOrg';

// ── Public surface ──────────────────────────────────────────────────────

export interface EntityListItem {
  logicalName: string;
  displayName: string;
  entitySetName: string;
  metadataId: string;
  /** Shown in Advanced Find; scopes the table pickers (#44). */
  isValidForAdvancedFind?: boolean;
}

export interface MetadataProvider {
  /** Lightweight entity list for the table picker. */
  listEntities(): Promise<EntityListItem[]>;
  /** Full `TableMeta` — attributes + relationships + alt keys. */
  getTable(logical: string): Promise<TableMeta | undefined>;
  /** Synchronous cached read for editors that need an instant value. */
  peekTable(logical: string): TableMeta | undefined;
  /** Invalidate one table; the next read refetches. */
  invalidateTable(logical: string): void;
  /** Nuke all caches + the live registry. The next read refetches. */
  invalidateAll(): void;
  /**
   * Refresh in place: drop caches, then re-fetch every currently-registered
   * table and re-publish it. Unlike `invalidateAll`, editors keep showing the
   * (briefly stale) current data until fresh data lands, with no need to
   * re-navigate. Wired to the Settings "Refresh metadata" button.
   */
  refreshAll(): Promise<void>;
}

// ── Raw → studio-shape mappers ─────────────────────────────────────────

function labelOf(
  d: EntityMetadata['DisplayName'] | AttributeMetadata['DisplayName'],
  fallback: string,
): string {
  return d?.UserLocalizedLabel?.Label || fallback;
}

const ATTRIBUTE_TYPE_MAP: Record<string, AttributeTypeCode | undefined> = {
  String: 'String',
  Memo: 'Memo',
  Integer: 'Integer',
  BigInt: 'BigInt',
  Decimal: 'Decimal',
  Double: 'Double',
  Money: 'Money',
  Boolean: 'Boolean',
  DateTime: 'DateTime',
  Picklist: 'Picklist',
  MultiSelectPicklist: 'MultiSelectPicklist',
  Virtual: undefined,
  State: 'State',
  Status: 'Status',
  Lookup: 'Lookup',
  Customer: 'Customer',
  Owner: 'Owner',
  Uniqueidentifier: 'Uniqueidentifier',
  EntityName: 'EntityName',
  File: 'File',
  Image: 'Image',
};

/**
 * Returns the OData attribute reference DRS should emit for `$select`,
 * `$filter` LHS, and inner `$expand($select=...)`. Lookup/Customer/Owner
 * surface their value at `_<logicalName>_value` (a primitive GUID property
 * in the OData schema); every other AttributeType is addressed by its
 * logical name as-is. This mirrors the DRB `Column.ODataName` rule but is
 * grounded in the OData metadata contract (the bare lookup name is a
 * navigation property, not an attribute — see
 * https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-web-api-functions#address-lookup-and-customer-fields-in-functions).
 */
function odataNameOf(attrType: AttributeTypeCode, logicalName: string): string {
  if (attrType === 'Lookup' || attrType === 'Customer' || attrType === 'Owner') {
    return `_${logicalName}_value`;
  }
  return logicalName;
}

/**
 * The studio type for a raw attribute. Image and File columns report
 * `AttributeType: Virtual`; their real type is in `AttributeTypeName` (#305).
 * Other virtual columns (`…name` label columns) stay unmapped.
 */
export function attributeTypeOf(
  raw: Pick<AttributeMetadata, 'AttributeType' | 'AttributeTypeName'>,
): AttributeTypeCode | undefined {
  if (raw.AttributeType === 'Virtual') {
    const name = raw.AttributeTypeName?.Value;
    if (name === 'ImageType') return 'Image';
    if (name === 'FileType') return 'File';
    return undefined;
  }
  return ATTRIBUTE_TYPE_MAP[raw.AttributeType];
}

export function mapAttribute(raw: AttributeMetadata): ColumnMeta | null {
  const attrType = attributeTypeOf(raw);
  if (!attrType) return null;

  // We intentionally do NOT drop `IsValidODataAttribute=false` columns.
  // That flag is unreliable on inheritance-based entities (task, email,
  // appointment, etc.) where most inherited attributes report false at
  // the child level even though they're queryable at runtime via the
  // parent. See pptbClient.getEntityAttributes for the full reasoning.
  //
  // The flag is still projected onto the column for diagnostics; some
  // future "show all attrs including legacy ones" diagnostic mode might
  // surface it.

  // Required-level + write-eligibility — base projection too. Drives the
  // FieldSetEditor's column filter (Create vs Update) and Create's
  // required-field gate. `required` stays as the back-compat boolean
  // (true iff SystemRequired or ApplicationRequired); the new
  // `requiredLevel` is the granular four-value enum.
  const requiredLevel = raw.RequiredLevel?.Value;
  const required = requiredLevel === 'SystemRequired' || requiredLevel === 'ApplicationRequired';

  const base = {
    logicalName: raw.LogicalName,
    displayName: labelOf(raw.DisplayName, raw.LogicalName),
    oDataName: odataNameOf(attrType, raw.LogicalName),
    required,
    requiredLevel,
    isValidForCreate: raw.IsValidForCreate !== false, // default true if undefined
    isValidForUpdate: raw.IsValidForUpdate !== false,
    isValidForRead: raw.IsValidForRead !== false,
    isValidForAdvancedFind: raw.IsValidForAdvancedFind?.Value !== false,
    // Carried for diagnostics. Don't filter on it — see comment above.
    isValidOData: raw.IsValidODataAttribute !== false,
    // Fetched on the basic projection alongside everything else — see
    // `getEntityAttributes` $select in pptbClient.ts. Both flags drive
    // the filter-rule antipattern indicator (calculated / logical column).
    isLogical: !!raw.IsLogical,
    sourceType: (raw.SourceType ?? undefined) as SourceType | undefined,
  } as const;

  switch (attrType) {
    case 'String':
      return { ...base, attributeType: 'String', format: 'Text', maxLength: raw.MaxLength ?? 100 };
    case 'Memo':
      return { ...base, attributeType: 'Memo', format: 'Text', maxLength: raw.MaxLength ?? 2000 };
    case 'Integer':
      return {
        ...base,
        attributeType: 'Integer',
        format: 'None',
        minValue: raw.MinValue,
        maxValue: raw.MaxValue,
      };
    case 'BigInt':
      return { ...base, attributeType: 'BigInt', minValue: raw.MinValue, maxValue: raw.MaxValue };
    case 'Decimal':
      return {
        ...base,
        attributeType: 'Decimal',
        precision: raw.Precision ?? 2,
        minValue: raw.MinValue,
        maxValue: raw.MaxValue,
      };
    case 'Double':
      return {
        ...base,
        attributeType: 'Double',
        precision: raw.Precision ?? 2,
        minValue: raw.MinValue,
        maxValue: raw.MaxValue,
      };
    case 'Money':
      return {
        ...base,
        attributeType: 'Money',
        precisionSource: 1,
        precision: raw.Precision ?? 2,
        minValue: raw.MinValue,
        maxValue: raw.MaxValue,
      };
    case 'Boolean':
      return {
        ...base,
        attributeType: 'Boolean',
        trueOption: { value: 1, label: labelOf(raw.OptionSet?.TrueOption?.Label, 'Yes') },
        falseOption: { value: 0, label: labelOf(raw.OptionSet?.FalseOption?.Label, 'No') },
      };
    case 'DateTime':
      return {
        ...base,
        attributeType: 'DateTime',
        format: (raw.Format as 'DateAndTime' | 'DateOnly') ?? 'DateAndTime',
        dateTimeBehavior:
          (raw.DateTimeBehavior?.Value as 'UserLocal' | 'DateOnly' | 'TimeZoneIndependent') ??
          'UserLocal',
      };
    case 'Picklist':
    case 'MultiSelectPicklist':
    case 'State':
    case 'EntityName': {
      const opts = (raw.OptionSet?.Options ?? []).map((o) => ({
        value: o.Value,
        label: labelOf(o.Label, String(o.Value)),
      }));
      if (
        (attrType === 'Picklist' || attrType === 'MultiSelectPicklist') &&
        typeof raw.DefaultFormValue === 'number'
      ) {
        return {
          ...base,
          attributeType: attrType,
          options: opts,
          defaultFormValue: raw.DefaultFormValue,
        };
      }
      return { ...base, attributeType: attrType, options: opts };
    }
    case 'Status': {
      const opts = (raw.OptionSet?.Options ?? []).map((o) => ({
        value: o.Value,
        label: labelOf(o.Label, String(o.Value)),
        state: 0,
      }));
      return { ...base, attributeType: 'Status', options: opts };
    }
    case 'Lookup':
      return { ...base, attributeType: 'Lookup', targets: raw.Targets ?? [] };
    case 'Customer':
      return { ...base, attributeType: 'Customer', targets: raw.Targets ?? ['account', 'contact'] };
    case 'Owner':
      return { ...base, attributeType: 'Owner', targets: raw.Targets ?? ['systemuser', 'team'] };
    case 'Uniqueidentifier':
      return { ...base, attributeType: 'Uniqueidentifier' };
    case 'File':
      return { ...base, attributeType: 'File' };
    case 'Image':
      return { ...base, attributeType: 'Image' };
  }
}

function mapRelationships(
  o2m: RelationshipMetadata[],
  m2o: RelationshipMetadata[],
  m2m: RelationshipMetadata[],
  selfEntity: string,
): NavProperty[] {
  const out: NavProperty[] = [];

  // ── N:1 (ManyToOne — lookups from this entity to a parent) ──
  // Use ReferencingEntityNavigationPropertyName as the nav name. Critical
  // for polymorphic Customer/Owner lookups: the OData schema does NOT
  // expose the bare attribute name (e.g. `customerid`) as a nav property —
  // it exposes target-disambiguated names (`customerid_account`,
  // `customerid_contact`). Using the bare name in $expand fails with
  // 0x80060888 "Could not find a property named 'customerid'…".
  for (const r of m2o) {
    const navName =
      r.ReferencingEntityNavigationPropertyName || r.ReferencingAttribute || r.SchemaName;
    out.push({
      name: navName,
      relationshipName: r.SchemaName,
      cardinality: 'ManyToOne',
      targetEntity: r.ReferencedEntity ?? '',
      // Source FK column name. Polymorphic lookups (Customer / Owner /
      // multi-target Lookup) have multiple ManyToOne relationships sharing
      // this same `referencingAttribute` but with different `targetEntity`
      // and different `name`. The write-body encoder matches by
      // (referencingAttribute, targetEntity) to pick the right
      // `@odata.bind` property name. See urlBuilder.bindPropertyFor.
      referencingAttribute: r.ReferencingAttribute,
      isValidForAdvancedFind: r.IsValidForAdvancedFind !== false,
    });
  }

  // ── 1:N (OneToMany — child collections off this entity) ──
  // Use ReferencedEntityNavigationPropertyName when present; falls back
  // to SchemaName for older metadata.
  for (const r of o2m) {
    const navName = r.ReferencedEntityNavigationPropertyName || r.SchemaName;
    out.push({
      name: navName,
      relationshipName: r.SchemaName,
      cardinality: 'OneToMany',
      targetEntity: r.ReferencingEntity ?? '',
      isValidForAdvancedFind: r.IsValidForAdvancedFind !== false,
    });
  }

  // ── N:N (ManyToMany) ──
  // Each side has its own nav property name; pick the one belonging to
  // self so the user sees "accountleads" not "leadaccounts" on `account`.
  for (const r of m2m) {
    const isEntity1 = r.Entity1LogicalName === selfEntity;
    const navName =
      (isEntity1 ? r.Entity1NavigationPropertyName : r.Entity2NavigationPropertyName) ||
      r.SchemaName;
    out.push({
      name: navName,
      relationshipName: r.SchemaName,
      cardinality: 'ManyToMany',
      targetEntity: isEntity1 ? (r.Entity2LogicalName ?? '') : (r.Entity1LogicalName ?? ''),
      isValidForAdvancedFind: r.IsValidForAdvancedFind !== false,
    });
  }
  return out;
}

function mapEntityListItem(e: EntityMetadata): EntityListItem {
  return {
    logicalName: e.LogicalName,
    displayName: labelOf(e.DisplayName, e.LogicalName),
    entitySetName: e.EntitySetName,
    metadataId: e.MetadataId,
    isValidForAdvancedFind: e.IsValidForAdvancedFind !== false,
  };
}

// ── Provider impl ──────────────────────────────────────────────────────

// Keyed by `${org}|${logical}` — a build for the same table in another org
// (after a connection switch) is a different build.
let buildInFlight = new Map<string, Promise<TableMeta | undefined>>();

// Fired after a refresh / invalidate drops the caches, so lists held in hook
// state (the entity picker) re-fetch too instead of waiting for a remount.
const resetListeners = new Set<() => void>();
function notifyReset(): void {
  for (const l of resetListeners) l();
}
export function subscribeMetadataReset(cb: () => void): () => void {
  resetListeners.add(cb);
  return () => resetListeners.delete(cb);
}

/**
 * Fill `maxSizeInKB` (File / Image) and `canStoreFullImage` / `isPrimaryImage`
 * (Image) from the Image/FileAttributeMetadata type casts. Best effort: on
 * failure the columns keep the editors' documented defaults.
 */
export async function enrichBinaryColumns(logical: string, columns: ColumnMeta[]): Promise<void> {
  const casts = [
    {
      type: 'Image' as const,
      cast: 'Microsoft.Dynamics.CRM.ImageAttributeMetadata',
      select: 'LogicalName,MaxSizeInKB,CanStoreFullImage,IsPrimaryImage',
    },
    {
      type: 'File' as const,
      cast: 'Microsoft.Dynamics.CRM.FileAttributeMetadata',
      select: 'LogicalName,MaxSizeInKB',
    },
  ].filter((c) => columns.some((col) => col.attributeType === c.type));
  await Promise.all(
    casts.map(async ({ type, cast, select }) => {
      try {
        const rows = (await loader.loadAttributesByTypeCast(
          logical,
          cast,
          select,
        )) as unknown as Array<{
          LogicalName: string;
          MaxSizeInKB?: number | null;
          CanStoreFullImage?: boolean | null;
          IsPrimaryImage?: boolean | null;
        }>;
        for (const r of rows) {
          const col = columns.find((c) => c.logicalName === r.LogicalName);
          if (!col) continue;
          if (col.attributeType === 'File' && type === 'File') {
            if (typeof r.MaxSizeInKB === 'number') col.maxSizeInKB = r.MaxSizeInKB;
          } else if (col.attributeType === 'Image' && type === 'Image') {
            if (typeof r.MaxSizeInKB === 'number') col.maxSizeInKB = r.MaxSizeInKB;
            if (typeof r.CanStoreFullImage === 'boolean')
              col.canStoreFullImage = r.CanStoreFullImage;
            if (r.IsPrimaryImage === true) col.isPrimaryImage = true;
          }
        }
      } catch (e) {
        console.warn(`[metadata] ${type} column details for ${logical} failed`, e);
      }
    }),
  );
}

async function buildTable(logical: string): Promise<TableMeta | undefined> {
  // The org this build belongs to. The result is registered under it even if
  // the connection switches mid-fetch, so it can't leak into another org.
  const org = activeOrg();
  const flightKey = `${org}|${logical}`;
  // Dedupe concurrent builds for the same entity.
  const existing = buildInFlight.get(flightKey);
  if (existing) return existing;

  // Assigned before the build's first await, so the build can tell whether
  // it's still the current one when it lands (#335).
  let promise!: Promise<TableMeta | undefined>;
  // eslint-disable-next-line prefer-const -- as `const`, reading it inside its own initializer is TS2454
  promise = (async () => {
    try {
      // Basic load: 3 HTTP requests per entity, period.
      //   1. Entity metadata (PrimaryIdAttribute, EntitySetName, …)
      //   2. Basic attribute projection (LogicalName + DisplayName +
      //      AttributeType — no Targets, no OptionSet, no MaxLength)
      //   3. Relationships (1:N + N:1 + N:N)
      //
      // Type-specific properties (Targets for lookups, OptionSet for
      // picklists, MaxLength/Format for strings, MinValue/MaxValue/Precision
      // for numerics, etc.) are NOT fetched here — they're loaded lazily
      // per-column on demand by `useColumnDetail` when an editor needs them.
      // Keeping this initial load small is what stops us tripping Dataverse's
      // 100-concurrent-request cap.
      const [entity, attrs, rels, keys] = await Promise.all([
        loader.loadEntityMetadata(logical),
        loader.loadEntityAttributes(logical),
        loader.loadEntityRelationships(logical),
        // Alternate keys — separate /Keys child collection. Cheap fetch
        // (most entities have 0-2 keys) and we need it eagerly so Upsert
        // mode's alt-key picker is populated when the user lands on it.
        loader.loadEntityKeys(logical),
      ]);

      const columns = attrs.map(mapAttribute).filter((c): c is ColumnMeta => c !== null);

      // Mark the primary key.
      const pk = columns.find((c) => c.logicalName === entity.PrimaryIdAttribute);
      if (pk && pk.attributeType === 'Uniqueidentifier') pk.isPrimaryKey = true;

      // Mark the primary image (thumbnail-only, settable on Create).
      const img = columns.find((c) => c.logicalName === entity.PrimaryImageAttribute);
      if (img && img.attributeType === 'Image') img.isPrimaryImage = true;

      // Size limits + full-image storage live on the type-cast subtypes,
      // not the basic projection. One extra request per binary type, and
      // only for tables that have such columns (#305).
      await enrichBinaryColumns(logical, columns);

      // Map the EntityKeyMetadata projection into our lightweight
      // AlternateKeyDef shape. We keep only Active keys — Pending /
      // Failed status means the index isn't ready and Upsert will fail
      // at runtime. (We could surface Pending as a disabled option later
      // with a "key not ready yet" tooltip; for now: hide.)
      const alternateKeys: AlternateKeyDef[] = keys
        .filter((k) => k.EntityKeyIndexStatus === 'Active')
        .map((k) => ({
          name: k.LogicalName,
          displayName: k.DisplayName?.UserLocalizedLabel?.Label ?? k.LogicalName,
          columns: k.KeyAttributes ?? [],
        }))
        .filter((k) => k.columns.length > 0);

      // #64 — key columns missing from the basic load (rare now that it's
      // unfiltered, #44; it used to be IsValidForAdvancedFind-only) would otherwise be typed as strings and single-quoted in the
      // key URL (e.g. canvasapp `componentstate` / `overwritetime`). Fetch
      // their AttributeType; anything that still fails stays unknown.
      const loaded = new Set(columns.map((c) => c.logicalName));
      const missing = Array.from(
        new Set(alternateKeys.flatMap((k) => k.columns).filter((c) => !loaded.has(c))),
      );
      if (missing.length > 0) {
        try {
          const typed = await loader.loadAttributeTypes(logical, missing);
          const types = new Map<string, AttributeTypeCode>();
          for (const t of typed) {
            // Through attributeTypeOf like every other column, so Image /
            // File (`Virtual` + AttributeTypeName) resolve the same (#318).
            const code = attributeTypeOf(t);
            if (code) types.set(t.LogicalName, code);
          }
          for (const k of alternateKeys) {
            const ct: Record<string, AttributeTypeCode> = {};
            for (const c of k.columns) {
              const code = types.get(c);
              if (code) ct[c] = code;
            }
            if (Object.keys(ct).length > 0) k.columnTypes = ct;
          }
        } catch (e) {
          console.warn(`[metadata] alternate-key attribute types for ${logical} failed`, e);
        }
      }

      const built: TableMeta = {
        logicalName: entity.LogicalName,
        entitySetName: entity.EntitySetName,
        displayName: labelOf(entity.DisplayName, entity.LogicalName),
        primaryKey: entity.PrimaryIdAttribute,
        primaryName: entity.PrimaryNameAttribute,
        columns,
        navigationProperties: mapRelationships(
          rels.oneToMany,
          rels.manyToOne,
          rels.manyToMany,
          logical,
        ),
        alternateKeys,
        supportsMerge: ['account', 'contact', 'lead', 'incident'].includes(logical),
      };

      // Publish into the synchronous registry — this is what makes the
      // ~22 editors that call `findTable(logical)` synchronously see the
      // live data (after a re-render triggered by mode-level
      // `useLiveTable`). Not when the table was invalidated while this
      // build ran (its slot is gone): its responses may predate whatever
      // prompted that, so it mustn't put the old table back, and what it
      // cached on the way is dropped too (#335).
      if (buildInFlight.get(flightKey) === promise) {
        __registerLiveTable(built, org);
      } else if (org === activeOrg()) {
        metadataCache.clearEntity(logical);
      }
      return built;
    } catch (e) {
      console.error(`[metadataProvider] getTable('${logical}') failed:`, e);
      return undefined;
    }
  })();

  buildInFlight.set(flightKey, promise);
  // Only drop our own slot — a refresh may have swapped in a newer build.
  void promise.finally(() => {
    if (buildInFlight.get(flightKey) === promise) buildInFlight.delete(flightKey);
  });
  return promise;
}

export const metadata: MetadataProvider = {
  async listEntities(): Promise<EntityListItem[]> {
    const all = await loader.loadAllEntities();
    return all.map(mapEntityListItem);
  },

  async getTable(logical: string): Promise<TableMeta | undefined> {
    // Synchronous cache hit (already-registered) → done.
    const live = findTable(logical);
    if (live) return live;
    return buildTable(logical);
  },

  peekTable(logical: string): TableMeta | undefined {
    return findTable(logical);
  },

  invalidateTable(logical: string): void {
    metadataCache.clearEntity(logical);
    buildInFlight.delete(`${activeOrg()}|${logical}`);
    // `getTable` reads the registry first, so drop the table there too, or
    // the next read would return the old one (#335). Listeners refetch, as
    // after `invalidateAll`.
    __unregisterLiveTable(logical);
  },

  invalidateAll(): void {
    metadataCache.clear();
    buildInFlight = new Map();
    // Also clear the synchronous registry so editors don't show stale data.
    // Synchronous on purpose: a caller's next `getTable` must miss the
    // registry and refetch (#53). `useLiveTable` refetches its table when
    // it disappears from the registry.
    __clearLiveTables();
    notifyReset();
  },

  async refreshAll(): Promise<void> {
    // Drain any currently in-flight builds before touching the cache.
    // Without this, an orphaned buildTable promise can settle AFTER the
    // cache clear and re-register a pre-refresh TableMeta, silently
    // overwriting the fresh data we're about to fetch (race on
    // __registerLiveTable). allSettled so a failing fetch doesn't abort.
    await Promise.allSettled([...buildInFlight.values()]);

    // Snapshot what's currently registered, drop caches, then rebuild each
    // table. `buildTable` always hits the network (cache just cleared) and
    // re-registers via `__registerLiveTable`, firing the registry listeners
    // so every editor re-renders with fresh data. We deliberately do NOT
    // clear the registry first, so the UI shows current data until each
    // table's refresh lands. The set is small (current target + warmed
    // related entities), keeping us clear of the 100-concurrent cap.
    const logicals = getLiveTables().map((t) => t.logicalName);
    metadataCache.clear();
    buildInFlight = new Map();
    notifyReset();
    await Promise.all(logicals.map((l) => buildTable(l)));
  },
};
