// Metadata provider — transforms raw Dataverse metadata into the studio's
// `TableMeta` / `ColumnMeta` shape, then publishes it into the synchronous
// live registry in `mock/metadata.ts` so every editor's `findTable(logical)`
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
import { __registerLiveTable, findTable, getLiveTables } from '../mock/metadata';
import type {
  TableMeta,
  ColumnMeta,
  NavProperty,
  AlternateKeyDef,
  AttributeTypeCode,
  SourceType,
} from '../mock/metadata';
import type { EntityMetadata, AttributeMetadata, RelationshipMetadata } from './pptbClient';
import { metadataCache } from './cache';

// ── Public surface ──────────────────────────────────────────────────────

export interface EntityListItem {
  logicalName: string;
  displayName: string;
  entitySetName: string;
  metadataId: string;
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

function mapAttribute(raw: AttributeMetadata): ColumnMeta | null {
  const attrType = ATTRIBUTE_TYPE_MAP[raw.AttributeType];
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
  };
}

// ── Provider impl ──────────────────────────────────────────────────────

let buildInFlight = new Map<string, Promise<TableMeta | undefined>>();

async function buildTable(logical: string): Promise<TableMeta | undefined> {
  // Dedupe concurrent builds for the same entity.
  const existing = buildInFlight.get(logical);
  if (existing) return existing;

  const promise = (async () => {
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
        supportsMerge: ['account', 'contact', 'incident'].includes(logical),
      };

      // Publish into the synchronous registry — this is what makes the
      // ~22 editors that call `findTable(logical)` synchronously see the
      // live data (after a re-render triggered by mode-level
      // `useLiveTable`).
      __registerLiveTable(built);
      return built;
    } catch (e) {
      console.error(`[metadataProvider] getTable('${logical}') failed:`, e);
      return undefined;
    } finally {
      buildInFlight.delete(logical);
    }
  })();

  buildInFlight.set(logical, promise);
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
    buildInFlight.delete(logical);
  },

  invalidateAll(): void {
    metadataCache.clear();
    buildInFlight = new Map();
    // Also clear the synchronous registry so editors don't show stale data.
    // The next mode-level `useLiveTable` fetch will repopulate.
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    import('../mock/metadata').then(({ __clearLiveTables }) => __clearLiveTables());
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
    await Promise.all(logicals.map((l) => buildTable(l)));
  },
};
