// Metadata cache — singleton, TTL-based, with in-flight promise dedup.
//
// Every metadata fetch (entities, attributes, relationships, publishers,
// solutions, components) flows through this cache. Consumers go:
//
//   1. Synchronous cache lookup        →  hits return immediately
//   2. In-flight promise lookup        →  dedupes concurrent reads of the
//                                          same key during a single fetch
//   3. Fresh fetch                     →  cache + clear in-flight on resolve
//
// Stale-after-TTL means re-fetches kick in 1 hour after the last write.
// Clear methods invalidate proactively (used by the "Refresh metadata"
// button in Settings).
//
// Anything stored here is RAW Dataverse-API shape — `EntityMetadata`,
// `AttributeMetadata`, `RelationshipMetadata`, `Publisher`, `Solution`.
// Studio-shape transformations (`TableMeta`, `ColumnMeta`) live one layer
// up in `metadataProvider.ts`.

import type {
  EntityMetadata, AttributeMetadata, RelationshipMetadata, EntityKeyMetadata,
  Publisher, Solution, PublisherWithSolutions, SolutionComponent,
} from './pptbClient';

interface CacheEntry<T> {
  data?: T;
  promise?: Promise<T>;
  timestamp: number;
}

// 1 hour. Metadata (entities/attributes/relationships) changes rarely in a
// released environment, so a long TTL keeps the tool snappy across a working
// session. Users can force a fresh pull anytime via the Settings → "Refresh
// metadata" button (metadata.refreshAll), which clears every cache entry.
const TTL_MS = 60 * 60 * 1000; // 1 hour

class MetadataCache {
  // ── Storage ─────────────────────────────────────────────────────────
  // Per-entity bag of related metadata (attributes/relationships/optionsets
  // are keyed inside the entities map and the sub-maps).
  private entities: Map<string, CacheEntry<EntityMetadata>> = new Map();
  private allEntitiesCache: CacheEntry<EntityMetadata[]> | undefined;
  // Entity attributes — key format: "${logicalName}_${advancedFindOnly}".
  // Type-cast attribute drills (OptionSet, detailed numeric) reuse this map
  // with keys like "${logical}_${attr}_optionset" / "_detailed".
  private entityAttributes: Map<string, CacheEntry<AttributeMetadata[]>> = new Map();
  // Entity relationships — key: "${logicalName}_${advancedFindOnly}".
  private entityRelationships: Map<string, CacheEntry<{
    oneToMany: RelationshipMetadata[];
    manyToOne: RelationshipMetadata[];
    manyToMany: RelationshipMetadata[];
  }>> = new Map();
  // Entity alternate keys — key: logicalName (advancedFindOnly doesn't
  // apply; Keys is its own child collection). Loaded eagerly when the
  // table loads so Upsert mode can populate its alt-key picker without
  // a separate fetch round-trip per mode-switch.
  private entityKeys: Map<string, CacheEntry<EntityKeyMetadata[]>> = new Map();
  // Publishers + solutions caches.
  private publishersCache: CacheEntry<Publisher[]> | undefined;
  private publishersWithSolutionsCache: CacheEntry<PublisherWithSolutions[]> | undefined;
  // Solutions by publisher-id-set — key: solutionIds.slice().sort().join(',').
  private solutionsCache: Map<string, CacheEntry<Solution[]>> = new Map();
  // Entities filtered by solution set — key: same as above.
  private filteredEntitiesCache: Map<string, CacheEntry<EntityMetadata[]>> = new Map();
  // Solution components (entities only) per solution-set key.
  private solutionComponentsCache: Map<string, CacheEntry<SolutionComponent[]>> = new Map();

  private fresh<T>(entry: CacheEntry<T> | undefined): T | undefined {
    if (!entry || entry.data === undefined) return undefined;
    if (Date.now() - entry.timestamp > TTL_MS) return undefined;
    return entry.data;
  }

  // ── allEntities ─────────────────────────────────────────────────────
  getAllEntities(): EntityMetadata[] | undefined {
    return this.fresh(this.allEntitiesCache);
  }
  setAllEntities(data: EntityMetadata[]): void {
    this.allEntitiesCache = { data, timestamp: Date.now() };
  }
  getAllEntitiesPromise(): Promise<EntityMetadata[]> | undefined {
    return this.allEntitiesCache?.promise;
  }
  setAllEntitiesPromise(promise: Promise<EntityMetadata[]>): void {
    this.allEntitiesCache = { ...this.allEntitiesCache, promise, timestamp: Date.now() };
  }
  clearAllEntitiesPromise(): void {
    if (this.allEntitiesCache) delete this.allEntitiesCache.promise;
  }

  // ── Per-entity metadata ─────────────────────────────────────────────
  getEntityMetadata(logical: string): EntityMetadata | undefined {
    return this.fresh(this.entities.get(logical));
  }
  setEntityMetadata(logical: string, data: EntityMetadata): void {
    this.entities.set(logical, { data, timestamp: Date.now() });
  }

  // ── Per-entity attributes (keyed by `${logical}_${af}` or detail keys) ─
  getEntityAttributes(cacheKey: string): AttributeMetadata[] | undefined {
    return this.fresh(this.entityAttributes.get(cacheKey));
  }
  setEntityAttributes(cacheKey: string, data: AttributeMetadata[]): void {
    this.entityAttributes.set(cacheKey, { data, timestamp: Date.now() });
  }
  getAttributesPromise(cacheKey: string): Promise<AttributeMetadata[]> | undefined {
    return this.entityAttributes.get(cacheKey)?.promise;
  }
  setAttributesPromise(cacheKey: string, promise: Promise<AttributeMetadata[]>): void {
    const existing = this.entityAttributes.get(cacheKey);
    this.entityAttributes.set(cacheKey, { ...existing, promise, timestamp: Date.now() });
  }
  clearAttributesPromise(cacheKey: string): void {
    const e = this.entityAttributes.get(cacheKey);
    if (e) delete e.promise;
  }

  // ── Per-entity relationships ────────────────────────────────────────
  getEntityRelationships(cacheKey: string) {
    return this.fresh(this.entityRelationships.get(cacheKey));
  }
  setEntityRelationships(
    cacheKey: string,
    data: {
      oneToMany: RelationshipMetadata[];
      manyToOne: RelationshipMetadata[];
      manyToMany: RelationshipMetadata[];
    },
  ): void {
    this.entityRelationships.set(cacheKey, { data, timestamp: Date.now() });
  }

  // ── Per-entity alternate keys ──────────────────────────────────────
  getEntityKeys(logicalName: string): EntityKeyMetadata[] | undefined {
    return this.fresh(this.entityKeys.get(logicalName));
  }
  setEntityKeys(logicalName: string, data: EntityKeyMetadata[]): void {
    this.entityKeys.set(logicalName, { data, timestamp: Date.now() });
  }

  // ── Publishers ──────────────────────────────────────────────────────
  getPublishers(): Publisher[] | undefined {
    return this.fresh(this.publishersCache);
  }
  setPublishers(data: Publisher[]): void {
    this.publishersCache = { data, timestamp: Date.now() };
  }
  getPublishersWithSolutions(): PublisherWithSolutions[] | undefined {
    return this.fresh(this.publishersWithSolutionsCache);
  }
  setPublishersWithSolutions(data: PublisherWithSolutions[]): void {
    this.publishersWithSolutionsCache = { data, timestamp: Date.now() };
  }

  // ── Solutions (by publisher id set; empty array = "all solutions") ──
  private solutionKey(publisherIds: string[]): string {
    return publisherIds.slice().sort().join(',');
  }
  getSolutions(publisherIds: string[]): Solution[] | undefined {
    return this.fresh(this.solutionsCache.get(this.solutionKey(publisherIds)));
  }
  setSolutions(publisherIds: string[], data: Solution[]): void {
    this.solutionsCache.set(this.solutionKey(publisherIds), { data, timestamp: Date.now() });
  }

  // ── Filtered entities (by solution set) ─────────────────────────────
  private filteredEntitiesKey(solutionIds: string[]): string {
    return solutionIds.slice().sort().join(',');
  }
  getFilteredEntities(solutionIds: string[]): EntityMetadata[] | undefined {
    return this.fresh(this.filteredEntitiesCache.get(this.filteredEntitiesKey(solutionIds)));
  }
  setFilteredEntities(solutionIds: string[], data: EntityMetadata[]): void {
    this.filteredEntitiesCache.set(
      this.filteredEntitiesKey(solutionIds),
      { data, timestamp: Date.now() },
    );
  }

  // ── Solution components ─────────────────────────────────────────────
  private solutionComponentsKey(solutionIds: string[]): string {
    return solutionIds.slice().sort().join(',');
  }
  getSolutionComponents(solutionIds: string[]): SolutionComponent[] | undefined {
    return this.fresh(this.solutionComponentsCache.get(this.solutionComponentsKey(solutionIds)));
  }
  setSolutionComponents(solutionIds: string[], data: SolutionComponent[]): void {
    this.solutionComponentsCache.set(
      this.solutionComponentsKey(solutionIds),
      { data, timestamp: Date.now() },
    );
  }

  // ── filterCachedEntitiesByNames — used by publisher/solution filter hooks
  filterCachedEntitiesByNames(logicalNames: string[]): EntityMetadata[] {
    const all = this.fresh(this.allEntitiesCache) ?? [];
    if (logicalNames.length === 0) return all;
    const set = new Set(logicalNames);
    return all.filter(e => set.has(e.LogicalName));
  }

  // ── Cache management ────────────────────────────────────────────────
  /** Nuke everything — clears every map + every entry. Wired to "Refresh metadata". */
  clear(): void {
    this.entities.clear();
    this.allEntitiesCache = undefined;
    this.entityAttributes.clear();
    this.entityRelationships.clear();
    this.entityKeys.clear();
    this.publishersCache = undefined;
    this.publishersWithSolutionsCache = undefined;
    this.solutionsCache.clear();
    this.filteredEntitiesCache.clear();
    this.solutionComponentsCache.clear();
  }

  /** Selectively invalidate one entity. */
  clearEntity(logical: string): void {
    this.entities.delete(logical);
    // Walk the attribute / relationship caches for any key that starts with
    // the entity's logical name — covers `${logical}_true`, `${logical}_false`,
    // and the per-attribute detail keys.
    for (const k of this.entityAttributes.keys()) {
      if (k.startsWith(`${logical}_`)) this.entityAttributes.delete(k);
    }
    for (const k of this.entityRelationships.keys()) {
      if (k.startsWith(`${logical}_`)) this.entityRelationships.delete(k);
    }
    this.entityKeys.delete(logical);
  }
}

export const metadataCache = new MetadataCache();
