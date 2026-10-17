// Three-tier metadata loaders — cache → in-flight → fetch.
//
// Every loader follows the same shape:
//
//   1. Synchronous cache lookup — if fresh, return immediately.
//   2. In-flight promise lookup — if a fetch is pending for the same key,
//      return its promise so concurrent reads collapse into one network call.
//   3. New fetch — fire the request, store the promise as in-flight, cache
//      the result on resolve, clear the in-flight slot.
//
// Errors bubble up to the caller — the cache only stores successful reads.
//
// Each loader pins the active org's cache before its first `await`, so a
// response that lands after a connection switch is stored under the org
// it came from, never the new one (#52).

import * as pptbClient from './pptbClient';
import { metadataCache, pinMetadataCache } from './cache';
import { activeOrg } from './activeOrg';
import type { EntityMetadata, AttributeMetadata, RelationshipMetadata } from './pptbClient';

/** All entities — always fetched with `advancedFindOnly=false`. Callers
 *  apply the AF filter at the display layer so the toggle is instant. */
export async function loadAllEntities(): Promise<EntityMetadata[]> {
  const cache = pinMetadataCache();
  const cached = cache.getAllEntities();
  if (cached) return cached;

  const inFlight = cache.getAllEntitiesPromise();
  if (inFlight) return inFlight;

  const promise = pptbClient.getAllEntities(false);
  cache.setAllEntitiesPromise(promise);

  try {
    const entities = await promise;
    cache.setAllEntities(entities);
    cache.clearAllEntitiesPromise();
    return entities;
  } catch (error) {
    console.error('loadAllEntities: API error:', error);
    cache.clearAllEntitiesPromise();
    throw error;
  }
}

/** Single-entity metadata. */
export async function loadEntityMetadata(logicalName: string): Promise<EntityMetadata> {
  const cache = pinMetadataCache();
  const cached = cache.getEntityMetadata(logicalName);
  if (cached) return cached;
  const metadata = await pptbClient.getEntityMetadata(logicalName);
  cache.setEntityMetadata(logicalName, metadata);
  return metadata;
}

/** Entity attributes — basic projection. */
export async function loadEntityAttributes(
  logicalName: string,
  advancedFindOnly: boolean = true,
): Promise<AttributeMetadata[]> {
  const cache = pinMetadataCache();
  const cacheKey = `${logicalName}_${advancedFindOnly}`;
  const cached = cache.getEntityAttributes(cacheKey);
  if (cached) return cached;

  const inFlight = cache.getAttributesPromise(cacheKey);
  if (inFlight) return inFlight;

  const promise = pptbClient.getEntityAttributes(logicalName, advancedFindOnly);
  cache.setAttributesPromise(cacheKey, promise);

  try {
    const attributes = await promise;
    cache.setEntityAttributes(cacheKey, attributes);
    cache.clearAttributesPromise(cacheKey);
    return attributes;
  } catch (error) {
    cache.clearAttributesPromise(cacheKey);
    throw error;
  }
}

/** Drill into a single attribute for OptionSet expansion. */
export async function loadAttributeWithOptionSet(
  entityLogicalName: string,
  attributeLogicalName: string,
): Promise<AttributeMetadata> {
  const cache = pinMetadataCache();
  const cacheKey = `${entityLogicalName}_${attributeLogicalName}_optionset`;
  const cached = cache.getEntityAttributes(cacheKey);
  if (cached && cached.length > 0) return cached[0];

  const attribute = await pptbClient.getAttributeWithOptionSet(
    entityLogicalName,
    attributeLogicalName,
  );
  cache.setEntityAttributes(cacheKey, [attribute]);
  return attribute;
}

/**
 * Bulk type-cast fetch — every attribute of a given subtype on an entity,
 * pulled in ONE request. Cached so repeat reads of the same (entity,
 * typeCast, selectProps) are free.
 *
 * Used by `metadataProvider.buildTable` to enrich the basic attribute
 * projection without firing one round-trip per attribute (which trips
 * Dataverse's 100-concurrent-request cap on wide entities).
 */
export async function loadAttributesByTypeCast(
  entityLogicalName: string,
  typeCast: string,
  selectProps: string,
  expandProps?: string,
): Promise<AttributeMetadata[]> {
  const cache = pinMetadataCache();
  const cacheKey = `${entityLogicalName}__cast__${typeCast}__${selectProps}__${expandProps ?? ''}`;
  const cached = cache.getEntityAttributes(cacheKey);
  if (cached) return cached;

  const inFlight = cache.getAttributesPromise(cacheKey);
  if (inFlight) return inFlight;

  const promise = pptbClient.getAttributesByTypeCast(
    entityLogicalName,
    typeCast,
    selectProps,
    expandProps,
  );
  cache.setAttributesPromise(cacheKey, promise);
  try {
    const list = await promise;
    cache.setEntityAttributes(cacheKey, list);
    cache.clearAttributesPromise(cacheKey);
    return list;
  } catch (error) {
    cache.clearAttributesPromise(cacheKey);
    throw error;
  }
}

/** Drill into a single attribute for type-specific properties. */
export async function loadAttributeDetailedMetadata(
  entityLogicalName: string,
  attributeLogicalName: string,
  attributeType: string,
): Promise<AttributeMetadata> {
  const cache = pinMetadataCache();
  const cacheKey = `${entityLogicalName}_${attributeLogicalName}_detailed`;
  const cached = cache.getEntityAttributes(cacheKey);
  if (cached && cached.length > 0) return cached[0];

  const attribute = await pptbClient.getAttributeDetailedMetadata(
    entityLogicalName,
    attributeLogicalName,
    attributeType,
  );
  cache.setEntityAttributes(cacheKey, [attribute]);
  return attribute;
}

/** Entity relationships — 1:N, N:1, N:N. */
export async function loadEntityRelationships(
  logicalName: string,
  advancedFindOnly: boolean = true,
): Promise<{
  oneToMany: RelationshipMetadata[];
  manyToOne: RelationshipMetadata[];
  manyToMany: RelationshipMetadata[];
}> {
  const cache = pinMetadataCache();
  const cacheKey = `${logicalName}_${advancedFindOnly}`;
  const cached = cache.getEntityRelationships(cacheKey);
  if (cached) return cached;

  const relationships = await pptbClient.getEntityRelationships(logicalName, advancedFindOnly);
  cache.setEntityRelationships(cacheKey, relationships);
  return relationships;
}

/** Entity alternate keys (EntityKeyMetadata). Cached per entity logical name.
 *  Powers the Upsert mode's alternate-key picker — without this fetch,
 *  the picker would always be empty for entities that DO have alt keys
 *  defined. */
export async function loadEntityKeys(logicalName: string) {
  const cache = pinMetadataCache();
  const cached = cache.getEntityKeys(logicalName);
  if (cached) return cached;
  const keys = await pptbClient.getEntityKeys(logicalName);
  cache.setEntityKeys(logicalName, keys);
  return keys;
}

/** Logical names of the tables in a set of solutions (solution-scoped entity
 *  pickers). Caches the raw components, never a list derived from the
 *  all-entities cache — that cache may not be loaded yet, which used to
 *  freeze an empty list for the solution set (#44). */
const solutionNamesInFlight = new Map<string, Promise<string[]>>();

export async function loadSolutionEntityNames(solutionIds: string[]): Promise<string[]> {
  const cache = pinMetadataCache();
  const toNames = (components: pptbClient.SolutionComponent[]) =>
    Array.from(new Set(components.map((c) => c.msdyn_name).filter(Boolean)));

  const cached = cache.getSolutionComponents(solutionIds);
  if (cached) return toNames(cached);

  const key = `${activeOrg()}|${solutionIds.slice().sort().join(',')}`;
  const inFlight = solutionNamesInFlight.get(key);
  if (inFlight) return inFlight;

  const promise = pptbClient
    .getSolutionComponents(solutionIds)
    .then((components) => {
      cache.setSolutionComponents(solutionIds, components);
      return toNames(components);
    })
    .finally(() => solutionNamesInFlight.delete(key));
  solutionNamesInFlight.set(key, promise);
  return promise;
}

/** Nuke everything — used by the "Refresh metadata" button. */
export function clearMetadataCache(): void {
  metadataCache.clear();
}

/** Invalidate one entity. */
export function clearEntityCache(logicalName: string): void {
  metadataCache.clearEntity(logicalName);
}
