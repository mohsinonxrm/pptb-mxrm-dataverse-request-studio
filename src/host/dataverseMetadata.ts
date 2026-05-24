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

import * as pptbClient from './pptbClient';
import { metadataCache } from './cache';
import type {
  EntityMetadata, AttributeMetadata, RelationshipMetadata,
} from './pptbClient';

/** All entities — always fetched with `advancedFindOnly=false`. Callers
 *  apply the AF filter at the display layer so the toggle is instant. */
export async function loadAllEntities(): Promise<EntityMetadata[]> {
  const cached = metadataCache.getAllEntities();
  if (cached) return cached;

  const inFlight = metadataCache.getAllEntitiesPromise();
  if (inFlight) return inFlight;

  const promise = pptbClient.getAllEntities(false);
  metadataCache.setAllEntitiesPromise(promise);

  try {
    const entities = await promise;
    metadataCache.setAllEntities(entities);
    metadataCache.clearAllEntitiesPromise();
    return entities;
  } catch (error) {
    console.error('loadAllEntities: API error:', error);
    metadataCache.clearAllEntitiesPromise();
    throw error;
  }
}

/** Single-entity metadata. */
export async function loadEntityMetadata(logicalName: string): Promise<EntityMetadata> {
  const cached = metadataCache.getEntityMetadata(logicalName);
  if (cached) return cached;
  const metadata = await pptbClient.getEntityMetadata(logicalName);
  metadataCache.setEntityMetadata(logicalName, metadata);
  return metadata;
}

/** Entity attributes — basic projection. */
export async function loadEntityAttributes(
  logicalName: string,
  advancedFindOnly: boolean = true,
): Promise<AttributeMetadata[]> {
  const cacheKey = `${logicalName}_${advancedFindOnly}`;
  const cached = metadataCache.getEntityAttributes(cacheKey);
  if (cached) return cached;

  const inFlight = metadataCache.getAttributesPromise(cacheKey);
  if (inFlight) return inFlight;

  const promise = pptbClient.getEntityAttributes(logicalName, advancedFindOnly);
  metadataCache.setAttributesPromise(cacheKey, promise);

  try {
    const attributes = await promise;
    metadataCache.setEntityAttributes(cacheKey, attributes);
    metadataCache.clearAttributesPromise(cacheKey);
    return attributes;
  } catch (error) {
    metadataCache.clearAttributesPromise(cacheKey);
    throw error;
  }
}

/** Drill into a single attribute for OptionSet expansion. */
export async function loadAttributeWithOptionSet(
  entityLogicalName: string,
  attributeLogicalName: string,
): Promise<AttributeMetadata> {
  const cacheKey = `${entityLogicalName}_${attributeLogicalName}_optionset`;
  const cached = metadataCache.getEntityAttributes(cacheKey);
  if (cached && cached.length > 0) return cached[0];

  const attribute = await pptbClient.getAttributeWithOptionSet(
    entityLogicalName,
    attributeLogicalName,
  );
  metadataCache.setEntityAttributes(cacheKey, [attribute]);
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
  const cacheKey = `${entityLogicalName}__cast__${typeCast}__${selectProps}__${expandProps ?? ''}`;
  const cached = metadataCache.getEntityAttributes(cacheKey);
  if (cached) return cached;

  const inFlight = metadataCache.getAttributesPromise(cacheKey);
  if (inFlight) return inFlight;

  const promise = pptbClient.getAttributesByTypeCast(
    entityLogicalName, typeCast, selectProps, expandProps,
  );
  metadataCache.setAttributesPromise(cacheKey, promise);
  try {
    const list = await promise;
    metadataCache.setEntityAttributes(cacheKey, list);
    metadataCache.clearAttributesPromise(cacheKey);
    return list;
  } catch (error) {
    metadataCache.clearAttributesPromise(cacheKey);
    throw error;
  }
}

/** Drill into a single attribute for type-specific properties. */
export async function loadAttributeDetailedMetadata(
  entityLogicalName: string,
  attributeLogicalName: string,
  attributeType: string,
): Promise<AttributeMetadata> {
  const cacheKey = `${entityLogicalName}_${attributeLogicalName}_detailed`;
  const cached = metadataCache.getEntityAttributes(cacheKey);
  if (cached && cached.length > 0) return cached[0];

  const attribute = await pptbClient.getAttributeDetailedMetadata(
    entityLogicalName,
    attributeLogicalName,
    attributeType,
  );
  metadataCache.setEntityAttributes(cacheKey, [attribute]);
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
  const cacheKey = `${logicalName}_${advancedFindOnly}`;
  const cached = metadataCache.getEntityRelationships(cacheKey);
  if (cached) return cached;

  const relationships = await pptbClient.getEntityRelationships(logicalName, advancedFindOnly);
  metadataCache.setEntityRelationships(cacheKey, relationships);
  return relationships;
}

/** Entity alternate keys (EntityKeyMetadata). Cached per entity logical name.
 *  Powers the Upsert mode's alternate-key picker — without this fetch,
 *  the picker would always be empty for entities that DO have alt keys
 *  defined. */
export async function loadEntityKeys(logicalName: string) {
  const cached = metadataCache.getEntityKeys(logicalName);
  if (cached) return cached;
  const keys = await pptbClient.getEntityKeys(logicalName);
  metadataCache.setEntityKeys(logicalName, keys);
  return keys;
}

/** Nuke everything — used by the "Refresh metadata" button. */
export function clearMetadataCache(): void {
  metadataCache.clear();
}

/** Invalidate one entity. */
export function clearEntityCache(logicalName: string): void {
  metadataCache.clearEntity(logicalName);
}
