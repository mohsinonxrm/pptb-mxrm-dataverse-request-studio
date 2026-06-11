// useLazyMetadata — component-facing wrapper over `dataverseMetadata` loaders.
//
// Exposes:
//   - Loader callbacks (loadEntities, loadEntityMetadata, loadAttributes, loadRelationships)
//   - Loading state + last error
//   - Cache utilities (clearCache, clearEntityCache)

import { useState, useCallback } from 'react';
import type { EntityMetadata, AttributeMetadata, RelationshipMetadata } from './pptbClient';
import * as metadataLoader from './dataverseMetadata';

interface UseLazyMetadataResult {
  loadEntities: () => Promise<EntityMetadata[]>;
  loadEntityMetadata: (logicalName: string) => Promise<EntityMetadata>;
  loadAttributes: (logicalName: string) => Promise<AttributeMetadata[]>;
  loadRelationships: (logicalName: string) => Promise<{
    oneToMany: RelationshipMetadata[];
    manyToOne: RelationshipMetadata[];
    manyToMany: RelationshipMetadata[];
  }>;
  isLoading: boolean;
  error: Error | null;
  clearCache: () => void;
  clearEntityCache: (logicalName: string) => void;
}

export function useLazyMetadata(): UseLazyMetadataResult {
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<Error | null>(null);

  const loadEntities = useCallback(async () => {
    setIsLoading(true);
    setError(null);
    try {
      return await metadataLoader.loadAllEntities();
    } catch (err) {
      const e = err instanceof Error ? err : new Error('Failed to load entities');
      setError(e);
      throw e;
    } finally {
      setIsLoading(false);
    }
  }, []);

  const loadEntityMetadata = useCallback(async (logicalName: string) => {
    setIsLoading(true);
    setError(null);
    try {
      return await metadataLoader.loadEntityMetadata(logicalName);
    } catch (err) {
      const e =
        err instanceof Error ? err : new Error(`Failed to load metadata for ${logicalName}`);
      setError(e);
      throw e;
    } finally {
      setIsLoading(false);
    }
  }, []);

  const loadAttributes = useCallback(async (logicalName: string) => {
    setIsLoading(true);
    setError(null);
    try {
      return await metadataLoader.loadEntityAttributes(logicalName);
    } catch (err) {
      const e =
        err instanceof Error ? err : new Error(`Failed to load attributes for ${logicalName}`);
      setError(e);
      throw e;
    } finally {
      setIsLoading(false);
    }
  }, []);

  const loadRelationships = useCallback(async (logicalName: string) => {
    setIsLoading(true);
    setError(null);
    try {
      return await metadataLoader.loadEntityRelationships(logicalName);
    } catch (err) {
      const e =
        err instanceof Error ? err : new Error(`Failed to load relationships for ${logicalName}`);
      setError(e);
      throw e;
    } finally {
      setIsLoading(false);
    }
  }, []);

  const clearCache = useCallback(() => {
    metadataLoader.clearMetadataCache();
  }, []);
  const clearEntityCache = useCallback((logicalName: string) => {
    metadataLoader.clearEntityCache(logicalName);
  }, []);

  return {
    loadEntities,
    loadEntityMetadata,
    loadAttributes,
    loadRelationships,
    isLoading,
    error,
    clearCache,
    clearEntityCache,
  };
}
