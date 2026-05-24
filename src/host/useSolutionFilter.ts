// useSolutionFilter — Solutions-Only cascade (Solution → Entities).
//
// Two layers:
//
//   1. All solutions are loaded once on mount via `getAllSolutionsWithEntities()`.
//      No publisher filter — every visible solution containing entities.
//
//   2. When the user selects solutions, entities are resolved by:
//      - fetching solution components for each
//      - extracting unique entity logical names
//      - filtering them out of the cached `loadAllEntities()` result

import { useState, useEffect, useCallback } from 'react';
import {
  getAllSolutionsWithEntities,
  getSolutionComponents,
  filterCachedEntitiesByNames,
  type Solution,
  type EntityMetadata,
} from './pptbClient';
import { metadataCache } from './cache';

export function useSolutionFilter() {
  const [solutions, setSolutions] = useState<Solution[]>([]);
  const [solutionsLoading, setSolutionsLoading] = useState(false);
  const [solutionsError, setSolutionsError] = useState<string | null>(null);

  const [selectedSolutionIds, setSelectedSolutionIds] = useState<string[]>([]);

  const [entities, setEntities] = useState<EntityMetadata[]>([]);
  const [entitiesLoading, setEntitiesLoading] = useState(false);
  const [entitiesError, setEntitiesError] = useState<string | null>(null);

  // ── Load all solutions on mount ──
  useEffect(() => {
    let mounted = true;

    async function loadSolutions() {
      // Cache key: empty array = "all solutions"
      const cached = metadataCache.getSolutions([]);
      if (cached) { setSolutions(cached); return; }

      try {
        setSolutionsLoading(true);
        setSolutionsError(null);
        const data = await getAllSolutionsWithEntities();
        if (mounted) {
          setSolutions(data);
          metadataCache.setSolutions([], data);
        }
      } catch (err) {
        if (mounted) setSolutionsError(err instanceof Error ? err.message : String(err));
      } finally {
        if (mounted) setSolutionsLoading(false);
      }
    }

    loadSolutions();
    return () => { mounted = false; };
  }, []);

  // ── Cascade solution → entities ──
  useEffect(() => {
    if (!selectedSolutionIds.length) {
      setEntities([]);
      setEntitiesError(null);
      return;
    }

    let mounted = true;

    async function loadEntities() {
      const cached = metadataCache.getFilteredEntities(selectedSolutionIds);
      if (cached) { setEntities(cached); return; }

      try {
        setEntitiesLoading(true);
        setEntitiesError(null);

        const components = await getSolutionComponents(selectedSolutionIds);
        const logicalNames = Array.from(
          new Set(components.map(c => c.msdyn_name).filter(Boolean))
        );
        const data = filterCachedEntitiesByNames(logicalNames);

        if (mounted) {
          setEntities(data);
          metadataCache.setFilteredEntities(selectedSolutionIds, data);
        }
      } catch (err) {
        if (mounted) setEntitiesError(err instanceof Error ? err.message : String(err));
      } finally {
        if (mounted) setEntitiesLoading(false);
      }
    }

    loadEntities();
    return () => { mounted = false; };
  }, [selectedSolutionIds]);

  const updateSelectedSolutions = useCallback((ids: string[]) => {
    setSelectedSolutionIds(ids);
  }, []);

  return {
    solutions,
    solutionsLoading,
    solutionsError,
    selectedSolutionIds,
    updateSelectedSolutions,
    entities,
    entitiesLoading,
    entitiesError,
  };
}
