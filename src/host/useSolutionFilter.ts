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
import { pinMetadataCache } from './cache';
import { useActiveOrg } from './activeOrg';
import { subscribeMetadataReset } from './metadataProvider';

const NONE: Solution[] = [];

/** `enabled: false` skips the solution fetch (the scope chip in another mode). */
export function useSolutionFilter(enabled = true) {
  // Tagged with the org it came from, so another org's list never shows.
  const [loadedFor, setLoadedFor] = useState<{ org: string; data: Solution[] }>({
    org: '',
    data: NONE,
  });
  const [solutionsLoading, setSolutionsLoading] = useState(false);
  const [solutionsError, setSolutionsError] = useState<string | null>(null);

  const [selectedSolutionIds, setSelectedSolutionIds] = useState<string[]>([]);

  const [entities, setEntities] = useState<EntityMetadata[]>([]);
  const [entitiesLoading, setEntitiesLoading] = useState(false);
  const [entitiesError, setEntitiesError] = useState<string | null>(null);

  // Reload when the PPTB connection switches to another org (#52).
  const org = useActiveOrg();
  const solutions = loadedFor.org === org ? loadedFor.data : NONE;
  // …and after "Refresh metadata" clears the cache.
  const [resets, setResets] = useState(0);
  useEffect(() => subscribeMetadataReset(() => setResets((n) => n + 1)), []);

  // ── Load all solutions on mount ──
  useEffect(() => {
    if (!enabled) return;
    let mounted = true;

    async function loadSolutions() {
      // Pinned: a response that lands after a connection switch is cached
      // under the org it came from (#52).
      const metadataCache = pinMetadataCache();
      // Cache key: empty array = "all solutions"
      const cached = metadataCache.getSolutions([]);
      if (cached) {
        setLoadedFor({ org, data: cached });
        setSolutionsLoading(false);
        setSolutionsError(null);
        return;
      }

      try {
        setSolutionsLoading(true);
        setSolutionsError(null);
        const data = await getAllSolutionsWithEntities();
        if (mounted) {
          setLoadedFor({ org, data });
          metadataCache.setSolutions([], data);
        }
      } catch (err) {
        if (mounted) setSolutionsError(err instanceof Error ? err.message : String(err));
      } finally {
        if (mounted) setSolutionsLoading(false);
      }
    }

    loadSolutions();
    return () => {
      mounted = false;
      // A fetch cut short (org switch, disabled) isn't loading any more.
      setSolutionsLoading(false);
    };
  }, [org, enabled, resets]);

  // ── Cascade solution → entities ──
  useEffect(() => {
    if (!selectedSolutionIds.length) {
      setEntities([]);
      setEntitiesError(null);
      return;
    }

    let mounted = true;

    async function loadEntities() {
      // Pinned: a response that lands after a connection switch is cached
      // under the org it came from (#52).
      const metadataCache = pinMetadataCache();
      const cached = metadataCache.getFilteredEntities(selectedSolutionIds);
      if (cached) {
        setEntities(cached);
        return;
      }

      try {
        setEntitiesLoading(true);
        setEntitiesError(null);

        const components = await getSolutionComponents(selectedSolutionIds);
        const logicalNames = Array.from(
          new Set(components.map((c) => c.msdyn_name).filter(Boolean)),
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
    return () => {
      mounted = false;
    };
  }, [selectedSolutionIds, org]);

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
