// usePublisherFilter — full-filter cascade (Publisher → Solution → Entities).
//
// Three layers of state:
//
//   1. Publishers + their solutions are loaded once on mount via
//      `getPublishersWithSolutions()` (a single $expand'd request — way
//      cheaper than separate publisher + solution fetches).
//
//   2. When the user selects publishers, solutions are FILTERED LOCALLY
//      from the combined response (no extra fetch). Selected solutions
//      that no longer map to a selected publisher get dropped.
//
//   3. When the user selects solutions, entities are resolved by:
//      - fetching solution components (componenttype eq 1 = Entity)
//      - extracting unique entity logical names
//      - filtering them out of the cached `loadAllEntities()` result
//        (no per-entity metadata fetch — instant once all-entities cache
//        is warm)
//
// Every layer has loading + error state for clean UI binding.

import { useState, useEffect, useCallback } from 'react';
import {
  getPublishersWithSolutions,
  getSolutionComponents,
  filterCachedEntitiesByNames,
  type Solution,
  type EntityMetadata,
  type PublisherWithSolutions,
} from './pptbClient';
import { metadataCache } from './cache';

export function usePublisherFilter() {
  // ── Layer 1: publishers + solutions (combined) ──
  const [publishersWithSolutions, setPublishersWithSolutions] = useState<PublisherWithSolutions[]>([]);
  const [publishersLoading, setPublishersLoading] = useState(false);
  const [publishersError, setPublishersError] = useState<string | null>(null);

  // ── User selections ──
  const [selectedPublisherIds, setSelectedPublisherIds] = useState<string[]>([]);
  const [selectedSolutionIds, setSelectedSolutionIds] = useState<string[]>([]);

  // ── Layer 2: solutions filtered by selected publishers ──
  const [solutions, setSolutions] = useState<Solution[]>([]);

  // ── Layer 3: entities resolved from selected solutions ──
  const [entities, setEntities] = useState<EntityMetadata[]>([]);
  const [entitiesLoading, setEntitiesLoading] = useState(false);
  const [entitiesError, setEntitiesError] = useState<string | null>(null);

  // ── Effect 1: load publishers + solutions on mount ──
  useEffect(() => {
    let mounted = true;

    async function loadPublishersWithSolutions() {
      const cached = metadataCache.getPublishersWithSolutions();
      if (cached) { setPublishersWithSolutions(cached); return; }

      try {
        setPublishersLoading(true);
        setPublishersError(null);
        const data = await getPublishersWithSolutions();
        if (mounted) {
          setPublishersWithSolutions(data);
          metadataCache.setPublishersWithSolutions(data);
        }
      } catch (err) {
        if (mounted) setPublishersError(err instanceof Error ? err.message : String(err));
      } finally {
        if (mounted) setPublishersLoading(false);
      }
    }

    loadPublishersWithSolutions();
    return () => { mounted = false; };
  }, []);

  // ── Effect 2: cascade publisher → solution selection ──
  useEffect(() => {
    if (!selectedPublisherIds.length) {
      setSolutions([]);
      setSelectedSolutionIds([]);
      return;
    }

    const selectedPublisherSet = new Set(selectedPublisherIds);
    const filteredSolutions = publishersWithSolutions
      .filter(pws => selectedPublisherSet.has(pws.publisher.publisherid))
      .flatMap(pws => pws.solutions);

    setSolutions(filteredSolutions);

    // Drop any selected solution that no longer maps to a selected publisher.
    setSelectedSolutionIds(current => {
      if (current.length === 0) return current;
      const availableIds = new Set(filteredSolutions.map(s => s.solutionid));
      const validIds = current.filter(id => availableIds.has(id));
      return validIds.length === current.length ? current : validIds;
    });
  }, [selectedPublisherIds, publishersWithSolutions]);

  // ── Effect 3: cascade solution → entities ──
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

  const updateSelectedPublishers = useCallback((ids: string[]) => {
    setSelectedPublisherIds(ids);
  }, []);
  const updateSelectedSolutions = useCallback((ids: string[]) => {
    setSelectedSolutionIds(ids);
  }, []);

  return {
    // Publishers
    publishers: publishersWithSolutions.map(pws => pws.publisher),
    publishersWithSolutions,
    publishersLoading,
    publishersError,
    selectedPublisherIds,
    updateSelectedPublishers,
    // Solutions
    solutions,
    selectedSolutionIds,
    updateSelectedSolutions,
    // Entities
    entities,
    entitiesLoading,
    entitiesError,
  };
}
