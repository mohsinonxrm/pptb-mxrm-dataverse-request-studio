// useScopedEntities — entity list narrowed by the user's Settings drawer
// selections. Switches data source by `settings.entityScopeMode`:
//
//   - 'all':                returns the unfiltered list from listEntities()
//                           (also honors the persisted advancedFindOnly toggle
//                           by filtering client-side when set).
//   - 'solution-only':      fetches solution components for the selected
//                           solutions and filters the all-entities list to
//                           that set.
//   - 'publisher-solution': same as solution-only after the publishers
//                           have narrowed which solutions are available.
//
// We always start from the cached `loadAllEntities()` result — per-component
// lookups never trigger a fresh entity metadata fetch.

import { useEffect, useMemo, useState } from "react";
import { getSolutionComponents, filterCachedEntitiesByNames } from "./pptbClient";
import { metadata, type EntityListItem } from "./metadataProvider";
import { metadataCache } from "./cache";
import { usePersistedSettings } from "./usePersistedSettings";

export interface UseScopedEntitiesResult {
	entities: EntityListItem[];
	loading: boolean;
	error: string | null;
	/**
	 * True when the entity list is empty because the user has not yet
	 * configured the Query Scope in Settings (scoped mode with no solutions
	 * selected). Used by TargetEditor to show a first-run guidance banner.
	 * Always false when entityScopeMode === 'all'.
	 */
	needsSetup: boolean;
	/** The active entityScopeMode — exposed so TargetEditor can tailor the hint. */
	scopeMode: import("../state/displaySettings").EntityScopeMode;
}

export function useScopedEntities(): UseScopedEntitiesResult {
	const [settings] = usePersistedSettings();
	const [all, setAll] = useState<EntityListItem[]>([]);
	const [allLoading, setAllLoading] = useState(true);
	const [allError, setAllError] = useState<string | null>(null);

	// ── Layer 1: unfiltered entity list ──
	useEffect(() => {
		let cancelled = false;
		metadata
			.listEntities()
			.then((list) => {
				if (!cancelled) setAll(list);
			})
			.catch((e) => {
				if (!cancelled) setAllError(e instanceof Error ? e.message : String(e));
			})
			.finally(() => {
				if (!cancelled) setAllLoading(false);
			});
		return () => {
			cancelled = true;
		};
	}, []);

	// ── Layer 2: solution-component cascade ──
	const solutionKey = settings.selectedSolutionIds.join("|");
	const [scopedNames, setScopedNames] = useState<string[] | null>(null);
	const [scopedLoading, setScopedLoading] = useState(false);
	const [scopedError, setScopedError] = useState<string | null>(null);

	useEffect(() => {
		if (settings.entityScopeMode === "all") {
			setScopedNames(null);
			setScopedError(null);
			return;
		}
		if (settings.selectedSolutionIds.length === 0) {
			setScopedNames([]);
			setScopedError(null);
			return;
		}

		let cancelled = false;
		setScopedLoading(true);
		setScopedError(null);

		// Cache hit → instant.
		const cached = metadataCache.getFilteredEntities(settings.selectedSolutionIds);
		if (cached) {
			setScopedNames(cached.map((e) => e.LogicalName));
			setScopedLoading(false);
			return;
		}

		(async () => {
			try {
				const components = await getSolutionComponents(settings.selectedSolutionIds);
				if (cancelled) return;
				const names = Array.from(new Set(components.map((c) => c.msdyn_name).filter(Boolean)));
				// Cache the filtered entity-meta list too so re-renders short-circuit.
				const data = filterCachedEntitiesByNames(names);
				metadataCache.setFilteredEntities(settings.selectedSolutionIds, data);
				setScopedNames(names);
			} catch (e) {
				if (!cancelled) setScopedError(e instanceof Error ? e.message : String(e));
			} finally {
				if (!cancelled) setScopedLoading(false);
			}
		})();
		return () => {
			cancelled = true;
		};
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [settings.entityScopeMode, solutionKey]);

	// ── Compose visible list ──
	const visible = useMemo(() => {
		if (settings.entityScopeMode === "all") return all;
		if (scopedNames === null) return [];
		const set = new Set(scopedNames);
		return all.filter((e) => set.has(e.logicalName));
	}, [all, scopedNames, settings.entityScopeMode]);

	// needsSetup: scoped mode but no solutions chosen yet (entity list will be
	// empty until the user configures Settings → Query Scope).
	const needsSetup =
		settings.entityScopeMode !== "all" && settings.selectedSolutionIds.length === 0;

	return {
		entities: visible,
		loading: allLoading || scopedLoading,
		error: allError ?? scopedError,
		needsSetup,
		scopeMode: settings.entityScopeMode,
	};
}
