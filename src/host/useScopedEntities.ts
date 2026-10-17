// useScopedEntities — entity list narrowed by the user's Settings drawer
// selections. Switches data source by `settings.entityScopeMode`:
//
//   - 'all':                returns the unfiltered list from listEntities().
//   - 'solution-only':      fetches solution components for the selected
//                           solutions and filters the all-entities list to
//                           that set.
//   - 'publisher-solution': same as solution-only after the publishers
//                           have narrowed which solutions are available.
//
// We always start from the cached `loadAllEntities()` result — per-component
// lookups never trigger a fresh entity metadata fetch.
//
// Both layers wait for an active PPTB connection and re-run when it arrives
// or changes. A fetch fired before the host handshake resolved used to fail
// once and never retry, leaving "0 entities" until a remount (#41).

import { useCallback, useEffect, useMemo, useState } from 'react';
import { loadSolutionEntityNames } from './dataverseMetadata';
import { metadata, type EntityListItem } from './metadataProvider';
import { useHostSession } from './HostContext';
import { usePersistedSettings } from './usePersistedSettings';

export interface UseScopedEntitiesResult {
  entities: EntityListItem[];
  /**
   * The full, unscoped entity list. For pickers whose choices are fixed by
   * the platform rather than by the user's Query Scope (e.g. Merge, which
   * only supports account / contact / lead / incident).
   */
  allEntities: EntityListItem[];
  /** True while the unscoped list is loading. */
  allLoading: boolean;
  loading: boolean;
  error: string | null;
  /** Re-run the entity list (and solution scope) fetch after an error. */
  retry: () => void;
  /**
   * True when the entity list is empty because the user has not yet
   * configured the Query Scope in Settings (scoped mode with no solutions
   * selected). Used by TargetEditor to show a first-run guidance banner.
   * Always false when entityScopeMode === 'all'.
   */
  needsSetup: boolean;
  /** The active entityScopeMode — exposed so TargetEditor can tailor the hint. */
  scopeMode: import('../state/displaySettings').EntityScopeMode;
}

const errorText = (e: unknown) => (e instanceof Error ? e.message : String(e));

const NO_CONNECTION =
  'No active Dataverse connection. Connect to an environment in Power Platform ToolBox, then Retry.';

export function useScopedEntities(): UseScopedEntitiesResult {
  const [settings] = usePersistedSettings();
  const session = useHostSession();
  // Inside PPTB, wait for the host handshake and an active connection before
  // querying. Standalone runs the fetch anyway so the "not available" error
  // is surfaced. Handshake done without a connection is an error state, not
  // an endless "Loading…".
  const ready = !session.embedded || session.connected;
  const noConnection = session.embedded && session.handshakeDone && !session.connected;
  const connectionKey = `${ready}|${noConnection}|${session.environmentUrl}`;
  const [attempt, setAttempt] = useState(0);
  const retry = useCallback(() => setAttempt((n) => n + 1), []);

  const [all, setAll] = useState<EntityListItem[]>([]);
  const [allLoading, setAllLoading] = useState(true);
  const [allError, setAllError] = useState<string | null>(null);

  // ── Layer 1: unfiltered entity list ──
  useEffect(() => {
    if (noConnection) {
      setAll([]);
      setAllLoading(false);
      setAllError(NO_CONNECTION);
      return;
    }
    if (!ready) {
      setAllLoading(true);
      return;
    }
    let cancelled = false;
    setAllLoading(true);
    setAllError(null);
    metadata
      .listEntities()
      .then((list) => {
        if (!cancelled) setAll(list);
      })
      .catch((e) => {
        if (!cancelled) setAllError(errorText(e));
      })
      .finally(() => {
        if (!cancelled) setAllLoading(false);
      });
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [connectionKey, attempt]);

  // ── Layer 2: solution-component cascade ──
  const solutionKey = settings.selectedSolutionIds.join('|');
  const [scopedNames, setScopedNames] = useState<string[] | null>(null);
  const [scopedLoading, setScopedLoading] = useState(false);
  const [scopedError, setScopedError] = useState<string | null>(null);

  useEffect(() => {
    if (settings.entityScopeMode === 'all') {
      setScopedNames(null);
      setScopedError(null);
      setScopedLoading(false);
      return;
    }
    if (settings.selectedSolutionIds.length === 0) {
      setScopedNames([]);
      setScopedError(null);
      setScopedLoading(false);
      return;
    }
    if (!ready) {
      // Layer 1 reports a missing connection; don't spin here as well.
      setScopedLoading(!noConnection);
      return;
    }

    let cancelled = false;
    setScopedLoading(true);
    setScopedError(null);
    loadSolutionEntityNames(settings.selectedSolutionIds)
      .then((names) => {
        if (!cancelled) setScopedNames(names);
      })
      .catch((e) => {
        if (!cancelled) setScopedError(errorText(e));
      })
      .finally(() => {
        if (!cancelled) setScopedLoading(false);
      });
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [settings.entityScopeMode, solutionKey, connectionKey, attempt]);

  // ── Compose visible list ──
  const visible = useMemo(() => {
    if (settings.entityScopeMode === 'all') return all;
    if (scopedNames === null) return [];
    const set = new Set(scopedNames);
    return all.filter((e) => set.has(e.logicalName));
  }, [all, scopedNames, settings.entityScopeMode]);

  // needsSetup: scoped mode but no solutions chosen yet (entity list will be
  // empty until the user configures Settings → Query Scope).
  const needsSetup =
    settings.entityScopeMode !== 'all' && settings.selectedSolutionIds.length === 0;

  return {
    entities: visible,
    allEntities: all,
    allLoading,
    loading: allLoading || scopedLoading,
    error: allError ?? scopedError,
    retry,
    needsSetup,
    scopeMode: settings.entityScopeMode,
  };
}
