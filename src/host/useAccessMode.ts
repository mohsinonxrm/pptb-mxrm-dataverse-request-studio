// useAccessMode — privilege check + access-mode derivation.
//
// On mount:
//   1. Calls WhoAmI to grab the current user id.
//   2. Parallel `RetrieveUserPrivilegeByPrivilegeName` for prvReadPublisher /
//      prvReadSolution / prvReadCustomization.
//   3. Derives 5 mutually-exclusive booleans for the SettingsDrawer's
//      scope-mode RadioGroup gating.
//   4. Pre-warms the entity metadata cache on success — speeds up the
//      first time the user opens the entity picker.

import { useState, useEffect } from 'react';
import { getAccessSummary, getAllAdvancedFindEntities, type AccessSummary } from './pptbClient';
import { metadataCache } from './cache';

export function useAccessMode() {
  const [accessSummary, setAccessSummary] = useState<AccessSummary | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let mounted = true;

    async function checkAccess() {
      try {
        setLoading(true);
        setError(null);

        const summary = await getAccessSummary();

        if (mounted) {
          setAccessSummary(summary);

          // Pre-warm: most users will open the entity picker right after
          // mount. Fire the all-entities fetch in the background so it's
          // hot when they get there.
          if (summary && !summary.noAccessMode) {
            getAllAdvancedFindEntities()
              .then((entities) => {
                metadataCache.setAllEntities(entities);
              })
              .catch((err) => {
                console.warn('[useAccessMode] failed to preload entity metadata:', err);
              });
          }
        }
      } catch (err) {
        if (mounted) setError(err instanceof Error ? err.message : String(err));
      } finally {
        if (mounted) setLoading(false);
      }
    }

    checkAccess();
    return () => {
      mounted = false;
    };
  }, []);

  return {
    accessSummary,
    loading,
    error,
    fullFilterMode: accessSummary?.fullFilterMode ?? false,
    solutionsOnlyMode: accessSummary?.solutionsOnlyMode ?? false,
    publishersOnlyMode: accessSummary?.publishersOnlyMode ?? false,
    metadataOnlyMode: accessSummary?.metadataOnlyMode ?? false,
    noAccessMode: accessSummary?.noAccessMode ?? false,
  };
}
