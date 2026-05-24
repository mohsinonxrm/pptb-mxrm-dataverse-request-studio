// React hooks that bridge live Dataverse metadata to the studio-shape
// `TableMeta` / `ColumnMeta` types editors expect.
//
// These hooks are the studio-shape wrappers around the raw loaders in
// `dataverseMetadata.ts` + the transformer in `metadataProvider.ts`.
// Components import these (not the loaders) so they get
// `findTable()`-friendly data + the registry subscription that re-renders
// downstream when a fetch resolves.
//
// Hooks:
//   - `useLiveTable(logical)` — fires `metadata.getTable(logical)` + subscribes.
//   - `useLiveEntities()` — `metadata.listEntities()`, returns EntityListItem[].

import { useEffect, useState, useCallback } from 'react';
import { metadata, type EntityListItem } from './metadataProvider';
import {
  __subscribeLiveTables, findTable, type TableMeta,
} from '../mock/metadata';

/**
 * Returns the TableMeta for a logical name. On mount/change it fires the
 * async fetch; when the metadata provider publishes the resulting TableMeta
 * into the synchronous registry, this hook re-renders (via the registry
 * subscription) so every consumer of `findTable(logical)` picks up live data.
 *
 * Returns `{ table, loading }` so editors can show a spinner during the
 * initial fetch.
 */
export function useLiveTable(logical: string | null | undefined): {
  table: TableMeta | undefined;
  loading: boolean;
} {
  // Bump on every registry update — that's what triggers re-render after
  // a metadata fetch resolves.
  const [, bump] = useState(0);
  useEffect(() => __subscribeLiveTables(() => bump(v => v + 1)), []);

  const [loading, setLoading] = useState(false);

  useEffect(() => {
    if (!logical) return;
    // Already cached → no fetch needed.
    if (findTable(logical)) return;
    setLoading(true);
    void metadata.getTable(logical).finally(() => setLoading(false));
  }, [logical]);

  return { table: logical ? findTable(logical) : undefined, loading };
}

/**
 * Returns the entity list for the table picker. Reads through
 * `metadata.listEntities()` which goes through the shared cache. The list
 * is unfiltered here — caller (or settings/scope hooks) can narrow it.
 */
export function useLiveEntities(): {
  entities: EntityListItem[];
  loading: boolean;
} {
  const [entities, setEntities] = useState<EntityListItem[]>([]);
  const [loading, setLoading] = useState<boolean>(true);

  useEffect(() => {
    let cancelled = false;
    metadata.listEntities()
      .then(list => { if (!cancelled) setEntities(list); })
      .catch(e => console.warn('[useLiveEntities] listEntities failed', e))
      .finally(() => { if (!cancelled) setLoading(false); });
    return () => { cancelled = true; };
  }, []);

  return { entities, loading };
}

/**
 * Imperative warmer — call after picking a table so synchronous reads of
 * `findTable(logical)` return live data on the next render.
 */
export function useWarmTable(): (logical: string) => Promise<void> {
  return useCallback(async (logical: string) => {
    if (!logical) return;
    try { await metadata.getTable(logical); }
    catch (e) { console.warn(`[useWarmTable] getTable('${logical}') failed`, e); }
  }, []);
}

// Note: `useWarmTables` (bulk pre-fetch) was intentionally removed.
// Pre-fetching every potential nav target on render trips Dataverse's
// 100-concurrent-request cap on wide entities. Editors now load the
// *selected* target only via `useLiveTable(name)`. Display labels for
// non-loaded targets gracefully fall back to the raw logical name.
