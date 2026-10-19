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

import { useEffect, useMemo, useState, useCallback } from 'react';
import { metadata, type EntityListItem } from './metadataProvider';
import { __subscribeLiveTables, findTable, type TableMeta } from '../mock/metadata';
import { collectReferencedEntities, type ReferencedTreeInput } from '../engine/navPath';

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
  useEffect(() => __subscribeLiveTables(() => bump((v) => v + 1)), []);

  const [loading, setLoading] = useState(false);
  const table = logical ? findTable(logical) : undefined;
  // Keyed on presence too, so the table is refetched when it drops out of
  // the registry (a metadata reset) instead of staying blank until the
  // user picks another table (#53).
  const present = !!table;

  useEffect(() => {
    if (!logical) return;
    // Already cached → no fetch needed.
    if (present) return;
    setLoading(true);
    void metadata.getTable(logical).finally(() => setLoading(false));
  }, [logical, present]);

  return { table, loading };
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
    metadata
      .listEntities()
      .then((list) => {
        if (!cancelled) setEntities(list);
      })
      .catch((e) => console.warn('[useLiveEntities] listEntities failed', e))
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
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
    try {
      await metadata.getTable(logical);
    } catch (e) {
      console.warn(`[useWarmTable] getTable('${logical}') failed`, e);
    }
  }, []);
}

// Note: `useWarmTables` (bulk pre-fetch) was intentionally removed.
// Pre-fetching every potential nav target on render trips Dataverse's
// 100-concurrent-request cap on wide entities. Editors now load the
// *selected* target only via `useLiveTable(name)`. Display labels for
// non-loaded targets gracefully fall back to the raw logical name.

/**
 * Pre-warm the RELATED entities a request actually references via nav-paths,
 * lambdas, and $expand — so the $filter/$apply/$expand encoders resolve leaf
 * column TYPES against the correct entity (issue #33) even when the request
 * arrives fully-formed (saved-request reload, pasted OData URL, or a direct
 * Execute) and no editor was ever opened to lazily trigger the loads.
 *
 * Targeted, not bulk: only the handful of entities the request mentions are
 * fetched (`metadata.getTable` dedupes + caches), so this stays well clear of
 * the 100-concurrent-request cap that motivated removing `useWarmTables`.
 *
 * Self-healing across hops: the walk only reaches as deep as the registry is
 * currently warm. Each fetch fires a registry update, which re-runs this hook
 * and lets the walk reach one level deeper, until the whole referenced graph
 * is present. Failures don't throw (best-effort warming; the editors still
 * render with whatever metadata is available) but are reported back.
 *
 * Returns:
 *   • `version` — bumps on every registry update. Modes put it in the deps
 *     of their `build*()` memo so the URL bar / Code tab re-encode when a
 *     related table lands, instead of showing the pre-metadata form (#63).
 *   • `pending` — referenced related tables not in the registry yet.
 *   • `failed`  — referenced related tables whose metadata fetch failed.
 */
export function useWarmReferencedTables(
  rootTable: string | null | undefined,
  trees: ReferencedTreeInput,
): { version: number; pending: string[]; failed: string[] } {
  const [registryVersion, setRegistryVersion] = useState(0);
  useEffect(() => __subscribeLiveTables(() => setRegistryVersion((v) => v + 1)), []);
  const [failedNames, setFailedNames] = useState<ReadonlySet<string>>(() => new Set());

  // Stable dependency key — the trees object is rebuilt by the caller each
  // render, so depend on its serialized content rather than its identity.
  const treeKey = useMemo(
    () => JSON.stringify([trees.filter, trees.expand, trees.apply, trees.orderby]),
    [trees.filter, trees.expand, trees.apply, trees.orderby],
  );

  // `trees` is read via the stable `treeKey`; `registryVersion` re-runs the
  // walk as deeper hops become resolvable.
  const referenced = useMemo(
    () => (rootTable ? collectReferencedEntities(rootTable, trees) : []),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [rootTable, treeKey, registryVersion],
  );

  useEffect(() => {
    for (const name of referenced) {
      if (findTable(name)) continue;
      void metadata
        .getTable(name)
        .then((t) => {
          if (!t) setFailedNames((s) => new Set(s).add(name));
        })
        .catch(() => setFailedNames((s) => new Set(s).add(name)));
    }
  }, [referenced]);

  const missing = referenced.filter((name) => !findTable(name));
  return {
    version: registryVersion,
    pending: missing.filter((name) => !failedNames.has(name)),
    failed: missing.filter((name) => failedNames.has(name)),
  };
}
