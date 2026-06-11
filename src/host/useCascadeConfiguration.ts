// useCascadeConfiguration — on-demand fetch of CascadeConfiguration for
// every 1:N relationship on a given entity.
//
// Why a separate hook (not in the basic metadata projection):
//
//   The basic NavProperty projection that warms when the user picks a
//   table (see metadataProvider) deliberately stays slim — we'd otherwise
//   pull MB of relationship metadata that 90% of modes never look at.
//   CascadeConfiguration only matters for Delete's "what gets cascaded?"
//   preview. So we fetch it lazily, once per (entity, connection), the
//   first time the Confirmation pane mounts.
//
// What we fetch:
//
//   GET /api/data/v9.2/EntityDefinitions(LogicalName='<entity>')
//       /OneToManyRelationships?$select=SchemaName,CascadeConfiguration,
//                                       ReferencedEntity,ReferencingEntity
//
//   Routed through PPTB's `dataverseAPI.getEntityRelatedMetadata` so the
//   bridge handles auth + headers + connection-target routing.
//
// What we return:
//
//   A `Map<SchemaName, CascadeBehavior>` where CascadeBehavior is the
//   `Delete` field from each CascadeConfiguration. SchemaName matches
//   `NavProperty.relationshipName` in our existing TableMeta projection,
//   so the caller can do a direct lookup per row.
//
// Caching:
//
//   Module-level Map keyed by entity logical name. Once warm, repeated
//   calls return instantly. The cache persists across mode switches but
//   resets on page reload — short-lived enough that a metadata change
//   isn't a real problem in practice.

import { useEffect, useState } from 'react';

export type CascadeBehavior =
  | 'NoCascade'
  | 'Cascade'
  | 'Active'
  | 'UserOwned'
  | 'RemoveLink'
  | 'Restrict';

export interface CascadeRow {
  schemaName: string;
  /** The "Delete" behavior — what happens to child records when the
   *  parent is deleted. The only one Delete mode cares about. */
  deleteBehavior: CascadeBehavior;
  /** Full CascadeConfiguration shape, kept for callers that want the
   *  Assign/Share/Reparent/Merge/etc. behaviors (none today). */
  full?: Record<string, CascadeBehavior>;
}

interface CascadeConfigShape {
  Assign?: CascadeBehavior;
  Delete?: CascadeBehavior;
  Share?: CascadeBehavior;
  Unshare?: CascadeBehavior;
  Reparent?: CascadeBehavior;
  Merge?: CascadeBehavior;
  Rollup?: CascadeBehavior;
}

interface OneToManyMeta {
  SchemaName: string;
  CascadeConfiguration?: CascadeConfigShape;
  ReferencedEntity?: string;
  ReferencingEntity?: string;
}

// Module-level cache. Map keyed by entity logical name → array of rows.
const __cache = new Map<string, CascadeRow[]>();

/**
 * Fetch (with cache) the cascade-delete behavior for every 1:N
 * relationship rooted at `entityLogical`. Returns `{ rows, loading,
 * error }`. `rows` is empty until the first fetch resolves; check
 * `loading` to distinguish "still warming up" from "actually empty".
 */
export function useCascadeConfiguration(entityLogical: string | null): {
  rows: CascadeRow[];
  loading: boolean;
  error: string | null;
} {
  const [rows, setRows] = useState<CascadeRow[]>(() =>
    entityLogical ? (__cache.get(entityLogical) ?? []) : [],
  );
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!entityLogical) {
      setRows([]);
      setError(null);
      return;
    }
    const cached = __cache.get(entityLogical);
    if (cached) {
      setRows(cached);
      setError(null);
      return;
    }
    if (typeof window === 'undefined' || !window.dataverseAPI?.getEntityRelatedMetadata) {
      // Standalone mode — no host. Return empty; the table will fall
      // back to its hedged "Cascade or Restrict" labels.
      setRows([]);
      return;
    }

    let cancelled = false;
    setLoading(true);
    setError(null);
    window.dataverseAPI
      .getEntityRelatedMetadata(entityLogical, 'OneToManyRelationships', [
        'SchemaName',
        'CascadeConfiguration',
        'ReferencedEntity',
        'ReferencingEntity',
      ])
      .then((res) => {
        if (cancelled) return;
        // Some PPTB hosts wrap the response in { value: [...] }; others
        // return the array directly. Normalize.
        const valueField = (res as { value?: unknown })?.value;
        const arr: OneToManyMeta[] = Array.isArray(res)
          ? (res as OneToManyMeta[])
          : Array.isArray(valueField)
            ? (valueField as OneToManyMeta[])
            : [];
        const mapped: CascadeRow[] = arr.map((r) => ({
          schemaName: r.SchemaName,
          deleteBehavior: (r.CascadeConfiguration?.Delete ?? 'NoCascade') as CascadeBehavior,
          full: r.CascadeConfiguration as Record<string, CascadeBehavior> | undefined,
        }));
        __cache.set(entityLogical, mapped);
        setRows(mapped);
        setLoading(false);
      })
      .catch((e: unknown) => {
        if (cancelled) return;
        const msg = e instanceof Error ? e.message : String(e);
        setError(msg);
        setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [entityLogical]);

  return { rows, loading, error };
}

/**
 * Severity ranking used by the Delete cascade preview to sort + highlight
 * relationships. LOWER number = MORE destructive / higher priority to
 * surface to the user.
 *
 *   1. Cascade       → child records WILL be deleted (parental — top!)
 *   2. Restrict      → delete BLOCKED if children exist (also critical)
 *   3. Active        → parent.active gets cascaded
 *   4. UserOwned     → cascades based on owning user match
 *   5. RemoveLink    → unlinks (less destructive, recoverable)
 *   6. NoCascade     → nothing happens (least interesting)
 */
export function cascadeSeverityRank(b: CascadeBehavior): number {
  switch (b) {
    case 'Cascade':
      return 1;
    case 'Restrict':
      return 2;
    case 'Active':
      return 3;
    case 'UserOwned':
      return 4;
    case 'RemoveLink':
      return 5;
    case 'NoCascade':
      return 6;
  }
}

/** True if a behavior implies cascade-deletion of child records — the
 *  rows that deserve a "parental" visual emphasis in the preview. */
export function isParental(b: CascadeBehavior): boolean {
  return b === 'Cascade' || b === 'Active' || b === 'UserOwned';
}
