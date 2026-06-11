// useLookupRecords — typeahead helper that fetches live Dataverse records
// for a Lookup-style column. Used by FilterValueInput's LookupTypeaheadCtl
// and any other place that needs to pick a target row for a lookup column
// without mock data.
//
// Strategy:
//   - Wait for the target entity's metadata to be in the live registry
//     (so we know primary key + primary name + entitySetName).
//   - On query change, fire a debounced contains() query against the
//     primaryName attribute. We cap to 50 results — enough for a typeahead.
//   - Cache per-(entity, query) so retyping a previous query is instant.

import { useEffect, useRef, useState } from 'react';
import { findTable } from '../mock/metadata';
import { metadata } from './metadataProvider';

export interface LookupRow {
  id: string;
  name: string;
}

interface CacheKey {
  entity: string;
  query: string;
}
const __cache = new Map<string, LookupRow[]>();
const cacheKey = (k: CacheKey): string => `${k.entity}::${k.query.toLowerCase()}`;

const DEBOUNCE_MS = 250;
const PAGE = 50;

export function useLookupRecords(
  targetEntity: string | null | undefined,
  query: string,
): {
  rows: LookupRow[];
  loading: boolean;
  error: string | null;
} {
  const [rows, setRows] = useState<LookupRow[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const lastReq = useRef(0);

  useEffect(() => {
    if (!targetEntity) {
      setRows([]);
      setError(null);
      return;
    }

    // Cache hit → instant.
    const key = cacheKey({ entity: targetEntity, query });
    const cached = __cache.get(key);
    if (cached) {
      setRows(cached);
      setError(null);
      return;
    }

    // Ensure target metadata is loaded — we need primaryKey/primaryName.
    let cancelled = false;
    setLoading(true);
    setError(null);
    const myReq = ++lastReq.current;

    const timer = setTimeout(async () => {
      try {
        const tbl = findTable(targetEntity) ?? (await metadata.getTable(targetEntity));
        if (cancelled || myReq !== lastReq.current) return;
        if (!tbl) {
          setError(`Couldn't load metadata for "${targetEntity}".`);
          setRows([]);
          return;
        }
        if (typeof window === 'undefined' || !window.dataverseAPI?.queryData) {
          // Standalone mode — no host bridge. Return empty and let the
          // caller render its empty state.
          setRows([]);
          return;
        }
        const q = query.trim();
        const filterPart = q
          ? `&$filter=contains(${tbl.primaryName},'${q.replace(/'/g, "''")}')`
          : '';
        const queryUrl =
          `${tbl.entitySetName}?$select=${tbl.primaryKey},${tbl.primaryName}` +
          `&$top=${PAGE}${filterPart}`;
        const result = await window.dataverseAPI.queryData(queryUrl);
        if (cancelled || myReq !== lastReq.current) return;
        const list = (result.value ?? []) as Record<string, unknown>[];
        const mapped: LookupRow[] = list.map((r) => ({
          id: String(r[tbl.primaryKey] ?? ''),
          name: String(r[tbl.primaryName] ?? r[tbl.primaryKey] ?? ''),
        }));
        __cache.set(key, mapped);
        setRows(mapped);
      } catch (e) {
        if (cancelled || myReq !== lastReq.current) return;
        setError(e instanceof Error ? e.message : String(e));
        setRows([]);
      } finally {
        if (!cancelled && myReq === lastReq.current) setLoading(false);
      }
    }, DEBOUNCE_MS);

    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
  }, [targetEntity, query]);

  return { rows, loading, error };
}
