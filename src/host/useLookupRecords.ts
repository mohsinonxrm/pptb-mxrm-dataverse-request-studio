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
//   - DRS's own writes invalidate the cache (`invalidateLookupRecords`,
//     called by the runtime dispatcher), and mounted pickers re-query, so
//     a renamed record doesn't keep its old name in the options (#81).

import { useEffect, useRef, useState, useSyncExternalStore } from 'react';
import { findTable, type TableMeta } from '../mock/metadata';
import { metadata } from './metadataProvider';
import { activeOrg, onActiveOrgChange, useActiveOrg } from './activeOrg';

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

// Bumped on every invalidation. Hooks depend on it to re-query, and a
// search that started before an invalidation doesn't write its (possibly
// pre-write) result back into the cache.
let __version = 0;
const __versionListeners = new Set<() => void>();
const subscribeVersion = (cb: () => void) => {
  __versionListeners.add(cb);
  return () => {
    __versionListeners.delete(cb);
  };
};

/**
 * Drop cached search results after DRS writes records. Pass the table that
 * was written, or nothing to drop every table (a delete can cascade, and a
 * deep insert, action or workflow can touch any table).
 */
export function invalidateLookupRecords(entity?: string): void {
  if (entity) {
    const prefix = `${entity}::`;
    for (const k of [...__cache.keys()]) if (k.startsWith(prefix)) __cache.delete(k);
  } else {
    __cache.clear();
  }
  __version++;
  for (const l of __versionListeners) l();
}

/** Changes after every `invalidateLookupRecords` call. */
export function useLookupRecordsVersion(): number {
  return useSyncExternalStore(subscribeVersion, () => __version);
}

/**
 * Re-read one record's primary name. Used to relabel a selected record
 * after a write when it isn't in the current result page. Resolves to
 * null when there is no host or the table has no primary name column.
 */
export async function fetchRecordName(tbl: TableMeta, id: string): Promise<string | null> {
  if (!tbl.primaryName) return null;
  if (typeof window === 'undefined' || !window.dataverseAPI?.retrieve) return null;
  const rec = await window.dataverseAPI.retrieve(tbl.logicalName, id, [tbl.primaryName]);
  const name = rec?.[tbl.primaryName];
  return name === undefined || name === null ? null : String(name);
}

const DEBOUNCE_MS = 250;
const PAGE = 50;

/**
 * `$filter` value for "name contains the text the user typed", ready to
 * append to a query string (#95):
 *   - `%`, `_` and `[` are Dataverse wildcards inside `contains()`; wrapping
 *     each in brackets makes it match literally ("50%" finds "50%", not
 *     "50 anything").
 *   - `'` is doubled — the OData string-literal escape.
 *   - The text is then percent-encoded (the column is a logical name and
 *     needs none). PPTB's `queryData` appends the query to the URL as
 *     given (Node's URL parser only escapes characters such as spaces and
 *     quotes), so a raw `&` would start a new query option, `#` would cut
 *     the rest off as a fragment, `+` would arrive as a space and `%` would
 *     start an escape sequence. Nothing decodes or re-encodes it on the
 *     way, so this is the only encoding pass.
 */
export function containsFilter(column: string, text: string): string {
  const literal = text.replace(/[[%_]/g, (c) => `[${c}]`).replace(/'/g, "''");
  return `contains(${column},'${encodeURIComponent(literal)}')`;
}

export function useLookupRecords(
  targetEntity: string | null | undefined,
  query: string,
): {
  rows: LookupRow[];
  loading: boolean;
  error: string | null;
} {
  // Re-query when the PPTB connection switches to another org (#52).
  const org = useActiveOrg();
  const [rows, setRows] = useState<LookupRow[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const lastReq = useRef(0);
  const version = useLookupRecordsVersion();

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
    const startVersion = __version;

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
        const filterPart = q ? `&$filter=${containsFilter(tbl.primaryName, q)}` : '';
        const queryUrl =
          `${tbl.entitySetName}?$select=${tbl.primaryKey},${tbl.primaryName}` +
          `&$top=${PAGE}${filterPart}`;
        const result = await window.dataverseAPI.queryData(queryUrl);
        // A result from before a connection switch belongs to the old org.
        if (cancelled || myReq !== lastReq.current || org !== activeOrg()) return;
        const list = (result.value ?? []) as Record<string, unknown>[];
        const mapped: LookupRow[] = list.map((r) => ({
          id: String(r[tbl.primaryKey] ?? ''),
          name: String(r[tbl.primaryName] ?? r[tbl.primaryKey] ?? ''),
        }));
        if (startVersion === __version) __cache.set(key, mapped);
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
  }, [targetEntity, query, org, version]);

  return { rows, loading, error };
}

// Cached searches belong to the org they ran against — drop them when the
// PPTB connection switches (#52). Record data is cheap to re-query, so
// unlike metadata it isn't kept per org.
onActiveOrgChange(() => __cache.clear());
