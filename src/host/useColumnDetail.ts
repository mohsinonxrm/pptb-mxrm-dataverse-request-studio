// useColumnDetail — lazy per-column metadata enrichment.
//
// The basic Attributes projection that buildTable fetches up-front carries
// only LogicalName / SchemaName / DisplayName / AttributeType / MetadataId
// — no Targets (lookups), no OptionSet (picklists), no MaxLength/Format
// (strings), no MinValue/MaxValue/Precision (numerics).
//
// When an editor renders one specific column whose UX needs those typed
// properties — e.g. the lookup typeahead needs `targets[0]`, the picklist
// value picker needs `options`, the integer SpinButton needs `minValue` —
// it calls this hook with that one column. We fetch via:
//
//   getAttributeWithOptionSet(entity, attr)      → Picklist / State / Status / Boolean
//   getAttributeDetailedMetadata(entity, attr, type) → Lookup / String / Memo / Integer / …
//
// On success we patch the column in-place inside the live registry and
// fire __registerLiveTable so every consumer re-renders with the fresh
// data. Each (entity, attribute) is fetched at most once because the
// loader-level cache short-circuits repeats.
//
// Enrichment is tracked per column OBJECT, not per (entity, attribute)
// name. A metadata refresh rebuilds tables from the basic projection with
// fresh column objects, and those must be enriched again — a name-keyed
// flag used to skip them, so refreshed picklists lost their options and
// lookups their targets (#53).

import { useEffect, useRef, useState, useSyncExternalStore } from 'react';
import {
  findTable,
  type ColumnMeta,
  type TableMeta,
  __registerLiveTable,
  __subscribeLiveTables,
  type PicklistColumnMeta,
  type StatusColumnMeta,
  type StateColumnMeta,
  type BooleanColumnMeta,
  type LookupColumnMeta,
  type CustomerColumnMeta,
  type OwnerColumnMeta,
  type StringColumnMeta,
  type MemoColumnMeta,
  type IntegerColumnMeta,
  type BigIntColumnMeta,
  type DecimalColumnMeta,
  type DoubleColumnMeta,
  type MoneyColumnMeta,
  type DateTimeColumnMeta,
  type MultiSelectPicklistColumnMeta,
} from '../metadata/metadata';
import {
  loadAttributeWithOptionSet,
  loadAttributeDetailedMetadata,
  loadAttributesByTypeCast,
} from './dataverseMetadata';
import { activeOrg } from './activeOrg';
import type { AttributeMetadata } from './pptbClient';

// Column objects that are enriched (or need no enrichment), so we don't
// re-fetch on every render. Weak, so replaced columns are collected.
const __enriched = new WeakSet<ColumnMeta>();

// Which AttributeType needs which loader? `null` means no enrichment.
function loaderFor(col: ColumnMeta): 'optionset' | 'detail' | null {
  switch (col.attributeType) {
    case 'Picklist':
    case 'MultiSelectPicklist':
    case 'State':
    case 'Status':
    case 'Boolean':
      return 'optionset';
    case 'Lookup':
    case 'Customer':
    case 'Owner':
    case 'String':
    case 'Memo':
    case 'Integer':
    case 'BigInt':
    case 'Decimal':
    case 'Double':
    case 'Money':
    case 'DateTime':
      return 'detail';
    default:
      return null;
  }
}

// Has this column already been enriched? Detect by presence of the
// type-specific property the loader fills in. If yes, skip the round trip.
function isAlreadyEnriched(col: ColumnMeta): boolean {
  switch (col.attributeType) {
    case 'Picklist':
    case 'MultiSelectPicklist':
    case 'State': {
      const c = col as PicklistColumnMeta | MultiSelectPicklistColumnMeta | StateColumnMeta;
      return Array.isArray(c.options) && c.options.length > 0;
    }
    case 'Status': {
      const c = col as StatusColumnMeta;
      return Array.isArray(c.options) && c.options.length > 0;
    }
    case 'Boolean': {
      const c = col as BooleanColumnMeta;
      return (
        !!c.trueOption &&
        !!c.falseOption &&
        (c.trueOption.label !== 'Yes' || c.falseOption.label !== 'No')
      );
    }
    case 'Lookup':
    case 'Customer':
    case 'Owner': {
      const c = col as LookupColumnMeta | CustomerColumnMeta | OwnerColumnMeta;
      return Array.isArray(c.targets) && c.targets.length > 0;
    }
    case 'String': {
      const c = col as StringColumnMeta;
      return typeof c.maxLength === 'number' && c.maxLength !== 100;
    }
    case 'Memo': {
      const c = col as MemoColumnMeta;
      return typeof c.maxLength === 'number' && c.maxLength !== 2000;
    }
    case 'Integer': {
      const c = col as IntegerColumnMeta;
      return c.minValue !== undefined || c.maxValue !== undefined;
    }
    case 'BigInt': {
      const c = col as BigIntColumnMeta;
      return c.minValue !== undefined || c.maxValue !== undefined;
    }
    case 'Decimal':
    case 'Double':
    case 'Money': {
      const c = col as DecimalColumnMeta | DoubleColumnMeta | MoneyColumnMeta;
      return c.minValue !== undefined || c.maxValue !== undefined;
    }
    case 'DateTime': {
      const c = col as DateTimeColumnMeta;
      return c.dateTimeBehavior !== 'UserLocal';
    }
    default:
      return true;
  }
}

// The column with detailed AttributeMetadata applied, as a new object.
function patchColumn(col: ColumnMeta, raw: AttributeMetadata): ColumnMeta {
  const patched: ColumnMeta = { ...col };

  switch (patched.attributeType) {
    case 'Lookup':
    case 'Customer':
    case 'Owner':
      (patched as LookupColumnMeta).targets = raw.Targets ?? [];
      break;
    case 'String':
      (patched as StringColumnMeta).maxLength =
        raw.MaxLength ?? (patched as StringColumnMeta).maxLength;
      if (raw.Format)
        (patched as StringColumnMeta).format = raw.Format as StringColumnMeta['format'];
      break;
    case 'Memo':
      (patched as MemoColumnMeta).maxLength =
        raw.MaxLength ?? (patched as MemoColumnMeta).maxLength;
      if (raw.Format) (patched as MemoColumnMeta).format = raw.Format as MemoColumnMeta['format'];
      break;
    case 'Integer':
      (patched as IntegerColumnMeta).minValue = raw.MinValue;
      (patched as IntegerColumnMeta).maxValue = raw.MaxValue;
      if (raw.Format)
        (patched as IntegerColumnMeta).format = raw.Format as IntegerColumnMeta['format'];
      break;
    case 'BigInt':
      (patched as BigIntColumnMeta).minValue = raw.MinValue;
      (patched as BigIntColumnMeta).maxValue = raw.MaxValue;
      break;
    case 'Decimal':
      (patched as DecimalColumnMeta).minValue = raw.MinValue;
      (patched as DecimalColumnMeta).maxValue = raw.MaxValue;
      if (typeof raw.Precision === 'number')
        (patched as DecimalColumnMeta).precision = raw.Precision;
      break;
    case 'Double':
      (patched as DoubleColumnMeta).minValue = raw.MinValue;
      (patched as DoubleColumnMeta).maxValue = raw.MaxValue;
      if (typeof raw.Precision === 'number')
        (patched as DoubleColumnMeta).precision = raw.Precision;
      break;
    case 'Money':
      (patched as MoneyColumnMeta).minValue = raw.MinValue;
      (patched as MoneyColumnMeta).maxValue = raw.MaxValue;
      if (typeof raw.Precision === 'number') (patched as MoneyColumnMeta).precision = raw.Precision;
      break;
    case 'DateTime':
      if (raw.Format)
        (patched as DateTimeColumnMeta).format = raw.Format as DateTimeColumnMeta['format'];
      if (raw.DateTimeBehavior?.Value) {
        (patched as DateTimeColumnMeta).dateTimeBehavior = raw.DateTimeBehavior
          .Value as DateTimeColumnMeta['dateTimeBehavior'];
      }
      break;
    case 'Boolean':
      if (raw.OptionSet?.TrueOption || raw.OptionSet?.FalseOption) {
        const t = raw.OptionSet.TrueOption?.Label?.UserLocalizedLabel?.Label;
        const f = raw.OptionSet.FalseOption?.Label?.UserLocalizedLabel?.Label;
        if (t) (patched as BooleanColumnMeta).trueOption = { value: 1, label: t };
        if (f) (patched as BooleanColumnMeta).falseOption = { value: 0, label: f };
      }
      break;
    case 'Picklist':
    case 'MultiSelectPicklist':
    case 'State': {
      const opts = (raw.OptionSet?.Options ?? []).map((o) => ({
        value: o.Value,
        label: o.Label?.UserLocalizedLabel?.Label ?? String(o.Value),
      }));
      (patched as PicklistColumnMeta | MultiSelectPicklistColumnMeta | StateColumnMeta).options =
        opts;
      if (patched.attributeType !== 'State' && typeof raw.DefaultFormValue === 'number') {
        (patched as PicklistColumnMeta | MultiSelectPicklistColumnMeta).defaultFormValue =
          raw.DefaultFormValue;
      }
      break;
    }
    case 'Status': {
      const opts = (raw.OptionSet?.Options ?? []).map((o) => ({
        value: o.Value,
        label: o.Label?.UserLocalizedLabel?.Label ?? String(o.Value),
        state: 0,
      }));
      (patched as StatusColumnMeta).options = opts;
      break;
    }
  }

  __enriched.add(patched);
  return patched;
}

// Apply detailed AttributeMetadata onto the column in the live registry
// and re-register the table so subscribers refresh.
function patchInto(tbl: TableMeta, columnLogical: string, raw: AttributeMetadata): void {
  const idx = tbl.columns.findIndex((c) => c.logicalName === columnLogical);
  if (idx < 0) return;
  const nextCols = [...tbl.columns];
  nextCols[idx] = patchColumn(tbl.columns[idx], raw);
  __registerLiveTable({ ...tbl, columns: nextCols });
}

/**
 * Hook: lazily load type-specific metadata for ONE column. Pass the parent
 * entity's logical name and the column's logical name. On mount it checks
 * the live column for the relevant typed property; if missing, fires the
 * appropriate `getAttributeWithOptionSet` / `getAttributeDetailedMetadata`
 * call, patches the column in the registry, and re-renders.
 *
 * Safe to call on every render of an editor; idempotent per (entity, attr).
 */
export function useColumnDetail(
  entityLogical: string | null | undefined,
  columnLogical: string | null | undefined,
): { loading: boolean; error: string | null } {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // The live column object. Re-renders when it is replaced in the registry
  // (table loaded, refreshed, or patched), which re-runs the effect below.
  const col = useSyncExternalStore(__subscribeLiveTables, () =>
    entityLogical && columnLogical
      ? findTable(entityLogical)?.columns.find((c) => c.logicalName === columnLogical)
      : undefined,
  );

  useEffect(() => {
    if (!entityLogical || !columnLogical) return;
    if (!col) return; // wait until table is loaded
    if (__enriched.has(col)) return;
    if (isAlreadyEnriched(col)) {
      __enriched.add(col);
      return;
    }

    const which = loaderFor(col);
    if (!which) {
      __enriched.add(col);
      return;
    }

    let cancelled = false;
    setLoading(true);
    setError(null);

    const run =
      which === 'optionset'
        ? loadAttributeWithOptionSet(entityLogical, columnLogical, col.attributeType)
        : loadAttributeDetailedMetadata(entityLogical, columnLogical, col.attributeType);

    run
      .then((raw) => {
        if (cancelled) return;
        // Patch only the column we fetched for. If a refresh replaced it in
        // the meantime, the effect re-runs for the new column instead.
        const tbl2 = findTable(entityLogical);
        if (tbl2?.columns.includes(col)) patchInto(tbl2, columnLogical, raw);
      })
      .catch((e) => {
        if (!cancelled) setError(e instanceof Error ? e.message : String(e));
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });

    return () => {
      cancelled = true;
    };
  }, [entityLogical, columnLogical, col]);

  return { loading, error };
}

// ── Whole-table enrichment ─────────────────────────────────────────────
//
// The JSON body editor's schema (#268) describes every writable column at
// once: options, lookup targets, date behavior. One type-cast request per
// attribute subtype fills them all, instead of one request per column.

const BULK_CASTS: { types: string[]; cast: string; select: string; expand?: string }[] = [
  {
    types: ['Picklist'],
    cast: 'Microsoft.Dynamics.CRM.PicklistAttributeMetadata',
    select: 'LogicalName,DefaultFormValue',
    expand: 'OptionSet($select=Options)',
  },
  {
    types: ['MultiSelectPicklist'],
    cast: 'Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata',
    select: 'LogicalName,DefaultFormValue',
    expand: 'OptionSet($select=Options)',
  },
  {
    types: ['State'],
    cast: 'Microsoft.Dynamics.CRM.StateAttributeMetadata',
    select: 'LogicalName',
    expand: 'OptionSet($select=Options)',
  },
  {
    types: ['Status'],
    cast: 'Microsoft.Dynamics.CRM.StatusAttributeMetadata',
    select: 'LogicalName',
    expand: 'OptionSet($select=Options)',
  },
  {
    types: ['Boolean'],
    cast: 'Microsoft.Dynamics.CRM.BooleanAttributeMetadata',
    select: 'LogicalName',
    expand: 'OptionSet($select=TrueOption,FalseOption)',
  },
  {
    // Customer and Owner columns are LookupAttributeMetadata too
    types: ['Lookup', 'Customer', 'Owner'],
    cast: 'Microsoft.Dynamics.CRM.LookupAttributeMetadata',
    select: 'LogicalName,Targets',
  },
  {
    types: ['DateTime'],
    cast: 'Microsoft.Dynamics.CRM.DateTimeAttributeMetadata',
    select: 'LogicalName,Format,DateTimeBehavior',
  },
];

// Column objects a bulk fetch covers (found or not), so a column the
// response leaves as it was doesn't trigger another round. Marked when the
// fetch starts, so re-renders meanwhile don't repeat it; unmarked if it
// fails or the connection switches, so the next render of the table tries
// again. Separate from `__enriched`, which would stop useColumnDetail too.
const __bulkTried = new WeakSet<ColumnMeta>();

function needsBulk(c: ColumnMeta): boolean {
  return (
    !__enriched.has(c) &&
    !__bulkTried.has(c) &&
    !isAlreadyEnriched(c) &&
    BULK_CASTS.some((b) => b.types.includes(c.attributeType))
  );
}

/** Enrich every column of `tbl` that needs it, one request per subtype. */
async function enrichTable(tbl: TableMeta): Promise<void> {
  const org = activeOrg();
  const cols = tbl.columns.filter(needsBulk);
  for (const c of cols) __bulkTried.add(c);
  const casts = BULK_CASTS.filter((b) => cols.some((c) => b.types.includes(c.attributeType)));
  const rows = new Map<string, AttributeMetadata>();
  await Promise.all(
    casts.map(async (b) => {
      try {
        const list = await loadAttributesByTypeCast(tbl.logicalName, b.cast, b.select, b.expand);
        for (const r of list) rows.set(r.LogicalName, r);
      } catch (e) {
        console.warn(`[metadata] ${b.cast} details for ${tbl.logicalName} failed`, e);
        for (const c of cols) if (b.types.includes(c.attributeType)) __bulkTried.delete(c);
      }
    }),
  );
  // Another org's registry now: its table isn't this one
  if (activeOrg() !== org) {
    for (const c of cols) __bulkTried.delete(c);
    return;
  }
  // Patch only the columns we fetched for. One a refresh or useColumnDetail
  // replaced in the meantime is left to them.
  const cur = findTable(tbl.logicalName);
  if (!cur) return;
  let changed = false;
  const nextCols = cur.columns.map((c) => {
    const raw = rows.get(c.logicalName);
    if (!raw || !cols.includes(c)) return c;
    changed = true;
    const patched = patchColumn(c, { ...raw, AttributeType: c.attributeType });
    __bulkTried.add(patched);
    return patched;
  });
  if (changed) __registerLiveTable({ ...cur, columns: nextCols });
}

/**
 * Hook: the live table, with every column's type-specific metadata loaded
 * (options, lookup targets, date behavior). `loading` until the fetches
 * land; on failure the columns stay as they were, and the next change to
 * the table (or remount) tries again.
 */
export function useTableDetail(entityLogical: string | null | undefined): {
  table: TableMeta | undefined;
  loading: boolean;
} {
  const tbl = useSyncExternalStore(__subscribeLiveTables, () =>
    entityLogical ? findTable(entityLogical) : undefined,
  );
  const [busy, setBusy] = useState(false);
  // The table object the last fetch ran for: a failed fetch leaves its
  // columns pending, but that table is done loading
  const [settled, setSettled] = useState<TableMeta>();
  const run = useRef(0);
  const pending = !!tbl && tbl !== settled && tbl.columns.some(needsBulk);

  useEffect(() => {
    if (!tbl || !tbl.columns.some(needsBulk)) return;
    const id = ++run.current;
    setBusy(true);
    // Not cancelled on cleanup: the patch re-registers the table, which
    // re-runs this effect, and `busy` must still clear.
    void enrichTable(tbl).finally(() => {
      if (run.current !== id) return;
      setSettled(tbl);
      setBusy(false);
    });
  }, [tbl]);

  return { table: tbl, loading: busy || pending };
}
