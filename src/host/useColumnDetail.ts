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

import { useEffect, useState } from 'react';
import {
  findTable,
  type ColumnMeta,
  type TableMeta,
  __registerLiveTable,
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
} from '../mock/metadata';
import { loadAttributeWithOptionSet, loadAttributeDetailedMetadata } from './dataverseMetadata';
import type { AttributeMetadata } from './pptbClient';

// Per-(entity, attribute) flag so we don't re-fetch on every render.
const __enriched = new Set<string>();
const key = (entity: string, attr: string) => `${entity}::${attr}`;

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

// Apply detailed AttributeMetadata onto the column in the live registry
// and re-register the table so subscribers refresh.
function patchInto(tbl: TableMeta, columnLogical: string, raw: AttributeMetadata): void {
  const idx = tbl.columns.findIndex((c) => c.logicalName === columnLogical);
  if (idx < 0) return;
  const col = tbl.columns[idx];
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

  const nextCols = [...tbl.columns];
  nextCols[idx] = patched;
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

  useEffect(() => {
    if (!entityLogical || !columnLogical) return;
    const cacheKey = key(entityLogical, columnLogical);
    if (__enriched.has(cacheKey)) return;

    const tbl = findTable(entityLogical);
    if (!tbl) return; // wait until table is loaded
    const col = tbl.columns.find((c) => c.logicalName === columnLogical);
    if (!col) return;
    if (isAlreadyEnriched(col)) {
      __enriched.add(cacheKey);
      return;
    }

    const which = loaderFor(col);
    if (!which) {
      __enriched.add(cacheKey);
      return;
    }

    let cancelled = false;
    setLoading(true);
    setError(null);

    const run =
      which === 'optionset'
        ? loadAttributeWithOptionSet(entityLogical, columnLogical)
        : loadAttributeDetailedMetadata(entityLogical, columnLogical, col.attributeType);

    run
      .then((raw) => {
        if (cancelled) return;
        const tbl2 = findTable(entityLogical);
        if (tbl2) patchInto(tbl2, columnLogical, raw);
        __enriched.add(cacheKey);
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
  }, [entityLogical, columnLogical]);

  return { loading, error };
}
