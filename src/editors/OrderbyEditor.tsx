import { useEffect, useMemo, useState } from 'react';
import { Button, Combobox, Option, tokens, MessageBar, MessageBarBody, mergeClasses, ToggleButton } from '@fluentui/react-components';
import { SegmentedToggle } from '../primitives/SegmentedToggle';
import {
  TextSortAscending20Filled, TextSortAscending20Regular, TextSortDescending20Regular,
  Add20Regular, Delete20Regular, ArrowUp20Regular, ArrowDown20Regular,
} from '@fluentui/react-icons';
import { useStudioStyles } from '../primitives/styles';
import { findTable, isCompanionLogicalReadOnly, resolveNavPath } from '../mock/metadata';
import { PaneHead } from './PaneHead';
import { ApplyOverridesBanner } from './ApplyOverridesBanner';
import { SortableList, SortableItem } from '../primitives/Sortable';
import type { RequestGroup } from '../registry/requestTypes';

export interface OrderbySpec {
  /** Stable id for keying / DnD reorder. Generated on `add` so rows that
   *  start with the same default column don't collide as identical items
   *  in the sortable list. Optional for backward compatibility with
   *  persisted state — falls back to `${col}-${i}` if absent. */
  id?: string;
  col: string;
  dir: 'asc' | 'desc';
}

export interface OrderbyEditorProps {
  table: string;
  items: OrderbySpec[];
  setItems: (items: OrderbySpec[]) => void;
  group?: RequestGroup;
  /** When true, $apply is on at the root and this $orderby is ignored. */
  applyActive?: boolean;
}

export function OrderbyEditor({ table, items, setItems, group = 'read', applyActive }: OrderbyEditorProps) {
  const s = useStudioStyles();
  const tbl = findTable(table);
  if (!tbl) return null;

  // Per-type orderby warnings — sourced from spec §21 (order-rows.md):
  //  - Picklist/State/Status: sorts by integer value, not label
  //  - MultiSelectPicklist:   sorts by raw semicolon-separated string (meaningless)
  //  - Memo:                  long text — slow, anti-pattern
  //  - Lookup/Customer/Owner: sorts by related row's primary name field (cross-table join)
  const warnings: string[] = [];
  for (const it of items) {
    const c = resolveNavPath(tbl, it.col).leaf;
    if (!c) continue;
    if (c.attributeType === 'Picklist' || c.attributeType === 'State' || c.attributeType === 'Status') {
      warnings.push(`${c.displayName}: choice columns sort by the underlying integer value, not the localized label. Use FetchXml if you need label order.`);
    }
    if (c.attributeType === 'MultiSelectPicklist') {
      warnings.push(`${c.displayName}: MultiSelectPicklist sorts by the raw semicolon-separated string — effectively meaningless.`);
    }
    if (c.attributeType === 'Memo') {
      warnings.push(`${c.displayName}: ordering long text is an anti-pattern (slow). Add a unique key as a tie-breaker for stable paging.`);
    }
    if (c.attributeType === 'Lookup' || c.attributeType === 'Customer' || c.attributeType === 'Owner') {
      warnings.push(`${c.displayName}: lookup columns sort by the related row's primary name field (join). Slower than scalar sorts.`);
    }
  }
  // Recommend a unique-key tie-breaker if no primary key is in the orderby
  const pkInSort = items.some(it => {
    const c = resolveNavPath(tbl, it.col).leaf;
    return c?.attributeType === 'Uniqueidentifier';
  });
  if (items.length > 0 && !pkInSort) {
    warnings.push(`Add ${tbl.primaryKey} as the last orderby column for deterministic paging.`);
  }

  const newOrderId = () =>
    (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function')
      ? `ord_${crypto.randomUUID()}`
      : `ord_${Date.now().toString(36)}_${Math.floor(Math.random() * 1e6).toString(36)}`;

  const add = () => setItems([...items, { id: newOrderId(), col: tbl.columns[0].logicalName, dir: 'asc' }]);
  const update = (i: number, patch: Partial<OrderbySpec>) =>
    setItems(items.map((it, j) => j === i ? { ...it, ...patch } : it));
  const remove = (i: number) => setItems(items.filter((_, j) => j !== i));
  const swap = (i: number, j: number) => {
    if (j < 0 || j >= items.length) return;
    const next = [...items];
    [next[i], next[j]] = [next[j], next[i]];
    setItems(next);
  };

  return (
    <div>
      <PaneHead
        icon={TextSortAscending20Filled}
        title="$orderby"
        sub="Multi-column sort. Note: choices/option-sets order by integer value, not label (use FetchXml for label order)."
        group={group}
      />

      {applyActive && <ApplyOverridesBanner clause="$orderby" />}

      {items.length === 0 && (
        <MessageBar layout="multiline" intent="info" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            No <code>$orderby</code> — server-default order. Add one to make pagination deterministic; recommended best practice is <code>orderby={tbl.primaryKey} asc</code> as a tie-breaker.
          </MessageBarBody>
        </MessageBar>
      )}

      {warnings.length > 0 && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 12 }}>
          {warnings.map((w, i) => (
            <MessageBar key={i} intent="warning">
              <MessageBarBody>{w}</MessageBarBody>
            </MessageBar>
          ))}
        </div>
      )}

      <div style={{ display: 'flex', flexDirection: 'column', gap: 8, maxWidth: 720 }}>
        {/* Drag-reorder via the grip column on each row. Up/Down arrow
            buttons preserved as an alternative for keyboard / mouse-only
            users. Each row carries a unique generated `id` (assigned by
            `add` or by the parser) — using the column logical name as id
            broke when multiple rows shared the same default column (e.g.
            three fresh rows all seeded with `accountid`). The id falls
            back to a `${col}-${i}` form for legacy persisted state. */}
        <SortableList
          ids={items.map((it, i) => it.id ?? `${it.col}-${i}`)}
          onReorder={(from, to) => {
            const next = [...items];
            const [moved] = next.splice(from, 1);
            next.splice(to, 0, moved);
            setItems(next);
          }}
        >
          {items.map((it, i) => {
            const col = resolveNavPath(tbl, it.col).leaf;
            const itemId = it.id ?? `${it.col}-${i}`;
            // Grid columns: grip · row-number · column picker · sort-direction toggle · ↑ · ↓ · 🗑
            return (
              <SortableItem key={itemId} id={itemId}>
                {({ gripProps, Grip }) => (
                  <div className={mergeClasses(s.inlineCard)} style={{ padding: '8px 10px', display: 'grid', gridTemplateColumns: '20px 24px 1fr auto 28px 28px 28px', gap: 8, alignItems: 'center' }}>
                    <span {...gripProps} aria-label={`Drag to reorder ${col?.displayName ?? it.col}`}>
                      <Grip />
                    </span>
                    <span style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11, color: tokens.colorNeutralForeground3 }}>
                      {i + 1}.
                    </span>
                    <OrderbyColumnCombo
                      value={it.col}
                      currentDisplay={col?.displayName}
                      // Read picker filter — same rule as SelectEditor /
                      // FilterEditor: honor IsValidForRead, hide companion
                      // *name / *yominame logical read-only columns.
                      columns={tbl.columns
                        .filter(c => c.isValidForRead !== false)
                        .filter(c => !isCompanionLogicalReadOnly(c))}
                      onChange={(logicalName) => update(i, { col: logicalName })}
                    />
                    <SegmentedToggle ariaLabel="Sort direction">
                      <ToggleButton
                        checked={it.dir === 'asc'}
                        icon={<TextSortAscending20Regular />}
                        onClick={() => update(i, { dir: 'asc' })}
                      >Asc</ToggleButton>
                      <ToggleButton
                        checked={it.dir === 'desc'}
                        icon={<TextSortDescending20Regular />}
                        onClick={() => update(i, { dir: 'desc' })}
                      >Desc</ToggleButton>
                    </SegmentedToggle>
                    <Button size="small" appearance="subtle" icon={<ArrowUp20Regular />}   onClick={() => swap(i, i - 1)} disabled={i === 0}                aria-label="Move up" />
                    <Button size="small" appearance="subtle" icon={<ArrowDown20Regular />} onClick={() => swap(i, i + 1)} disabled={i === items.length - 1} aria-label="Move down" />
                    <Button size="small" appearance="subtle" icon={<Delete20Regular />}    onClick={() => remove(i)} aria-label="Remove" />
                  </div>
                )}
              </SortableItem>
            );
          })}
        </SortableList>
      </div>

      <div style={{ marginTop: 12 }}>
        <Button icon={<Add20Regular />} appearance="outline" size="small" onClick={add}>Add sort</Button>
      </div>
    </div>
  );
}

export const orderbyToOData = (items: OrderbySpec[]): string =>
  items.map(it => it.dir === 'asc' ? it.col : `${it.col} desc`).join(',');

// ────────────────────────────────────────────────────────────
// Searchable column picker for the orderby row
// ────────────────────────────────────────────────────────────
/**
 * Freeform Combobox that filters the table's column list as the user
 * types. Drop-in replacement for the previous always-show-all Combobox.
 * Same UX as the filter editor's column picker for consistency.
 */
function OrderbyColumnCombo({
  value, currentDisplay, columns, onChange,
}: {
  value: string;
  currentDisplay?: string;
  columns: { logicalName: string; displayName: string }[];
  onChange: (logicalName: string) => void;
}) {
  const [query, setQuery] = useState<string>(currentDisplay ?? value);
  useEffect(() => { setQuery(currentDisplay ?? value); }, [currentDisplay, value]);
  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q || q === (currentDisplay ?? '').toLowerCase()) return columns;
    return columns.filter(c =>
      c.displayName.toLowerCase().includes(q) ||
      c.logicalName.toLowerCase().includes(q),
    );
  }, [columns, query, currentDisplay]);

  return (
    <Combobox
      freeform
      clearable
      size="small"
      value={query}
      selectedOptions={value ? [value] : []}
      onChange={(e) => setQuery((e.target as HTMLInputElement).value)}
      onOptionSelect={(_, d) => {
        if (d.optionValue) {
          const picked = columns.find(c => c.logicalName === d.optionValue);
          onChange(d.optionValue);
          setQuery(picked?.displayName ?? d.optionValue);
        } else {
          setQuery('');
        }
      }}
      placeholder="Search columns…"
      listbox={{ style: { maxHeight: 320 } }}
    >
      {filtered.length === 0 && (
        <Option value="__none" text="" disabled>
          <span style={{ color: tokens.colorNeutralForeground3, fontSize: 11 }}>
            No matches for "{query}"
          </span>
        </Option>
      )}
      {filtered.map(c => (
        <Option key={c.logicalName} value={c.logicalName} text={c.displayName}>
          {c.displayName}{' '}
          <span style={{ color: tokens.colorNeutralForeground3, fontFamily: tokens.fontFamilyMonospace, fontSize: 10 }}>
            · {c.logicalName}
          </span>
        </Option>
      ))}
    </Combobox>
  );
}
