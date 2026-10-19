// Type-to-filter helpers for long Combobox option lists (#21).
//
// Fluent v9's Combobox mounts every <Option> child — and registers each one
// in the listbox's option collection — as soon as the input gains focus,
// and re-renders all of them on every keystroke. Real environments have
// ~2,000 tables and tables with 500–800 columns, which made the table and
// column pickers take seconds to open and stutter while typing. Capping the
// rendered options (with a "keep typing" hint) keeps the list cheap without
// changing how filtering works: matches still come from the whole list.

import { useDeferredValue, useMemo } from 'react';
import { Caption1, Option, tokens } from '@fluentui/react-components';

/** Most options a picker's listbox renders at once. */
export const MAX_VISIBLE_OPTIONS = 100;

export interface OptionSearchResult<T> {
  /** The first `limit` matches, in list order. */
  visible: T[];
  /** Every match for the query (`visible` plus the hidden remainder). */
  matchCount: number;
  /** Matches not rendered because of the cap. */
  hiddenCount: number;
}

/**
 * Case-insensitive substring filter over `items`, capped at `limit`.
 *
 * - The lower-cased search text per item is built once per `items` array,
 *   not once per keystroke. `keys` is read only then, so it must be a pure
 *   function of the item.
 * - The query is deferred (`useDeferredValue`): the input repaints with the
 *   typed character first and the option list follows in a lower-priority
 *   render, so typing stays responsive on long lists.
 * - An empty (or whitespace) query matches everything.
 */
export function useOptionSearch<T>(
  items: readonly T[],
  query: string,
  keys: (item: T) => ReadonlyArray<string | undefined>,
  limit: number = MAX_VISIBLE_OPTIONS,
): OptionSearchResult<T> {
  const haystacks = useMemo(
    () => items.map((it) => keys(it).filter(Boolean).join('\u0000').toLowerCase()),
    // `keys` is documented as pure; re-deriving on its identity would
    // rebuild the index every render for inline arrow functions.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [items],
  );
  const deferredQuery = useDeferredValue(query);
  return useMemo(() => {
    const q = deferredQuery.trim().toLowerCase();
    const visible: T[] = [];
    let matchCount = 0;
    for (let i = 0; i < items.length; i++) {
      if (q && !haystacks[i].includes(q)) continue;
      matchCount++;
      if (visible.length < limit) visible.push(items[i]);
    }
    return { visible, matchCount, hiddenCount: matchCount - visible.length };
  }, [items, haystacks, deferredQuery, limit]);
}

/**
 * Disabled trailing option shown when the cap hid some matches. Render it
 * as the LAST child of the Combobox. Its value (`__more`) is never a real
 * logical name, and disabled options can't be selected.
 */
export function MoreOptionsHint({ hiddenCount }: { hiddenCount: number }) {
  if (hiddenCount <= 0) return null;
  return (
    <Option value="__more" text="" disabled>
      <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
        {hiddenCount.toLocaleString()} more — keep typing to narrow the list
      </Caption1>
    </Option>
  );
}
