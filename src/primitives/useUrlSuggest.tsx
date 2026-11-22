// Suggestions for a hand-typed Web API URL field (#268): tables, operations,
// navigation properties, query options and columns, from the metadata.
//
// Works on the existing Fluent Input / Textarea: spread `fieldProps` onto
// the field (they go to the native element) and render `list` next to it.
// The list is rendered inline, so it stays inside a dialog's focus trap,
// and positioned `fixed`, so a scrolling container doesn't clip it.
//
// The list opens on a typed edit (not on focus, a click or a paste) or
// Ctrl+Space. While it's open: ↑ / ↓ move, Enter or Tab inserts, Esc
// closes it until the next edit.

import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type ChangeEvent,
  type FocusEvent,
  type KeyboardEvent,
  type SyntheticEvent,
} from 'react';
import { Caption1, tokens, useId, useMergedRefs, usePositioning } from '@fluentui/react-components';
import {
  applySuggestion,
  urlCaretContext,
  urlSuggestions,
  type UrlMode,
} from '../engine/urlCompletion';
import { useScopedEntities } from '../host/useScopedEntities';
import { metadata } from '../host/metadataProvider';
import { actions } from '../host/csdlProvider';
import { isColumnInScope, isNavInScope, useAdvancedFindOnly } from '../state/metadataScope';
import type { CsdlAction } from '../metadata/actionsCsdl';
import { findTable } from '../metadata/metadata';

type FieldElement = HTMLInputElement | HTMLTextAreaElement;

export interface UrlSuggestOptions {
  value: string;
  onChange: (value: string) => void;
  /** `url`: a path or whole URL. `query`: query options for `entitySet`. */
  mode: UrlMode;
  entitySet?: string;
  /** A Textarea: no combobox role, which a textarea can't take. */
  multiline?: boolean;
}

/** Inserting one of these leaves the caret where the next suggestion goes. */
const CHAINS = /[=(/]$/;

const visuallyHidden = {
  position: 'absolute',
  width: 1,
  height: 1,
  overflow: 'hidden',
  clip: 'rect(0 0 0 0)',
  whiteSpace: 'nowrap',
} as const;

export function useUrlSuggest({ value, onChange, mode, entitySet, multiline }: UrlSuggestOptions) {
  const listId = useId('drs-url-suggest');
  const fieldRef = useRef<FieldElement | null>(null);
  const [focused, setFocused] = useState(false);
  const [caret, setCaret] = useState<number | null>(null);
  // Closed until a typed edit or Ctrl+Space
  const [dismissed, setDismissed] = useState(true);
  const [active, setActive] = useState(0);
  const [operations, setOperations] = useState<readonly CsdlAction[]>([]);
  // Bumped when a table the suggestions wait for has loaded
  const [loaded, setLoaded] = useState(0);
  const pendingCaret = useRef<number | null>(null);

  const { entities, allEntities } = useScopedEntities();
  const advancedFindOnly = useAdvancedFindOnly();

  const live = focused && !dismissed && caret !== null;
  const ctx = useMemo(
    () => (live && caret !== null ? urlCaretContext(value, caret, mode, entitySet) : null),
    [live, value, caret, mode, entitySet],
  );
  const result = useMemo(
    () =>
      ctx
        ? urlSuggestions(ctx, {
            entities,
            allEntities,
            tableOf: findTable,
            operations,
            columnInScope: (c) => isColumnInScope(c, advancedFindOnly),
            navInScope: (n) => isNavInScope(n, advancedFindOnly),
          })
        : null,
    // `loaded`: findTable reads the registry, which a load has filled
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [ctx, entities, allEntities, operations, advancedFindOnly, loaded],
  );

  // Load the table the suggestions wait for. Not useLiveTable: that
  // re-renders on every table load anywhere, for every field.
  const needTable = result?.needTable;
  useEffect(() => {
    if (!needTable) return;
    let current = true;
    void metadata
      .getTable(needTable)
      .catch(() => undefined)
      .finally(() => current && setLoaded((n) => n + 1));
    return () => {
      current = false;
    };
  }, [needTable]);

  const items = result?.items ?? [];
  const open = !!ctx && (items.length > 0 || !!needTable);
  const activeIndex = Math.min(active, Math.max(0, items.length - 1));

  const { targetRef, containerRef } = usePositioning({
    position: 'below',
    align: 'start',
    strategy: 'fixed',
    offset: 2,
  });
  const ref = useMergedRefs(fieldRef, targetRef);

  // After an insert, put the caret after it once the new value renders
  useLayoutEffect(() => {
    const el = fieldRef.current;
    if (pendingCaret.current === null || !el) return;
    el.setSelectionRange(pendingCaret.current, pendingCaret.current);
    setCaret(pendingCaret.current);
    pendingCaret.current = null;
  }, [value]);

  // A new list, or a narrower one, starts at its best match
  const listKey = ctx ? `${ctx.want}:${ctx.from}:${ctx.prefix}` : '';
  useEffect(() => setActive(0), [listKey]);

  // Keep the active item in view
  useEffect(() => {
    if (!open) return;
    document.getElementById(`${listId}-${activeIndex}`)?.scrollIntoView({ block: 'nearest' });
  }, [open, activeIndex, listId]);

  const readCaret = useCallback((e: SyntheticEvent<FieldElement>) => {
    const el = e.currentTarget;
    setCaret(el.selectionStart === el.selectionEnd ? el.selectionStart : null);
  }, []);

  const insert = (i: number) => {
    const item = items[i];
    if (!ctx || !item) return;
    const next = applySuggestion(value, ctx, item);
    setDismissed(!CHAINS.test(item.insert));
    if (next.value === value) {
      // Nothing to re-render: place the caret now
      fieldRef.current?.setSelectionRange(next.caret, next.caret);
      setCaret(next.caret);
      return;
    }
    pendingCaret.current = next.caret;
    onChange(next.value);
  };

  const onKeyDown = (e: KeyboardEvent<FieldElement>) => {
    if (e.nativeEvent.isComposing) return;
    if (e.key === ' ' && e.ctrlKey) {
      e.preventDefault();
      setDismissed(false);
      readCaret(e);
      return;
    }
    if (!open) return;
    const plain = !e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      if (items.length) {
        const step = e.key === 'ArrowDown' ? 1 : -1;
        setActive((activeIndex + step + items.length) % items.length);
      }
    } else if ((e.key === 'Enter' || e.key === 'Tab') && plain && items.length) {
      e.preventDefault();
      insert(activeIndex);
    } else if (e.key === 'Escape') {
      // Close the list, not the dialog around the field
      e.preventDefault();
      e.stopPropagation();
      setDismissed(true);
    }
  };

  const fieldProps = {
    ref,
    onChange: (e: ChangeEvent<FieldElement>) => {
      // A pasted URL is complete; typing is what gets suggestions
      setDismissed((e.nativeEvent as InputEvent).inputType === 'insertFromPaste');
      readCaret(e);
      onChange(e.target.value);
    },
    onKeyDown,
    onSelect: readCaret,
    onFocus: (e: FocusEvent<FieldElement>) => {
      setFocused(true);
      readCaret(e);
      // Cached per connection by the CSDL provider
      if (mode === 'url') {
        void actions
          .loadAll()
          .then(setOperations)
          .catch(() => setOperations([]));
      }
    },
    onBlur: () => {
      setFocused(false);
      setDismissed(true);
    },
    autoComplete: 'off',
    spellCheck: false,
    'aria-autocomplete': 'list' as const,
    'aria-controls': open ? listId : undefined,
    'aria-activedescendant': open && items.length ? `${listId}-${activeIndex}` : undefined,
    ...(multiline ? {} : { role: 'combobox', 'aria-expanded': open }),
  };

  const hidden = result?.hiddenCount ?? 0;
  const total = items.length + hidden;
  const footer = { display: 'block', padding: '4px 8px', color: tokens.colorNeutralForeground3 };
  const list = (
    <>
      {/* Announces the list, which a textarea has no other way to say */}
      <span role="status" style={visuallyHidden}>
        {open && total ? `${total} suggestion${total === 1 ? '' : 's'}` : ''}
      </span>
      {open && (
        <div
          ref={containerRef}
          // Clicks keep the focus in the field
          onMouseDown={(e) => e.preventDefault()}
          style={{
            position: 'fixed',
            zIndex: 1000000,
            minWidth: 280,
            maxWidth: 560,
            padding: 4,
            background: tokens.colorNeutralBackground1,
            border: `1px solid ${tokens.colorNeutralStroke1}`,
            borderRadius: tokens.borderRadiusMedium,
            boxShadow: tokens.shadow16,
          }}
        >
          <div
            id={listId}
            role="listbox"
            aria-label="Suggestions"
            style={{ maxHeight: 280, overflowY: 'auto' }}
          >
            {items.map((s, i) => (
              <div
                key={`${s.kind}:${s.label}`}
                id={`${listId}-${i}`}
                role="option"
                aria-selected={i === activeIndex}
                aria-label={s.detail ? `${s.label}, ${s.detail}` : s.label}
                onMouseEnter={() => setActive(i)}
                onClick={() => insert(i)}
                style={{
                  display: 'flex',
                  alignItems: 'baseline',
                  gap: 12,
                  padding: '4px 8px',
                  borderRadius: tokens.borderRadiusSmall,
                  cursor: 'pointer',
                  background:
                    i === activeIndex ? tokens.colorNeutralBackground1Selected : undefined,
                }}
              >
                <span
                  style={{
                    fontFamily: tokens.fontFamilyMonospace,
                    fontSize: tokens.fontSizeBase200,
                    whiteSpace: 'nowrap',
                  }}
                >
                  {s.label}
                </span>
                {s.detail && (
                  <Caption1
                    style={{
                      marginLeft: 'auto',
                      color: tokens.colorNeutralForeground3,
                      whiteSpace: 'nowrap',
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                    }}
                  >
                    {s.detail}
                  </Caption1>
                )}
              </div>
            ))}
          </div>
          {items.length === 0 && needTable && (
            <Caption1 style={footer}>Loading {needTable} metadata…</Caption1>
          )}
          {hidden > 0 && (
            <Caption1 style={footer}>
              {hidden.toLocaleString()} more — keep typing to narrow the list
            </Caption1>
          )}
        </div>
      )}
    </>
  );

  return { fieldProps, list };
}
