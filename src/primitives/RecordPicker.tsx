// RecordPicker — live typeahead OR direct-GUID-entry for a single record.
//
// Used by:
//   • TargetEditor (Retrieve Single, Update, Delete, Upsert, Merge — anywhere
//     the request needs a primary key).
//   • Standalone in modes that pick a relationship target (Associate /
//     Disassociate).
//
// Two modes, toggled by the user. A picker given a `memoryKey` remembers the
// toggle and the raw GUID text in the enclosing RecordPickerMemoryScope
// (one per mode, in ModeShell), so switching sidebar panes — which unmounts
// the pane — doesn't reset it to Search (#189):
//
//   ┌── Search ──┬── By GUID ──┐
//   ┌──────────────────────────┐
//   │ <typeahead Combobox>     │       ← live results from useLookupRecords
//   │  or                      │
//   │ <Input + inline validate>│       ← raw GUID, no queries fired
//   └──────────────────────────┘
//
// Why a toggle:
//   • In Search mode, every keystroke fires a debounced `contains()` query.
//     If you've already picked a record and want to clear it, the search
//     re-fires on each backspace, which is both wasteful and confusing.
//   • Power users frequently have a GUID on hand (copied from a URL bar in
//     the maker, a Power Automate run, a SQL query). They don't want to
//     name-search at all.
//   • Switching modes preserves the selection — the GUID input mirrors the
//     picked record's id, the search field re-displays the picked record's
//     primary name.
//
// Live data: rides on `useLookupRecords` — same hook FilterValueInput uses
// for lookup-column typeaheads. `contains(<primaryName>, …)`, 250 ms debounce,
// capped at 50 rows, cached per (entity, query).

import {
  createContext,
  useContext,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from 'react';
import {
  Combobox,
  Option,
  Persona,
  Spinner,
  Input,
  RadioGroup,
  Radio,
  tokens,
  Caption1,
  Button,
  Tooltip,
} from '@fluentui/react-components';
import { Dismiss16Regular, Warning20Filled, CheckmarkCircle16Filled } from '@fluentui/react-icons';
import { parseGuid } from './guid';
import { findTable } from '../mock/metadata';
import {
  fetchRecordName,
  useLookupRecords,
  useLookupRecordsVersion,
} from '../host/useLookupRecords';

export interface PickedRecord {
  id: string; // primary key value
  primary: string; // primary name value
}

type PickerMode = 'search' | 'guid';

// ── Per-mode memory (#189) ───────────────────────────────────────────
//
// Sidebar panes unmount when the user switches between them, so local state
// is lost. ModeShell stays mounted for the life of a mode, so it holds this
// map; switching modes drops it, the same as the mode's own state.

interface PickerMemory {
  mode?: PickerMode;
  /** GUID-mode text as typed (may be invalid or not yet resolved). */
  raw?: string;
}

const PickerMemoryContext = createContext<Map<string, PickerMemory> | null>(null);

export function RecordPickerMemoryScope({ children }: { children: ReactNode }) {
  const [memory] = useState(() => new Map<string, PickerMemory>());
  return <PickerMemoryContext.Provider value={memory}>{children}</PickerMemoryContext.Provider>;
}

/** useState that also writes through to the scope's memory when keyed. */
function useRemembered<K extends keyof PickerMemory>(
  memoryKey: string | undefined,
  field: K,
  initial: (stored: PickerMemory[K] | undefined) => NonNullable<PickerMemory[K]>,
) {
  const memory = useContext(PickerMemoryContext);
  const entry = memoryKey && memory ? memory.get(memoryKey) : undefined;
  const [value, setValue] = useState(() => initial(entry?.[field]));
  const set = (next: NonNullable<PickerMemory[K]>) => {
    setValue(next);
    if (memoryKey && memory) memory.set(memoryKey, { ...memory.get(memoryKey), [field]: next });
  };
  return [value, set] as const;
}

// For comparing ids from different sources (picked row vs. saved request).
const normalizeGuid = (s: string): string => parseGuid(s) ?? s.trim().toLowerCase();

export function RecordPicker({
  table,
  selectedId,
  onPick,
  placeholder = 'Search records by name…',
  memoryKey,
}: {
  table: string; // logical name
  selectedId: string | null;
  onPick: (rec: PickedRecord | null) => void;
  placeholder?: string;
  /** Remember the Search / By GUID toggle and GUID text across pane
   *  switches under this key (unique within the mode). */
  memoryKey?: string;
}) {
  const tbl = findTable(table);

  // The mode toggle defaults to 'search'. If a caller hydrates with just
  // a GUID (e.g. a saved request that only has the id, no name), they
  // can hit "By GUID" to see the raw value and edit it.
  const [mode, setMode] = useRemembered(memoryKey, 'mode', (stored) => stored ?? 'search');

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
      <RadioGroup
        layout="horizontal"
        value={mode}
        onChange={(_, d) => setMode(d.value as PickerMode)}
        style={{ gap: 12 }}
      >
        <Radio value="search" label="Search" />
        <Radio value="guid" label="By GUID" />
      </RadioGroup>
      {mode === 'search' ? (
        <SearchPicker
          table={table}
          tbl={tbl}
          selectedId={selectedId}
          onPick={onPick}
          placeholder={placeholder}
        />
      ) : (
        <GuidPicker selectedId={selectedId} onPick={onPick} memoryKey={memoryKey} />
      )}
    </div>
  );
}

// ── Search mode ──────────────────────────────────────────────────────
//
// Pure typeahead. The previous design conflated "typed query" and
// "selected record label" in one input value, which made clearing
// awkward (deleting chars seemed to do nothing because the selected-
// record label fell back into view). This split version keeps the
// fields' semantics tight:
//
//   • `search` is the user's CURRENT input text — what they're typing.
//     Empty string means "show me all matches" (which the lookup hook
//     translates into "no filter, top 50").
//   • Selected state is fully owned by the parent via `selectedId`.
//   • When the parent has a selection AND the user hasn't started typing,
//     we display the resolved primary name as the field value.
//   • An explicit Clear button (small X) drops both the search text and
//     the selection — same affordance as the combobox's `clearable` prop
//     but more discoverable, and side-steps the Fluent v9 quirk where
//     `clearable` doesn't fire a separate clear event the parent can hook.

function SearchPicker({
  table,
  tbl,
  selectedId,
  onPick,
  placeholder,
}: {
  table: string;
  tbl: ReturnType<typeof findTable>;
  selectedId: string | null;
  onPick: (rec: PickedRecord | null) => void;
  placeholder: string;
}) {
  const [search, setSearch] = useState('');
  // True once the user has typed anything since the last selection /
  // clear. While false, the field displays the resolved name of the
  // currently-selected record instead of the search text. This is what
  // lets a saved-request load show "Litware, Inc." in the field instead
  // of an empty string while the typeahead warms up.
  const [userIsTyping, setUserIsTyping] = useState(false);

  // Live lookup. When `search` is empty AND the user hasn't started
  // typing, fire with empty query so the listbox shows the first page
  // of records — useful as a "browse" mode.
  const { rows, loading, error } = useLookupRecords(table || null, userIsTyping ? search : '');

  // Resolved label for the currently-selected record. We keep the last
  // resolved value across re-fetches so the field doesn't flicker when
  // the user is just hovering options.
  const lastResolvedRef = useRef<PickedRecord | null>(null);
  const matchedRow = selectedId
    ? rows.find((r) => normalizeGuid(r.id) === normalizeGuid(selectedId))
    : null;
  useEffect(() => {
    if (matchedRow) {
      lastResolvedRef.current = { id: matchedRow.id, primary: matchedRow.name };
    } else if (!selectedId) {
      lastResolvedRef.current = null;
    }
  }, [matchedRow, selectedId]);

  // After DRS writes records (#81) the selected record may have a new name.
  // The refreshed rows relabel it when it's on the current page; otherwise
  // re-read its name directly so the field doesn't keep the old one.
  const lookupVersion = useLookupRecordsVersion();
  const [, setRelabeled] = useState(0);
  useEffect(() => {
    if (lookupVersion === 0 || !selectedId || !tbl) return;
    let cancelled = false;
    fetchRecordName(tbl, selectedId)
      .then((name) => {
        if (cancelled || name === null) return;
        lastResolvedRef.current = { id: selectedId, primary: name };
        setRelabeled((n) => n + 1);
      })
      .catch(() => {
        // Best effort — e.g. the record was just deleted. Keep the label.
      });
    return () => {
      cancelled = true;
    };
    // Only on invalidation; selection changes resolve through the rows.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lookupVersion]);

  // Reset search state whenever the parent clears the selection externally
  // (e.g. switching tables, loading a saved request that targets a
  // different record). Without this, the stale search string would
  // persist across selection changes.
  useEffect(() => {
    if (!selectedId) {
      setSearch('');
      setUserIsTyping(false);
    }
  }, [selectedId]);

  // Only trust the cached label when it belongs to the current selection —
  // loading a saved request swaps `selectedId` without a pick, and the
  // previous record's name must not stay in the field (#46).
  const resolvedForSelection =
    lastResolvedRef.current &&
    selectedId &&
    normalizeGuid(lastResolvedRef.current.id) === normalizeGuid(selectedId)
      ? lastResolvedRef.current.primary
      : null;
  const selectedLabel = resolvedForSelection ?? (selectedId ? `(${selectedId.slice(0, 8)}…)` : '');
  const displayValue = userIsTyping ? search : selectedLabel;

  const onClear = () => {
    setSearch('');
    setUserIsTyping(false);
    lastResolvedRef.current = null;
    onPick(null);
  };

  return (
    <div style={{ position: 'relative', display: 'flex', gap: 4 }}>
      <Combobox
        placeholder={placeholder}
        value={displayValue}
        selectedOptions={selectedId ? [normalizeGuid(selectedId)] : []}
        onOptionSelect={(_, d) => {
          const rec = rows.find((r) => r.id === d.optionValue);
          if (rec) {
            const picked = { id: rec.id, primary: rec.name };
            lastResolvedRef.current = picked;
            onPick(picked);
            setSearch('');
            setUserIsTyping(false);
          }
        }}
        onChange={(e) => {
          const next = (e.target as HTMLInputElement).value;
          setSearch(next);
          setUserIsTyping(true);
          // Empty input → drop selection. Done here (not on blur) so the
          // URL bar reflects the clear immediately, not on next focus loss.
          if (!next && selectedId) {
            lastResolvedRef.current = null;
            onPick(null);
          }
        }}
        freeform
        style={{ flex: 1 }}
        listbox={{ style: { maxHeight: 360 } }}
      >
        {loading && (
          <Option value="__loading" text="" disabled>
            <span
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: 6,
                color: tokens.colorNeutralForeground3,
              }}
            >
              <Spinner size="tiny" /> Searching…
            </span>
          </Option>
        )}
        {!loading && error && (
          <Option value="__error" text="" disabled>
            <span style={{ color: tokens.colorPaletteRedForeground1, fontSize: 12 }}>{error}</span>
          </Option>
        )}
        {!loading && !error && rows.length === 0 && (
          <Option value="__none" text="" disabled>
            <span style={{ color: tokens.colorNeutralForeground3, fontSize: 12 }}>
              {tbl
                ? search
                  ? `No ${tbl.displayName} records match "${search}"`
                  : `Type to search ${tbl.displayName}`
                : 'Pick a table first'}
            </span>
          </Option>
        )}
        {!loading &&
          !error &&
          rows.map((r) => (
            <Option key={r.id} value={r.id} text={r.name}>
              <Persona
                size="small"
                name={r.name}
                secondaryText={r.id}
                avatar={{ color: 'colorful' }}
              />
            </Option>
          ))}
      </Combobox>
      {selectedId && (
        <Tooltip content="Clear selection" relationship="label">
          <Button
            icon={<Dismiss16Regular />}
            appearance="subtle"
            size="small"
            onClick={onClear}
            aria-label="Clear selected record"
          />
        </Tooltip>
      )}
    </div>
  );
}

// ── GUID mode ────────────────────────────────────────────────────────
//
// Plain Input. Validates inline. Emits onPick(null) when invalid OR empty,
// so the mode never keeps targeting the previously picked record while the
// field shows a different (invalid) value (#94). Valid input is emitted in
// the dashed 8-4-4-4-12 form the OData key needs — braces and the 32-digit
// dash-less form are accepted and normalized. No metadata queries fire —
// this mode is for users who already have an id and just want to paste it.

function GuidPicker({
  selectedId,
  onPick,
  memoryKey,
}: {
  selectedId: string | null;
  onPick: (rec: PickedRecord | null) => void;
  memoryKey?: string;
}) {
  // Local string so the user can type freely; only a parsed GUID (or
  // null) goes upstream. Remembered text is reused only while it still
  // agrees with the selection (an invalid entry ↔ no selection); if the
  // selection changed while the pane was closed, start from the selection.
  const [raw, setRaw] = useRemembered(memoryKey, 'raw', (stored) =>
    stored !== undefined &&
    parseGuid(stored) === (selectedId ? (parseGuid(selectedId) ?? selectedId) : null)
      ? stored
      : (selectedId ?? ''),
  );
  // The id we last sent up. When the parent echoes it back we keep the
  // user's text as typed — otherwise emitting null for an invalid entry
  // would wipe the field mid-typing.
  const lastEmittedRef = useRef<string | null>(selectedId);
  // If the parent's selectedId changes from elsewhere (mode switch,
  // saved-request load, table switch), sync our local state.
  useEffect(() => {
    if (selectedId === lastEmittedRef.current) return;
    lastEmittedRef.current = selectedId;
    setRaw(selectedId ?? '');
  }, [selectedId]);

  const normalized = useMemo(() => parseGuid(raw), [raw]);
  const valid = !raw.trim() || normalized !== null;

  const onChange = (next: string) => {
    setRaw(next);
    const id = parseGuid(next);
    if (id === lastEmittedRef.current) return;
    lastEmittedRef.current = id;
    onPick(id ? { id, primary: '' } : null);
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
      <Input
        value={raw}
        onChange={(_, d) => onChange(d.value)}
        placeholder="00000000-0000-0000-0000-000000000000"
        contentAfter={
          !raw ? null : !valid ? (
            <Warning20Filled style={{ color: tokens.colorPaletteRedForeground1 }} />
          ) : (
            <CheckmarkCircle16Filled style={{ color: tokens.colorPaletteGreenForeground1 }} />
          )
        }
        style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 12 }}
      />
      {raw && !valid && (
        <Caption1 style={{ color: tokens.colorPaletteRedForeground1 }}>
          Not a valid GUID — use the 8-4-4-4-12 form (braces optional) or 32 hex digits. No record
          is targeted until this is fixed.
        </Caption1>
      )}
      {raw && valid && normalized && (
        <Caption1
          style={{ color: tokens.colorNeutralForeground3, fontFamily: tokens.fontFamilyMonospace }}
        >
          {normalized}
        </Caption1>
      )}
    </div>
  );
}
