// RecordPicker — live typeahead OR direct-GUID-entry for a single record.
//
// Used by:
//   • TargetEditor (Retrieve Single, Update, Delete, Upsert, Merge — anywhere
//     the request needs a primary key).
//   • Standalone in modes that pick a relationship target (Associate /
//     Disassociate).
//
// Two modes, toggled by the user (and remembered locally):
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

import { useEffect, useMemo, useRef, useState } from 'react';
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
import { findTable } from '../mock/metadata';
import { useLookupRecords } from '../host/useLookupRecords';

export interface PickedRecord {
  id: string; // primary key value
  primary: string; // primary name value
}

type PickerMode = 'search' | 'guid';

const GUID_RE =
  /^[{(]?[0-9a-fA-F]{8}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{12}[)}]?$/;
const isValidGuid = (s: string): boolean => GUID_RE.test(s.trim());
const normalizeGuid = (s: string): string =>
  s
    .replace(/[{}()]/g, '')
    .trim()
    .toLowerCase();

export function RecordPicker({
  table,
  selectedId,
  onPick,
  placeholder = 'Search records by name…',
}: {
  table: string; // logical name
  selectedId: string | null;
  onPick: (rec: PickedRecord | null) => void;
  placeholder?: string;
}) {
  const tbl = findTable(table);

  // The mode toggle defaults to 'search'. If a caller hydrates with just
  // a GUID (e.g. a saved request that only has the id, no name), they
  // can hit "By GUID" to see the raw value and edit it.
  const [mode, setMode] = useState<PickerMode>('search');

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
        <GuidPicker selectedId={selectedId} onPick={onPick} />
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

  const selectedLabel =
    lastResolvedRef.current?.primary ?? (selectedId ? `(${selectedId.slice(0, 8)}…)` : '');
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
// Plain Input. Validates inline. Emits onPick(null) when invalid OR empty.
// No metadata queries fire — this mode is for users who already have an
// id and just want to paste it.

function GuidPicker({
  selectedId,
  onPick,
}: {
  selectedId: string | null;
  onPick: (rec: PickedRecord | null) => void;
}) {
  // Local string so the user can type freely; we only commit upstream
  // when the value parses as a GUID (or clears).
  const [raw, setRaw] = useState(selectedId ?? '');
  // If the parent's selectedId changes from elsewhere (mode switch,
  // saved-request load), sync our local state.
  useEffect(() => {
    setRaw(selectedId ?? '');
  }, [selectedId]);

  const valid = !raw || isValidGuid(raw);
  const normalized = useMemo(() => (valid && raw ? normalizeGuid(raw) : null), [raw, valid]);

  const onChange = (next: string) => {
    setRaw(next);
    if (!next) {
      onPick(null);
      return;
    }
    if (isValidGuid(next)) {
      onPick({ id: normalizeGuid(next), primary: '' });
    }
    // Invalid intermediate state — don't propagate yet. The mode shell
    // surfaces an advisory via the regular validation path.
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
          Not a valid GUID — Dataverse record ids are 36-char UUIDs (braces optional).
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
