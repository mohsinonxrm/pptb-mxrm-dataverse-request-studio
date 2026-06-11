// Update Record — PATCH /<entitySet>(<id>) (multi-field) or PUT
// /<entitySet>(<id>)/<column> (single column).
//
// Sidebar:
//   * Target         → RecordPick
//   * Diff           [3 changed]     ← primary view, NOT Field set
//   * Field set      (collapsed)
//   * Prefer
//   * Headers
//
//   "Update single column": When user drills into a single column from the
//    Field set OR the Diff, the URL bar method pill switches from PATCH to
//    PUT and the path becomes /accounts(id)/name. This is the only mode
//    where the method changes based on selection.
//
// Grounded in:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/update-delete-entities-using-web-api

import { useEffect, useMemo, useState } from 'react';
import {
  Table20Regular,
  Table20Filled,
  BranchFork20Regular,
  BranchFork20Filled,
  FormNew20Regular,
  FormNew20Filled,
  TextBulletList20Regular,
  TextBulletList20Filled,
  Settings20Regular,
  Settings20Filled,
  LineHorizontal320Regular,
  LineHorizontal320Filled,
  ShieldLock20Regular,
  ShieldLock20Filled,
  ArrowSwap20Regular,
  ArrowSwap20Filled,
  Checkmark16Filled,
  Warning20Filled,
} from '@fluentui/react-icons';
import {
  Caption1,
  Combobox,
  Option,
  tokens,
  Field,
  MessageBar,
  MessageBarBody,
  Button,
} from '@fluentui/react-components';
import { Sidebar } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { PaneHead } from '../editors/PaneHead';
import { TargetEditor } from '../editors/TargetEditor';
import { SelectEditor } from '../editors/SelectEditor';
import { FieldSetEditor } from '../editors/FieldSetEditor';
import { PreferEditor, emptyPrefer, preferToHeaderString } from '../editors/PreferEditor';
import { HeadersEditor, defaultWriteHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { BypassEditor, summarize as summarizeBypass } from '../editors/BypassEditor';
import { applyBypassToHeaders } from '../engine/bypassHeaders';
import { detectBypassAdvisories } from '../engine/bypassAdvisories';
import { disabledReasonFromAdvisories, type Advisory } from '../primitives/advisories';
import { PreconditionEditor } from '../editors/PreconditionEditor';
import { UpdateDiffPane } from '../editors/UpdateDiffPane';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { findTable, findColumn } from '../mock/metadata';
import { findRequestType } from '../registry/requestTypes';
import { buildUpdate, buildUpdateBody } from '../engine/urlBuilder';
import { runtime, type ExecResult } from '../engine/runtime';
import {
  defaultBypassOptions,
  type UpdateState,
  type CreateFieldValue,
  type UpdateMethod,
} from '../state/writeState';
import type { RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useLiveTable } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import {
  serializeUpdate,
  deserializeUpdate,
  hashState,
  type SavedRequest,
  type SerializedUpdateState,
} from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';

// ──────────────────────────────────────────────────────────────
// Initial state
// ──────────────────────────────────────────────────────────────
const initialState = (): UpdateState => ({
  table: '',
  recordId: null,
  method: 'PATCH',
  putColumn: null,
  fieldValues: {},
  nullFields: [],
  prefer: { ...emptyPrefer(), formattedValues: true, returnRepresentation: false },
  headers: defaultWriteHeaders(),
  returnSelect: [],
  concurrency: { kind: 'update-only' }, // If-Match: * default
  bypass: defaultBypassOptions(),
  dirty: new Set(),
});

// GUID shape — match what RecordPicker accepts. Invalid GUIDs short-circuit
// the live fetch (no point sending a malformed id over the wire) and
// surface an advisory instead.
const GUID_RE =
  /^[{(]?[0-9a-fA-F]{8}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{12}[)}]?$/;
const isValidGuid = (s: string | null | undefined): boolean => !!s && GUID_RE.test(s.trim());

type RootClauseId =
  | 'target'
  | 'method'
  | 'diff'
  | 'fieldset'
  | 'precondition'
  | 'prefer'
  | 'returnselect'
  | 'headers'
  | 'bypass';

export function UpdateMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('update');
  const [state, setState] = useState(initialState);
  // Live-metadata subscription so child editors re-render when columns/relationships land.
  // We also need the loading flag so the Diff pane doesn't flash "Couldn't load"
  // during the brief window between (a) picker firing → recordId set, and
  // (b) tbl metadata resolving from the live registry. Before this, the
  // original-row fetch effect would early-return on the missing tbl, leaving
  // originalLoading=false + originalRow=null → diff renders the error state.
  const { loading: tableLoading } = useLiveTable(state.table || null);
  const { entities } = useScopedEntities();
  const [activePath, setActivePath] = useState<string>('target');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);

  // Last-saved tracking → Save button dirty/clean state, library checkmark.
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);
  // Primary name from the RecordPicker — used as a fallback label until
  // the live row fetch completes (Search-mode picks deliver this for free).
  const [pickedPrimaryFromPicker, setPickedPrimaryFromPicker] = useState<string>('');

  const built = useMemo(() => buildUpdate(state), [state]);
  const body = useMemo(() => buildUpdateBody(state), [state]);
  const tbl = findTable(state.table);

  const markDirty = (id: string) =>
    setState((s) => {
      const d = new Set(s.dirty);
      d.add(id);
      return { ...s, dirty: d };
    });
  const set = <K extends keyof UpdateState>(k: K, v: UpdateState[K], dirtyId?: string) => {
    setState((s) => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  // ── Live original row + ETag ──
  //
  // Diff pane needs the record's CURRENT persisted values to compare against
  // the user's edits. The If-Match card needs the real ETag so "Use current
  // ETag" can populate the header for optimistic concurrency.
  //
  // ONE round-trip per (table, recordId) change. We fetch with NO $select
  // so the response includes every editable column — the user can diff/edit
  // any of them without forcing a re-fetch. Response carries `@odata.etag`
  // as an annotation; we extract it directly. For very wide entities this
  // could be a noticeable payload, but it's still one shot per pick vs.
  // N shots per "first time you touch column X".
  const [originalRow, setOriginalRow] = useState<Record<string, unknown> | null>(null);
  const [originalLoading, setOriginalLoading] = useState(false);
  const [originalError, setOriginalError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    setOriginalError(null);
    if (!state.recordId || !state.table || !tbl) {
      setOriginalRow(null);
      return;
    }
    if (!isValidGuid(state.recordId)) {
      setOriginalRow(null);
      return;
    }
    if (typeof window === 'undefined' || !window.dataverseAPI?.queryData) {
      setOriginalRow(null);
      return;
    }
    setOriginalLoading(true);
    // Use the same leading-slash URL form as executeRetrieveSingle — PPTB's
    // queryData is reliable with this form for single-record fetches.
    const url = `/${tbl.entitySetName}(${state.recordId})`;
    window.dataverseAPI
      .queryData(url)
      .then((res) => {
        if (cancelled) return;
        // PPTB returns single-record responses in one of THREE shapes
        // depending on the host build + whether $select was applied:
        //   1. `{ value: <recordObject> }`             — wrapped, object
        //   2. `{ value: [<recordObject>] }`           — wrapped, single-elt array
        //   3. `<recordObject>` (no wrapper)            — raw, common when no $select
        // The signature says (1) but empirically PPTB without $select sometimes
        // returns shape (3). Defend against all three.
        const wrappedValue = (res as { value?: unknown } | null)?.value;
        let row: Record<string, unknown> | null = null;
        if (Array.isArray(wrappedValue)) {
          row = (wrappedValue[0] as Record<string, unknown>) ?? null;
        } else if (wrappedValue && typeof wrappedValue === 'object') {
          row = wrappedValue as Record<string, unknown>;
        } else if (res && typeof res === 'object') {
          // No `value` wrapper. If the response itself has the entity's
          // primary key OR an `@odata.*` annotation, treat it as the record.
          const direct = res as Record<string, unknown>;
          const looksLikeRecord =
            tbl.primaryKey in direct || Object.keys(direct).some((k) => k.startsWith('@odata.'));
          if (looksLikeRecord) row = direct;
        }
        setOriginalRow(row);
        setOriginalLoading(false);
      })
      .catch((e: unknown) => {
        if (cancelled) return;
        const msg = e instanceof Error ? e.message : String(e);
        setOriginalRow(null);
        setOriginalLoading(false);
        setOriginalError(msg);
      });
    return () => {
      cancelled = true;
    };
  }, [state.recordId, state.table, tbl]);

  // Pull the real ETag out of the fetched row (Dataverse responses carry it
  // as `@odata.etag`). When we don't have a row yet, currentEtag stays null —
  // the If-Match card disables its "Use current ETag" button accordingly.
  const currentEtag = useMemo<string | null>(() => {
    if (!originalRow) return null;
    const raw = originalRow['@odata.etag'];
    return typeof raw === 'string' && raw ? raw : null;
  }, [originalRow]);

  // ── Drill into a single column → PUT mode ──
  const onDrillColumn = (column: string) => {
    setState((s) => ({ ...s, method: 'PUT', putColumn: column }));
    setActivePath('method');
    markDirty('method');
  };
  const exitPutMode = () => {
    setState((s) => ({ ...s, method: 'PATCH', putColumn: null }));
    markDirty('method');
  };

  // ── Save / Load ──
  const currentSerialized = useMemo(() => serializeUpdate(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;

  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };

  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== 'update') return;
    const snap = entry.state as SerializedUpdateState;
    if (entities.length > 0 && !entities.some((e) => e.logicalName === snap.table)) {
      window.alert(
        `Can't load "${entry.name}": entity \`${snap.table}\` ` +
          `isn't available in this environment. The solution may have been ` +
          `removed or you may be connected to a different org.`,
      );
      return;
    }
    setState(deserializeUpdate(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setActivePath('target');
  };

  // Publish save context. Hidden until the user has a table picked —
  // no point persisting an empty shell.
  usePublishSaveContext(
    useMemo(() => {
      if (!state.table) return null;
      return {
        state: currentSerialized,
        modeId: 'update' as const,
        dirty: isDirty,
        lastSavedId,
        onSaved,
        onLoadSaved,
      };
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [currentSerialized, isDirty, lastSavedId, state.table]),
  );

  // ── Execute guards ──
  const fieldCount = Object.keys(state.fieldValues).length;
  const putValueOk =
    state.method !== 'PUT' ||
    (state.putColumn != null && state.fieldValues[state.putColumn] != null);
  const disabledReason = !tbl
    ? 'Pick a target table first.'
    : !state.recordId
      ? 'Pick a record to update.'
      : !isValidGuid(state.recordId)
        ? 'Record id is not a valid GUID.'
        : state.method === 'PUT'
          ? !state.putColumn
            ? 'Pick a column for the PUT.'
            : !putValueOk
              ? `Set a value for ${state.putColumn}.`
              : null
          : (fieldCount === 0
              ? 'No field changes — nothing to PATCH. Pick at least one field to update.'
              : Object.keys(body).length === 0
                ? 'Body is empty — set at least one value.'
                : null) ||
            (state.concurrency.kind === 'etag' && !state.concurrency.etag.trim()
              ? 'Provide an etag value or switch the concurrency mode.'
              : state.headers.some((h) => h.enabled && !h.name)
                ? 'Fix empty header name.'
                : null);

  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.update(state);
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents((rs) =>
      [
        {
          id: `r-${Date.now()}`,
          modeId: 'update',
          url: built.relativeUrl,
          method: state.method,
          ts: Date.now(),
          status: res.status,
          ms: res.ms,
          rowCount: res.ok ? 1 : 0,
        },
        ...rs,
      ].slice(0, 8),
    );
    setState((s) => ({ ...s, dirty: new Set() }));
  };

  // ── Compose effective headers ──
  const effectiveHeaders = useMemo(() => {
    let h = [...state.headers];
    // Mirror the IfMatch card editor's choice onto a header for execute time
    const m = state.concurrency;
    if (m.kind === 'update-only') {
      h.push({
        id: '__cc__If-Match-any',
        name: 'If-Match',
        value: '*',
        enabled: true,
        builtin: true,
        hint: 'Auto-composed from the If-Match pane.',
      });
    } else if (m.kind === 'etag' && m.etag.trim()) {
      h.push({
        id: '__cc__If-Match-etag',
        name: 'If-Match',
        value: m.etag,
        enabled: true,
        builtin: true,
        hint: 'Auto-composed from the If-Match pane.',
      });
    }
    h = applyBypassToHeaders(h, state.bypass);
    return h;
  }, [state.headers, state.concurrency, state.bypass]);

  // ── Advisories — bypass family + PUT-method unsupported + GUID shape ──
  //
  // Same family as Delete's single-property case: PPTB's dataverseAPI.update
  // is PATCH-only. There's no method to issue `PUT /<entityset>(<id>)/<col>`.
  // We still let the user author the request (correct URL + body in the Code
  // tab) but block Execute and explain why.
  const advisories = useMemo<Advisory[]>(() => {
    const out: Advisory[] = [...detectBypassAdvisories(state.bypass, 'bypass')];

    if (state.table && state.recordId && !isValidGuid(state.recordId)) {
      out.push({
        id: 'upd-bad-guid',
        severity: 'error',
        source: 'validation',
        focusNode: 'target',
        title: 'Invalid record id',
        body: `"${state.recordId}" isn't a valid GUID. The update will fail at the wire.`,
      });
    }

    if (state.method === 'PUT') {
      out.push({
        id: 'upd-put-unsupported',
        severity: 'error',
        source: 'validation',
        focusNode: 'method',
        title: "PUT single-column isn't supported by the PPTB host",
        body: (
          <>
            DRS authors the correct request{' '}
            <code>
              PUT /{tbl?.entitySetName ?? '&lt;set&gt;'}(&lt;id&gt;)/
              {state.putColumn ?? '&lt;column&gt;'}
            </code>{' '}
            with body <code>{'{ "value": <scalar> }'}</code>, but PPTB&apos;s{' '}
            <code>dataverseAPI.update</code> is PATCH-only — there&apos;s no raw-request hook for
            property-path URLs. To set one column from inside PPTB, switch back to{' '}
            <strong>PATCH</strong> with a single field in the body — it&apos;s functionally
            equivalent. To use the PUT pattern, copy the URL + body from the Code tab and run it
            from Postman / curl / the JS SDK / Power Automate.
          </>
        ),
      });
    }

    return out;
  }, [state.bypass, state.table, state.recordId, state.method, state.putColumn, tbl]);
  const bypassBlocker = disabledReasonFromAdvisories(advisories);

  // ── Sidebar ──
  const returnRep = state.prefer.returnRepresentation;
  const isPut = state.method === 'PUT';
  const methodBadge = isPut
    ? `PUT /${state.putColumn ?? '?'}`
    : `PATCH · ${fieldCount} field${fieldCount === 1 ? '' : 's'}`;
  const ifMatchBadge =
    state.concurrency.kind === 'update-only'
      ? '* (any)'
      : state.concurrency.kind === 'etag' &&
          currentEtag &&
          state.concurrency.etag.includes(currentEtag)
        ? 'current etag'
        : state.concurrency.kind === 'etag'
          ? 'custom etag'
          : 'none';

  const sections = [
    {
      id: 'target',
      label: 'Target',
      meta: tbl
        ? originalError
          ? 'fetch failed — see Diff'
          : `${tbl.displayName} record`
        : 'Pick a table',
      items: [
        {
          id: 'target',
          icon: Table20Regular,
          iconFilled: Table20Filled,
          // Source-of-truth chain for the sidebar label:
          //   1. Picker primary (free, immediate after pick)
          //   2. Live row's primaryName (after fetch resolves)
          //   3. Fallback hint
          label:
            pickedPrimaryFromPicker ||
            (originalRow && tbl ? String(originalRow[tbl.primaryName] ?? '') : '') ||
            (state.recordId ? '(loading record…)' : 'Pick a record'),
          dirty: state.dirty.has('target'),
        },
      ],
    },
    {
      id: 'method',
      label: 'Method',
      meta: methodBadge,
      items: [
        {
          id: 'method',
          icon: ArrowSwap20Regular,
          iconFilled: ArrowSwap20Filled,
          label: isPut ? 'PUT — single column' : 'PATCH — multi-field',
          badge: methodBadge,
          badgeAppearance: 'ghost' as const,
          dirty: state.dirty.has('method'),
        },
      ],
    },
    {
      // Field set goes FIRST (you have to set the values before there's
      // anything to diff against). The natural flow is set → diff, not
      // diff → set; Diff stays empty until Field set is touched.
      id: 'body',
      label: 'Field set',
      meta: isPut ? 'single column' : `${fieldCount} changed`,
      items: [
        {
          id: 'fieldset',
          icon: FormNew20Regular,
          iconFilled: FormNew20Filled,
          label: isPut ? `Value (${state.putColumn ?? '?'})` : 'Field set',
          badge: isPut ? (putValueOk ? '✓' : '?') : fieldCount || null,
          badgeAppearance: 'ghost' as const,
          dirty: state.dirty.has('fieldset'),
        },
        // Diff is collapsed under Field set — surfaces what's about to
        // be PATCHed once the user has actually set values.
        ...(!isPut
          ? [
              {
                id: 'diff',
                icon: BranchFork20Regular,
                iconFilled: BranchFork20Filled,
                label: 'Diff',
                badge: fieldCount || null,
                badgeAppearance: 'tint' as const,
                badgeColor: fieldCount > 0 ? ('brand' as const) : ('subtle' as const),
                dirty: state.dirty.has('diff'),
              },
            ]
          : []),
      ],
    },
    {
      id: 'precondition',
      label: 'Precondition',
      meta: ifMatchBadge,
      items: [
        {
          id: 'precondition',
          icon: ShieldLock20Regular,
          iconFilled: ShieldLock20Filled,
          label: 'Optimistic concurrency',
          badge: ifMatchBadge,
          badgeAppearance: 'ghost' as const,
          dirty: state.dirty.has('precondition'),
        },
      ],
    },
    ...(!isPut
      ? [
          {
            id: 'prefer',
            label: 'Prefer',
            meta: returnRep ? 'return=representation' : '204 No Content',
            items: [
              {
                id: 'prefer',
                icon: Settings20Regular,
                iconFilled: Settings20Filled,
                label: 'Prefer header',
                badge: preferToHeaderString(state.prefer) ? 'on' : null,
                badgeAppearance: 'ghost' as const,
                dirty: state.dirty.has('prefer'),
              },
              ...(returnRep
                ? [
                    {
                      id: 'returnselect',
                      icon: TextBulletList20Regular,
                      iconFilled: TextBulletList20Filled,
                      label: '$select (response)',
                      code: true,
                      badge: state.returnSelect.length || null,
                      dirty: state.dirty.has('returnselect'),
                    },
                  ]
                : []),
            ],
          },
        ]
      : []),
    {
      id: 'headers',
      label: 'Headers',
      meta: `${state.headers.filter((h) => h.enabled).length} active`,
      items: [
        {
          id: 'headers',
          icon: LineHorizontal320Regular,
          iconFilled: LineHorizontal320Filled,
          label: 'HTTP headers',
          badge: state.headers.filter((h) => h.enabled).length || null,
          dirty: state.dirty.has('headers'),
        },
        {
          id: 'bypass',
          icon: Warning20Filled,
          iconFilled: Warning20Filled,
          label: 'Bypass logic',
          badge: summarizeBypass(state.bypass),
          badgeAppearance:
            state.bypass.businessLogic !== 'none' || state.bypass.suppressFlows
              ? ('tint' as const)
              : ('ghost' as const),
          badgeColor:
            state.bypass.businessLogic !== 'none' || state.bypass.suppressFlows
              ? ('warning' as const)
              : ('subtle' as const),
          dirty: state.dirty.has('bypass'),
        },
      ],
    },
  ];

  // ── Builder pane router ──
  let pane: React.ReactNode;
  const root = activePath as RootClauseId;
  switch (root) {
    case 'target':
      pane = (
        <TargetEditor
          table={state.table}
          onTableChange={(t) => {
            // Switching/clearing the target invalidates every column-bound
            // clause for Update: recordId (not valid on new entity), body
            // fields, the PUT column choice, and the response $select.
            setState((s) => ({
              ...s,
              table: t,
              recordId: null,
              fieldValues: {},
              method: 'PATCH',
              putColumn: null,
              returnSelect: [],
              dirty: new Set(['target', 'fieldset', 'method']),
            }));
            setResult(null);
            setPickedPrimaryFromPicker('');
          }}
          recordId={state.recordId}
          onRecordChange={(id, primary) => {
            set('recordId', id, 'target');
            setPickedPrimaryFromPicker(primary ?? '');
          }}
          group="write"
          sub="Single-record update — pick the entity set and the record GUID."
        />
      );
      break;
    case 'method':
      pane = (
        <MethodPane
          method={state.method}
          putColumn={state.putColumn}
          setMethod={(m) => set('method', m, 'method')}
          setPutColumn={(c) => set('putColumn', c, 'method')}
          table={state.table}
          onExitPut={exitPutMode}
        />
      );
      break;
    case 'diff':
      pane = (
        <UpdateDiffPane
          table={state.table}
          recordId={state.recordId}
          fieldValues={state.fieldValues}
          original={originalRow}
          // Combine both load signals — the diff should show "loading"
          // when EITHER the table metadata is still resolving OR the
          // row fetch is in flight. Without the `tableLoading` half,
          // the diff briefly renders the "couldn't load" error in the
          // gap between recordId being set and tbl becoming available.
          originalLoading={originalLoading || tableLoading}
          onDrillColumn={onDrillColumn}
        />
      );
      break;
    case 'fieldset':
      pane = isPut ? (
        <SingleColumnEditor
          table={state.table}
          column={state.putColumn}
          value={state.putColumn ? state.fieldValues[state.putColumn] : undefined}
          onChange={(v) => {
            if (!state.putColumn) return;
            set('fieldValues', { ...state.fieldValues, [state.putColumn]: v }, 'fieldset');
          }}
          onSwapMethod={exitPutMode}
          themeMode={themeMode}
        />
      ) : (
        <FieldSetEditor
          table={state.table}
          values={state.fieldValues}
          setValues={(next) =>
            set('fieldValues', next as Record<string, CreateFieldValue>, 'fieldset')
          }
          nullFields={state.nullFields}
          setNullFields={(n) => set('nullFields', n, 'fieldset')}
          group="write"
          themeMode={themeMode}
          purpose="update"
        />
      );
      break;
    case 'precondition':
      pane = (
        <PreconditionEditor
          mode={state.concurrency}
          setMode={(m) => set('concurrency', m, 'precondition')}
          // Update supports None / require-existing / etag. No create-only
          // (that's an Upsert pattern — "create the row if absent").
          available={['none', 'update-only', 'etag']}
          currentEtag={currentEtag}
          group="write"
        />
      );
      break;
    case 'prefer':
      pane = (
        <PreferEditor
          spec={state.prefer}
          setSpec={(p) => set('prefer', p, 'prefer')}
          group="write"
        />
      );
      break;
    case 'returnselect':
      pane = (
        <SelectEditor
          table={state.table}
          selectedIds={state.returnSelect}
          setSelectedIds={(ids) => set('returnSelect', ids, 'returnselect')}
          group="write"
        />
      );
      break;
    case 'headers':
      pane = (
        <HeadersEditor
          items={state.headers}
          setItems={(h) => set('headers', h, 'headers')}
          group="write"
        />
      );
      break;
    case 'bypass':
      pane = (
        <BypassEditor
          value={state.bypass}
          onChange={(b) => set('bypass', b, 'bypass')}
          group="write"
        />
      );
      break;
  }

  const headersMap = headerItemsToObject(effectiveHeaders, preferToHeaderString(state.prefer));
  const codeInputs = {
    method: state.method,
    built,
    headers: headersMap,
    body,
    entityLogical: built.entityLogical,
  };

  return (
    <ModeShell
      sidebar={
        <Sidebar
          type={type}
          urlPreview={built.relativeNoBase}
          sections={sections}
          activeNode={activePath}
          onSelect={(id) => {
            setActivePath(id);
            setTab('builder');
          }}
          recents={recents}
        />
      }
      urlBar={
        <UrlBar
          method={state.method}
          url={built.relativeUrl}
          executeVerb={type.executeVerb}
          disabledReason={disabledReason ?? bypassBlocker}
          loading={loading}
          onExecute={onExecute}
          advisories={advisories}
          onAdvisoryFocus={setActivePath}
        />
      }
    >
      <MainTabs tab={tab} onTabChange={setTab} resultCount={result?.ok ? 1 : null}>
        {tab === 'builder' && pane}
        {tab === 'code' && <CodeView themeMode={themeMode} inputs={codeInputs} />}
        {tab === 'results' && (
          <ResultsView
            result={result}
            mode="single"
            table={state.table}
            writeContext={{
              operation: 'update',
              table: state.table,
              recordId: state.recordId,
              recordName:
                pickedPrimaryFromPicker ||
                (originalRow && tbl ? String(originalRow[tbl.primaryName] ?? '') : null),
            }}
          />
        )}
      </MainTabs>
    </ModeShell>
  );
}

// ──────────────────────────────────────────────────────────────
// Method pane — PATCH multi-field vs PUT single-column radio cards
// ──────────────────────────────────────────────────────────────
function MethodPane({
  method,
  putColumn,
  setMethod,
  setPutColumn,
  table,
  onExitPut,
}: {
  method: UpdateMethod;
  putColumn: string | null;
  setMethod: (m: UpdateMethod) => void;
  setPutColumn: (c: string | null) => void;
  table: string;
  onExitPut: () => void;
}) {
  const tbl = findTable(table);
  const pickable = useMemo(
    () =>
      (tbl?.columns ?? [])
        .filter(
          (c) =>
            c.attributeType !== 'Uniqueidentifier' &&
            c.attributeType !== 'File' &&
            c.attributeType !== 'Image' &&
            // Lookups via PUT use $ref form (Associate) — not the property URL path
            c.attributeType !== 'Lookup' &&
            c.attributeType !== 'Customer' &&
            c.attributeType !== 'Owner' &&
            // Honor IsValidForUpdate — system-managed audit columns, formula/
            // rollup outputs, etc. can't be PUT either. `undefined` treated as
            // permissive in case the host hasn't returned the flag.
            c.isValidForUpdate !== false,
        )
        .slice()
        .sort((a, b) => a.displayName.localeCompare(b.displayName)),
    [tbl],
  );
  // Same freeform + clearable Combobox pattern as Delete's ScopeEditor —
  // a vanilla Combobox is unusable on tables with 100+ columns; without
  // freeform the user can't type to filter, and without clearable they
  // can't backspace through the field to start over.
  const [search, setSearch] = useState('');
  const [userIsTyping, setUserIsTyping] = useState(false);

  // Reset typed search whenever the parent clears or changes the column.
  useEffect(() => {
    if (!putColumn) {
      setSearch('');
      setUserIsTyping(false);
    }
  }, [putColumn]);

  const selectedColumn = putColumn ? pickable.find((c) => c.logicalName === putColumn) : undefined;
  const selectedLabel = selectedColumn?.displayName ?? putColumn ?? '';

  const filtered = useMemo(() => {
    if (!search) return pickable;
    const q = search.toLowerCase();
    return pickable.filter(
      (c) =>
        c.displayName.toLowerCase().includes(q) ||
        c.logicalName.toLowerCase().includes(q) ||
        c.attributeType.toLowerCase().includes(q),
    );
  }, [pickable, search]);

  const displayValue = userIsTyping ? search : selectedLabel;

  return (
    <div>
      <PaneHead
        icon={ArrowSwap20Filled}
        title="Method"
        sub="PATCH for multi-field updates, PUT for single-column writes via the property URL."
        group="write"
      />

      <div
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(2, 1fr)',
          gap: 12,
          maxWidth: 760,
        }}
      >
        <MethodCard
          title="PATCH — multi-field"
          description="Send a JSON body of changed fields. Body shape: { col1: val1, col2: val2, ... }"
          syntax={`/{entitySet}({id})`}
          selected={method === 'PATCH'}
          onClick={() => {
            setMethod('PATCH');
            onExitPut();
          }}
        />
        <MethodCard
          title="PUT — single column"
          description='Set ONE column via the property URL. Body: { "value": <scalar> }. Lookups use $ref (Associate) instead.'
          syntax={`/{entitySet}({id})/{column}`}
          selected={method === 'PUT'}
          onClick={() => setMethod('PUT')}
        />
      </div>

      {method === 'PUT' && (
        <div style={{ marginTop: 18, maxWidth: 480 }}>
          <Field
            label="Column to PUT"
            hint={`${pickable.length} eligible column${pickable.length === 1 ? '' : 's'}. Type to filter. Lookups aren't supported via this path — use Associate (POST /$ref) instead.`}
          >
            <Combobox
              freeform
              clearable
              value={displayValue}
              selectedOptions={putColumn ? [putColumn] : []}
              onChange={(e) => {
                const next = (e.target as HTMLInputElement).value;
                setSearch(next);
                setUserIsTyping(true);
                // Empty input → drop the selection (matches RecordPicker /
                // Delete-ScopeEditor semantics).
                if (!next && putColumn) {
                  setPutColumn(null);
                }
              }}
              onOptionSelect={(_, d) => {
                if (!d.optionValue) {
                  setPutColumn(null);
                  setSearch('');
                  setUserIsTyping(false);
                  return;
                }
                setPutColumn(d.optionValue);
                setSearch('');
                setUserIsTyping(false);
              }}
              placeholder="Type to search columns…"
              listbox={{ style: { maxHeight: 360 } }}
            >
              {filtered.length === 0 && (
                <Option value="__none" text="" disabled>
                  <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                    No columns match &ldquo;{search}&rdquo;
                  </Caption1>
                </Option>
              )}
              {filtered.map((c) => (
                <Option key={c.logicalName} value={c.logicalName} text={c.displayName}>
                  <div style={{ display: 'flex', flexDirection: 'column' }}>
                    <span>{c.displayName}</span>
                    <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                      {c.logicalName} ·{' '}
                      <span style={{ color: tokens.colorBrandForeground2 }}>{c.attributeType}</span>
                    </Caption1>
                  </div>
                </Option>
              ))}
            </Combobox>
          </Field>
          <MessageBar layout="multiline" intent="info" style={{ marginTop: 14 }}>
            <MessageBarBody>
              In PUT mode, the URL changes to{' '}
              <code style={{ fontFamily: tokens.fontFamilyMonospace }}>
                /{tbl?.entitySetName ?? '<set>'}(&lt;id&gt;)/{putColumn ?? '<col>'}
              </code>{' '}
              and the body is{' '}
              <code style={{ fontFamily: tokens.fontFamilyMonospace }}>
                {'{ "value": <scalar> }'}
              </code>
              . Switch to <strong>Field set</strong> in the sidebar to edit the value.
            </MessageBarBody>
          </MessageBar>
        </div>
      )}
    </div>
  );
}

function MethodCard({
  title,
  description,
  syntax,
  selected,
  onClick,
}: {
  title: string;
  description: string;
  syntax: string;
  selected: boolean;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      style={{
        textAlign: 'left',
        padding: 14,
        cursor: 'pointer',
        background: selected ? tokens.colorBrandBackground2 : tokens.colorNeutralBackground1,
        border: `1px solid ${selected ? tokens.colorBrandStroke1 : tokens.colorNeutralStroke2}`,
        borderRadius: tokens.borderRadiusMedium,
        outline: 'none',
        display: 'flex',
        flexDirection: 'column',
        gap: 8,
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
        <strong style={{ fontSize: 13 }}>{title}</strong>
        {/* Use an iconic checkmark, not a "SELECTED" caps label. The
            bg/border already encode the state; a tiny 10px text label is
            redundant chrome and below the legibility floor. */}
        {selected && (
          <Checkmark16Filled
            style={{ marginLeft: 'auto', color: tokens.colorBrandForeground1 }}
            aria-label="Selected"
          />
        )}
      </div>
      <Caption1 style={{ color: tokens.colorNeutralForeground2, lineHeight: 1.4 }}>
        {description}
      </Caption1>
      <code
        style={{
          marginTop: 'auto',
          fontFamily: tokens.fontFamilyMonospace,
          fontSize: 10,
          color: tokens.colorBrandForeground2,
          background: tokens.colorBrandBackground2Hover,
          padding: '4px 6px',
          borderRadius: tokens.borderRadiusSmall,
          alignSelf: 'flex-start',
        }}
      >
        {syntax}
      </code>
    </button>
  );
}

// ──────────────────────────────────────────────────────────────
// SingleColumnEditor — value input for PUT mode
// Renders the same per-type input dispatch as the FieldSetEditor row, but
// scoped to one column.
// ──────────────────────────────────────────────────────────────
function SingleColumnEditor({
  table,
  column,
  value,
  onChange,
  onSwapMethod,
  themeMode,
}: {
  table: string;
  column: string | null;
  value: CreateFieldValue | undefined;
  onChange: (v: CreateFieldValue) => void;
  onSwapMethod: () => void;
  themeMode: ThemeMode;
}) {
  const tbl = findTable(table);
  const col = tbl && column ? findColumn(tbl, column) : undefined;

  if (!col) {
    return (
      <MessageBar layout="multiline" intent="warning">
        <MessageBarBody>
          Pick a column on the <strong>Method</strong> pane first.
        </MessageBarBody>
      </MessageBar>
    );
  }

  return (
    <div>
      <PaneHead
        icon={FormNew20Filled}
        title={`${col.displayName} → PUT`}
        // Fluent `Button` (transparent appearance) instead of a click-only
        // `<a>` — keeps the inline visual weight but gets proper button
        // semantics: Enter/Space activation, focus ring, ARIA role.
        sub={
          <span>
            Setting one column via the property URL.{' '}
            <Button
              appearance="transparent"
              size="small"
              onClick={onSwapMethod}
              style={{ padding: '0 2px', minWidth: 0, height: 'auto', verticalAlign: 'baseline' }}
            >
              Switch back to PATCH (multi-field)
            </Button>
          </span>
        }
        group="write"
      />
      <div style={{ maxWidth: 720 }}>
        <FieldSetEditor
          table={table}
          values={{ [col.logicalName]: value ?? '' }}
          setValues={(next) => onChange(next[col.logicalName] ?? '')}
          group="write"
          themeMode={themeMode}
          purpose="update"
        />
      </div>
    </div>
  );
}
