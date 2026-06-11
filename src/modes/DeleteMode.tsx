// Delete Record — DELETE /<entitySet>(<id>) or DELETE /<entitySet>(<id>)/<col>.
//
// Layout:
//   1. Big record-summary card with danger-red border showing target row + ID
//   2. Cascade-impact table — relationship list + cardinality + behavior NOTE
//      (we don't fetch counts; that's a per-relationship Dataverse query that
//      would cost N round-trips per pick. Behavior is what matters anyway.)
//   3. Type-the-record-name confirmation input (gates Execute)
//   4. "I understand this can't be undone" checkbox (gates Execute)
//
// Sidebar:
//   • Target       — entity + record + scope
//   • Confirmation — type-to-confirm + ack
//   • Headers      — HTTP / concurrency (If-Match) / bypass family
//
// Per docs:
//   • Whole-row delete:    DELETE /<entitySet>(<id>)            → 204 / 404
//   • Single property:     DELETE /<entitySet>(<id>)/<column>   → 204
//     (Doesn't support single-valued navigation properties — use
//      Disassociate via $ref for those.)
//
// Save / Load:
//   • Saved snapshot lives in the per-org saved-request library alongside
//     the read modes (modeId='delete'). We DO NOT persist `confirmText` or
//     `acknowledged` — those are one-shot safety affordances that must be
//     re-entered every time a saved delete is loaded.
//
// Reference: https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/update-delete-entities-using-web-api

import { useEffect, useMemo, useState } from 'react';
import {
  Table20Regular,
  Table20Filled,
  Delete20Regular,
  Delete20Filled,
  LineHorizontal320Regular,
  LineHorizontal320Filled,
  ShieldLock20Regular,
  ShieldLock20Filled,
  Warning20Filled,
  Checkmark20Filled,
} from '@fluentui/react-icons';
import {
  Field,
  RadioGroup,
  Radio,
  Combobox,
  Option,
  Caption1,
  tokens,
  Input,
  Checkbox,
  Badge,
  Persona,
  Spinner,
  Button,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  mergeClasses,
  Table,
  TableHeader,
  TableHeaderCell,
  TableBody,
  TableRow,
  TableCell,
} from '@fluentui/react-components';
import { Sidebar } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { useStudioStyles } from '../primitives/styles';
import { TargetEditor } from '../editors/TargetEditor';
import { PaneHead } from '../editors/PaneHead';
import {
  HeadersEditor,
  defaultWriteHeaders,
  headerItemsToObject,
  type HeaderItem,
} from '../editors/HeadersEditor';
import { PreconditionEditor, preconditionToHeader } from '../editors/PreconditionEditor';
import { BypassEditor, summarize as summarizeBypass } from '../editors/BypassEditor';
import { applyBypassToHeaders } from '../engine/bypassHeaders';
import { detectBypassAdvisories } from '../engine/bypassAdvisories';
import { disabledReasonFromAdvisories, type Advisory } from '../primitives/advisories';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { findTable } from '../mock/metadata';
import { findRequestType } from '../registry/requestTypes';
import { buildDelete } from '../engine/urlBuilder';
import { runtime, type ExecResult } from '../engine/runtime';
import { defaultBypassOptions, type DeleteState, type DeleteScope } from '../state/writeState';
import type { RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useLiveTable } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import {
  useCascadeConfiguration,
  cascadeSeverityRank,
  isParental,
  type CascadeBehavior,
} from '../host/useCascadeConfiguration';
import {
  serializeDelete,
  deserializeDelete,
  hashState,
  type SavedRequest,
  type SerializedDeleteState,
} from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';

const initialState = (): DeleteState => ({
  table: '',
  recordId: null,
  scope: { kind: 'whole-row' },
  concurrency: { kind: 'none' },
  headers: defaultWriteHeaders(),
  bypassCustomPlugins: false,
  bypass: defaultBypassOptions(),
  confirmText: '',
  acknowledged: false,
  dirty: new Set(),
});

type RootClauseId = 'target' | 'scope' | 'confirmation' | 'precondition' | 'headers' | 'bypass';

// GUID shape — Dataverse rejects malformed ids at the wire level, but we
// surface the mismatch via an advisory so the user finds out before the
// DELETE is sent (we don't want an irreversible round-trip on a typo).
const GUID_RE =
  /^[{(]?[0-9a-fA-F]{8}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{12}[)}]?$/;
const isValidGuid = (s: string | null | undefined): boolean => !!s && GUID_RE.test(s.trim());

export function DeleteMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('delete');
  const [state, setState] = useState(initialState);
  // Live-metadata subscription so child editors re-render when columns/relationships land.
  useLiveTable(state.table || null);
  const { entities } = useScopedEntities();
  const [activePath, setActivePath] = useState<string>('target');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);

  // Last-saved tracking (Save button dirty/clean + library checkmark).
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);

  const built = useMemo(() => buildDelete(state), [state]);
  const tbl = findTable(state.table);

  // ── Live row preview ──
  //
  // We have TWO sources of truth for the record's primary name:
  //
  //   1. The RecordPicker's `primary` callback — populated immediately
  //      when the user picks via the typeahead. Free, no extra fetch.
  //   2. A targeted `queryData` fetch — only needed when the picker
  //      can't supply the name (GUID-paste mode, saved-request load,
  //      etc.). Also pulls a small "tell me who this is" projection
  //      (accountnumber, email, etc.) for the danger card.
  //
  // The picker primary wins when available; the fetch fills the rest
  // (and can fail silently — the typed-confirm still works off the
  // picker primary).
  const [pickedPrimaryFromPicker, setPickedPrimaryFromPicker] = useState<string>('');
  const [pickedRow, setPickedRow] = useState<Record<string, unknown> | null>(null);
  const [pickedRowLoading, setPickedRowLoading] = useState(false);
  const [pickedRowError, setPickedRowError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    setPickedRowError(null);
    if (!state.recordId || !state.table || !tbl) {
      setPickedRow(null);
      return;
    }
    if (!isValidGuid(state.recordId)) {
      setPickedRow(null);
      return;
    }
    if (typeof window === 'undefined' || !window.dataverseAPI?.queryData) {
      setPickedRow(null);
      return;
    }
    // Build a tight $select. Filter by columns we KNOW exist on this
    // entity — sending a column the schema doesn't have makes Dataverse
    // 400 and breaks the entire preview.
    const cols = new Set<string>([tbl.primaryKey, tbl.primaryName]);
    for (const candidate of [
      'accountnumber',
      'emailaddress1',
      'fullname',
      'firstname',
      'lastname',
      'createdon',
      'modifiedon',
    ]) {
      if (tbl.columns.some((c) => c.logicalName === candidate)) cols.add(candidate);
    }
    // Use the same URL form as executeRetrieveSingle (leading slash).
    // PPTB's queryData accepts both forms in different paths but the
    // single-record form is more reliable with the leading slash.
    const url = `/${tbl.entitySetName}(${state.recordId})?$select=${[...cols].join(',')}`;
    setPickedRowLoading(true);
    window.dataverseAPI
      .queryData(url)
      .then((res) => {
        if (cancelled) return;
        // PPTB normalizes single-record responses into either
        // `{ value: <recordObject> }` OR `{ value: [<recordObject>] }`
        // depending on the host build. Handle both.
        const v = (res as { value?: unknown } | null)?.value;
        const row = Array.isArray(v)
          ? ((v[0] as Record<string, unknown> | undefined) ?? null)
          : v && typeof v === 'object'
            ? (v as Record<string, unknown>)
            : null;
        setPickedRow(row);
        setPickedRowLoading(false);
        setPickedRowError(null);
      })
      .catch((e: unknown) => {
        if (cancelled) return;
        // Capture the actual error so the danger card can explain WHY
        // the fetch didn't work — much more helpful than the previous
        // silent "couldn't fetch this record" message.
        const msg = e instanceof Error ? e.message : String(e);
        setPickedRow(null);
        setPickedRowLoading(false);
        setPickedRowError(msg);
      });
    return () => {
      cancelled = true;
    };
  }, [state.recordId, state.table, tbl]);

  // Picker-primary wins over fetch-derived primary. The fetch is only
  // for the secondary identifier (accountnumber etc.) in the danger card.
  const primaryName =
    pickedPrimaryFromPicker || (pickedRow && tbl ? String(pickedRow[tbl.primaryName] ?? '') : '');

  const markDirty = (id: string) =>
    setState((s) => {
      const d = new Set(s.dirty);
      d.add(id);
      return { ...s, dirty: d };
    });
  const set = <K extends keyof DeleteState>(k: K, v: DeleteState[K], dirtyId?: string) => {
    setState((s) => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  // ── Gates ──
  const requireTypedConfirm = state.scope.kind === 'whole-row';
  const typedConfirmOk =
    !requireTypedConfirm ||
    (state.confirmText.trim() === primaryName.trim() && primaryName.length > 0);

  // ── Save / Load ──
  const currentSerialized = useMemo(() => serializeDelete(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;

  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };

  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== 'delete') return;
    const snap = entry.state as SerializedDeleteState;
    // Lenient table check — entity list may still be warming up.
    if (entities.length > 0 && !entities.some((e) => e.logicalName === snap.table)) {
      window.alert(
        `Can't load "${entry.name}": entity \`${snap.table}\` ` +
          `isn't available in this environment. The solution may have been ` +
          `removed or you may be connected to a different org.`,
      );
      return;
    }
    setState(deserializeDelete(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setActivePath('target');
  };

  // Publish save context — hide when no table set (nothing meaningful
  // to save yet).
  usePublishSaveContext(
    useMemo(() => {
      if (!state.table) return null;
      return {
        state: currentSerialized,
        modeId: 'delete' as const,
        dirty: isDirty,
        lastSavedId,
        onSaved,
        onLoadSaved,
      };
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [currentSerialized, isDirty, lastSavedId, state.table]),
  );

  // ── Advisories — bypass family + GUID validation + missing record.
  // Same pattern as the read modes: a flat Advisory[] aggregated into the
  // URL-bar drawer, with `focusNode` driving "Open" navigation.
  const advisories = useMemo<Advisory[]>(() => {
    const out: Advisory[] = [];
    // Bypass-pane signals (legacy detector — composes business-logic +
    // step-id + flow-suppress advisories into the same Advisory shape).
    out.push(...detectBypassAdvisories(state.bypass, 'bypass'));

    if (state.table && state.recordId && !isValidGuid(state.recordId)) {
      out.push({
        id: 'del-bad-guid',
        severity: 'error',
        source: 'validation',
        focusNode: 'target',
        title: 'Invalid record id',
        body: `"${state.recordId}" isn't a valid GUID. DELETE will fail at the wire.`,
      });
    }

    // Single-property delete + embedded-in-PPTB = unsupported by the host
    // (see executeDelete for the full reasoning). We still let the user
    // author the request and copy the URL / Code snippet, but block
    // Execute so we don't silently whole-row-delete instead. In standalone
    // mode there's no host at all — Execute is already disabled by the
    // missing dataverseAPI — but we surface the advisory anyway so the
    // limitation is documented.
    if (state.scope.kind === 'single-property') {
      out.push({
        id: 'del-single-property-unsupported',
        severity: 'error',
        source: 'validation',
        focusNode: 'scope',
        title: "Single-property DELETE isn't supported by the PPTB host",
        body: (
          <>
            DRS authors the correct URL{' '}
            <code>DELETE /{state.table ? `&lt;set&gt;` : ''}(&lt;id&gt;)/&lt;column&gt;</code>, but
            PPTB&apos;s <code>dataverseAPI</code> only exposes whole-row delete — there&apos;s no
            raw-request hook for property-path URLs. To clear one column from inside PPTB, switch to{' '}
            <strong>Update mode</strong> and set the column to <code>null</code>. To use the
            property-DELETE pattern, copy the URL from the URL bar (or the Code-tab{' '}
            <code>fetch</code> snippet) and run it from Postman / curl / the JS SDK / Power
            Automate.
          </>
        ),
      });
    }

    return out;
  }, [state.bypass, state.table, state.recordId, state.scope]);

  const bypassBlocker = disabledReasonFromAdvisories(advisories);

  const disabledReason = !tbl
    ? 'Pick a target table first.'
    : !state.recordId
      ? 'Pick a record to delete.'
      : !isValidGuid(state.recordId)
        ? 'Record id is not a valid GUID.'
        : state.scope.kind === 'single-property' && !state.scope.column
          ? 'Pick a column to clear.'
          : state.concurrency.kind === 'etag' && !state.concurrency.etag.trim()
            ? 'Provide an etag value or switch the concurrency mode.'
            : state.headers.some((h) => h.enabled && !h.name)
              ? 'Fix empty header name.'
              : requireTypedConfirm && !primaryName
                ? 'Waiting for record metadata — pick again or check the GUID.'
                : requireTypedConfirm && !typedConfirmOk
                  ? `Type the record name "${primaryName}" to confirm.`
                  : requireTypedConfirm && !state.acknowledged
                    ? 'Acknowledge that this action cannot be undone.'
                    : bypassBlocker;

  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.delete(state);
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents((rs) =>
      [
        {
          id: `r-${Date.now()}`,
          modeId: 'delete',
          url: built.relativeUrl,
          method: 'DELETE',
          ts: Date.now(),
          status: res.status,
          ms: res.ms,
          rowCount: res.ok ? 1 : 0,
        },
        ...rs,
      ].slice(0, 8),
    );
    // Reset typed confirmation so a follow-up delete requires a fresh ack.
    // Also clear `dirty` (Save button settles to clean if the user wants
    // to persist the now-completed shape).
    setState((s) => ({ ...s, dirty: new Set(), confirmText: '', acknowledged: false }));
  };

  const effectiveHeaders = useMemo<HeaderItem[]>(() => {
    // Order: user headers → concurrency → bypass family (composed centrally).
    let h: HeaderItem[] = [...state.headers];
    const cc = preconditionToHeader(state.concurrency);
    if (cc) {
      h.push({
        id: `__cc__${cc.name}`,
        name: cc.name,
        value: cc.value,
        enabled: true,
        builtin: true,
        hint: 'Auto-composed from the Concurrency pane.',
      });
    }
    // Legacy field — kept in state for backward-compat. When true and the new
    // bypass UI is at 'none', we mirror it forward to bypass.businessLogic='sync'
    // + bypass.useLegacyHeader=true at compose time.
    const effectiveBypass =
      state.bypassCustomPlugins && state.bypass.businessLogic === 'none'
        ? { ...state.bypass, businessLogic: 'sync' as const, useLegacyHeader: true }
        : state.bypass;
    h = applyBypassToHeaders(h, effectiveBypass);
    return h;
  }, [state.headers, state.concurrency, state.bypassCustomPlugins, state.bypass]);

  const scopeBadge =
    state.scope.kind === 'whole-row' ? 'row' : `clear ${state.scope.column ?? '?'}`;
  // DELETE only supports `If-Match: "<etag>"` per docs — no `*` form,
  // no `If-None-Match`. Two badge states only.
  const concurrencyBadge = state.concurrency.kind === 'etag' ? 'If-Match: etag' : 'none';

  const isWholeRow = state.scope.kind === 'whole-row';
  const confirmOk =
    !isWholeRow ||
    (state.confirmText.trim() === primaryName.trim() &&
      primaryName.length > 0 &&
      state.acknowledged);
  const sections = [
    {
      id: 'target',
      label: 'Target',
      meta: tbl ? `${tbl.displayName} record` : 'Pick a table',
      items: [
        {
          id: 'target',
          icon: Table20Regular,
          iconFilled: Table20Filled,
          label: state.recordId
            ? primaryName || `${tbl?.displayName ?? ''} (selected)`
            : 'Pick a record',
          dirty: state.dirty.has('target'),
        },
        {
          id: 'scope',
          icon: Delete20Regular,
          iconFilled: Delete20Filled,
          label:
            state.scope.kind === 'whole-row' ? 'Whole row' : `Clear ${state.scope.column || '?'}`,
          code: state.scope.kind === 'single-property',
          badge: scopeBadge,
          badgeAppearance: 'ghost' as const,
          dirty: state.dirty.has('scope'),
        },
      ],
    },
    {
      id: 'confirmation',
      label: 'Confirmation',
      meta: isWholeRow ? (confirmOk ? '✓ ready' : 'incomplete') : 'n/a for single property',
      items: [
        {
          id: 'confirmation',
          icon: ShieldLock20Regular,
          iconFilled: ShieldLock20Filled,
          label: 'Type-to-confirm + ack',
          badge: isWholeRow ? (confirmOk ? '✓' : 'required') : 'skipped',
          badgeAppearance: 'tint' as const,
          badgeColor: isWholeRow
            ? confirmOk
              ? ('success' as const)
              : ('danger' as const)
            : ('subtle' as const),
          dirty: state.dirty.has('confirmation'),
        },
      ],
    },
    // Per unified write-mode layout: Precondition is a top-level section
    // (consistent with Update / Upsert) rather than a sub-item under
    // Headers. Same `state.concurrency` field underneath — just promoted
    // visually.
    {
      id: 'precondition',
      label: 'Precondition',
      meta: concurrencyBadge,
      items: [
        {
          id: 'precondition',
          icon: ShieldLock20Regular,
          iconFilled: ShieldLock20Filled,
          label: 'Optimistic concurrency',
          badge: concurrencyBadge,
          badgeAppearance: 'ghost' as const,
          dirty: state.dirty.has('precondition'),
        },
      ],
    },
    {
      id: 'headers',
      label: 'Headers',
      meta: `${state.headers.filter((h) => h.enabled).length + (state.concurrency.kind !== 'none' ? 1 : 0)} active`,
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

  let pane: React.ReactNode;
  const root = activePath as RootClauseId;
  switch (root) {
    case 'target':
      // Target = identification only. "What am I deleting?"
      // No relationships, no typed-confirm — those live on Confirmation.
      pane = (
        <DeleteTargetPane
          state={state}
          row={pickedRow}
          rowLoading={pickedRowLoading}
          rowError={pickedRowError}
          primaryName={primaryName}
          tbl={tbl}
          onTableChange={(t) => {
            setState((s) => ({
              ...s,
              table: t,
              recordId: null,
              scope: { kind: 'whole-row' },
              confirmText: '',
              acknowledged: false,
              dirty: new Set(['target', 'scope']),
            }));
            setResult(null);
            setPickedPrimaryFromPicker('');
          }}
          onRecordChange={(id, primary) => {
            set('recordId', id, 'target');
            setPickedPrimaryFromPicker(primary ?? '');
          }}
          onProceedToConfirmation={() => setActivePath('confirmation')}
        />
      );
      break;
    case 'confirmation':
      // Confirmation = consequences + ack. "Are you sure?"
      // Cascade-relationships table is here (not on Target) because it's
      // about what HAPPENS, not what's being targeted.
      pane = (
        <DeleteConfirmationPane
          state={state}
          row={pickedRow}
          primaryName={primaryName}
          tbl={tbl}
          onConfirmText={(v) => set('confirmText', v, 'confirmation')}
          onAck={(v) => set('acknowledged', v, 'confirmation')}
          onBackToTarget={() => setActivePath('target')}
        />
      );
      break;
    case 'scope':
      pane = (
        <ScopeEditor
          table={state.table}
          scope={state.scope}
          setScope={(sc) => set('scope', sc, 'scope')}
        />
      );
      break;
    case 'precondition':
      pane = (
        <PreconditionEditor
          mode={state.concurrency}
          setMode={(m) => set('concurrency', m, 'precondition')}
          // Per Microsoft docs (Apply optimistic concurrency on delete):
          // DELETE only supports `If-Match: "<etag>"`. The `If-Match: *` form
          // and `If-None-Match: *` form are Upsert-specific (PATCH) — DELETE
          // already fails 404 on missing records without any header, so
          // "require existing" adds no value. Restrict to None + ETag.
          available={['none', 'etag']}
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

  const headersMap = headerItemsToObject(effectiveHeaders, null);
  const codeInputs = {
    method: 'DELETE',
    built,
    headers: headersMap,
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
          method="DELETE"
          url={built.relativeUrl}
          executeVerb={type.executeVerb}
          // Delete uses the trash icon as a safety affordance so the
          // destructiveness reads before clicking.
          executeIcon={Delete20Filled}
          disabledReason={disabledReason}
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
              operation: 'delete',
              table: state.table,
              recordId: state.recordId,
              recordName: primaryName,
              clearedColumn: state.scope.kind === 'single-property' ? state.scope.column : null,
            }}
          />
        )}
      </MainTabs>
    </ModeShell>
  );
}

// ──────────────────────────────────────────────────────────────
// DeleteTargetPane — pick what to delete (no consequences shown here).
//
// Renders:
//   • Table picker
//   • Record picker
//   • Compact identity card (Persona + GUID + entity name)
//   • "Review consequences →" button → Confirmation pane
//
// Does NOT render the relationships table or the typed-confirm form.
// Splitting those out keeps Target focused on identification and
// Confirmation focused on safety. Prior design routed both sidebar
// entries to one combined pane and duplicated everything.
// ──────────────────────────────────────────────────────────────
function DeleteTargetPane({
  state,
  row,
  rowLoading,
  rowError,
  primaryName,
  tbl,
  onTableChange,
  onRecordChange,
  onProceedToConfirmation,
}: {
  state: DeleteState;
  row: Record<string, unknown> | null;
  rowLoading: boolean;
  rowError: string | null;
  primaryName: string;
  tbl: import('../mock/metadata').TableMeta | undefined;
  onTableChange: (t: string) => void;
  onRecordChange: (id: string | null, primary?: string) => void;
  onProceedToConfirmation: () => void;
}) {
  const s = useStudioStyles();
  const isWholeRow = state.scope.kind === 'whole-row';
  // Secondary identifier — first non-primary-name column we managed to
  // fetch with the row preview. Helps disambiguate two records with the
  // same primary name (common for contacts).
  const secondary: string | null = (() => {
    if (!row || !tbl) return null;
    const candidates = ['accountnumber', 'emailaddress1', tbl.primaryKey];
    for (const k of candidates) {
      const v = row[k];
      if (typeof v === 'string' && v.length > 0) return v;
    }
    return null;
  })();

  return (
    <div>
      <PaneHead
        icon={Delete20Filled}
        title={
          isWholeRow
            ? primaryName
              ? `Delete ${primaryName}`
              : 'Delete record'
            : `Clear ${state.scope.kind === 'single-property' ? state.scope.column : ''}`
        }
        sub={
          isWholeRow
            ? `This will issue DELETE on the row with ID ${state.recordId ?? '<no id>'}`
            : `This will clear a single column on row ${state.recordId ?? '<no id>'}`
        }
        group="write"
      />

      <MessageBar
        layout="multiline"
        intent={isWholeRow ? 'error' : 'warning'}
        icon={<Warning20Filled />}
        style={{ marginBottom: 14 }}
      >
        <MessageBarBody>
          <MessageBarTitle>
            {isWholeRow
              ? 'This is a destructive, irreversible request.'
              : 'Clearing a column is also irreversible.'}
          </MessageBarTitle>
          {isWholeRow
            ? 'Pick the record below. The Confirmation pane will show which relationships may cascade and require a typed name + acknowledgement before Execute lights up.'
            : 'The column value is set to null on the row. Other columns are unaffected.'}
        </MessageBarBody>
      </MessageBar>

      {/* Target picker — table + record (live typeahead / GUID). */}
      <div style={{ maxWidth: 720, marginBottom: 14 }}>
        <TargetEditor
          table={state.table}
          onTableChange={onTableChange}
          recordId={state.recordId}
          onRecordChange={onRecordChange}
          group="write"
          sub="Single-record delete — pick the entity set and the record."
        />
      </div>

      {state.recordId && (
        <div
          className={mergeClasses(s.inlineCard, s.inlineCardDanger)}
          style={{
            padding: 14,
            marginBottom: 14,
            maxWidth: 880,
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 12 }}>
            {/*
              Three states for the card header:

                1. We have a primary name (either from the picker OR a
                   successful fetch) → render the Persona, even while
                   the secondary fields are still loading. This is the
                   happy path for both Search-mode AND GUID-paste.
                2. No primary name yet AND we're fetching → spinner.
                3. No primary name AND fetch errored / returned empty →
                   warning block with the actual error message.
            */}
            {primaryName && tbl ? (
              <>
                {/* Persona handles name + secondary + tertiary natively
                    via its own slots — avoids rendering the primary name
                    twice (once in Persona, once in a sibling div). The
                    tertiary line includes the entity context + the
                    optional secondary identifier (accountnumber, email)
                    when the row preview has loaded it. */}
                {/* No "DELETE" pill on the right — the user feedback was
                    that it reads as an actionable button rather than a
                    status indicator. The pane title, URL-bar method pill,
                    and Execute button's trash icon already communicate
                    the verb. */}
                <Persona
                  size="huge"
                  name={primaryName}
                  primaryText={primaryName}
                  secondaryText={`${tbl.primaryKey}: ${String((row && row[tbl.primaryKey]) ?? state.recordId)}`}
                  tertiaryText={[
                    `${tbl.displayName} · ${tbl.entitySetName}`,
                    secondary,
                    rowLoading ? 'resolving extra fields…' : null,
                  ]
                    .filter(Boolean)
                    .join(' · ')}
                  avatar={{ color: 'red' }}
                />
              </>
            ) : rowLoading ? (
              <>
                <Spinner size="small" />
                <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                  Resolving record metadata…
                </Caption1>
              </>
            ) : (
              <>
                <Warning20Filled
                  style={{ color: tokens.colorPaletteRedForeground1, width: 32, height: 32 }}
                />
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 14, fontWeight: 600 }}>
                    Couldn&apos;t resolve this record
                  </div>
                  <Caption1 style={{ color: tokens.colorNeutralForeground3, marginTop: 4 }}>
                    GUID{' '}
                    <code style={{ fontFamily: tokens.fontFamilyMonospace }}>{state.recordId}</code>{' '}
                    didn&apos;t resolve in <code>{state.table}</code>. The record may already be
                    deleted, you might not have read access, or the metadata is still loading.
                    {rowError && (
                      <>
                        {' '}
                        Server response:{' '}
                        <code style={{ fontFamily: tokens.fontFamilyMonospace }}>{rowError}</code>
                      </>
                    )}
                  </Caption1>
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {/* Forward action to Confirmation. Only surfaces once the user has a
          target locked in AND it's a whole-row delete (single-property
          doesn't need a confirm step). */}
      {state.recordId && isWholeRow && primaryName && (
        <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginTop: 8 }}>
          <Button appearance="primary" onClick={onProceedToConfirmation}>
            Continue &rarr;
          </Button>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            Next: review the relationships that may be affected, then type the record name to
            confirm.
          </Caption1>
        </div>
      )}
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// DeleteConfirmationPane — consequences + safety checks.
//
// Renders:
//   • Compact record-identity strip (smaller than Target's, just for context)
//   • Cascade-impact relationships table
//   • Typed-name confirmation
//   • "I understand" ack checkbox
//   • Back button to return to Target if the user picked the wrong row
//
// Surfaces only when the record is selected. Single-property scope
// doesn't have meaningful cascades AND doesn't gate via typed-confirm,
// so this pane shows a short note + Back button instead.
// ──────────────────────────────────────────────────────────────
function DeleteConfirmationPane({
  state,
  row,
  primaryName,
  tbl,
  onConfirmText,
  onAck,
  onBackToTarget,
}: {
  state: DeleteState;
  row: Record<string, unknown> | null;
  primaryName: string;
  tbl: import('../mock/metadata').TableMeta | undefined;
  onConfirmText: (v: string) => void;
  onAck: (v: boolean) => void;
  onBackToTarget: () => void;
}) {
  const s = useStudioStyles();
  const isWholeRow = state.scope.kind === 'whole-row';

  // No record picked yet → bounce back to Target.
  if (!state.recordId) {
    return (
      <div>
        <PaneHead
          icon={ShieldLock20Filled}
          title="Confirmation"
          sub="Pick a record on the Target pane first — there's nothing to confirm yet."
          group="write"
        />
        <Button appearance="primary" onClick={onBackToTarget}>
          &larr; Go to Target
        </Button>
      </div>
    );
  }

  // Single-property scope → no typed-confirm, no cascade. Just a brief
  // note + Back. The actual column to clear is set on the Scope sidebar
  // entry, not here.
  if (!isWholeRow) {
    return (
      <div>
        <PaneHead
          icon={ShieldLock20Filled}
          title="Confirmation"
          sub="Single-property delete doesn't require type-to-confirm — clearing one column is reversible by re-entering the value."
          group="write"
        />
        <MessageBar layout="multiline" intent="info" style={{ marginBottom: 14, maxWidth: 720 }}>
          <MessageBarBody>
            Execute clears{' '}
            <code>{state.scope.kind === 'single-property' ? state.scope.column : ''}</code> on{' '}
            <strong>{primaryName || state.recordId}</strong> ({tbl?.displayName ?? state.table}).
            Other columns are untouched.
          </MessageBarBody>
        </MessageBar>
        <Button appearance="outline" onClick={onBackToTarget}>
          &larr; Change target
        </Button>
      </div>
    );
  }

  // Layout: full-height flex column. Top (head + identity strip) and
  // bottom (typed-confirm + ack) are fixed-size; the relationships table
  // in the middle gets the remaining vertical space and scrolls internally
  // so the confirm form is ALWAYS visible without the user scrolling the
  // outer pane. `minHeight: 0` on the middle is the canonical flex trick
  // that lets an `overflow: auto` child shrink below its content size.
  return (
    <div
      style={{
        display: 'flex',
        flexDirection: 'column',
        height: '100%',
        minHeight: 0,
        gap: 0,
      }}
    >
      <div style={{ flexShrink: 0 }}>
        <PaneHead
          icon={ShieldLock20Filled}
          title={primaryName ? `Confirm: delete ${primaryName}` : 'Confirm delete'}
          sub="Review what may be affected, then type the record name and acknowledge."
          group="write"
        />
      </div>

      {/* Compact identity strip — context-only. The big Persona lives on Target. */}
      {primaryName && tbl && (
        <div
          className={mergeClasses(s.inlineCard, s.inlineCardDanger)}
          style={{
            padding: '10px 14px',
            marginBottom: 14,
            display: 'flex',
            alignItems: 'center',
            gap: 12,
            flexShrink: 0,
          }}
        >
          <Persona
            size="medium"
            name={primaryName}
            primaryText={primaryName}
            secondaryText={`${tbl.primaryKey}: ${String((row && row[tbl.primaryKey]) ?? state.recordId)} · ${tbl.displayName}`}
            avatar={{ color: 'red' }}
          />
          <span style={{ flex: 1 }} />
          <Button size="small" appearance="subtle" onClick={onBackToTarget}>
            Change target
          </Button>
        </div>
      )}

      {/* Cascade impact preview — scrolls internally so the confirm form
          stays pinned at the bottom of the pane regardless of how many
          relationships the entity has. */}
      {tbl && tbl.navigationProperties.some((n) => n.cardinality !== 'ManyToOne') && (
        <div
          style={{
            display: 'flex',
            flexDirection: 'column',
            flex: '1 1 0',
            minHeight: 0,
            marginBottom: 14,
          }}
        >
          <strong style={{ fontSize: 12, display: 'block', marginBottom: 6, flexShrink: 0 }}>
            Related relationships that may be affected
          </strong>
          <div
            style={{
              flex: '1 1 0',
              minHeight: 120, // never collapse below useful height
              overflow: 'auto',
              border: `1px solid ${tokens.colorNeutralStroke2}`,
              borderRadius: tokens.borderRadiusMedium,
              background: tokens.colorNeutralBackground1,
            }}
          >
            <CascadeTable table={state.table} />
          </div>
          <Caption1
            style={{
              display: 'block',
              marginTop: 6,
              color: tokens.colorNeutralForeground3,
              flexShrink: 0,
            }}
          >
            Sorted by severity — <strong>Cascade</strong> (parental — children will be deleted) and{' '}
            <strong>Restrict</strong> (delete blocked if children exist) are highlighted at the top.
            Behaviors come from each relationship&apos;s <code>CascadeConfiguration.Delete</code>;
            row counts aren&apos;t shown.
          </Caption1>
        </div>
      )}

      {/* Typed confirmation — ALWAYS pinned at the bottom of the pane. */}
      {primaryName && (
        <div
          className={s.inlineCard}
          style={{
            padding: 14,
            maxWidth: 720,
            flexShrink: 0,
          }}
        >
          <strong style={{ fontSize: 12, display: 'block', marginBottom: 8 }}>
            Type the record name to confirm
          </strong>
          <Field
            label={
              <span>
                Expected:{' '}
                <strong style={{ fontFamily: tokens.fontFamilyMonospace }}>{primaryName}</strong>
              </span>
            }
            validationState={
              state.confirmText && state.confirmText.trim() === primaryName.trim()
                ? 'success'
                : state.confirmText
                  ? 'error'
                  : 'none'
            }
            validationMessage={
              state.confirmText && state.confirmText.trim() !== primaryName.trim()
                ? "Doesn't match — copy the name above exactly."
                : state.confirmText.trim() === primaryName.trim() && primaryName.length > 0
                  ? 'Match — Execute is unlocked once you check the box below.'
                  : undefined
            }
          >
            <Input
              value={state.confirmText}
              onChange={(_, d) => onConfirmText(d.value)}
              placeholder="Type to confirm…"
              contentAfter={
                state.confirmText.trim() === primaryName.trim() && primaryName.length > 0 ? (
                  <Checkmark20Filled style={{ color: tokens.colorPaletteGreenForeground1 }} />
                ) : undefined
              }
            />
          </Field>
          <div style={{ marginTop: 10 }}>
            <Checkbox
              checked={state.acknowledged}
              onChange={(_, d) => onAck(!!d.checked)}
              label="I understand this action cannot be undone."
            />
          </div>
        </div>
      )}
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// CascadeTable — relationship list view (metadata only, no live counts).
//
// What we DROPPED from the prototype version:
//   • Random/fake row counts (Math.random) — actively misleading.
//   • Per-relationship behavior badges hard-coded to 'Cascade'/'Remove Link'.
//     The actual behavior depends on each relationship's CascadeConfiguration
//     metadata, which is NOT in the basic navigation-property projection
//     we cache. Fetching it for every relationship would be N round-trips
//     per pick.
//
// What we KEEP:
//   • Listing every non-N:1 nav so the user sees what's downstream.
//   • Cardinality badge — the user can already infer "1:N typically
//     cascades, N:N typically un-links" without us pretending to know.
//   • Loud caveat below the table directing them to verify in the maker.
// ──────────────────────────────────────────────────────────────
function CascadeTable({ table }: { table: string }) {
  const tbl = findTable(table);
  // Fetch live CascadeConfiguration.Delete for every 1:N relationship
  // rooted at this entity. Cached at module level so re-mounting the
  // pane doesn't refetch. N:N relationships don't have configurable
  // cascade behavior on delete — they always RemoveLink — so we don't
  // fetch ManyToManyRelationships separately; the table renders those
  // with a fixed 'RemoveLink' label.
  const cascade = useCascadeConfiguration(table || null);

  // Build the row list and sort by severity. Rules:
  //   1. Within 1:N: rank by Delete behavior (Cascade → Restrict → ...)
  //   2. N:N rows sort to the bottom — always RemoveLink, never parental
  //   3. Ties broken by target display name (alphabetical)
  const rels = useMemo(() => {
    if (!tbl) return [];
    const byName = new Map(cascade.rows.map((r) => [r.schemaName, r.deleteBehavior]));
    const items = tbl.navigationProperties
      .filter((n) => n.cardinality !== 'ManyToOne')
      .map((n) => {
        const targetTbl = findTable(n.targetEntity);
        const display = targetTbl?.displayName ?? n.targetEntity;
        const behavior: CascadeBehavior | null =
          n.cardinality === 'ManyToMany' ? 'RemoveLink' : (byName.get(n.relationshipName) ?? null);
        return { nav: n, display, behavior };
      });
    items.sort((a, b) => {
      // Known parental + restrict first, then everything else by rank.
      // Unknown (null) behavior — typically because the cascade fetch
      // hasn't completed yet or this is a 1:N where the config didn't
      // come back — gets a middling rank so it doesn't dominate but
      // isn't hidden at the bottom either.
      const aRank = a.behavior ? cascadeSeverityRank(a.behavior) : 3.5;
      const bRank = b.behavior ? cascadeSeverityRank(b.behavior) : 3.5;
      if (aRank !== bRank) return aRank - bRank;
      return a.display.localeCompare(b.display);
    });
    return items;
  }, [tbl, cascade.rows]);

  if (!tbl || rels.length === 0) return null;

  return (
    <Table
      size="small"
      aria-label="Cascade impact preview"
      // `max-content` widths + `minWidth: 100%` is the same pattern
      // CollectionSubgrid uses — each column gets its natural width;
      // when total exceeds the container, horizontal scroll handles it
      // instead of columns squishing into each other.
      style={{ width: 'max-content', minWidth: '100%' }}
    >
      {/* Sticky header. Fluent v9 Table doesn't have a built-in sticky
          mode (DataGrid does, but DataGrid is heavier and we don't need
          sort/select primitives here). Two CSS tricks compose:
            1. `position: sticky; top: 0` on each TableHeaderCell so it
               stays visible when the scroll container scrolls.
            2. A solid `background` on the header cell — otherwise the
               cells go transparent and row content shows through during
               scroll. We use the same `colorNeutralBackground1` the
               table body uses, plus a bottom border for visual anchor.
          z-index keeps the header above row content. */}
      <TableHeader>
        <TableRow
          style={{
            position: 'sticky',
            top: 0,
            zIndex: 1,
            backgroundColor: tokens.colorNeutralBackground1,
          }}
        >
          <TableHeaderCell
            style={{
              minWidth: 180,
              position: 'sticky',
              top: 0,
              backgroundColor: tokens.colorNeutralBackground1,
              borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
            }}
          >
            Related table
          </TableHeaderCell>
          <TableHeaderCell
            style={{
              minWidth: 80,
              position: 'sticky',
              top: 0,
              backgroundColor: tokens.colorNeutralBackground1,
              borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
            }}
          >
            Cardinality
          </TableHeaderCell>
          <TableHeaderCell
            style={{
              minWidth: 260,
              position: 'sticky',
              top: 0,
              backgroundColor: tokens.colorNeutralBackground1,
              borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
            }}
          >
            Schema name
          </TableHeaderCell>
          <TableHeaderCell
            style={{
              minWidth: 180,
              position: 'sticky',
              top: 0,
              backgroundColor: tokens.colorNeutralBackground1,
              borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
            }}
          >
            Delete behavior
          </TableHeaderCell>
        </TableRow>
      </TableHeader>
      <TableBody>
        {rels.map(({ nav, display, behavior }) => {
          // Resolved label + color for the Delete behavior column.
          // Parental relationships (Cascade / Active / UserOwned) get
          // a red emphasis to signal "this WILL delete children".
          // Restrict gets a warning color (delete will be blocked).
          // RemoveLink / NoCascade stay neutral.
          const isP = !!behavior && isParental(behavior);
          const isRestrict = behavior === 'Restrict';
          const behaviorLabel = behavior ?? (cascade.loading ? 'loading…' : 'unknown');
          const behaviorColor: 'danger' | 'warning' | 'subtle' = isP
            ? 'danger'
            : isRestrict
              ? 'warning'
              : 'subtle';
          return (
            <TableRow
              key={nav.name}
              style={{
                // Highlight parental rows so the user sees the rows that
                // will cascade-delete children before the others. A red
                // left rule is more scannable than a full-row tint and
                // composes with Table's existing hover styles.
                borderLeft: isP
                  ? `3px solid ${tokens.colorPaletteRedBorder2}`
                  : isRestrict
                    ? `3px solid ${tokens.colorPaletteYellowBorder2}`
                    : '3px solid transparent',
              }}
            >
              <TableCell style={{ minWidth: 180 }}>
                <span style={{ fontWeight: 600 }}>{display}</span>
              </TableCell>
              <TableCell style={{ minWidth: 80 }}>
                <Badge appearance="ghost">
                  {nav.cardinality === 'OneToMany'
                    ? '1:N'
                    : nav.cardinality === 'ManyToMany'
                      ? 'N:N'
                      : 'N:1'}
                </Badge>
              </TableCell>
              <TableCell style={{ minWidth: 260 }}>
                <code
                  style={{
                    fontFamily: tokens.fontFamilyMonospace,
                    fontSize: 11,
                    color: tokens.colorNeutralForeground3,
                    whiteSpace: 'nowrap',
                  }}
                >
                  {nav.name}
                </code>
              </TableCell>
              <TableCell style={{ minWidth: 180 }}>
                <Badge
                  appearance="tint"
                  color={behaviorColor}
                  size="small"
                  style={{ fontFamily: tokens.fontFamilyMonospace }}
                >
                  {behaviorLabel}
                </Badge>
              </TableCell>
            </TableRow>
          );
        })}
      </TableBody>
    </Table>
  );
}

// ──────────────────────────────────────────────────────────────
// Scope editor — whole-row delete vs single-property clear.
// ──────────────────────────────────────────────────────────────
function ScopeEditor({
  table,
  scope,
  setScope,
}: {
  table: string;
  scope: DeleteScope;
  setScope: (s: DeleteScope) => void;
}) {
  const tbl = findTable(table);
  const clearable = useMemo(
    () =>
      (tbl?.columns ?? [])
        .filter(
          (c) =>
            c.attributeType !== 'Uniqueidentifier' &&
            c.attributeType !== 'File' &&
            c.attributeType !== 'Image' &&
            !c.required,
        )
        .slice()
        .sort((a, b) => a.displayName.localeCompare(b.displayName)),
    [tbl],
  );
  // `search` is the user's currently-typed filter text. While typing, the
  // input shows what they're typing; once they pick (or clear), it falls
  // back to the resolved displayName for the selected column.
  const [search, setSearch] = useState('');
  const [userIsTyping, setUserIsTyping] = useState(false);

  // Reset typed search whenever the parent clears the selection (e.g.
  // switching the scope back to whole-row and then to single-property
  // again). Without this, a stale search string would carry over.
  useEffect(() => {
    if (scope.kind === 'whole-row' || !scope.column) {
      setSearch('');
      setUserIsTyping(false);
    }
  }, [scope]);

  const selectedColumn =
    scope.kind === 'single-property'
      ? clearable.find((c) => c.logicalName === scope.column)
      : undefined;
  const selectedLabel =
    selectedColumn?.displayName ?? (scope.kind === 'single-property' ? scope.column : '');

  const filtered = useMemo(() => {
    if (!search) return clearable;
    const q = search.toLowerCase();
    return clearable.filter(
      (c) =>
        c.displayName.toLowerCase().includes(q) ||
        c.logicalName.toLowerCase().includes(q) ||
        c.attributeType.toLowerCase().includes(q),
    );
  }, [clearable, search]);

  const displayValue = userIsTyping ? search : selectedLabel;

  return (
    <div>
      <PaneHead
        icon={Delete20Filled}
        title="Delete scope"
        sub="Remove the entire row, or just clear one column's value."
        group="write"
      />
      {scope.kind === 'single-property' && (
        <MessageBar layout="multiline" intent="warning" style={{ marginBottom: 14, maxWidth: 720 }}>
          <MessageBarBody>
            <MessageBarTitle>Author-only inside PPTB</MessageBarTitle>
            PPTB&apos;s <code>dataverseAPI</code> doesn&apos;t expose property-path DELETE — only
            whole-row. DRS shows the correct URL + code so you can copy it for use outside (Postman,
            curl, SDK, Flow). To clear one column from inside PPTB, use <strong>Update mode</strong>{' '}
            and set the column to <code>null</code>.
          </MessageBarBody>
        </MessageBar>
      )}
      <div style={{ maxWidth: 720 }}>
        <Field label="Scope">
          <RadioGroup
            value={scope.kind}
            onChange={(_, d) => {
              if (d.value === 'whole-row') setScope({ kind: 'whole-row' });
              else
                setScope({
                  kind: 'single-property',
                  column: scope.kind === 'single-property' ? scope.column : '',
                });
            }}
          >
            <Radio
              value="whole-row"
              label={
                <span>
                  <strong>Whole row</strong> ·{' '}
                  <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>
                    DELETE /{tbl?.entitySetName ?? 'set'}(&lt;id&gt;)
                  </code>
                  <Caption1
                    style={{
                      display: 'block',
                      color: tokens.colorNeutralForeground3,
                      marginTop: 2,
                    }}
                  >
                    Removes the entire record. Cascading delete behavior follows the relationship
                    configuration.
                  </Caption1>
                </span>
              }
            />
            <Radio
              value="single-property"
              label={
                <span>
                  <strong>Single property</strong> ·{' '}
                  <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>
                    DELETE /{tbl?.entitySetName ?? 'set'}(&lt;id&gt;)/&lt;column&gt;
                  </code>
                  <Caption1
                    style={{
                      display: 'block',
                      color: tokens.colorNeutralForeground3,
                      marginTop: 2,
                    }}
                  >
                    Clears the value of one column. Required columns and lookups (single-valued
                    navigation properties) aren&apos;t supported via this path.
                  </Caption1>
                </span>
              }
            />
          </RadioGroup>
        </Field>

        {scope.kind === 'single-property' && (
          <div style={{ marginTop: 14, maxWidth: 480 }}>
            <Field
              label="Column to clear"
              hint={`${clearable.length} clearable column${clearable.length === 1 ? '' : 's'} on this entity. Type to filter. Required columns and lookup navigation properties are excluded.`}
            >
              <Combobox
                freeform
                clearable
                value={displayValue}
                selectedOptions={scope.column ? [scope.column] : []}
                onChange={(e) => {
                  const next = (e.target as HTMLInputElement).value;
                  setSearch(next);
                  setUserIsTyping(true);
                  // Empty input → drop the selection (matches the
                  // RecordPicker semantics: clearing the field clears
                  // the upstream choice immediately).
                  if (!next && scope.column) {
                    setScope({ kind: 'single-property', column: '' });
                  }
                }}
                onOptionSelect={(_, d) => {
                  if (!d.optionValue) {
                    setScope({ kind: 'single-property', column: '' });
                    setSearch('');
                    setUserIsTyping(false);
                    return;
                  }
                  setScope({ kind: 'single-property', column: d.optionValue });
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
                        <span style={{ color: tokens.colorBrandForeground2 }}>
                          {c.attributeType}
                        </span>
                      </Caption1>
                    </div>
                  </Option>
                ))}
              </Combobox>
            </Field>
          </div>
        )}
      </div>
    </div>
  );
}
