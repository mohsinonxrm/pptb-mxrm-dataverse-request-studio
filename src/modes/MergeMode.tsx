// Merge Records — POST /Merge — bound action that merges Subordinate INTO
// Target. Restricted to account / contact / lead / incident.
//
// Modeled on the Power Apps "Merge duplicate records" UX:
//   https://learn.microsoft.com/en-us/power-apps/user/merge-duplicate-records
//
// Sidebar:
//   • Entity        — pick which table (account / contact / incident)
//   • Target        — winner record (survives)
//   • Subordinate   — loser record (deactivated, re-parented)
//   • Field diff    — per-field choice (Keep Target / Use Subordinate / Custom)
//   • Options       — PerformParentingChecks, SuppressDuplicateDetection
//   • Headers
//
// Main pane (Builder tab):
//   1. Banner — irreversibility warning
//   2. Side-by-side record cards — Target ← Subordinate, with avatars + counts
//   3. Field comparison table — three-way choice per row
//   4. Merge options card
//   5. Request body preview
//
// Grounded in:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/merge
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/merge-entity-using-web-api

import { useEffect, useMemo, useState } from 'react';
import {
  Table20Regular,
  Table20Filled,
  PersonAccounts20Regular,
  PersonAccounts20Filled,
  BranchFork20Regular,
  BranchFork20Filled,
  Settings20Regular,
  Settings20Filled,
  LineHorizontal320Regular,
  LineHorizontal320Filled,
  Warning20Filled,
  ArrowLeft20Regular,
} from '@fluentui/react-icons';
import {
  Switch,
  Caption1,
  tokens,
  Field,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Combobox,
  Option,
  Badge,
  Persona,
  Button,
} from '@fluentui/react-components';
import { Sidebar } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { useStudioStyles } from '../primitives/styles';
import { PaneHead } from '../editors/PaneHead';
import { MergeFieldDiff } from '../editors/MergeFieldDiff';
import { HeadersEditor, defaultWriteHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { BypassEditor, summarize as summarizeBypass } from '../editors/BypassEditor';
import { applyBypassToHeaders } from '../engine/bypassHeaders';
import { detectBypassAdvisories } from '../engine/bypassAdvisories';
import { disabledReasonFromAdvisories, type Advisory } from '../primitives/advisories';
import { RecordPicker } from '../primitives/RecordPicker';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { TABLES, findTable } from '../mock/metadata';
import { findRequestType } from '../registry/requestTypes';
import { buildMerge, buildMergeBody } from '../engine/urlBuilder';
import { runtime, type ExecResult } from '../engine/runtime';
import {
  defaultBypassOptions,
  type MergeState,
  type MergeFieldChoice,
  type CreateFieldValue,
} from '../state/writeState';
import type { RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useLiveTable } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import {
  serializeMerge,
  deserializeMerge,
  hashState,
  type SavedRequest,
  type SerializedMergeState,
} from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';

// ──────────────────────────────────────────────────────────────
// Initial state
// ──────────────────────────────────────────────────────────────
// Empty initial state. The user picks one of account / contact / incident
// (the only entities the Merge action supports per docs), then picks both
// records via the live RecordPicker.
const initialState = (): MergeState => ({
  table: '',
  targetId: null,
  subordinateId: null,
  fieldChoices: {},
  customValues: {},
  performParentingChecks: true,
  suppressDuplicateDetection: false,
  bypass: defaultBypassOptions(),
  headers: defaultWriteHeaders(),
  dirty: new Set(),
  targetSnapshot: null,
  subordinateSnapshot: null,
});

const GUID_RE =
  /^[{(]?[0-9a-fA-F]{8}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{12}[)}]?$/;
const isValidGuid = (s: string | null | undefined): boolean => !!s && GUID_RE.test(s.trim());

type RootClauseId = 'target' | 'subordinate' | 'fielddiff' | 'options' | 'headers' | 'bypass';

export function MergeMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('merge');
  const [state, setState] = useState(initialState);
  // Live-metadata subscription so child editors re-render when columns/relationships land.
  useLiveTable(state.table || null);
  const { entities } = useScopedEntities();
  const [activePath, setActivePath] = useState<string>('target');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);

  // Save / Load tracking
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);

  // Loading flags for the two record fetches.
  const [targetLoading, setTargetLoading] = useState(false);
  const [subordinateLoading, setSubordinateLoading] = useState(false);

  const built = useMemo(() => buildMerge(state), [state]);
  const body = useMemo(() => buildMergeBody(state), [state]);
  const tbl = findTable(state.table);

  // ── Live row fetches ──
  //
  // The field-diff pane needs BOTH rows' values to render the side-by-side
  // comparison, and buildMergeBody reads from `subordinateSnapshot` to copy
  // values when the user picks "use subordinate" for a field.
  //
  // We fire two parallel fetches whenever (table, id) changes for either
  // role. The "tell-me-who-this-is" projection in Delete/Update used a
  // tight $select; for Merge we WANT the full row because every column
  // appears in the diff table. Heavy entities (200+ columns) will take a
  // beat but it's one fetch per role per pick.
  useEffect(() => {
    let cancelled = false;
    if (!state.table || !state.targetId || !tbl || !isValidGuid(state.targetId)) {
      setState((s) => ({ ...s, targetSnapshot: null }));
      return;
    }
    if (typeof window === 'undefined' || !window.dataverseAPI?.queryData) return;
    setTargetLoading(true);
    const url = `/${tbl.entitySetName}(${state.targetId})`;
    window.dataverseAPI
      .queryData(url)
      .then((res) => {
        if (cancelled) return;
        const v = (res as { value?: unknown } | null)?.value;
        const row = Array.isArray(v)
          ? ((v[0] as Record<string, unknown> | undefined) ?? null)
          : v && typeof v === 'object'
            ? (v as Record<string, unknown>)
            : res && typeof res === 'object' && tbl.primaryKey in (res as object)
              ? (res as Record<string, unknown>)
              : null;
        setState((s) => ({ ...s, targetSnapshot: row }));
        setTargetLoading(false);
      })
      .catch(() => {
        if (cancelled) return;
        setState((s) => ({ ...s, targetSnapshot: null }));
        setTargetLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [state.table, state.targetId, tbl]);

  useEffect(() => {
    let cancelled = false;
    if (!state.table || !state.subordinateId || !tbl || !isValidGuid(state.subordinateId)) {
      setState((s) => ({ ...s, subordinateSnapshot: null }));
      return;
    }
    if (typeof window === 'undefined' || !window.dataverseAPI?.queryData) return;
    setSubordinateLoading(true);
    const url = `/${tbl.entitySetName}(${state.subordinateId})`;
    window.dataverseAPI
      .queryData(url)
      .then((res) => {
        if (cancelled) return;
        const v = (res as { value?: unknown } | null)?.value;
        const row = Array.isArray(v)
          ? ((v[0] as Record<string, unknown> | undefined) ?? null)
          : v && typeof v === 'object'
            ? (v as Record<string, unknown>)
            : res && typeof res === 'object' && tbl.primaryKey in (res as object)
              ? (res as Record<string, unknown>)
              : null;
        setState((s) => ({ ...s, subordinateSnapshot: row }));
        setSubordinateLoading(false);
      })
      .catch(() => {
        if (cancelled) return;
        setState((s) => ({ ...s, subordinateSnapshot: null }));
        setSubordinateLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [state.table, state.subordinateId, tbl]);

  const markDirty = (id: string) =>
    setState((s) => {
      const d = new Set(s.dirty);
      d.add(id);
      return { ...s, dirty: d };
    });
  const set = <K extends keyof MergeState>(k: K, v: MergeState[K], dirtyId?: string) => {
    setState((s) => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };
  const setFieldChoice = (field: string, choice: MergeFieldChoice) => {
    setState((s) => {
      const next = { ...s.fieldChoices };
      if (choice === 'target') delete next[field];
      else next[field] = choice;
      return { ...s, fieldChoices: next, dirty: new Set(s.dirty).add('fielddiff') };
    });
  };
  const setCustomValue = (field: string, v: CreateFieldValue) => {
    setState((s) => ({
      ...s,
      customValues: { ...s.customValues, [field]: v },
      dirty: new Set(s.dirty).add('fielddiff'),
    }));
  };

  // ── Resolved records ──
  // Read from the live snapshots in state (populated by the effects above).
  // `null` either means "not picked yet" OR "fetch in flight" — disambiguated
  // by the loading flags.
  const targetRow = state.targetSnapshot;
  const subRow = state.subordinateSnapshot;

  const supportedTables = TABLES.filter((t) => t.supportsMerge);
  const overrideCount = Object.values(state.fieldChoices).filter((c) => c !== 'target').length;

  const disabledReason = !tbl
    ? 'Pick a target table first.'
    : !tbl.supportsMerge
      ? 'Merge is only supported for account, contact, and incident.'
      : !state.targetId
        ? 'Pick a Target record (the winner).'
        : !isValidGuid(state.targetId)
          ? 'Target record id is not a valid GUID.'
          : !state.subordinateId
            ? 'Pick a Subordinate record (the loser).'
            : !isValidGuid(state.subordinateId)
              ? 'Subordinate record id is not a valid GUID.'
              : state.targetId === state.subordinateId
                ? 'Target and Subordinate must be different records.'
                : state.headers.some((h) => h.enabled && !h.name)
                  ? 'Fix empty header name.'
                  : null;

  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.merge(state);
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents((rs) =>
      [
        {
          id: `r-${Date.now()}`,
          modeId: 'merge',
          url: built.relativeUrl,
          method: 'POST',
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

  // ── Compose effective headers (auto-toggle MSCRM.SuppressDuplicateDetection + bypass family) ──
  const effectiveHeaders = useMemo(() => {
    let h = [...state.headers];
    if (state.suppressDuplicateDetection) {
      h.push({
        id: '__sup-dupe',
        name: 'MSCRM.SuppressDuplicateDetection',
        value: 'true',
        enabled: true,
        builtin: true,
        hint: 'Skips duplicate detection rules when applying UpdateContent overrides.',
      });
    }
    h = applyBypassToHeaders(h, state.bypass);
    return h;
  }, [state.headers, state.suppressDuplicateDetection, state.bypass]);

  // ── Advisories — bypass family + GUID validation + same-record guard ──
  const advisories = useMemo<Advisory[]>(() => {
    const out: Advisory[] = [...detectBypassAdvisories(state.bypass, 'bypass')];
    if (state.targetId && !isValidGuid(state.targetId)) {
      out.push({
        id: 'merge-bad-target-guid',
        severity: 'error',
        source: 'validation',
        focusNode: 'target',
        title: 'Invalid Target record id',
        body: `"${state.targetId}" isn't a valid GUID. The merge will fail at the wire.`,
      });
    }
    if (state.subordinateId && !isValidGuid(state.subordinateId)) {
      out.push({
        id: 'merge-bad-sub-guid',
        severity: 'error',
        source: 'validation',
        focusNode: 'subordinate',
        title: 'Invalid Subordinate record id',
        body: `"${state.subordinateId}" isn't a valid GUID. The merge will fail at the wire.`,
      });
    }
    return out;
  }, [state.bypass, state.targetId, state.subordinateId]);
  const bypassBlocker = disabledReasonFromAdvisories(advisories);

  // ── Save / Load ──
  const currentSerialized = useMemo(() => serializeMerge(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;

  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };

  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== 'merge') return;
    const snap = entry.state as SerializedMergeState;
    if (entities.length > 0 && !entities.some((e) => e.logicalName === snap.table)) {
      window.alert(
        `Can't load "${entry.name}": entity \`${snap.table}\` ` +
          `isn't available in this environment. The solution may have been ` +
          `removed or you may be connected to a different org.`,
      );
      return;
    }
    setState(deserializeMerge(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setActivePath('target');
  };

  usePublishSaveContext(
    useMemo(() => {
      if (!state.table) return null;
      return {
        state: currentSerialized,
        modeId: 'merge' as const,
        dirty: isDirty,
        lastSavedId,
        onSaved,
        onLoadSaved,
      };
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [currentSerialized, isDirty, lastSavedId, state.table]),
  );

  // Merge has its own canonical section structure:
  //   * Target          → Master (winner) record picker
  //   * Subordinate     → Record to be merged in
  //   * Override fields → Columns where subordinate wins
  //   * Settings        → PerformParentingChecks + Suppress duplicate detection
  //   * Headers
  // The entity-type picker lives inside the Target pane (combobox up top).
  const sections = [
    {
      id: 'target',
      label: 'Target',
      meta: tbl?.displayName ?? 'Pick a table',
      items: [
        {
          id: 'target',
          icon: PersonAccounts20Regular,
          iconFilled: PersonAccounts20Filled,
          label: 'Master (survives)',
          badge: state.targetId ? '✓' : null,
          badgeAppearance: 'ghost' as const,
          dirty: state.dirty.has('target') || state.dirty.has('entity'),
        },
      ],
    },
    {
      id: 'subordinate-section',
      label: 'Subordinate',
      meta: state.subordinateId ? 'will be deactivated' : 'pick a record',
      items: [
        {
          id: 'subordinate',
          icon: BranchFork20Regular,
          iconFilled: BranchFork20Filled,
          label: 'Subordinate record',
          badge: state.subordinateId ? '✓' : null,
          badgeAppearance: 'ghost' as const,
          dirty: state.dirty.has('subordinate'),
        },
      ],
    },
    {
      id: 'override',
      label: 'Override fields',
      meta: `${overrideCount} override${overrideCount === 1 ? '' : 's'}`,
      items: [
        {
          id: 'fielddiff',
          icon: BranchFork20Regular,
          iconFilled: BranchFork20Filled,
          label: 'Field comparison',
          badge: overrideCount || null,
          badgeAppearance: 'tint' as const,
          badgeColor: overrideCount > 0 ? ('brand' as const) : ('subtle' as const),
          dirty: state.dirty.has('fielddiff'),
        },
      ],
    },
    {
      id: 'settings',
      label: 'Settings',
      meta:
        (state.performParentingChecks ? 1 : 0) + (state.suppressDuplicateDetection ? 1 : 0) === 0
          ? 'defaults'
          : `${(state.performParentingChecks ? 1 : 0) + (state.suppressDuplicateDetection ? 1 : 0)} set`,
      items: [
        {
          id: 'options',
          icon: Settings20Regular,
          iconFilled: Settings20Filled,
          label: 'Merge options',
          badge:
            (state.performParentingChecks ? 1 : 0) + (state.suppressDuplicateDetection ? 1 : 0) ||
            null,
          badgeAppearance: 'ghost' as const,
          dirty: state.dirty.has('options'),
        },
      ],
    },
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

  let pane: React.ReactNode;
  const root = activePath as RootClauseId;
  switch (root) {
    case 'target':
      // Target combines the entity-type picker (top) with the master
      // record picker. Subordinate is its own section.
      pane = (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
          <EntityPane state={state} setState={setState} markDirty={markDirty} />
          <RecordPickPane
            title="Master record"
            sub="The record that keeps its identity. Subordinate is merged into it."
            table={state.table}
            selectedId={state.targetId}
            onPick={(id) => set('targetId', id, 'target')}
          />
        </div>
      );
      break;
    case 'subordinate':
      pane = (
        <RecordPickPane
          title="Subordinate (deactivated)"
          sub="The duplicate to merge away. Its relationships re-parent to the Target; the row gets statecode=Inactive."
          table={state.table}
          selectedId={state.subordinateId}
          onPick={(id) => set('subordinateId', id, 'subordinate')}
          conflictWith={state.targetId}
        />
      );
      break;
    case 'fielddiff':
      // The field comparison pane is the centerpiece — also surfaces the
      // side-by-side header card above the diff table.
      pane = (
        <div>
          <MergeBanner
            overrideCount={overrideCount}
            subordinateName={primaryNameOf(tbl, subRow ?? undefined)}
          />
          <MergeRecordsHeader
            table={state.table}
            target={targetRow ?? undefined}
            subordinate={subRow ?? undefined}
            targetLoading={targetLoading}
            subordinateLoading={subordinateLoading}
          />
          <MergeFieldDiff
            table={state.table}
            target={state.targetSnapshot ?? null}
            sub={state.subordinateSnapshot ?? null}
            fieldChoices={state.fieldChoices}
            customValues={state.customValues}
            setFieldChoice={setFieldChoice}
            setCustomValue={setCustomValue}
          />
        </div>
      );
      break;
    case 'options':
      pane = (
        <OptionsPane
          performParentingChecks={state.performParentingChecks}
          setPerformParentingChecks={(v) => set('performParentingChecks', v, 'options')}
          suppressDuplicateDetection={state.suppressDuplicateDetection}
          setSuppressDuplicateDetection={(v) => set('suppressDuplicateDetection', v, 'options')}
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
    method: 'POST',
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
          method="POST"
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
              operation: 'merge',
              table: state.table,
              recordId: state.targetId,
              recordName: primaryNameOf(tbl, targetRow ?? undefined),
              targetName: primaryNameOf(tbl, targetRow ?? undefined),
              subordinateName: primaryNameOf(tbl, subRow ?? undefined),
            }}
          />
        )}
      </MainTabs>
    </ModeShell>
  );
}

// ──────────────────────────────────────────────────────────────
// Banner — irreversibility warning + override summary
// ──────────────────────────────────────────────────────────────
function MergeBanner({
  overrideCount,
  subordinateName,
}: {
  overrideCount: number;
  subordinateName: string | null;
}) {
  return (
    <MessageBar
      layout="multiline"
      intent="warning"
      icon={<Warning20Filled />}
      style={{ marginBottom: 14 }}
    >
      <MessageBarBody>
        <MessageBarTitle>Merge is irreversible.</MessageBarTitle>
        {subordinateName ? (
          <>
            <strong>{subordinateName}</strong> will be deactivated,{' '}
          </>
        ) : (
          'The subordinate will be deactivated, '
        )}
        its related records re-parent to the Target, and <strong>{overrideCount}</strong> field
        {overrideCount === 1 ? '' : 's'} on the Target will be overwritten by your overrides. The
        server also auto-fills any empty Target fields from the Subordinate.
      </MessageBarBody>
    </MessageBar>
  );
}

// ──────────────────────────────────────────────────────────────
// Side-by-side header card — Target ← Subordinate
// ──────────────────────────────────────────────────────────────
function MergeRecordsHeader({
  table,
  target,
  subordinate,
  targetLoading,
  subordinateLoading,
}: {
  table: string;
  target: Record<string, unknown> | undefined;
  subordinate: Record<string, unknown> | undefined;
  targetLoading?: boolean;
  subordinateLoading?: boolean;
}) {
  const s = useStudioStyles();
  const tbl = findTable(table);
  if (!tbl) return null;

  return (
    <div
      className={s.inlineCard}
      style={{
        padding: 14,
        marginBottom: 14,
        display: 'grid',
        gridTemplateColumns: '1fr 44px 1fr',
        gap: 10,
        alignItems: 'stretch',
      }}
    >
      {/* Target */}
      <RecordCard
        tone="target"
        chip="TARGET"
        chipDescription="survives the merge"
        row={target}
        tbl={tbl}
        emptyHint={targetLoading ? 'Loading Target record…' : 'Pick a Target record.'}
      />

      {/* Arrow is a static directional cue (subordinate flows INTO target),
          not an actionable element — Regular weight + foreground3 keeps it
          as ambient chrome instead of competing with the Execute button for
          the brand color. */}
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          color: tokens.colorNeutralForeground3,
        }}
      >
        <ArrowLeft20Regular style={{ width: 28, height: 28 }} />
      </div>

      {/* Subordinate */}
      <RecordCard
        tone="sub"
        chip="SUBORDINATE"
        chipDescription="will be deactivated"
        row={subordinate}
        tbl={tbl}
        emptyHint={
          subordinateLoading ? 'Loading Subordinate record…' : 'Pick a Subordinate record.'
        }
      />
    </div>
  );
}

function RecordCard({
  tone,
  chip,
  chipDescription,
  row,
  tbl,
  emptyHint,
}: {
  tone: 'target' | 'sub';
  chip: string;
  chipDescription: string;
  row: Record<string, unknown> | undefined;
  tbl: import('../mock/metadata').TableMeta;
  emptyHint: string;
}) {
  // Uses Fluent palette tokens (not hand-mixed rgba) so the accent shifts
  // automatically when the theme switches between light and dark.
  const accent =
    tone === 'target'
      ? {
          border: `1px solid ${tokens.colorPaletteGreenBorderActive}`,
          background: tokens.colorPaletteGreenBackground1,
        }
      : {
          border: `1px solid ${tokens.colorPaletteDarkOrangeBorderActive}`,
          background: tokens.colorPaletteDarkOrangeBackground1,
        };

  const chipBg =
    tone === 'target'
      ? tokens.colorPaletteGreenForeground1
      : tokens.colorPaletteDarkOrangeForeground1;

  const name = primaryNameOf(tbl, row);
  const id = row ? String(row[tbl.primaryKey] ?? '') : null;

  // Mock-data context — show created date + a couple of relationship counts when possible.
  const createdOn = row?.createdon ? String(row.createdon) : null;
  const createdLabel = createdOn
    ? new Date(createdOn).toLocaleDateString(undefined, { year: 'numeric', month: 'short' })
    : null;

  return (
    <div style={{ ...accent, borderRadius: tokens.borderRadiusMedium, padding: 12 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
        <Badge
          appearance="filled"
          style={{
            background: chipBg,
            color: tokens.colorNeutralForegroundOnBrand,
            fontWeight: 700,
            letterSpacing: 0.4,
          }}
        >
          {chip}
        </Badge>
        <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>{chipDescription}</Caption1>
      </div>

      {row ? (
        <>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 8 }}>
            <Persona size="small" name={name ?? '?'} avatar={{ color: 'colorful' }} />
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontSize: 13, fontWeight: 600 }}>{name ?? '(no name)'}</div>
              <div
                style={{
                  fontFamily: tokens.fontFamilyMonospace,
                  fontSize: 10,
                  color: tokens.colorNeutralForeground3,
                }}
              >
                {id ?? ''}
              </div>
            </div>
          </div>
          <div
            style={{
              display: 'flex',
              flexDirection: 'column',
              gap: 2,
              fontSize: 11,
              color: tokens.colorNeutralForeground3,
            }}
          >
            {createdLabel && (
              <div>
                Created:{' '}
                <strong style={{ color: tokens.colorNeutralForeground2 }}>{createdLabel}</strong>
              </div>
            )}
            {row.accountnumber != null && (
              <div>
                Account #:{' '}
                <strong
                  style={{
                    color: tokens.colorNeutralForeground2,
                    fontFamily: tokens.fontFamilyMonospace,
                  }}
                >
                  {String(row.accountnumber)}
                </strong>
              </div>
            )}
            {row.revenue != null && (
              <div>
                Revenue:{' '}
                <strong style={{ color: tokens.colorNeutralForeground2 }}>
                  {Number(row.revenue).toLocaleString(undefined, {
                    style: 'currency',
                    currency: 'USD',
                    maximumFractionDigits: 0,
                  })}
                </strong>
              </div>
            )}
            {row.emailaddress1 != null && (
              <div>
                Email:{' '}
                <strong style={{ color: tokens.colorNeutralForeground2 }}>
                  {String(row.emailaddress1)}
                </strong>
              </div>
            )}
          </div>
        </>
      ) : (
        <div
          style={{
            padding: '20px 4px',
            color: tokens.colorNeutralForeground3,
            fontSize: 12,
            fontStyle: 'italic',
            textAlign: 'center',
          }}
        >
          {emptyHint}
        </div>
      )}
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Sidebar-driven panes
// ──────────────────────────────────────────────────────────────
// Per the docs Merge is restricted to these three entities (and lead in
// older versions; current docs name account/contact/incident). Hardcoded
// because it's a server-side constraint of the OOB Merge action — not
// configurable per environment.
//   Reference: https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/merge
const MERGE_SUPPORTED_LOGICAL_NAMES = ['account', 'contact', 'incident'] as const;

function EntityPane({
  state,
  setState,
  markDirty,
}: {
  state: MergeState;
  setState: React.Dispatch<React.SetStateAction<MergeState>>;
  markDirty: (id: string) => void;
}) {
  const tbl = findTable(state.table);
  const { entities } = useScopedEntities();
  // Cross-reference the three docs-supported logical names against the
  // live entity list — so we show the env's actual display names (which
  // can be customized) and skip entities the env doesn't have provisioned.
  const supportedTables = useMemo(
    () =>
      MERGE_SUPPORTED_LOGICAL_NAMES.map((ln) => entities.find((e) => e.logicalName === ln)).filter(
        (e): e is NonNullable<typeof e> => !!e,
      ),
    [entities],
  );
  return (
    <div>
      <PaneHead
        icon={Table20Filled}
        title="Merge entity type"
        sub="Pick the table to merge. The Merge action is restricted to account, contact, and incident."
        group="write"
      />
      <div style={{ maxWidth: 560 }}>
        <Field label="Entity">
          <Combobox
            value={tbl?.displayName ?? ''}
            selectedOptions={[state.table]}
            onOptionSelect={(_, d) => {
              if (!d.optionValue) return;
              setState((s) => ({
                ...s,
                table: d.optionValue!,
                targetId: null,
                subordinateId: null,
                fieldChoices: {},
                customValues: {},
                targetSnapshot: null,
                subordinateSnapshot: null,
              }));
              markDirty('entity');
              markDirty('target');
              markDirty('subordinate');
              markDirty('fielddiff');
            }}
          >
            {supportedTables.map((t) => (
              <Option key={t.logicalName} value={t.logicalName} text={t.displayName}>
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                  <span>{t.displayName}</span>
                  <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                    {t.entitySetName} · supports Merge
                  </Caption1>
                </div>
              </Option>
            ))}
          </Combobox>
        </Field>
        <MessageBar layout="multiline" intent="info" style={{ marginTop: 14 }}>
          <MessageBarBody>
            Merge requires both records be the <strong>same entity type</strong>.{' '}
            <code>incident</code> behaves differently — per docs, the <code>UpdateContent</code>{' '}
            parameter is <strong>ignored</strong> for incident, and the merge runs in the user's
            security context (not system).
          </MessageBarBody>
        </MessageBar>
      </div>
    </div>
  );
}

function RecordPickPane({
  title,
  sub,
  table,
  selectedId,
  onPick,
  conflictWith,
}: {
  title: string;
  sub: string;
  table: string;
  selectedId: string | null;
  onPick: (id: string | null) => void;
  conflictWith?: string | null;
}) {
  const tbl = findTable(table);
  const conflict = conflictWith && selectedId && conflictWith === selectedId;
  return (
    <div>
      <PaneHead icon={PersonAccounts20Filled} title={title} sub={sub} group="write" />
      <div style={{ maxWidth: 480 }}>
        <Field label="Record">
          <RecordPicker
            table={table}
            selectedId={selectedId}
            onPick={(r) => onPick(r?.id ?? null)}
            placeholder={`Search ${tbl?.displayName ?? ''} records…`}
          />
        </Field>
        {conflict && (
          <MessageBar layout="multiline" intent="error" style={{ marginTop: 14 }}>
            <MessageBarBody>Target and Subordinate must be different records.</MessageBarBody>
          </MessageBar>
        )}
        {!selectedId && (
          <Caption1
            style={{ display: 'block', marginTop: 8, color: tokens.colorNeutralForeground3 }}
          >
            Tip: use the search box to filter by primary name (e.g. account name, contact full
            name).
          </Caption1>
        )}
      </div>
    </div>
  );
}

function OptionsPane({
  performParentingChecks,
  setPerformParentingChecks,
  suppressDuplicateDetection,
  setSuppressDuplicateDetection,
}: {
  performParentingChecks: boolean;
  setPerformParentingChecks: (v: boolean) => void;
  suppressDuplicateDetection: boolean;
  setSuppressDuplicateDetection: (v: boolean) => void;
}) {
  return (
    <div>
      <PaneHead
        icon={Settings20Filled}
        title="Merge options"
        sub="Controls applied during the merge operation."
        group="write"
      />
      <div style={{ maxWidth: 720, display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div
          style={{
            padding: 12,
            border: `1px solid ${tokens.colorNeutralStroke2}`,
            borderRadius: tokens.borderRadiusMedium,
          }}
        >
          <Switch
            checked={performParentingChecks}
            onChange={(_, d) => setPerformParentingChecks(d.checked)}
            label={
              <span>
                <strong>PerformParentingChecks</strong>
                <Caption1
                  style={{ display: 'block', color: tokens.colorNeutralForeground3, marginTop: 2 }}
                >
                  When enabled, the merge fails if the Target and Subordinate have different parent
                  records. Useful for child entities (e.g. contacts under different accounts).
                </Caption1>
              </span>
            }
          />
        </div>

        <div
          style={{
            padding: 12,
            border: `1px solid ${tokens.colorNeutralStroke2}`,
            borderRadius: tokens.borderRadiusMedium,
          }}
        >
          <Switch
            checked={suppressDuplicateDetection}
            onChange={(_, d) => setSuppressDuplicateDetection(d.checked)}
            label={
              <span>
                <strong>MSCRM.SuppressDuplicateDetection</strong>{' '}
                <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>= true</code>
                <Caption1
                  style={{ display: 'block', color: tokens.colorNeutralForeground3, marginTop: 2 }}
                >
                  Skip duplicate detection rules when applying <code>UpdateContent</code> to the
                  Target. Adds the <code>MSCRM.SuppressDuplicateDetection: true</code> request
                  header.
                </Caption1>
              </span>
            }
          />
        </div>

        <MessageBar layout="multiline" intent="warning">
          <MessageBarBody>
            <MessageBarTitle>Merge is destructive.</MessageBarTitle>
            The Subordinate is deactivated (<code>statecode=Inactive</code>) and most of its
            relationships re-parent to the Target. Some Target fields take Subordinate's value only
            when the Target's value is empty. Use <strong>UpdateContent</strong> (the Field
            comparison pane) to override that.
          </MessageBarBody>
        </MessageBar>
      </div>
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Helpers
// ──────────────────────────────────────────────────────────────
function primaryNameOf(
  tbl: import('../mock/metadata').TableMeta | undefined,
  row: Record<string, unknown> | undefined,
): string | null {
  if (!tbl || !row) return null;
  const v = row[tbl.primaryName];
  return v != null ? String(v) : null;
}

// Suppress unused-import warning
void Button;
