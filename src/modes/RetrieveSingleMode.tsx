// RetrieveSingle — fetch ONE record by primary key.
//
// State shape: { table, recordId, select, expand, prefer, headers }.
// No filter / orderby / top / count / apply — those are collection-only.
//
// Mostly composes RetrieveMultiple's primitives:
//   • TargetEditor — table picker + RecordPicker (live typeahead)
//   • SelectEditor — DataGrid with multi-select
//   • ExpandRouter — $expand tree
//   • PreferEditor + HeadersEditor — shared
//
// Adds the record-picker affordance and the only new validation:
//   • recordId is required (Execute blocks until set)
//   • recordId must look like a GUID
//
// Save/Load: wired via the unified SaveContext. RetrieveSingle's saved
// snapshots live in the same per-org bucket as RetrieveMultiple's and
// are distinguished by `modeId: 'retrieve-single'`. The library popover
// filters by mode so each mode sees only its own saves.

import { useMemo, useState } from 'react';
import {
  Table20Regular, Table20Filled,
  TextBulletList20Regular, TextBulletList20Filled,
  Filter20Regular, Filter20Filled,
  TextSortAscending20Regular, TextSortAscending20Filled,
  NumberSymbol20Regular, NumberSymbol20Filled,
  BranchFork20Regular, BranchFork20Filled,
  Settings20Regular, Settings20Filled,
  LineHorizontal320Regular, LineHorizontal320Filled,
  Link20Regular,
} from '@fluentui/react-icons';
import { Sidebar, type SidebarClauseItem } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { TargetEditor } from '../editors/TargetEditor';
import { SelectEditor } from '../editors/SelectEditor';
import { PreferEditor, emptyPrefer, preferToHeaderString } from '../editors/PreferEditor';
import { HeadersEditor, defaultReadHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { ExpandRouter } from '../editors/expand/ExpandRouter';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { countRules } from '../editors/filter/filterTree';
import { findTable } from '../mock/metadata';
import { findRequestType } from '../registry/requestTypes';
import { buildRetrieveSingle } from '../engine/urlBuilder';
import { runtime, type ExecResult } from '../engine/runtime';
import type { ExpandSpec } from '../editors/ExpandEditor';
import type { RetrieveSingleState, RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useLiveTable, useWarmReferencedTables } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import {
  serializeRetrieveSingle, hashState, deserializeRetrieveSingle,
  type SavedRequest, type SerializedRetrieveSingleState,
} from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';
import type { Advisory } from '../primitives/advisories';

const initialState = (): RetrieveSingleState => ({
  table: '',
  recordId: null,
  select: [],
  expand: [],
  prefer: { ...emptyPrefer(), formattedValues: true },
  headers: defaultReadHeaders().map(h => h.name === 'If-None-Match' ? { ...h, enabled: true } : h),
  dirty: new Set(),
});

type RootClauseId = 'target' | 'select' | 'prefer' | 'headers';

// GUID validator — Dataverse accepts braced, hyphenated, or hex-only forms.
const GUID_RE = /^[{(]?[0-9a-fA-F]{8}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{12}[)}]?$/;
const isValidGuid = (s: string | null | undefined): boolean => !!s && GUID_RE.test(s.trim());

export function RetrieveSingleMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('retrieve-single');
  const [state, setState] = useState(initialState);
  // Live-metadata subscription so child editors re-render when columns/relationships land.
  useLiveTable(state.table || null);
  // Pre-warm entities referenced by nested $expand (inner $select/$filter/
  // $orderby) so encoders resolve against the right related entity — even on
  // saved-request reload / pasted URL where no expand editor was opened.
  useWarmReferencedTables(state.table || null, { expand: state.expand });
  const { entities } = useScopedEntities();
  const [activePath, setActivePath] = useState<string>('target');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);

  // Last-saved tracking — drives the Save button's primary/subtle state
  // and lights up the checkmark next to the current entry in the library.
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);

  const built = useMemo(() => buildRetrieveSingle(state), [state]);
  const tbl = findTable(state.table);

  // Hash + dirty are derived from the live state every render. Same identity
  // discipline as RetrieveMultiple: a string compare is cheap and avoids
  // stale-baseline bugs that creep in if we cache a serialized snapshot.
  const currentSerialized = useMemo(() => serializeRetrieveSingle(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;

  const markDirty = (id: string) => setState(s => { const d = new Set(s.dirty); d.add(id); return { ...s, dirty: d }; });
  const set = <K extends keyof RetrieveSingleState>(k: K, v: RetrieveSingleState[K], dirtyId?: string) => {
    setState(s => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  // ── Save / Load ───────────────────────────────────────────────────
  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };

  const onLoadSaved = (entry: SavedRequest) => {
    // Defensive routing — the library popover already filters by modeId,
    // but if someone hand-edits the bucket we don't want to feed the
    // wrong shape into our deserializer.
    if (entry.modeId !== 'retrieve-single') return;
    const snap = entry.state as SerializedRetrieveSingleState;

    // Lenient validation — only block when entities have actually loaded
    // and the saved table is provably gone. Empty entities list ⇒ still
    // warming up; hydrate optimistically and let the metadata fetch
    // surface a clearer error if the entity really is missing.
    const entitiesLoaded = entities.length > 0;
    if (entitiesLoaded) {
      const known = entities.some(e => e.logicalName === snap.table);
      if (!known) {
        window.alert(
          `Can't load "${entry.name}": entity \`${snap.table}\` ` +
          `isn't available in this environment. The solution may have been ` +
          `removed or you may be connected to a different org.`,
        );
        return;
      }
    }

    setState(deserializeRetrieveSingle(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setActivePath('target');
  };

  // Publish our save context to the FrameHeader's Save / Saved-library
  // buttons. Hidden automatically when no table is picked (nothing useful
  // to save yet).
  usePublishSaveContext(useMemo(() => {
    if (!state.table) return null;
    return {
      state: currentSerialized,
      modeId: 'retrieve-single' as const,
      dirty: isDirty,
      lastSavedId,
      onSaved,
      onLoadSaved,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentSerialized, isDirty, lastSavedId, state.table]));

  // ── Advisories ────────────────────────────────────────────────────
  // RetrieveSingle's gates are tight — a missing or malformed recordId is
  // the only common error class. Surfacing it via the URL bar drawer keeps
  // the failure mode consistent with the bigger modes (no silent disable).
  const advisories: Advisory[] = useMemo(() => {
    const out: Advisory[] = [];
    if (state.table && !state.recordId) {
      out.push({
        id: 'rs-no-record',
        severity: 'error',
        source: 'validation',
        focusNode: 'target',
        title: 'No record selected',
        body: `Pick a ${tbl?.displayName ?? 'record'} or paste a GUID before executing.`,
      });
    }
    if (state.recordId && !isValidGuid(state.recordId)) {
      out.push({
        id: 'rs-bad-guid',
        severity: 'error',
        source: 'validation',
        focusNode: 'target',
        title: 'Invalid record id',
        body: `"${state.recordId}" doesn't look like a GUID. Dataverse record ids are 36-char UUIDs (with or without braces).`,
      });
    }
    return out;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [state.table, state.recordId, tbl?.displayName]);

  // ── Execute ───────────────────────────────────────────────────────
  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.retrieveSingle(state);
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents(rs => [{
      id: `r-${Date.now()}`, modeId: 'retrieve-single',
      url: built.relativeUrl, method: 'GET', ts: Date.now(),
      status: res.status, ms: res.ms, rowCount: res.outcome === 'ok' ? 1 : 0,
    }, ...rs].slice(0, 8));
    setState(s => ({ ...s, dirty: new Set() }));
  };

  const disabledReason =
    !state.table ? 'Pick a table first.' :
    !state.recordId ? 'Pick a record first.' :
    !isValidGuid(state.recordId) ? 'Record id is not a valid GUID.' :
    state.headers.some(h => h.enabled && !h.name) ? 'Fix empty header name.' :
    null;

  // ── Sidebar tree for $expand (recursive) ──────────────────────────
  // Same shape + visibility rules as Retrieve Multiple (see comment block
  // there for the source-of-truth rationale). Kept duplicated here rather
  // than extracted because the two modes will diverge later (Single might
  // gain $apply-style affordances; Multiple already has).
  function expandTreeItems(items: ExpandSpec[], pathPrefix: string, parentEntityLogical: string): SidebarClauseItem[] {
    const parentTbl = findTable(parentEntityLogical);
    return items.map(it => {
      const nav = parentTbl?.navigationProperties.find(n => n.name === it.nav);
      const target = nav ? findTable(nav.targetEntity) : undefined;
      const card = nav?.cardinality;
      const isCollection = card === 'OneToMany' || card === 'ManyToMany';
      const isNN = card === 'ManyToMany';
      const navPath = `${pathPrefix}/${it.id}`;
      const innerFilterCount = it.filter ? countRules(it.filter) : 0;
      const children: SidebarClauseItem[] = [];

      children.push({
        id: `${navPath}/select`,
        icon: TextBulletList20Regular, iconFilled: TextBulletList20Filled,
        label: '$select', code: true, badge: it.select.length || null,
      });
      children.push({
        id: `${navPath}/filter`,
        icon: Filter20Regular, iconFilled: Filter20Filled,
        label: '$filter', code: true, badge: innerFilterCount || null,
      });
      if (isCollection) {
        children.push({
          id: `${navPath}/orderby`,
          icon: TextSortAscending20Regular, iconFilled: TextSortAscending20Filled,
          label: '$orderby', code: true, badge: it.orderby.length || null,
        });
        children.push({
          id: `${navPath}/top`,
          icon: NumberSymbol20Regular, iconFilled: NumberSymbol20Filled,
          label: '$top', code: true,
          badge: it.top ? it.top.toString() : null, badgeAppearance: 'ghost' as const,
        });
      }
      if (!isNN) {
        children.push({
          id: `${navPath}/expand`,
          icon: BranchFork20Regular, iconFilled: BranchFork20Filled,
          label: '$expand', code: true,
          badge: it.nestedExpand?.length || null,
          children: target && it.nestedExpand?.length
            ? expandTreeItems(it.nestedExpand, `${navPath}/expand`, target.logicalName)
            : undefined,
        });
      }
      return {
        id: navPath,
        icon: Link20Regular,
        label: (
          <span>
            {it.nav}{' '}
            <span style={{ opacity: 0.55, fontWeight: 400 }}>
              → {target?.displayName ?? nav?.targetEntity ?? '?'}
            </span>
          </span>
        ),
        badge: isNN ? 'N:N' : isCollection ? 'many' : 'one',
        badgeAppearance: 'ghost' as const,
        children,
      };
    });
  }

  const expandSidebarChildren = expandTreeItems(state.expand, 'expand', state.table);

  const sections = [
    {
      id: 'target', label: 'Target',
      meta: tbl ? `${tbl.displayName} record` : 'Pick a table',
      items: [{
        id: 'target',
        icon: Table20Regular, iconFilled: Table20Filled,
        label: state.recordId ? `${tbl?.displayName ?? ''} (selected)` : 'Pick a record',
        badge: state.recordId
          ? (isValidGuid(state.recordId) ? '✓' : '⚠')
          : null,
        badgeAppearance: 'ghost' as const,
        badgeColor: state.recordId && !isValidGuid(state.recordId) ? ('danger' as const) : undefined,
        dirty: state.dirty.has('target'),
      }],
    },
    {
      id: 'anatomy', label: 'Request anatomy',
      items: [
        { id: 'select', icon: TextBulletList20Regular, iconFilled: TextBulletList20Filled, label: '$select', code: true, badge: state.select.length || null, dirty: state.dirty.has('select') },
        {
          id: 'expand', icon: BranchFork20Regular, iconFilled: BranchFork20Filled,
          label: '$expand', code: true,
          badge: state.expand.length || null,
          dirty: state.dirty.has('expand'),
          children: expandSidebarChildren.length ? expandSidebarChildren : undefined,
        },
      ],
    },
    {
      id: 'prefer', label: 'Prefer',
      items: [{
        id: 'prefer',
        icon: Settings20Regular, iconFilled: Settings20Filled,
        label: 'Prefer header',
        badge: preferToHeaderString(state.prefer) ? 'on' : null, badgeAppearance: 'ghost' as const,
        dirty: state.dirty.has('prefer'),
      }],
    },
    {
      id: 'headers', label: 'Headers',
      meta: `${state.headers.filter(h => h.enabled).length} active`,
      items: [{
        id: 'headers',
        icon: LineHorizontal320Regular, iconFilled: LineHorizontal320Filled,
        label: 'HTTP headers',
        badge: state.headers.filter(h => h.enabled).length || null,
        dirty: state.dirty.has('headers'),
      }],
    },
  ];

  let pane: React.ReactNode;
  if (activePath.startsWith('expand')) {
    pane = (
      <ExpandRouter
        path={activePath}
        rootEntity={state.table}
        rootExpand={state.expand}
        setRootExpand={(items) => set('expand', items, 'expand')}
        setActivePath={setActivePath}
      />
    );
  } else {
    const root = activePath as RootClauseId;
    switch (root) {
      case 'target':  pane = <TargetEditor
        table={state.table}
        onTableChange={t => {
          // Switching/clearing the target entity invalidates every
          // column-bound clause for this mode (select + expand) plus
          // the record id (no longer valid on the new entity).
          setState(s => ({
            ...s,
            table: t,
            recordId: null,
            select: [],
            expand: [],
            dirty: new Set(['target']),
          }));
          setResult(null);
        }}
        recordId={state.recordId}
        onRecordChange={id => set('recordId', id, 'target')}
        sub="Single-record fetch — pick the entity set and the record GUID."
        onResetRequest={() => {
          setState(s => ({
            ...s,
            select: [],
            expand: [],
            dirty: new Set(['select', 'expand']),
          }));
        }}
      />; break;
      case 'select':  pane = <SelectEditor table={state.table} selectedIds={state.select} setSelectedIds={ids => set('select', ids, 'select')} />; break;
      case 'prefer':  pane = <PreferEditor spec={state.prefer} setSpec={p => set('prefer', p, 'prefer')} />; break;
      case 'headers': pane = <HeadersEditor items={state.headers} setItems={h => set('headers', h, 'headers')} />; break;
    }
  }

  const headersMap = headerItemsToObject(state.headers, preferToHeaderString(state.prefer));
  const codeInputs = { method: 'GET', built, headers: headersMap };

  return (
    <ModeShell
      sidebar={
        <Sidebar
          type={type}
          urlPreview={built.relativeNoBase}
          sections={sections}
          activeNode={activePath}
          onSelect={(id) => setActivePath(id)}
          recents={recents}
        />
      }
      urlBar={
        <UrlBar
          method="GET"
          url={built.relativeUrl}
          executeVerb={type.executeVerb}
          disabledReason={disabledReason}
          loading={loading}
          onExecute={onExecute}
          advisories={advisories}
          onAdvisoryFocus={(nodeId) => setActivePath(nodeId)}
        />
      }
    >
      <MainTabs tab={tab} onTabChange={setTab} resultCount={result?.outcome === 'ok' ? 1 : null}>
        {tab === 'builder' && pane}
        {tab === 'code'    && <CodeView themeMode={themeMode} inputs={codeInputs} />}
        {tab === 'results' && <ResultsView result={result} mode="single" table={state.table} />}
      </MainTabs>
    </ModeShell>
  );
}
