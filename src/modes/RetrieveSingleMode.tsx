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

import { useEffect, useMemo, useState } from 'react';
import {
  Table20Regular,
  Table20Filled,
  TextBulletList20Regular,
  TextBulletList20Filled,
  Filter20Regular,
  Filter20Filled,
  TextSortAscending20Regular,
  TextSortAscending20Filled,
  NumberSymbol20Regular,
  NumberSymbol20Filled,
  BranchFork20Regular,
  BranchFork20Filled,
  Settings20Regular,
  Settings20Filled,
  LineHorizontal320Regular,
  LineHorizontal320Filled,
  Link20Regular,
} from '@fluentui/react-icons';
import { tokens } from '@fluentui/react-components';
import { Sidebar, type SidebarClauseItem } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { notify } from '../shell/dialogs';
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
import { relatedMetadataAdvisories } from '../engine/navPath';
import { runtime, type ExecResult } from '../engine/runtime';
import type { ExpandSpec } from '../editors/ExpandEditor';
import type { RetrieveSingleState, RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useLiveTable, useWarmReferencedTables } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import {
  serializeRetrieveSingle,
  hashState,
  deserializeRetrieveSingle,
  type SavedRequest,
  type SerializedRetrieveSingleState,
  recentRunAsSavedEntry,
} from '../state/savedRequests';
import { useModeDraft, useStashModeDraft } from '../state/modeDrafts';
import { usePublishSaveContext } from '../state/SaveContext';
import { useLatestCallback } from '../primitives/useLatestCallback';
import { takeModeHandoff, useImportedRequest } from '../state/modeHandoff';
import type { Advisory } from '../primitives/advisories';
import { isRecordGuid } from '../primitives/guid';
import { useRecentRuns } from '../state/runHistory';
import { importRetrieveSingle } from '../engine/curlImport/applyImport';

const initialState = (): RetrieveSingleState => ({
  table: '',
  recordId: null,
  select: [],
  expand: [],
  prefer: { ...emptyPrefer(), formattedValues: true },
  // If-None-Match used to be on (`null`) by default. Execute sends it now
  // (#109), and `null` never matches an etag, so it's off until you add one.
  headers: defaultReadHeaders(),
  dirty: new Set(),
});

type RootClauseId = 'target' | 'select' | 'prefer' | 'headers';

// GUID validator — the OData key needs the dashed 8-4-4-4-12 form (RecordPicker
// normalizes what the user types to it).
const isValidGuid = isRecordGuid;

export function RetrieveSingleMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('retrieve-single');
  const draft = useModeDraft('retrieve-single', deserializeRetrieveSingle, initialState);
  const [state, setState] = useState(draft.state);
  // Live-metadata subscription so child editors re-render when columns/relationships land.
  // The live table is a memo dependency below: build*() reads metadata via
  // findTable(), which is empty until the fetch lands — keying on `state`
  // alone left the URL / Code tab blank after picking a table (#78).
  const { table: liveTable } = useLiveTable(state.table || null);
  // Pre-warm entities referenced by nested $expand (inner $select/$filter/
  // $orderby) so encoders resolve against the right related entity — even on
  // saved-request reload / pasted URL where no expand editor was opened.
  // `related.version` re-encodes `built` when a related table lands (#63).
  const related = useWarmReferencedTables(state.table || null, { expand: state.expand });
  // Every table, not just the pickers' scope: a saved request's table only
  // has to exist (#44).
  const { allEntities: entities } = useScopedEntities();
  const [activePath, setActivePath] = useState<string>('target');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  // The URL the shown result came from: the builder may have changed since (#291).
  const [runUrl, setRunUrl] = useState<string | undefined>(undefined);
  const [loading, setLoading] = useState(false);
  // Kept across reloads, per environment (#226).
  const [recents, setRecents] = useRecentRuns('retrieve-single');

  // Last-saved tracking — drives the Save button's primary/subtle state
  // and lights up the checkmark next to the current entry in the library.
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(draft.lastSavedId);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(draft.lastSavedHash);

  const built = useMemo(
    () => buildRetrieveSingle(state),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [state, liveTable, related.version],
  );
  const tbl = findTable(state.table);

  // Hash + dirty are derived from the live state every render. Same identity
  // discipline as RetrieveMultiple: a string compare is cheap and avoids
  // stale-baseline bugs that creep in if we cache a serialized snapshot.
  const currentSerialized = useMemo(() => serializeRetrieveSingle(state), [state]);
  useStashModeDraft('retrieve-single', currentSerialized, lastSavedId, lastSavedHash);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;

  const markDirty = (id: string) =>
    setState((s) => {
      const d = new Set(s.dirty);
      d.add(id);
      return { ...s, dirty: d };
    });
  const set = <K extends keyof RetrieveSingleState>(
    k: K,
    v: RetrieveSingleState[K],
    dirtyId?: string,
  ) => {
    setState((s) => ({ ...s, [k]: v }));
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
      const known = entities.some((e) => e.logicalName === snap.table);
      if (!known) {
        notify(
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

  // Recent runs (#2) — restore a run's snapshot through the same loader as
  // the Saved library, but keep the current save baseline: a run isn't a
  // saved entry, so Save keeps targeting whatever it targeted before.
  const onRecentClick = (run: RecentRun) => {
    const entry = recentRunAsSavedEntry(run, 'retrieve-single');
    if (!entry) return;
    const keepId = lastSavedId;
    const keepHash = lastSavedHash;
    onLoadSaved(entry);
    setLastSavedId(keepId);
    setLastSavedHash(keepHash);
  };

  // Publish our save context to the FrameHeader's Save / Saved-library
  // buttons. Hidden automatically when no table is picked (nothing useful
  // to save yet).
  // The header holds on to these; the stable wrappers always run this
  // render's handlers, not the ones from when the context was published.
  const onSavedLatest = useLatestCallback(onSaved);
  // cURL import (#271): a command pasted anywhere opens here. Like a recent
  // run, it isn't a saved entry, so the save baseline stays as it was.
  useImportedRequest('retrieve-single', (request) => {
    if (request.modeId !== 'retrieve-single') return;
    setState(importRetrieveSingle(initialState(), request));
    setResult(null);
    setTab('builder');
    setActivePath('target');
  });
  const onLoadSavedLatest = useLatestCallback(onLoadSaved);
  usePublishSaveContext(
    useMemo(() => {
      if (!state.table) return null;
      return {
        state: currentSerialized,
        modeId: 'retrieve-single' as const,
        dirty: isDirty,
        lastSavedId,
        onSaved: onSavedLatest,
        onLoadSaved: onLoadSavedLatest,
      };
    }, [currentSerialized, isDirty, lastSavedId, state.table, onSavedLatest, onLoadSavedLatest]),
  );

  // ── Advisories ────────────────────────────────────────────────────
  // RetrieveSingle's gates are tight — a missing or malformed recordId is
  // the only common error class. Surfacing it via the URL bar drawer keeps
  // the failure mode consistent with the bigger modes (no silent disable).
  const relatedPendingKey = related.pending.join(',');
  const relatedFailedKey = related.failed.join(',');
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
    out.push(...relatedMetadataAdvisories(related.pending, related.failed));
    return out;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [state.table, state.recordId, tbl?.displayName, relatedPendingKey, relatedFailedKey]);

  // ── Execute ───────────────────────────────────────────────────────
  const onExecute = async () => {
    setLoading(true);
    setRunUrl(built.relativeUrl);
    const res = await runtime.retrieveSingle(state, headerItemsToObject(state.headers, null));
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents((rs) =>
      [
        {
          id: `r-${Date.now()}`,
          snapshot: currentSerialized,
          modeId: 'retrieve-single',
          url: built.relativeUrl,
          method: 'GET',
          ts: Date.now(),
          status: res.status,
          ms: res.ms,
          rowCount: res.outcome === 'ok' ? 1 : 0,
        },
        ...rs,
      ].slice(0, 8),
    );
    setState((s) => ({ ...s, dirty: new Set() }));
  };

  const disabledReason = !state.table
    ? 'Pick a table first.'
    : !state.recordId
      ? 'Pick a record first.'
      : !isValidGuid(state.recordId)
        ? 'Record id is not a valid GUID.'
        : state.headers.some((h) => h.enabled && !h.name)
          ? 'Fix empty header name.'
          : null;

  // ── Handoff from another mode (#3) ────────────────────────────────
  // WriteResultCard's "Retrieve this record" switches here with a table +
  // record id waiting. Take it once on mount, start a fresh request for that
  // record (not saved, so no Save baseline), and run it as soon as the
  // table's metadata has landed — the URL needs its entity set name.
  const [autoRunPending, setAutoRunPending] = useState(false);
  useEffect(() => {
    const handoff = takeModeHandoff('retrieve-single');
    if (!handoff) return;
    setState({ ...initialState(), table: handoff.table, recordId: handoff.recordId });
    setLastSavedId(undefined);
    setLastSavedHash(null);
    setResult(null);
    setActivePath('target');
    setAutoRunPending(!!handoff.autoExecute);
  }, []);
  useEffect(() => {
    if (!autoRunPending || loading) return;
    if (!liveTable || !tbl?.entitySetName || disabledReason) return;
    setAutoRunPending(false);
    void onExecute();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [autoRunPending, liveTable, tbl?.entitySetName, disabledReason, loading]);

  // ── Sidebar tree for $expand (recursive) ──────────────────────────
  // Same shape + visibility rules as Retrieve Multiple (see comment block
  // there for the source-of-truth rationale). Kept duplicated here rather
  // than extracted because the two modes will diverge later (Single might
  // gain $apply-style affordances; Multiple already has).
  function expandTreeItems(
    items: ExpandSpec[],
    pathPrefix: string,
    parentEntityLogical: string,
  ): SidebarClauseItem[] {
    const parentTbl = findTable(parentEntityLogical);
    return items.map((it) => {
      const nav = parentTbl?.navigationProperties.find((n) => n.name === it.nav);
      const target = nav ? findTable(nav.targetEntity) : undefined;
      const card = nav?.cardinality;
      const isCollection = card === 'OneToMany' || card === 'ManyToMany';
      const isNN = card === 'ManyToMany';
      const navPath = `${pathPrefix}/${it.id}`;
      const innerFilterCount = it.filter ? countRules(it.filter) : 0;
      const children: SidebarClauseItem[] = [];

      children.push({
        id: `${navPath}/select`,
        icon: TextBulletList20Regular,
        iconFilled: TextBulletList20Filled,
        label: '$select',
        code: true,
        badge: it.select.length || null,
      });
      children.push({
        id: `${navPath}/filter`,
        icon: Filter20Regular,
        iconFilled: Filter20Filled,
        label: '$filter',
        code: true,
        badge: innerFilterCount || null,
      });
      if (isCollection) {
        children.push({
          id: `${navPath}/orderby`,
          icon: TextSortAscending20Regular,
          iconFilled: TextSortAscending20Filled,
          label: '$orderby',
          code: true,
          badge: it.orderby.length || null,
        });
        children.push({
          id: `${navPath}/top`,
          icon: NumberSymbol20Regular,
          iconFilled: NumberSymbol20Filled,
          label: '$top',
          code: true,
          badge: it.top ? it.top.toString() : null,
          badgeAppearance: 'ghost' as const,
        });
      }
      if (!isNN) {
        children.push({
          id: `${navPath}/expand`,
          icon: BranchFork20Regular,
          iconFilled: BranchFork20Filled,
          label: '$expand',
          code: true,
          badge: it.nestedExpand?.length || null,
          children:
            target && it.nestedExpand?.length
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
            <span style={{ color: tokens.colorNeutralForeground3, fontWeight: 400 }}>
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
      id: 'target',
      label: 'Target',
      meta: tbl ? `${tbl.displayName} record` : 'Pick a table',
      items: [
        {
          id: 'target',
          icon: Table20Regular,
          iconFilled: Table20Filled,
          label: state.recordId ? `${tbl?.displayName ?? ''} (selected)` : 'Pick a record',
          badge: state.recordId ? (isValidGuid(state.recordId) ? '✓' : '⚠') : null,
          badgeAppearance: 'ghost' as const,
          badgeColor:
            state.recordId && !isValidGuid(state.recordId) ? ('danger' as const) : undefined,
          dirty: state.dirty.has('target'),
        },
      ],
    },
    {
      id: 'anatomy',
      label: 'Request anatomy',
      items: [
        {
          id: 'select',
          icon: TextBulletList20Regular,
          iconFilled: TextBulletList20Filled,
          label: '$select',
          code: true,
          badge: state.select.length || null,
          dirty: state.dirty.has('select'),
        },
        {
          id: 'expand',
          icon: BranchFork20Regular,
          iconFilled: BranchFork20Filled,
          label: '$expand',
          code: true,
          badge: state.expand.length || null,
          dirty: state.dirty.has('expand'),
          children: expandSidebarChildren.length ? expandSidebarChildren : undefined,
        },
      ],
    },
    {
      id: 'prefer',
      label: 'Prefer',
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
      ],
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
      case 'target':
        pane = (
          <TargetEditor
            table={state.table}
            onTableChange={(t) => {
              // Switching/clearing the target entity invalidates every
              // column-bound clause for this mode (select + expand) plus
              // the record id (no longer valid on the new entity).
              setState((s) => ({
                ...s,
                table: t,
                recordId: null,
                select: [],
                expand: [],
                dirty: new Set(['target']),
              }));
              setResult(null);
              // New table → fresh draft. Forget the last save so Save can't
              // overwrite an entry that belongs to the old table (#58).
              setLastSavedId(undefined);
              setLastSavedHash(null);
            }}
            recordId={state.recordId}
            onRecordChange={(id) => set('recordId', id, 'target')}
            sub="Single-record fetch — pick the entity set and the record GUID."
            onResetRequest={() => {
              setState((s) => ({
                ...s,
                select: [],
                expand: [],
                dirty: new Set(['select', 'expand']),
              }));
            }}
          />
        );
        break;
      case 'select':
        pane = (
          <SelectEditor
            table={state.table}
            selectedIds={state.select}
            setSelectedIds={(ids) => set('select', ids, 'select')}
          />
        );
        break;
      case 'prefer':
        pane = <PreferEditor spec={state.prefer} setSpec={(p) => set('prefer', p, 'prefer')} />;
        break;
      case 'headers':
        pane = (
          <HeadersEditor
            policy="read"
            items={state.headers}
            setItems={(h) => set('headers', h, 'headers')}
          />
        );
        break;
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
          onSelect={(id) => {
            setActivePath(id);
            setTab('builder');
          }}
          recents={recents}
          onRecentClick={onRecentClick}
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
        {tab === 'code' && <CodeView themeMode={themeMode} inputs={codeInputs} />}
        {tab === 'results' && (
          <ResultsView result={result} mode="single" table={state.table} requestUrl={runUrl} />
        )}
      </MainTabs>
    </ModeShell>
  );
}
