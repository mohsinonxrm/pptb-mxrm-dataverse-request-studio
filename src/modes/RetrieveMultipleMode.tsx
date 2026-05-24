import { useMemo, useState } from 'react';
import {
  Table20Regular, Table20Filled,
  TextBulletList20Regular, TextBulletList20Filled,
  Filter20Regular, Filter20Filled,
  TextSortAscending20Regular, TextSortAscending20Filled,
  NumberSymbol20Regular, NumberSymbol20Filled,
  Tag20Regular, Tag20Filled,
  BranchFork20Regular, BranchFork20Filled,
  Settings20Regular, Settings20Filled,
  LineHorizontal320Regular, LineHorizontal320Filled,
  ChartMultiple20Regular, ChartMultiple20Filled,
  Link20Regular,
} from '@fluentui/react-icons';
import { Sidebar, type SidebarClauseItem } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { detectRetrieveMultipleAntipatterns } from '../engine/antipatterns';
import { collectStrippedWildcards } from '../editors/filter/filterTree';
import { disabledReasonFromAdvisories, type Advisory } from '../primitives/advisories';
import { validateRequest } from '../state/validateRequest';
import { ModeShell } from '../shell/ModeShell';
import { TargetEditor } from '../editors/TargetEditor';
import { SelectEditor } from '../editors/SelectEditor';
import { FilterEditor } from '../editors/filter/FilterEditor';
import { OrderbyEditor } from '../editors/OrderbyEditor';
import { TopEditor, CountEditor } from '../editors/TopCountEditors';
import { ApplyEditor, emptyApply } from '../editors/ApplyEditor';
import { PreferEditor, emptyPrefer, preferToHeaderString } from '../editors/PreferEditor';
import { HeadersEditor, defaultReadHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { ExpandRouter } from '../editors/expand/ExpandRouter';
import { hasCollectionInvolvedNestedExpand } from '../editors/expand/expandTree';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { countRules, newId } from '../editors/filter/filterTree';
import { findTable } from '../mock/metadata';
import { findRequestType } from '../registry/requestTypes';
import { buildRetrieveMultiple } from '../engine/urlBuilder';
// Phase 1 wiring: routes through `runtime.retrieveMultiple` which auto-picks
// the live (dvHost.queryData) or mock executor based on `isEmbedded()`. The
// mode doesn't need to know which.
import { runtime, type ExecResult } from '../engine/runtime';
import type { ExpandSpec } from '../editors/ExpandEditor';
import type { RetrieveMultipleState, RecentRun } from '../state/readState';
import { serializeRetrieveMultiple, hashState, deserializeRetrieveMultiple, type SavedRequest, type SerializedRetrieveMultipleState } from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';
import type { ThemeMode } from '../theme/theme';
// Live-metadata hook: subscribes to the registry so child editors
// (SelectEditor, FilterEditor, OrderbyEditor, ExpandEditor, lambdas)
// re-render when columns + relationships land from Dataverse.
import { useLiveTable } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';

// Empty initial state — the studio is a request *builder*; seeding it with
// mock data confused users in PPTB ("why are there pre-filled rules against
// a table I don't recognize?"). User picks the target table first; column-
// driven clauses ($select / $filter / $orderby / $expand) come from live
// metadata once a table is chosen.
//
// `top: 50` stays as a sensible default — it's a query-shape preference,
// not metadata that would mislead anyone.
const emptyFilter = () => ({
  id: newId('root'), type: 'group' as const, combinator: 'and' as const, rules: [],
});

const initialState = (): RetrieveMultipleState => ({
  table: '',
  select: [],
  filter: emptyFilter(),
  orderby: [],
  top: 50,
  countOn: false,
  expand: [],
  apply: emptyApply(),
  prefer: emptyPrefer(),
  headers: defaultReadHeaders(),
  dirty: new Set(),
});

// Active sidebar paths — strings so the $expand tree can encode arbitrary depth
// via slash-delimited segments like `expand/<id>/select` or
// `expand/<id>/expand/<id>/filter`.
type RootClauseId =
  | 'target' | 'select' | 'filter' | 'orderby' | 'top' | 'count'
  | 'apply' | 'prefer' | 'headers';

export function RetrieveMultipleMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('retrieve-multiple');
  const [state, setState] = useState(initialState);
  // Default the active sidebar to 'target' so the user lands on the table
  // picker. Previously this opened on '$filter', which made no sense when
  // the initial state is empty (you can't filter against a table you haven't
  // picked yet).
  const [activePath, setActivePath] = useState<string>('target');

  // ── Live-metadata subscription ──────────────────────────────────────
  // Every child editor (SelectEditor, FilterEditor, OrderbyEditor,
  // ExpandEditor, lambda inner-editors, …) reads columns + relationships
  // via the synchronous `findTable()` API. That API consults a live cache
  // populated by `metadata.getTable(...)` — but the cache is async, so the
  // first render hits the mock fallback (or empty when the table doesn't
  // exist in mock). Without a subscription at the MODE level, child
  // editors don't re-render when live data lands.
  //
  // `useLiveTable(state.table)` does two things:
  //   1. Fires the async fetch on mount + whenever `state.table` changes.
  //   2. Subscribes to the live registry, triggering a re-render here when
  //      the fetch resolves — which cascades to every child editor.
  //
  // The hook intentionally returns nothing we use here; its side effect
  // (subscription + warming) is the point.
  useLiveTable(state.table || null);
  // Entity list — used to validate saved-request loads against the
  // current org (block load if the entity has been uninstalled).
  const { entities } = useScopedEntities();
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);

  // ── Saved-request persistence (localStorage) ──
  // Track the last-saved id + the hash of state at that save so we can
  // tell "the request has unsaved changes since last save". When dirty,
  // the SaveButton in the URL bar lights up; when clean, it disables.
  // Reset id on entity change (saved entries are scoped to one entity
  // at a time semantically — switching entities effectively starts a
  // fresh draft that can be saved separately).
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);
  const currentSerialized = useMemo(() => serializeRetrieveMultiple(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  // Dirty when: never saved (no hash baseline) AND state has any meaningful
  // content (table picked), OR saved but the current hash diverges.
  const isDirty = useMemo(() => {
    if (!state.table) return false;          // nothing to save yet
    if (lastSavedHash === null) return true; // fresh draft with a table
    return currentHash !== lastSavedHash;
  }, [state.table, lastSavedHash, currentHash]);

  const built = useMemo(() => buildRetrieveMultiple(state), [state]);
  const tbl = findTable(state.table);
  const filterCount = useMemo(() => countRules(state.filter), [state.filter]);

  const markDirty = (id: string) => setState(s => {
    const d = new Set(s.dirty); d.add(id); return { ...s, dirty: d };
  });

  const set = <K extends keyof RetrieveMultipleState>(k: K, v: RetrieveMultipleState[K], dirtyId?: string) => {
    setState(s => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  // ── Saved-request save / load handlers ──
  const onSaved = (entry: SavedRequest) => {
    // Remember which entry this matches so the next Save (without "As")
    // overwrites instead of creating a new entry, and reset dirty state.
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(entry.state));
  };
  const onLoadSaved = (entry: SavedRequest) => {
    // The library is mode-routed already (modeId filter in the popover),
    // but defend against a manually-hand-crafted bucket import: refuse
    // to hydrate from a non-retrieve-multiple snapshot.
    if (entry.modeId !== 'retrieve-multiple') return;
    // SaveContext erases the state type so storage stays mode-agnostic.
    // Cast back to our serialized shape now that we've routed by modeId.
    const snap = entry.state as SerializedRetrieveMultipleState;

    // Validate the saved entity against the live entity list — BUT only
    // when that list has actually loaded. Without this guard, the load
    // gets blocked right after a fresh page load / mode switch because
    // `entities` is still []: a false negative that infuriates users
    // who JUST saved the request a minute ago. When the list is empty
    // it's a "not loaded yet" signal, not a "entity is gone" signal.
    // Hydrate optimistically; if the entity really is missing, the
    // metadata fetch will fail with a clearer downstream error.
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
    // Hydrate the state. `deserializeRetrieveMultiple` re-inits the
    // `dirty` Set; useLiveTable picks up the entity change.
    setState(deserializeRetrieveMultiple(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setActivePath('target');
  };

  // Publish our save context to the top-right Save / Saved-library
  // buttons in FrameHeader. Each render re-publishes the current values;
  // the hook clears on unmount so a stale mode doesn't bleed across
  // mode switches. The setter is stable, so React only re-runs the
  // effect when the value object identity changes — which we want.
  usePublishSaveContext(useMemo(() => ({
    state: currentSerialized,
    modeId: 'retrieve-multiple' as const,
    dirty: isDirty,
    lastSavedId,
    onSaved,
    onLoadSaved,
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }), [currentSerialized, isDirty, lastSavedId]));

  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.retrieveMultiple(state);
    setResult(res);
    setLoading(false);
    setTab('results');
    const rows = (res.body as { value?: unknown[] } | null)?.value;
    setRecents(rs => [{
      id: `r-${Date.now()}`,
      modeId: 'retrieve-multiple',
      url: built.relativeUrl,
      method: 'GET',
      ts: Date.now(),
      status: res.status,
      ms: res.ms,
      rowCount: Array.isArray(rows) ? rows.length : undefined,
    }, ...rs].slice(0, 8));
    setState(s => ({ ...s, dirty: new Set() }));
  };

  // ── Infinite scroll + Retrieve-All ──
  // Dataverse paginates large result sets via `@odata.nextLink` (the URL
  // for the next page sits on the response envelope). We append next-page
  // rows onto the existing body so the grid keeps everything in one
  // virtualized list — no client-side paging needed.
  const [isLoadingMore, setIsLoadingMore] = useState(false);

  const fetchOneMorePage = async () => {
    const nextLink = (result?.body as { '@odata.nextLink'?: string } | null)?.['@odata.nextLink'];
    if (!nextLink || isLoadingMore) return;
    setIsLoadingMore(true);
    try {
      const more = await runtime.absoluteUrl(nextLink);
      if (!more.ok) { setIsLoadingMore(false); return; }
      const prevRows = ((result?.body as { value?: Record<string, unknown>[] } | null)?.value) ?? [];
      const moreRows = ((more.body as { value?: Record<string, unknown>[] } | null)?.value) ?? [];
      const moreNext = (more.body as { '@odata.nextLink'?: string } | null)?.['@odata.nextLink'];
      // Merge: keep existing envelope, replace `value` + carry the new nextLink (or drop it if exhausted).
      const mergedBody: Record<string, unknown> = {
        ...(result?.body as Record<string, unknown> | null ?? {}),
        value: [...prevRows, ...moreRows],
      };
      if (moreNext) mergedBody['@odata.nextLink'] = moreNext;
      else delete mergedBody['@odata.nextLink'];
      setResult(r => r ? { ...r, body: mergedBody, bytes: JSON.stringify(mergedBody).length } : r);
    } finally {
      setIsLoadingMore(false);
    }
  };

  const retrieveAllPages = async () => {
    if (isLoadingMore) return;
    setIsLoadingMore(true);
    try {
      // Walk the envelope chain; mergedBody mutates in place inside the loop.
      let mergedBody = result?.body as Record<string, unknown> | null;
      let nextLink = (mergedBody as { '@odata.nextLink'?: string } | null)?.['@odata.nextLink'];
      const allRows = ((mergedBody as { value?: Record<string, unknown>[] } | null)?.value ?? []).slice();
      let guard = 0;
      while (nextLink && guard++ < 200) { // hard safety cap
        const more = await runtime.absoluteUrl(nextLink);
        if (!more.ok) break;
        const moreRows = ((more.body as { value?: Record<string, unknown>[] } | null)?.value) ?? [];
        allRows.push(...moreRows);
        nextLink = (more.body as { '@odata.nextLink'?: string } | null)?.['@odata.nextLink'];
      }
      mergedBody = { ...(mergedBody ?? {}), value: allRows };
      if (nextLink) (mergedBody as Record<string, unknown>)['@odata.nextLink'] = nextLink;
      else delete (mergedBody as Record<string, unknown>)['@odata.nextLink'];
      setResult(r => r ? { ...r, body: mergedBody, bytes: JSON.stringify(mergedBody).length } : r);
    } finally {
      setIsLoadingMore(false);
    }
  };

  // ── Advisories — query antipatterns + stripped wildcards ────────────
  // Detection runs on every render (cheap — pure tree walks); the
  // AdvisoryDrawer + URL-bar consume the result.
  const advisories = useMemo(() => {
    // When $apply is on, the top-level $filter is silently ignored by Dataverse
    // and the active filter is the apply.prefilter (filter() stage). Feed the
    // correct one to the wildcard / antipattern detectors.
    const activeFilter = state.apply.enabled ? state.apply.prefilter : state.filter;
    const stripped = collectStrippedWildcards(activeFilter).map(w => ({
      col: w.col, raw: w.raw, cleaned: w.cleaned, kinds: w.kinds.slice(),
    }));
    // Antipatterns + wildcard rewrites.
    const antipatterns = detectRetrieveMultipleAntipatterns({
      table: state.table,
      select: state.select,
      filter: activeFilter,
      strippedWildcards: stripped,
    });
    // Spec-grounded validation (Gap C of the doc audit). Walks the
    // filter / expand / orderby / apply trees and flags constructs that
    // Dataverse rejects — duplicate aggregate aliases, `not` on plain
    // comparisons, lambdas on N:1 navs, the $orderby+$apply conflict,
    // etc. Most are caught at construction time by UI gates; this catches
    // anything that drifts in (saved requests, hand edits).
    const validations = validateRequest({
      table: state.table,
      filter: state.filter,
      orderby: state.orderby,
      expand: state.expand,
      apply: state.apply,
    });

    // Dataverse rule (empirically confirmed against the docs):
    //   When the query contains a nested $expand AND ANY expand in the
    //   tree targets a collection-valued nav (1:N or N:N), the top-level
    //   $top and $orderby are forbidden. Server returns:
    //     0x80060888 "Only $select, $filter and $orderby clauses can be
    //                provided at top level while doing $expand on nested
    //                one-to-many relationship."
    //   (yes — the error names $orderby as both allowed and forbidden;
    //    empirically $orderby is also rejected when $top is also set,
    //    so we treat both as blockers.)
    //
    // Workaround: drop $top (and any non-empty $orderby) and use
    // `Prefer: odata.maxpagesize=N` instead. The mode's PreferEditor
    // already has the slot.
    const collectionNested = state.table
      && hasCollectionInvolvedNestedExpand(state.expand, state.table);
    const nestingAdvisories: Advisory[] = [];
    if (collectionNested && state.top != null) {
      nestingAdvisories.push({
        id: 'nested-collection-top',
        severity: 'error',
        source: 'validation',
        focusNode: 'top',
        title: '$top is forbidden with a nested $expand on a collection',
        body: (
          <>
            Your query has a nested <code>$expand</code> involving a 1:N or N:N
            relationship. Dataverse rejects top-level <code>$top</code> in this
            shape: <em>"Only $select, $filter and $orderby clauses can be
            provided at top level while doing $expand on nested one-to-many
            relationship."</em>
            <br />
            Drop <code>$top</code> and use{' '}
            <code>Prefer: odata.maxpagesize</code> in the Prefer pane instead —
            paging will then apply correctly to each expanded collection.
          </>
        ),
      });
    }
    if (collectionNested && state.orderby.length > 0) {
      nestingAdvisories.push({
        id: 'nested-collection-orderby',
        severity: 'error',
        source: 'validation',
        focusNode: 'orderby',
        title: '$orderby is forbidden with a nested $expand on a collection',
        body: (
          <>
            Same Dataverse rule as the <code>$top</code> case — top-level{' '}
            <code>$orderby</code> isn't allowed when a nested <code>$expand</code>{' '}
            involves a 1:N or N:N. Either clear <code>$orderby</code>, or sort
            client-side after the request returns.
          </>
        ),
      });
    }

    return [...antipatterns, ...validations, ...nestingAdvisories];
  }, [state.table, state.select, state.filter, state.orderby, state.expand, state.apply, state.top]);

  const blockerReason = disabledReasonFromAdvisories(advisories);

  const disabledReason =
    !state.table ? 'Pick a target table first.' :
    state.headers.some(h => h.enabled && !h.name) ? 'Fix empty header name.' :
    blockerReason ??
    null;

  // Per the aggregate-data docs, $apply overrides $select / $filter / $orderby / $expand / $count.
  // Hoisted here so the sidebar + right-pane router both reference the same flag.
  const applyActive = state.apply.enabled;

  // ── Sidebar tree for $expand (recursive) ────────────────────
  // Sub-clause visibility per the join-tables.md rules:
  //   - ManyToOne (single-valued): $select + $filter + $expand (no $top/$orderby — Dataverse rejects)
  //   - OneToMany (1:N collection): $select + $filter + $orderby + $top + $expand
  //   - ManyToMany (N:N collection): $select + $filter + $orderby + $top — NO $expand
  //     (Dataverse rejects any nested expand inside N:N with error 0x80060888)
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

      // $select — always
      children.push({
        id: `${navPath}/select`,
        icon: TextBulletList20Regular, iconFilled: TextBulletList20Filled,
        label: '$select', code: true, badge: it.select.length || null,
      });
      // $filter — single-valued AND collection both accept it
      children.push({
        id: `${navPath}/filter`,
        icon: Filter20Regular, iconFilled: Filter20Filled,
        label: '$filter', code: true, badge: innerFilterCount || null,
      });
      // $orderby / $top — collection-valued only (not meaningful on a single record)
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
      // $expand — every cardinality EXCEPT N:N (which rejects all nesting)
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

  // When $apply is active, $select / $filter / $orderby / $expand / $count
  // are NOT applied by Dataverse. Show those sidebar items with the "ignored"
  // appearance — keep them clickable so the user can still see/edit, but
  // flag the trade-off visually.
  // Use a short '∅' glyph (with a Tooltip on the row) rather than a long
  // 'overridden by $apply' badge — the long text wraps / clips on narrow
  // viewports and conveys the same state with worse legibility.
  const overrideBadge = (existing: React.ReactNode): { badge: React.ReactNode; badgeAppearance: 'ghost' | 'outline' } =>
    applyActive
      ? { badge: '∅', badgeAppearance: 'outline' }
      : { badge: existing, badgeAppearance: 'ghost' };

  // Sidebar config — section list driving the left-rail tree.
  const sections = [
    {
      id: 'target', label: 'Target', meta: tbl?.displayName,
      items: [{
        id: 'target',
        icon: Table20Regular, iconFilled: Table20Filled,
        label: tbl?.displayName ?? '(none)',
        badge: tbl ? 'table' : null, badgeAppearance: 'ghost' as const,
        dirty: state.dirty.has('target'),
      }],
    },
    {
      id: 'qstruct', label: 'Query structure',
      items: [
        { id: 'select',  icon: TextBulletList20Regular,    iconFilled: TextBulletList20Filled,    label: '$select',  code: true, dirty: state.dirty.has('select'),  ...overrideBadge(state.select.length || null) },
        { id: 'filter',  icon: Filter20Regular,            iconFilled: Filter20Filled,            label: '$filter',  code: true, dirty: state.dirty.has('filter'),  ...overrideBadge(filterCount || null) },
        { id: 'orderby', icon: TextSortAscending20Regular, iconFilled: TextSortAscending20Filled, label: '$orderby', code: true, dirty: state.dirty.has('orderby'), ...overrideBadge(state.orderby.length || null) },
        { id: 'top',     icon: NumberSymbol20Regular,      iconFilled: NumberSymbol20Filled,      label: '$top',     code: true, badge: state.top ? state.top.toString() : 'default', badgeAppearance: 'ghost' as const, dirty: state.dirty.has('top') },
        { id: 'count',   icon: Tag20Regular,               iconFilled: Tag20Filled,               label: '$count',   code: true, dirty: state.dirty.has('count'),   ...overrideBadge(state.countOn ? 'on' : null) },
        {
          id: 'expand',  icon: BranchFork20Regular, iconFilled: BranchFork20Filled,
          label: '$expand', code: true,
          dirty: state.dirty.has('expand'),
          ...overrideBadge(state.expand.length || null),
          children: expandSidebarChildren.length ? expandSidebarChildren : undefined,
        },
        { id: 'apply',   icon: ChartMultiple20Regular,     iconFilled: ChartMultiple20Filled,     label: '$apply',   code: true, badge: state.apply.enabled ? 'on' : null, badgeAppearance: applyActive ? 'filled' as const : 'ghost' as const, badgeColor: applyActive ? 'brand' as const : undefined, dirty: state.dirty.has('apply') },
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

  // ── Right pane router ───────────────────────────────────────
  // (applyActive is hoisted above — it gates the sidebar badges too.)
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
      case 'target':  pane = (
        <TargetEditor
          table={state.table}
          onTableChange={t => {
            // Switching (or clearing) the target entity invalidates EVERY
            // column-bound clause — select, filter, orderby, expand, apply,
            // top. Selecting 'account' after 'incident' would otherwise
            // keep `incidentid` in $select, `_customerid_value` in $filter,
            // etc., which the new entity doesn't have. Reset all of them
            // back to empty state and dump any cached results.
            setState(s => ({
              ...s,
              table: t,
              select: [],
              filter: emptyFilter(),
              orderby: [],
              top: 50,
              countOn: false,
              expand: [],
              apply: emptyApply(),
              dirty: new Set(['target']),
            }));
            setResult(null);
            // Entity change → effectively a fresh draft. Clear the
            // "last saved" reference so the Save button doesn't try to
            // overwrite a save belonging to a different entity.
            setLastSavedId(undefined);
            setLastSavedHash(null);
          }}
          onResetRequest={() => {
            // Same shape as onTableChange but keeps the table. The user
            // explicitly asked for "start over without changing the
            // target/root entity" — useful between test iterations.
            setState(s => ({
              ...s,
              select: [],
              filter: emptyFilter(),
              orderby: [],
              top: 50,
              countOn: false,
              expand: [],
              apply: emptyApply(),
              dirty: new Set(['target']),
            }));
            setResult(null);
          }}
          onApplyParsed={parsed => {
            // Apply a parsed OData URL. If the parser resolved the entity
            // set to a known table we set it; otherwise we keep the user's
            // current table and let column validation surface mismatches.
            setState(s => ({
              ...s,
              table: parsed.table || s.table,
              select: parsed.select,
              filter: parsed.filter,
              orderby: parsed.orderby,
              top: parsed.top,
              countOn: parsed.countOn,
              expand: parsed.expand,
              apply: parsed.apply ?? emptyApply(),
              dirty: new Set(['target']),
            }));
            setResult(null);
            // Bring the user to the filter pane if there's a filter; else
            // to select; else stay on target so they see what landed.
            if (parsed.filter.rules.length > 0) setActivePath('filter');
            else if (parsed.select.length > 0) setActivePath('select');
          }}
        />
      ); break;
      case 'select':  pane = <SelectEditor table={state.table} selectedIds={state.select} setSelectedIds={ids => set('select', ids, 'select')} applyActive={applyActive} />; break;
      case 'filter':  pane = <FilterEditor  table={state.table} tree={state.filter} setTree={t => set('filter', t, 'filter')} urlBytes={built.bytes} applyActive={applyActive} />; break;
      case 'orderby': pane = <OrderbyEditor table={state.table} items={state.orderby} setItems={i => set('orderby', i, 'orderby')} applyActive={applyActive} />; break;
      case 'top':     pane = <TopEditor     top={state.top} setTop={n => set('top', n, 'top')} maxPageSize={state.prefer.maxpagesize} />; break;
      case 'count':   pane = <CountEditor   countOn={state.countOn} setCountOn={b => set('countOn', b, 'count')} applyActive={applyActive} />; break;
      case 'apply':   pane = <ApplyEditor   table={state.table} spec={state.apply} setSpec={a => set('apply', a, 'apply')} />; break;
      case 'prefer':  pane = <PreferEditor  spec={state.prefer} setSpec={p => set('prefer', p, 'prefer')} />; break;
      case 'headers': pane = <HeadersEditor items={state.headers} setItems={h => set('headers', h, 'headers')} />; break;
    }
  }

  const headersMap = headerItemsToObject(state.headers, preferToHeaderString(state.prefer));
  const codeInputs = { method: 'GET', built, headers: headersMap };

  const rows = (result?.body as { value?: unknown[] } | null)?.value;
  const resultCount = Array.isArray(rows) ? rows.length : (result?.outcome === 'ok' ? 1 : null);

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
          onAdvisoryFocus={setActivePath}
        />
      }
    >
      <MainTabs tab={tab} onTabChange={setTab} resultCount={resultCount}>
        {tab === 'builder' && pane}
        {tab === 'code'    && <CodeView themeMode={themeMode} inputs={codeInputs} />}
        {tab === 'results' && <ResultsView
          result={result}
          mode="multi"
          table={state.table}
          select={state.select}
          expand={state.expand}
          orderby={state.orderby}
          requestUrl={built.relativeUrl}
          isLoadingMore={isLoadingMore}
          onRefresh={onExecute}
          onLoadMore={fetchOneMorePage}
          onRetrieveAll={retrieveAllPages}
        />}
      </MainTabs>
    </ModeShell>
  );
}
