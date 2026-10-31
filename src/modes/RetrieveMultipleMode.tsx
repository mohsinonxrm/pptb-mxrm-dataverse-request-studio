import { useEffect, useMemo, useRef, useState } from 'react';
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
  Tag20Regular,
  Tag20Filled,
  BranchFork20Regular,
  BranchFork20Filled,
  Settings20Regular,
  Settings20Filled,
  LineHorizontal320Regular,
  LineHorizontal320Filled,
  ChartMultiple20Regular,
  ChartMultiple20Filled,
  Link20Regular,
} from '@fluentui/react-icons';
import {
  Button,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  tokens,
} from '@fluentui/react-components';
import { Sidebar, type SidebarClauseItem } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { detectRetrieveMultipleAntipatterns } from '../engine/antipatterns';
import { collectStrippedWildcards } from '../editors/filter/filterTree';
import { disabledReasonFromAdvisories, type Advisory } from '../primitives/advisories';
import { validateRequest } from '../state/validateRequest';
import { ModeShell } from '../shell/ModeShell';
import { notify } from '../shell/dialogs';
import { TargetEditor } from '../editors/TargetEditor';
import { SelectEditor } from '../editors/SelectEditor';
import { FilterEditor } from '../editors/filter/FilterEditor';
import { OrderbyEditor } from '../editors/OrderbyEditor';
import { TopEditor, CountEditor } from '../editors/TopCountEditors';
import { ApplyEditor, emptyApply } from '../editors/ApplyEditor';
import { PreferEditor, emptyPrefer, preferToHeaderString } from '../editors/PreferEditor';
import { HeadersEditor, defaultReadHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { ExpandRouter } from '../editors/expand/ExpandRouter';
import {
  hasCollectionExpandWithNestedExpand,
  hasCollectionInvolvedNestedExpand,
} from '../editors/expand/expandTree';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { countRules, newId } from '../editors/filter/filterTree';
import { findTable } from '../mock/metadata';
import { findRequestType } from '../registry/requestTypes';
import { buildRetrieveMultiple } from '../engine/urlBuilder';
import { relatedMetadataAdvisories } from '../engine/navPath';
// Phase 1 wiring: routes through `runtime.retrieveMultiple` which auto-picks
// the live (dvHost.queryData) or mock executor based on `isEmbedded()`. The
// mode doesn't need to know which.
import { runtime, type ExecResult } from '../engine/runtime';
import type { ExpandSpec } from '../editors/ExpandEditor';
import type { RetrieveMultipleState, RecentRun } from '../state/readState';
import {
  serializeRetrieveMultiple,
  hashState,
  deserializeRetrieveMultiple,
  type SavedRequest,
  type SerializedRetrieveMultipleState,
  recentRunAsSavedEntry,
} from '../state/savedRequests';
import { useModeDraft, useStashModeDraft } from '../state/modeDrafts';
import { usePublishSaveContext } from '../state/SaveContext';
import { useLatestCallback } from '../primitives/useLatestCallback';
import type { ThemeMode } from '../theme/theme';
// Live-metadata hook: subscribes to the registry so child editors
// (SelectEditor, FilterEditor, OrderbyEditor, ExpandEditor, lambdas)
// re-render when columns + relationships land from Dataverse.
import { useLiveTable, useWarmReferencedTables } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import { pageSizeIsSent } from '../host/pptbBridge';
import { serviceProtectionLimit } from '../engine/serviceProtection';
import { topLevelRowCount } from '../views/results/flattenExpansions';
import {
  mergePage,
  pageOf,
  retrieveAll,
  type PageBatch,
  type PagingStatus,
} from '../views/results/pagedResults';
import { useRecentRuns } from '../state/runHistory';

// Empty initial state — the studio is a request *builder*; seeding it with
// mock data confused users in PPTB ("why are there pre-filled rules against
// a table I don't recognize?"). User picks the target table first; column-
// driven clauses ($select / $filter / $orderby / $expand) come from live
// metadata once a table is chosen.
//
// `top: 50` stays as a sensible default — it's a query-shape preference,
// not metadata that would mislead anyone.
const emptyFilter = () => ({
  id: newId('root'),
  type: 'group' as const,
  combinator: 'and' as const,
  rules: [],
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
  | 'target'
  | 'select'
  | 'filter'
  | 'orderby'
  | 'top'
  | 'count'
  | 'apply'
  | 'prefer'
  | 'headers';

export function RetrieveMultipleMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('retrieve-multiple');
  const draft = useModeDraft('retrieve-multiple', deserializeRetrieveMultiple, initialState);
  const [state, setState] = useState(draft.state);
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
  // The live table is a memo dependency below: build*() reads metadata via
  // findTable(), which is empty until the fetch lands — keying on `state`
  // alone left the URL / Code tab blank after picking a table (#78).
  const { table: liveTable } = useLiveTable(state.table || null);
  // Pre-warm RELATED entities referenced by nav-path filters / lambdas /
  // $expand / $apply / $orderby so the encoders resolve leaf column TYPES
  // against the correct entity (issue #33) — even on saved-request reload,
  // pasted-URL apply, or a direct Execute where no editor was opened to
  // lazily trigger the loads.
  // `relatedVersion` bumps when a related table lands so `built` re-encodes
  // (#63); `pending` / `failed` drive the metadata advisory.
  const related = useWarmReferencedTables(state.table || null, {
    filter: state.filter,
    expand: state.expand,
    apply: state.apply,
    orderby: state.orderby,
  });
  // Entity list — used to validate saved-request loads against the
  // current org (block load if the entity has been uninstalled).
  // Every table, not just the pickers' scope: a saved request's table only
  // has to exist (#44).
  const { allEntities: entities } = useScopedEntities();
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  // Kept across reloads, per environment (#226).
  const [recents, setRecents] = useRecentRuns('retrieve-multiple');
  // The $filter "Run without $filter" cleared, until Restore or the next run (#291).
  const [removedFilter, setRemovedFilter] = useState<RetrieveMultipleState['filter'] | null>(null);
  const runWithoutFilterRef = useRef(false);
  // The URL the shown result came from, for Copy request URL and the error
  // report: the builder may have changed since (#291).
  const [runUrl, setRunUrl] = useState<string | undefined>(undefined);

  // ── Saved-request persistence (localStorage) ──
  // Track the last-saved id + the hash of state at that save so we can
  // tell "the request has unsaved changes since last save". When dirty,
  // the SaveButton in the URL bar lights up; when clean, it disables.
  // Reset id on entity change (saved entries are scoped to one entity
  // at a time semantically — switching entities effectively starts a
  // fresh draft that can be saved separately).
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(draft.lastSavedId);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(draft.lastSavedHash);
  const currentSerialized = useMemo(() => serializeRetrieveMultiple(state), [state]);
  useStashModeDraft('retrieve-multiple', currentSerialized, lastSavedId, lastSavedHash);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  // Dirty when: never saved (no hash baseline) AND state has any meaningful
  // content (table picked), OR saved but the current hash diverges.
  const isDirty = useMemo(() => {
    if (!state.table) return false; // nothing to save yet
    if (lastSavedHash === null) return true; // fresh draft with a table
    return currentHash !== lastSavedHash;
  }, [state.table, lastSavedHash, currentHash]);

  const built = useMemo(
    () => buildRetrieveMultiple(state),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [state, liveTable, related.version],
  );
  const tbl = findTable(state.table);
  const filterCount = useMemo(() => countRules(state.filter), [state.filter]);

  const markDirty = (id: string) =>
    setState((s) => {
      const d = new Set(s.dirty);
      d.add(id);
      return { ...s, dirty: d };
    });

  const set = <K extends keyof RetrieveMultipleState>(
    k: K,
    v: RetrieveMultipleState[K],
    dirtyId?: string,
  ) => {
    setState((s) => ({ ...s, [k]: v }));
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
    // Hydrate the state. `deserializeRetrieveMultiple` re-inits the
    // `dirty` Set; useLiveTable picks up the entity change.
    setState(deserializeRetrieveMultiple(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    clearResult();
    setActivePath('target');
  };

  // Recent runs (#2) — restore a run's snapshot through the same loader as
  // the Saved library, but keep the current save baseline: a run isn't a
  // saved entry, so Save keeps targeting whatever it targeted before.
  const onRecentClick = (run: RecentRun) => {
    const entry = recentRunAsSavedEntry(run, 'retrieve-multiple');
    if (!entry) return;
    const keepId = lastSavedId;
    const keepHash = lastSavedHash;
    onLoadSaved(entry);
    setLastSavedId(keepId);
    setLastSavedHash(keepHash);
  };

  // Publish our save context to the top-right Save / Saved-library
  // buttons in FrameHeader. Each render re-publishes the current values;
  // the hook clears on unmount so a stale mode doesn't bleed across
  // mode switches. The setter is stable, so React only re-runs the
  // effect when the value object identity changes — which we want.
  // The header holds on to these; the stable wrappers always run this
  // render's handlers, not the ones from when the context was published.
  const onSavedLatest = useLatestCallback(onSaved);
  const onLoadSavedLatest = useLatestCallback(onLoadSaved);
  usePublishSaveContext(
    useMemo(
      () => ({
        state: currentSerialized,
        modeId: 'retrieve-multiple' as const,
        dirty: isDirty,
        lastSavedId,
        onSaved: onSavedLatest,
        onLoadSaved: onLoadSavedLatest,
      }),
      [currentSerialized, isDirty, lastSavedId, onSavedLatest, onLoadSavedLatest],
    ),
  );

  // ── Run identity ──
  // Every Execute (and every reset that clears the results) starts a new
  // "run". Paging handlers capture the run id when they start and drop
  // their merge if it changed while they were awaiting — otherwise a long
  // Retrieve All for the previous query overwrote the newer query's
  // results when it finished (#98). Executing therefore also cancels any
  // in-flight paging.
  const runIdRef = useRef(0);
  // The page size the run's first page asked for. Later pages must ask for
  // the same one, even if the Prefer pane has changed since (#219).
  const runPageSizeRef = useRef<number | null>(null);
  // The headers the run was sent with: later pages send the same ones, or a
  // page after the first would run as the signed-in user (#109).
  const runHeadersRef = useRef<Record<string, string> | undefined>(undefined);
  // Bumped by Stop to end a Retrieve All, keeping the rows already merged.
  const pagingIdRef = useRef(0);
  const [isLoadingMore, setIsLoadingMore] = useState(false);
  // Last next-page failure, shown above the grid. Auto-load on scroll is
  // paused while it's set so a failing page isn't retried on every scroll.
  const [pageError, setPageError] = useState<string | null>(null);
  // Why the last Retrieve all paused with pages left (#223).
  const [pagingNotice, setPagingNotice] = useState<string | null>(null);
  // Rows in the current result, and the run's Recent runs entry, so paging
  // keeps that entry's row count in step with the grid (#211).
  const rowCountRef = useRef(0);
  const runRecentIdRef = useRef<string | null>(null);

  const startRun = () => {
    runIdRef.current += 1;
    setIsLoadingMore(false);
    setPageError(null);
    setPagingNotice(null);
    rowCountRef.current = 0;
    runRecentIdRef.current = null;
    return runIdRef.current;
  };
  const clearResult = () => {
    // Also drops an in-flight Execute's response (it belonged to the
    // request being reset), so its spinner must stop here.
    startRun();
    setLoading(false);
    setResult(null);
    setRemovedFilter(null);
  };

  // `keepRemovedFilter` is only `true` for the run that "Run without
  // $filter" starts; a click handler passes its event here instead.
  const onExecute = async (keepRemovedFilter?: unknown) => {
    if (keepRemovedFilter !== true) setRemovedFilter(null);
    const run = startRun();
    setLoading(true);
    setRunUrl(built.relativeUrl);
    runPageSizeRef.current = state.prefer.maxpagesize;
    runHeadersRef.current = headerItemsToObject(state.headers, null);
    const res = await runtime.retrieveMultiple(state, runHeadersRef.current);
    // A newer Execute (e.g. Refresh clicked mid-flight) owns the results now.
    if (runIdRef.current !== run) return;
    setResult(res);
    setLoading(false);
    setTab('results');
    const recentId = `r-${Date.now()}`;
    runRecentIdRef.current = recentId;
    rowCountRef.current = topLevelRowCount(res.body) ?? 0;
    setRecents((rs) =>
      [
        {
          id: recentId,
          snapshot: currentSerialized,
          modeId: 'retrieve-multiple',
          url: built.relativeUrl,
          method: 'GET',
          ts: Date.now(),
          status: res.status,
          ms: res.ms,
          rowCount: topLevelRowCount(res.body),
        },
        ...rs,
      ].slice(0, 8),
    );
    setState((s) => ({ ...s, dirty: new Set() }));
  };

  // ── 0 rows: next steps (#291) ──
  // "Run without $filter" clears the builder's $filter, so the builder, the
  // URL and Recent runs all show what ran, and runs once that state has
  // rendered. The results then offer Restore, which puts the $filter back,
  // until the next run, a reset, or a new $filter in the builder.
  const filterActionsOn = filterCount > 0 && !state.apply.enabled;
  const runWithoutFilter = () => {
    setRemovedFilter(state.filter);
    runWithoutFilterRef.current = true;
    setState((s) => ({ ...s, filter: emptyFilter() }));
  };
  const restoreFilter = () => {
    const f = removedFilter;
    setRemovedFilter(null);
    if (f) setState((s) => ({ ...s, filter: f }));
  };
  useEffect(() => {
    if (runWithoutFilterRef.current) {
      runWithoutFilterRef.current = false;
      void onExecute(true);
    } else if (countRules(state.filter) > 0) {
      setRemovedFilter(null);
    }
    // Runs on the render that has the cleared $filter; onExecute reads it.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [state.filter]);

  // ── Infinite scroll + Retrieve-All ──
  // Dataverse paginates large result sets via `@odata.nextLink` (the URL
  // for the next page sits on the response envelope). We append next-page
  // rows onto the existing body so the grid keeps everything in one
  // virtualized list — no client-side paging needed.
  //
  // Pages are merged into the CURRENT result (functional update), and only
  // while the run that requested them is still current. A failed page
  // keeps the rows + nextLink already loaded, so the user can retry.
  const commitPages = (batch: PageBatch) => {
    setResult((r) => (r ? mergePage(r, batch) : r));
    rowCountRef.current += batch.rows.length;
    // The Recent runs entry follows the grid (#211).
    const id = runRecentIdRef.current;
    const n = rowCountRef.current;
    const summary = `${n.toLocaleString()} rows · ${batch.nextLink ? 'more available' : 'all pages'}`;
    if (id) setRecents((rs) => rs.map((x) => (x.id === id ? { ...x, rowCount: n, summary } : x)));
  };

  const describePageError = (more: ExecResult) => {
    const msg = (more.body as { error?: { message?: string } } | null)?.error?.message;
    // Status 0 means the host didn't report one (#59): show the status text
    // (or a generic label) rather than a misleading "HTTP 0".
    const status = more.status ? `HTTP ${more.status}` : more.statusText || 'Request failed';
    return `${status}${msg ? ` — ${msg}` : ''}`;
  };

  // Shared start-up for both paging handlers: claims a paging id under the
  // current run and returns whether it's still current, stopped (Stop bumps
  // the paging id) or superseded (a newer run or a reset).
  const beginPaging = () => {
    const run = runIdRef.current;
    const paging = ++pagingIdRef.current;
    setIsLoadingMore(true);
    setPageError(null);
    setPagingNotice(null);
    return (): PagingStatus =>
      runIdRef.current !== run
        ? 'superseded'
        : pagingIdRef.current !== paging
          ? 'stopped'
          : 'current';
  };

  const fetchOneMorePage = async () => {
    const nextLink = (result?.body as { '@odata.nextLink'?: string } | null)?.['@odata.nextLink'];
    if (!nextLink || isLoadingMore) return;
    const status = beginPaging();
    try {
      const more = await runtime.absoluteUrl(
        nextLink,
        runPageSizeRef.current,
        runHeadersRef.current,
      );
      if (status() !== 'current') return;
      if (!more.ok) {
        const throttled = serviceProtectionLimit(more);
        setPageError(
          throttled
            ? `Couldn't load the next page: Dataverse is throttling requests (${throttled.limit}). ${throttled.advice}`
            : `Couldn't load the next page: ${describePageError(more)}`,
        );
        return;
      }
      commitPages(pageOf(more));
    } finally {
      if (status() === 'current') setIsLoadingMore(false);
    }
  };

  // Pages reach the grid in batches and the loop pauses at a row / page
  // budget with a Continue prompt, so a large table can't freeze the tool
  // or fill memory unasked (#223).
  const retrieveAllPages = async () => {
    const nextLink = (result?.body as { '@odata.nextLink'?: string } | null)?.['@odata.nextLink'];
    if (!nextLink || isLoadingMore) return;
    const status = beginPaging();
    try {
      const outcome = await retrieveAll({
        nextLink,
        fetchPage: (url) => runtime.absoluteUrl(url, runPageSizeRef.current, runHeadersRef.current),
        commit: commitPages,
        status,
      });
      const throttled = outcome.kind === 'failed' ? serviceProtectionLimit(outcome.error) : null;
      if (throttled) {
        // Throttling passes (#269): a pause with Continue, from the page
        // that was refused, rather than an error. Rows so far are kept.
        setPagingNotice(
          `Retrieve all paused: Dataverse is throttling requests (${throttled.limit}). ` +
            `${rowCountRef.current.toLocaleString()} rows are loaded. ${throttled.advice} ` +
            'Continue picks up from the next page.',
        );
      } else if (outcome.kind === 'failed') {
        setPageError(`Retrieve all stopped at a failed page: ${describePageError(outcome.error)}`);
      } else if (outcome.kind === 'paused') {
        setPagingNotice(
          `Retrieve all paused after ${outcome.rows.toLocaleString()} more rows ` +
            `(${rowCountRef.current.toLocaleString()} loaded). More pages remain: Continue ` +
            'loads the next batch.',
        );
      }
    } finally {
      if (status() === 'current') setIsLoadingMore(false);
    }
  };

  // Stop keeps every page merged so far; the remaining nextLink stays on
  // the envelope so scrolling / Retrieve all can resume.
  const stopLoadingMore = () => {
    pagingIdRef.current += 1;
    setIsLoadingMore(false);
  };

  const retryLoadMore = () => {
    setPageError(null);
    void fetchOneMorePage();
  };

  // ── Advisories — query antipatterns + stripped wildcards ────────────
  // Detection runs on every render (cheap — pure tree walks); the
  // AdvisoryDrawer + URL-bar consume the result.
  const relatedPendingKey = related.pending.join(',');
  const relatedFailedKey = related.failed.join(',');
  const advisories = useMemo(() => {
    // When $apply is on, the top-level $filter is silently ignored by Dataverse
    // and the active filter is the apply.prefilter (filter() stage). Feed the
    // correct one to the wildcard / antipattern detectors.
    const activeFilter = state.apply.enabled ? state.apply.prefilter : state.filter;
    const stripped = collectStrippedWildcards(activeFilter).map((w) => ({
      col: w.col,
      raw: w.raw,
      cleaned: w.cleaned,
      kinds: w.kinds.slice(),
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

    // Dataverse rule (#126, confirmed live): when a collection-valued
    // $expand (1:N / N:N) itself has a nested $expand, top-level $top is
    // rejected with 400 0x80060888 "Only $select, $filter and $orderby
    // clauses can be provided at top level while doing $expand on nested
    // one-to-many relationship." — a blocker. Workaround: drop $top and use
    // `Prefer: odata.maxpagesize` (the PreferEditor has the slot).
    // Other nested shapes that merely involve a collection (e.g. N:1 → 1:N)
    // keep a non-blocking warning; the server is the arbiter there.
    const collectionNested =
      state.table && hasCollectionInvolvedNestedExpand(state.expand, state.table);
    const collectionWithNested =
      !!state.table && hasCollectionExpandWithNestedExpand(state.expand, state.table);
    const nestingAdvisories: Advisory[] = [];
    // On a ToolBox older than 1.2.7, Execute can't send the page size (#219).
    const pageSizeNote = pageSizeIsSent() ? null : (
      <>
        {' '}
        This ToolBox version can&apos;t send it on Execute (ToolBox 1.2.7+ can), so Execute gets
        Dataverse&apos;s 5,000-row pages.
      </>
    );
    if (collectionWithNested && state.top != null) {
      nestingAdvisories.push({
        id: 'nested-collection-top',
        severity: 'error',
        source: 'validation',
        focusNode: 'top',
        title: '$top is forbidden when a collection $expand has a nested $expand',
        body: (
          <>
            A 1:N or N:N <code>$expand</code> in your query has its own nested <code>$expand</code>.
            Dataverse rejects top-level <code>$top</code> in this shape (<code>0x80060888</code>):{' '}
            <em>
              "Only $select, $filter and $orderby clauses can be provided at top level while doing
              $expand on nested one-to-many relationship."
            </em>
            <br />
            Drop <code>$top</code> and use <code>Prefer: odata.maxpagesize</code> in the Prefer pane
            instead.{pageSizeNote}
          </>
        ),
        learnMoreUrl:
          'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query/join-tables#nested-expand-on-collection-valued-navigation-properties',
      });
    } else if (collectionNested && state.top != null) {
      nestingAdvisories.push({
        id: 'nested-collection-top',
        severity: 'warning',
        source: 'validation',
        focusNode: 'top',
        title: '$top with a nested $expand involving a collection',
        body: (
          <>
            Your query has a nested <code>$expand</code> involving a 1:N or N:N relationship. If
            Dataverse rejects <code>$top</code> (<code>0x80060888</code>), drop it and use{' '}
            <code>Prefer: odata.maxpagesize</code> in the Prefer pane instead.{pageSizeNote}
          </>
        ),
        learnMoreUrl:
          'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query/join-tables#nested-expand-on-collection-valued-navigation-properties',
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
            Top-level <code>$orderby</code> isn't allowed when a nested <code>$expand</code>{' '}
            involves a 1:N or N:N. Either clear <code>$orderby</code>, or sort client-side after the
            request returns.
          </>
        ),
      });
    }

    return [
      ...antipatterns,
      ...validations,
      ...nestingAdvisories,
      ...relatedMetadataAdvisories(related.pending, related.failed),
    ];
    // eslint-disable-next-line react-hooks/exhaustive-deps -- related.pending / failed are new arrays every render, so keyed on their joined contents; liveTable re-runs it when the table's metadata lands
  }, [
    state.table,
    state.select,
    state.filter,
    state.orderby,
    state.expand,
    state.apply,
    state.top,
    relatedPendingKey,
    relatedFailedKey,
    // The detectors and validateRequest read the table through findTable().
    // A saved request loaded before its table's metadata arrives would
    // otherwise keep the empty result, and no Execute blocker, until the
    // request changed.
    liveTable,
  ]);

  const blockerReason = disabledReasonFromAdvisories(advisories);

  const disabledReason = !state.table
    ? 'Pick a target table first.'
    : state.headers.some((h) => h.enabled && !h.name)
      ? 'Fix empty header name.'
      : (blockerReason ?? null);

  // Per the aggregate-data docs, $apply overrides $select / $filter / $orderby / $expand / $count.
  // Hoisted here so the sidebar + right-pane router both reference the same flag.
  const applyActive = state.apply.enabled;

  // ── Sidebar tree for $expand (recursive) ────────────────────
  // Sub-clause visibility per the join-tables.md rules:
  //   - ManyToOne (single-valued): $select + $filter + $expand (no $top/$orderby — Dataverse rejects)
  //   - OneToMany (1:N collection): $select + $filter + $orderby + $top + $expand
  //   - ManyToMany (N:N collection): $select + $filter + $orderby + $top — NO $expand
  //     (Dataverse rejects any nested expand inside N:N with error 0x80060888)
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

      // $select — always
      children.push({
        id: `${navPath}/select`,
        icon: TextBulletList20Regular,
        iconFilled: TextBulletList20Filled,
        label: '$select',
        code: true,
        badge: it.select.length || null,
      });
      // $filter — single-valued AND collection both accept it
      children.push({
        id: `${navPath}/filter`,
        icon: Filter20Regular,
        iconFilled: Filter20Filled,
        label: '$filter',
        code: true,
        badge: innerFilterCount || null,
      });
      // $orderby / $top — collection-valued only (not meaningful on a single record)
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
      // $expand — every cardinality EXCEPT N:N (which rejects all nesting)
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

  // When $apply is active, $select / $filter / $orderby / $expand / $count
  // are NOT applied by Dataverse. Show those sidebar items with the "ignored"
  // appearance — keep them clickable so the user can still see/edit, but
  // flag the trade-off visually.
  // Use a short '∅' glyph (with a Tooltip on the row) rather than a long
  // 'overridden by $apply' badge — the long text wraps / clips on narrow
  // viewports and conveys the same state with worse legibility.
  const overrideBadge = (
    existing: React.ReactNode,
  ): { badge: React.ReactNode; badgeAppearance: 'ghost' | 'outline' } =>
    applyActive
      ? { badge: '∅', badgeAppearance: 'outline' }
      : { badge: existing, badgeAppearance: 'ghost' };

  // Sidebar config — section list driving the left-rail tree.
  const sections = [
    {
      id: 'target',
      label: 'Target',
      meta: tbl?.displayName,
      items: [
        {
          id: 'target',
          icon: Table20Regular,
          iconFilled: Table20Filled,
          label: tbl?.displayName ?? '(none)',
          badge: tbl ? 'table' : null,
          badgeAppearance: 'ghost' as const,
          dirty: state.dirty.has('target'),
        },
      ],
    },
    {
      id: 'qstruct',
      label: 'Query structure',
      items: [
        {
          id: 'select',
          icon: TextBulletList20Regular,
          iconFilled: TextBulletList20Filled,
          label: '$select',
          code: true,
          dirty: state.dirty.has('select'),
          ...overrideBadge(state.select.length || null),
        },
        {
          id: 'filter',
          icon: Filter20Regular,
          iconFilled: Filter20Filled,
          label: '$filter',
          code: true,
          dirty: state.dirty.has('filter'),
          ...overrideBadge(filterCount || null),
        },
        {
          id: 'orderby',
          icon: TextSortAscending20Regular,
          iconFilled: TextSortAscending20Filled,
          label: '$orderby',
          code: true,
          dirty: state.dirty.has('orderby'),
          ...overrideBadge(state.orderby.length || null),
        },
        {
          id: 'top',
          icon: NumberSymbol20Regular,
          iconFilled: NumberSymbol20Filled,
          label: '$top',
          code: true,
          badge: state.top ? state.top.toString() : 'default',
          badgeAppearance: 'ghost' as const,
          dirty: state.dirty.has('top'),
        },
        {
          id: 'count',
          icon: Tag20Regular,
          iconFilled: Tag20Filled,
          label: '$count',
          code: true,
          dirty: state.dirty.has('count'),
          ...overrideBadge(state.countOn ? 'on' : null),
        },
        {
          id: 'expand',
          icon: BranchFork20Regular,
          iconFilled: BranchFork20Filled,
          label: '$expand',
          code: true,
          dirty: state.dirty.has('expand'),
          ...overrideBadge(state.expand.length || null),
          children: expandSidebarChildren.length ? expandSidebarChildren : undefined,
        },
        {
          id: 'apply',
          icon: ChartMultiple20Regular,
          iconFilled: ChartMultiple20Filled,
          label: '$apply',
          code: true,
          badge: state.apply.enabled ? 'on' : null,
          badgeAppearance: applyActive ? ('filled' as const) : ('ghost' as const),
          badgeColor: applyActive ? ('brand' as const) : undefined,
          dirty: state.dirty.has('apply'),
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
      case 'target':
        pane = (
          <TargetEditor
            table={state.table}
            onTableChange={(t) => {
              // Switching (or clearing) the target entity invalidates EVERY
              // column-bound clause — select, filter, orderby, expand, apply,
              // top. Selecting 'account' after 'incident' would otherwise
              // keep `incidentid` in $select, `_customerid_value` in $filter,
              // etc., which the new entity doesn't have. Reset all of them
              // back to empty state and dump any cached results.
              setState((s) => ({
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
              clearResult();
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
              setState((s) => ({
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
              clearResult();
            }}
            onApplyParsed={(parsed) => {
              // Apply a parsed OData URL. If the parser resolved the entity
              // set to a known table we set it; otherwise we keep the user's
              // current table and let column validation surface mismatches.
              setState((s) => ({
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
              clearResult();
              // Bring the user to the filter pane if there's a filter; else
              // to select; else stay on target so they see what landed.
              if (parsed.filter.rules.length > 0) setActivePath('filter');
              else if (parsed.select.length > 0) setActivePath('select');
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
            applyActive={applyActive}
          />
        );
        break;
      case 'filter':
        pane = (
          <FilterEditor
            table={state.table}
            tree={state.filter}
            setTree={(t) => set('filter', t, 'filter')}
            urlBytes={built.bytes}
            applyActive={applyActive}
          />
        );
        break;
      case 'orderby':
        pane = (
          <OrderbyEditor
            table={state.table}
            items={state.orderby}
            setItems={(i) => set('orderby', i, 'orderby')}
            applyActive={applyActive}
          />
        );
        break;
      case 'top':
        pane = (
          <TopEditor
            top={state.top}
            setTop={(n) => set('top', n, 'top')}
            maxPageSize={state.prefer.maxpagesize}
          />
        );
        break;
      case 'count':
        pane = (
          <CountEditor
            countOn={state.countOn}
            setCountOn={(b) => set('countOn', b, 'count')}
            applyActive={applyActive}
          />
        );
        break;
      case 'apply':
        pane = (
          <ApplyEditor
            table={state.table}
            spec={state.apply}
            setSpec={(a) => set('apply', a, 'apply')}
          />
        );
        break;
      case 'prefer':
        pane = (
          <PreferEditor
            spec={state.prefer}
            setSpec={(p) => set('prefer', p, 'prefer')}
            sendsPageSize
          />
        );
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

  const rows = (result?.body as { value?: unknown[] } | null)?.value;
  const resultCount = Array.isArray(rows) ? rows.length : result?.outcome === 'ok' ? 1 : null;

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
          onAdvisoryFocus={setActivePath}
        />
      }
    >
      <MainTabs tab={tab} onTabChange={setTab} resultCount={resultCount}>
        {tab === 'builder' && pane}
        {tab === 'code' && <CodeView themeMode={themeMode} inputs={codeInputs} />}
        {tab === 'results' && (
          <ResultsView
            result={result}
            mode="multi"
            table={state.table}
            select={state.select}
            expand={state.expand}
            orderby={state.orderby}
            requestUrl={runUrl}
            isLoadingMore={isLoadingMore}
            onRefresh={onExecute}
            onLoadMore={pageError ? undefined : fetchOneMorePage}
            onRetrieveAll={retrieveAllPages}
            onStopLoadingMore={stopLoadingMore}
            loadMoreError={pageError}
            onRetryLoadMore={retryLoadMore}
            onDismissLoadMoreError={() => setPageError(null)}
            loadMoreNotice={pagingNotice}
            onContinueLoadMore={retrieveAllPages}
            onDismissLoadMoreNotice={() => setPagingNotice(null)}
            onEditFilter={
              filterActionsOn
                ? () => {
                    setActivePath('filter');
                    setTab('builder');
                  }
                : undefined
            }
            // Same gate as Execute: nothing runs while a run is in flight
            // or the request has a blocker.
            onRunWithoutFilter={
              filterActionsOn && !loading && !disabledReason ? runWithoutFilter : undefined
            }
            notice={
              removedFilter && filterCount === 0 ? (
                <MessageBar intent="info" layout="multiline">
                  <MessageBarBody>
                    Ran without $filter. Your $filter was removed from the builder.
                  </MessageBarBody>
                  <MessageBarActions>
                    <Button onClick={restoreFilter}>Restore $filter</Button>
                  </MessageBarActions>
                </MessageBar>
              ) : undefined
            }
          />
        )}
      </MainTabs>
    </ModeShell>
  );
}
