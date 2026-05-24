// Predefined Query — GET /<entitySet>?savedQuery=<id> | ?userQuery=<id>
//
// Saved queries (system views) and user queries (personal views) embed
// FetchXml that owns $select / $filter / $orderby / $expand. Only $top and
// Prefer headers stay editable — everything else is locked.
//
// Modeled on the Power Apps "Saved views" / "My views" UX. Cross-references:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/retrieve-and-execute-predefined-queries
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/savedquery-entity
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/userquery-entity
//
// Sidebar:
//   • Target     — pick a table, then a savedQuery or userQuery on it
//   • Conditions — read-only FetchXml preview (locked)
//   • Top        — $top override (one of the few clauses that still works)
//   • Prefer     — header preferences
//   • Headers    — generic HTTP headers
//
// The user view list is the queries this user can READ (their own +
// shared/team-granted). We don't filter by ownerid; Dataverse handles
// access control at the request boundary.

import { useEffect, useMemo, useState } from 'react';
import Editor from '@monaco-editor/react';
import {
  Filter20Regular, Filter20Filled,
  AppsList20Regular, AppsList20Filled,
  NumberSymbol20Regular, NumberSymbol20Filled,
  Settings20Regular, Settings20Filled,
  LineHorizontal320Regular, LineHorizontal320Filled,
} from '@fluentui/react-icons';
import {
  Field, Combobox, Option, OptionGroup, MessageBar, MessageBarBody, MessageBarTitle,
  Caption1, tokens, mergeClasses, Spinner, Body1, Badge, Tooltip,
} from '@fluentui/react-components';
import { Sidebar } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { PaneHead } from '../editors/PaneHead';
import { TargetEditor } from '../editors/TargetEditor';
import { TopEditor } from '../editors/TopCountEditors';
import { PreferEditor, emptyPrefer, preferToHeaderString } from '../editors/PreferEditor';
import { HeadersEditor, defaultReadHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { findRequestType } from '../registry/requestTypes';
import { findTable } from '../mock/metadata';
import { buildPredefinedQuery } from '../engine/urlBuilder';
import { runtime, type ExecResult } from '../engine/runtime';
import type { PredefinedQueryState, RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useLiveTable } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import { useStudioStyles } from '../primitives/styles';
import {
  serializePredefinedQuery, deserializePredefinedQuery, hashState,
  type SavedRequest, type SerializedPredefinedQueryState,
} from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';

// Empty initial state — the studio is a request *builder*. Pre-seeding a
// table or a query confused users in PPTB ("why is this random view
// pre-selected?"). User picks the target table first; queries come from
// the live fetch once a table is chosen.
const initialState = (): PredefinedQueryState => ({
  table: '',
  queryId: null,
  queryType: 'savedQuery',
  top: 50,
  prefer: { ...emptyPrefer(), formattedValues: true },
  headers: defaultReadHeaders(),
  dirty: new Set(),
});

type ClauseId = 'target' | 'conditions' | 'top' | 'prefer' | 'headers';

// ── Live query shapes ──────────────────────────────────────────────────
//
// Both savedquery and userquery share the same useful fields. We project a
// tight $select so the picker dropdown stays small; FetchXml is the
// big one and only needed for the chosen query.
//
// `layoutxml` defines what columns/widths are returned in the result grid —
// even if the underlying FetchXml selects MORE attrs, the layout is the
// canonical "this is what you see". Model-driven advanced-find consumes
// the same field. `layoutjson` is the parsed form; we parse layoutxml
// ourselves for compatibility (older queries may have only the XML).

interface QueryRow {
  id: string;
  name: string;
  description: string | null;
  fetchXml: string | null;
  layoutXml: string | null;
  layoutColumns: LayoutColumn[];
  returnedTypeCode: string;
  queryType: number;
  isDefault: boolean;
  isQuickFind: boolean;
  isCustom: boolean;
  isManaged: boolean;
  kind: 'savedQuery' | 'userQuery';
}

interface LayoutColumn {
  /** Full cell `name` attr — for joined columns this contains `<linkAlias>.<attr>`. */
  fullName: string;
  /** The actual attribute name (after the dot for joined cells). */
  attr: string;
  /** Pixel width from layoutxml — used for sort ordering and visual weight. */
  width: number;
  /** True when the cell is a linked-entity column (`alias.attr`). */
  isLink: boolean;
  /** Link alias when isLink — useful for showing "from <link>" pills. */
  linkAlias?: string;
}

/**
 * Parse the standard layoutxml shape into a list of result-grid columns.
 * Tolerant of XML idiosyncrasies — layoutxml is a small, well-known schema
 * (no namespaces, no comments), so a regex extract is sufficient and avoids
 * pulling in DOMParser (which behaves differently in test envs).
 *
 *   <grid …><row …><cell name="X" width="100" /></row></grid>
 *
 * Linked-entity columns use a `linkAlias.attr` form in `name=`. We split on
 * the first `.` to surface the alias separately.
 */
function parseLayoutColumns(layoutXml: string | null): LayoutColumn[] {
  if (!layoutXml) return [];
  const cells = [...layoutXml.matchAll(/<cell\s+([^>]+?)\s*\/?>/g)];
  const out: LayoutColumn[] = [];
  for (const m of cells) {
    const attrs = m[1];
    const nameMatch = /name="([^"]+)"/.exec(attrs);
    if (!nameMatch) continue;
    const fullName = nameMatch[1];
    const widthMatch = /width="([^"]+)"/.exec(attrs);
    const width = widthMatch ? parseInt(widthMatch[1], 10) || 100 : 100;
    const dotIdx = fullName.indexOf('.');
    if (dotIdx >= 0) {
      out.push({
        fullName, width, isLink: true,
        linkAlias: fullName.slice(0, dotIdx),
        attr: fullName.slice(dotIdx + 1),
      });
    } else {
      out.push({ fullName, width, isLink: false, attr: fullName });
    }
  }
  return out;
}

/**
 * Map SavedQuery.QueryType integer → short human label.
 *
 * The full enum is defined in `Microsoft.Crm.Sdk.SavedQueryQueryType`:
 *   https://learn.microsoft.com/dotnet/api/microsoft.crm.sdk.savedqueryquerytype
 *
 * Values are bitfield-like (powers of 2 plus 0 for the default) so unknown
 * values get a `Type N` fallback rather than an opaque number badge.
 */
function queryTypeLabel(qt: number): string {
  switch (qt) {
    case 0:    return 'Main view';
    case 1:    return 'Advanced Find';
    case 2:    return 'Subgrid';
    case 4:    return 'Quick Find';
    case 8:    return 'Reserved';
    case 16:   return 'Lookup';
    case 32:   return 'Marketing list';
    case 64:   return 'Address book';
    case 128:  return 'Main (no subj)';
    case 256:  return 'Outlook filter';
    case 512:  return 'Reserved';
    case 1024: return 'Offline filter';
    case 2048: return 'Outlook template';
    case 4096: return 'Reserved';
    case 8192: return 'Svc mgmt app';
    default:   return `Type ${qt}`;
  }
}

export function PredefinedQueryMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('predefined-query');
  const s = useStudioStyles();
  const [state, setState] = useState(initialState);
  // Live-metadata subscription so the entity-set name resolves once the
  // table cache lands.
  useLiveTable((state).table || null);
  const { entities } = useScopedEntities();
  const [activeNode, setActiveNode] = useState<ClauseId>('target');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);

  // Save / Load
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);

  // Live-fetched queries scoped to the chosen table.
  const [queries, setQueries] = useState<QueryRow[]>([]);
  const [queriesLoading, setQueriesLoading] = useState(false);
  const [queriesError, setQueriesError] = useState<string | null>(null);

  // Picker typed-input state. Fluent v9 Combobox is `freeform` here so the
  // user can type to filter; the input value isn't necessarily a selected
  // option name, so we track it separately. When a selection lands we sync
  // it to the picked query's name; when the user types we don't touch the
  // selection.
  const [comboInput, setComboInput] = useState<string>('');

  const built = useMemo(() => buildPredefinedQuery(state), [state]);
  const tbl = findTable(state.table);
  const selectedQuery = queries.find(q => q.id === state.queryId) ?? null;

  // Two-group dropdown layout — system views (savedquery) above personal
  // views (userquery). Sorted by isDefault desc + name asc inside each
  // group (the fetch effect already enforces this).
  const systemQueries = queries.filter(q => q.kind === 'savedQuery');
  const personalQueries = queries.filter(q => q.kind === 'userQuery');

  // Client-side typed filter — Fluent v9 Combobox doesn't auto-filter.
  // Empty input → show everything. Non-empty → substring match on the
  // query name (case-insensitive). This keeps the dropdown manageable
  // on tables with 30+ system views (most large entities).
  const filteredSystem = useMemo(() => {
    const q = comboInput.trim().toLowerCase();
    if (!q) return systemQueries;
    return systemQueries.filter(x => x.name.toLowerCase().includes(q));
  }, [systemQueries, comboInput]);
  const filteredPersonal = useMemo(() => {
    const q = comboInput.trim().toLowerCase();
    if (!q) return personalQueries;
    return personalQueries.filter(x => x.name.toLowerCase().includes(q));
  }, [personalQueries, comboInput]);

  // Sync the combobox input text to the selected query's name (and reset
  // on table change). User typing overrides this via onChange.
  useEffect(() => {
    setComboInput(selectedQuery?.name ?? '');
  }, [selectedQuery?.name, state.table]);

  // ── Live fetch: savedqueries + userqueries scoped to (table) ──
  //
  // We issue both in parallel — the picker tabs ("System views" /
  // "Personal views") work off the union so switching tabs is free after
  // the initial load.
  //
  // savedquery and userquery use `returnedtypecode` which on the modern
  // Web API is a string (`EntityName` type) matching the table's
  // LogicalName. Older docs show it as an int (object type code) — that
  // form still works for legacy reasons but the string form is canonical.
  //
  // statecode eq 0 filters out deactivated views. We DON'T filter on
  // ownerid for userqueries — Dataverse already gates by share-access at
  // the request boundary, so what we receive IS what the user can run.
  useEffect(() => {
    let cancelled = false;
    if (!state.table) {
      setQueries([]);
      setQueriesError(null);
      return;
    }
    if (typeof window === 'undefined' || !window.dataverseAPI?.queryData) {
      // Embedded host not available — leave the list empty so the picker
      // shows the "no queries" empty state with the right hint.
      setQueries([]);
      return;
    }
    setQueriesLoading(true);
    setQueriesError(null);

    // Project layoutxml + querytype + the badge-driving flags. layoutxml is
    // the biggest field by far but it's also the most useful — we render the
    // result columns as pills the moment a query is picked, no extra round-
    // trip needed. The list typically has <50 queries per entity so the
    // payload stays manageable.
    const savedSelect =
      `savedqueryid,name,description,fetchxml,layoutxml,returnedtypecode,` +
      `querytype,isdefault,isquickfindquery,iscustomizable,ismanaged`;
    const userSelect =
      `userqueryid,name,description,fetchxml,layoutxml,returnedtypecode,` +
      `querytype`;
    const savedUrl =
      `savedqueries?$select=${savedSelect}` +
      `&$filter=returnedtypecode eq '${state.table}' and statecode eq 0` +
      `&$orderby=name asc`;
    const userUrl =
      `userqueries?$select=${userSelect}` +
      `&$filter=returnedtypecode eq '${state.table}' and statecode eq 0` +
      `&$orderby=name asc`;

    Promise.all([
      window.dataverseAPI.queryData(savedUrl).catch((e) => {
        // savedqueries should always exist; bubble the error up
        throw e instanceof Error ? e : new Error(String(e));
      }),
      window.dataverseAPI.queryData(userUrl).catch(() => ({ value: [] })),
    ]).then(([sv, uv]) => {
      if (cancelled) return;
      const svRows = (sv as { value?: Array<Record<string, unknown>> })?.value ?? [];
      const uvRows = (uv as { value?: Array<Record<string, unknown>> })?.value ?? [];
      const all: QueryRow[] = [
        ...svRows.map((r) => {
          const layoutXml = r.layoutxml != null ? String(r.layoutxml) : null;
          return {
            id: String(r.savedqueryid ?? ''),
            name: String(r.name ?? '(unnamed)'),
            description: r.description != null ? String(r.description) : null,
            fetchXml: r.fetchxml != null ? String(r.fetchxml) : null,
            layoutXml,
            layoutColumns: parseLayoutColumns(layoutXml),
            returnedTypeCode: String(r.returnedtypecode ?? state.table),
            queryType: typeof r.querytype === 'number' ? r.querytype : 0,
            isDefault: r.isdefault === true,
            isQuickFind: r.isquickfindquery === true,
            // iscustomizable is a ManagedProperty wrapper `{ Value: bool, CanBeChanged: bool }`
            isCustom: (r.iscustomizable as { Value?: boolean } | undefined)?.Value === true,
            isManaged: r.ismanaged === true,
            kind: 'savedQuery' as const,
          };
        }),
        ...uvRows.map((r) => {
          const layoutXml = r.layoutxml != null ? String(r.layoutxml) : null;
          return {
            id: String(r.userqueryid ?? ''),
            name: String(r.name ?? '(unnamed)'),
            description: r.description != null ? String(r.description) : null,
            fetchXml: r.fetchxml != null ? String(r.fetchxml) : null,
            layoutXml,
            layoutColumns: parseLayoutColumns(layoutXml),
            returnedTypeCode: String(r.returnedtypecode ?? state.table),
            queryType: typeof r.querytype === 'number' ? r.querytype : 0,
            // userquery doesn't have these flags — personal views are user-
            // owned by definition and never "default" / "quick find".
            isDefault: false,
            isQuickFind: false,
            isCustom: true,
            isManaged: false,
            kind: 'userQuery' as const,
          };
        }),
      ];
      // Default views float to the top of their kind — they're what most
      // users mean when they pick "the Account view".
      all.sort((a, b) => {
        if (a.kind !== b.kind) return a.kind === 'savedQuery' ? -1 : 1;
        if (a.isDefault !== b.isDefault) return a.isDefault ? -1 : 1;
        return a.name.localeCompare(b.name);
      });
      setQueries(all);
      setQueriesLoading(false);
    }).catch(e => {
      if (cancelled) return;
      const msg = e instanceof Error ? e.message : String(e);
      setQueriesError(msg);
      setQueries([]);
      setQueriesLoading(false);
    });

    return () => { cancelled = true; };
  }, [state.table]);

  const markDirty = (id: ClauseId) => setState(ss => { const d = new Set(ss.dirty); d.add(id); return { ...ss, dirty: d }; });
  const set = <K extends keyof PredefinedQueryState>(k: K, v: PredefinedQueryState[K], dirtyId?: ClauseId) => {
    setState(ss => ({ ...ss, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  const onTableChange = (logical: string) => {
    setState(ss => ({ ...ss, table: logical, queryId: null }));
    markDirty('target');
  };

  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.predefined(state);
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents(rs => [{
      id: `r-${Date.now()}`, modeId: 'predefined-query',
      url: built.relativeUrl, method: 'GET', ts: Date.now(),
      status: res.status, ms: res.ms,
      rowCount: ((res.body as { value?: unknown[] } | null)?.value?.length) ?? 0,
    }, ...rs].slice(0, 8));
    setState(ss => ({ ...ss, dirty: new Set() }));
  };

  const disabledReason =
    !state.table ? 'Pick a table first.' :
    !state.queryId ? 'Pick a saved query or user view.' :
    null;

  // ── Save / Load ──
  const currentSerialized = useMemo(() => serializePredefinedQuery(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;

  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };

  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== 'predefined-query') return;
    const snap = entry.state as SerializedPredefinedQueryState;
    if (entities.length > 0 && !entities.some(e => e.logicalName === snap.table)) {
      window.alert(
        `Can't load "${entry.name}": entity \`${snap.table}\` ` +
        `isn't available in this environment. The solution may have been ` +
        `removed or you may be connected to a different org.`,
      );
      return;
    }
    setState(deserializePredefinedQuery(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setActiveNode('target');
  };

  usePublishSaveContext(useMemo(() => {
    if (!state.table) return null;
    return {
      state: currentSerialized,
      modeId: 'predefined-query' as const,
      dirty: isDirty,
      lastSavedId,
      onSaved,
      onLoadSaved,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentSerialized, isDirty, lastSavedId, state.table]));

  const sections = [
    {
      id: 'target', label: 'Target', meta: tbl?.displayName,
      items: [{
        id: 'target', icon: AppsList20Regular, iconFilled: AppsList20Filled,
        label: selectedQuery?.name ?? (state.table ? 'Pick a query' : 'Pick a table'),
        badge: selectedQuery?.kind === 'savedQuery' ? 'system' : selectedQuery?.kind === 'userQuery' ? 'user' : null,
        badgeAppearance: 'ghost' as const,
        dirty: state.dirty.has('target'),
      }],
    },
    {
      id: 'conditions', label: 'Conditions',
      items: [{
        id: 'conditions', icon: Filter20Regular, iconFilled: Filter20Filled,
        label: 'Filter (read-only)',
        badge: selectedQuery ? 'inherited' : null, badgeAppearance: 'ghost' as const,
      }],
    },
    {
      id: 'top', label: 'Top',
      items: [{
        id: 'top', icon: NumberSymbol20Regular, iconFilled: NumberSymbol20Filled,
        label: '$top', code: true,
        badge: state.top ? state.top.toString() : 'default', badgeAppearance: 'ghost' as const,
        dirty: state.dirty.has('top'),
      }],
    },
    {
      id: 'prefer', label: 'Prefer',
      items: [{
        id: 'prefer', icon: Settings20Regular, iconFilled: Settings20Filled,
        label: 'Prefer header',
        badge: preferToHeaderString(state.prefer) ? 'on' : null, badgeAppearance: 'ghost' as const,
        dirty: state.dirty.has('prefer'),
      }],
    },
    {
      id: 'headers', label: 'Headers', meta: `${state.headers.filter(h => h.enabled).length} active`,
      items: [{
        id: 'headers', icon: LineHorizontal320Regular, iconFilled: LineHorizontal320Filled,
        label: 'HTTP headers',
        badge: state.headers.filter(h => h.enabled).length || null,
        dirty: state.dirty.has('headers'),
      }],
    },
  ];

  let pane: React.ReactNode;
  switch (activeNode) {
    case 'target': pane = (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
        <TargetEditor
          table={state.table}
          onTableChange={onTableChange}
          group="read"
          sub="Pick the table whose saved queries (system views) or user queries (personal views) you want to execute."
        />

        {state.table && (
          <div style={{ maxWidth: 640, display: 'flex', flexDirection: 'column', gap: 14 }}>
            {/* One picker, two OptionGroups. Drops the tabs entirely — the
                group headers ("System views" / "Personal views") communicate
                kind without needing a second navigation level. Per the
                Fluent v9 storybook Combobox grouping pattern. */}
            <Field
              label="View"
              hint={
                queries.length > 0
                  ? `${systemQueries.length} system · ${personalQueries.length} personal`
                  : undefined
              }
            >
              <Combobox
                freeform
                clearable
                value={comboInput}
                selectedOptions={state.queryId ? [state.queryId] : []}
                placeholder={
                  queriesLoading ? 'Loading queries…' :
                  queries.length === 0 ? 'No queries on this table' :
                  'Search views by name…'
                }
                disabled={queriesLoading}
                onChange={(e) => {
                  // User typed — treat as a filter, not a selection. If they
                  // clear the input we also clear the selection (matches
                  // the `clearable` X behavior).
                  const next = (e.target as HTMLInputElement).value;
                  setComboInput(next);
                  if (next === '' && state.queryId) {
                    setState(ss => ({ ...ss, queryId: null }));
                    markDirty('target');
                  }
                }}
                onOptionSelect={(_, d) => {
                  if (!d.optionValue) {
                    // Native clearable X dispatches a no-option select.
                    setState(ss => ({ ...ss, queryId: null }));
                    setComboInput('');
                    markDirty('target');
                    return;
                  }
                  const q = queries.find(x => x.id === d.optionValue);
                  if (!q) return;
                  setState(ss => ({ ...ss, queryId: q.id, queryType: q.kind }));
                  setComboInput(q.name);
                  markDirty('target');
                }}
                listbox={{ style: { maxHeight: 420 } }}
              >
                {filteredSystem.length === 0 && filteredPersonal.length === 0 && comboInput.trim() && (
                  <Option value="__none" text="" disabled>
                    <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                      No views match "{comboInput.trim()}"
                    </Caption1>
                  </Option>
                )}
                {filteredSystem.length > 0 && (
                  <OptionGroup label={`System views (${filteredSystem.length})`}>
                    {filteredSystem.map(q => (
                      <Option key={q.id} value={q.id} text={q.name}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 1, minWidth: 0 }}>
                          <Body1 style={{ fontWeight: q.isDefault ? 600 : 400 }}>{q.name}</Body1>
                          {q.description && (
                            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                              {q.description}
                            </Caption1>
                          )}
                        </div>
                      </Option>
                    ))}
                  </OptionGroup>
                )}
                {filteredPersonal.length > 0 && (
                  <OptionGroup label={`Personal views (${filteredPersonal.length})`}>
                    {filteredPersonal.map(q => (
                      <Option key={q.id} value={q.id} text={q.name}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 1, minWidth: 0 }}>
                          <Body1>{q.name}</Body1>
                          {q.description && (
                            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                              {q.description}
                            </Caption1>
                          )}
                        </div>
                      </Option>
                    ))}
                  </OptionGroup>
                )}
              </Combobox>
            </Field>

            {queriesLoading && (
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, color: tokens.colorNeutralForeground3 }}>
                <Spinner size="tiny" />
                <Caption1>Fetching saved + user queries…</Caption1>
              </div>
            )}

            {queriesError && (
              <MessageBar intent="error" layout="multiline">
                <MessageBarBody>
                  <MessageBarTitle>Couldn't load queries</MessageBarTitle>
                  {queriesError}
                </MessageBarBody>
              </MessageBar>
            )}

            {selectedQuery && (
              <div className={mergeClasses(s.inlineCard)} style={{ padding: 14, display: 'flex', flexDirection: 'column', gap: 10 }}>
                {/* Header row — name + flags pile (default/quick-find/managed/custom) */}
                <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
                  <Body1 style={{ fontWeight: 600 }}>{selectedQuery.name}</Body1>
                  <Badge appearance="ghost" size="small">{queryTypeLabel(selectedQuery.queryType)}</Badge>
                  {selectedQuery.isDefault && (
                    <Badge appearance="tint" color="brand" size="small">default view</Badge>
                  )}
                  {selectedQuery.isQuickFind && (
                    <Badge appearance="tint" color="informative" size="small">quick find</Badge>
                  )}
                  {selectedQuery.kind === 'savedQuery' && selectedQuery.isManaged && (
                    <Badge appearance="ghost" size="small">managed</Badge>
                  )}
                  {selectedQuery.kind === 'savedQuery' && !selectedQuery.isManaged && (
                    <Badge appearance="ghost" size="small">unmanaged</Badge>
                  )}
                </div>

                {/* GUID */}
                <div>
                  <Caption1 style={{ fontWeight: 600, color: tokens.colorNeutralForeground2 }}>Query GUID</Caption1>
                  <div>
                    <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11, color: tokens.colorNeutralForeground2 }}>
                      {selectedQuery.id}
                    </code>
                  </div>
                </div>

                {/* Layout-driven columns. The FetchXml may select more attrs than
                    these — anything outside the layout still lands on the wire,
                    but the layout is what the maker sees in the model-driven
                    grid. Surfacing it gives the user a quick "this is what's
                    coming back" preview before they hit Execute. */}
                {selectedQuery.layoutColumns.length > 0 && (
                  <div>
                    <Caption1 style={{ fontWeight: 600, color: tokens.colorNeutralForeground2 }}>
                      Result columns ({selectedQuery.layoutColumns.length}) — from layoutxml
                    </Caption1>
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4, marginTop: 4 }}>
                      {selectedQuery.layoutColumns.map(c => (
                        <Tooltip
                          key={c.fullName}
                          relationship="label"
                          content={
                            c.isLink
                              ? `${c.fullName} — joined column via link alias \`${c.linkAlias}\` · width ${c.width}px`
                              : `${c.fullName} — width ${c.width}px`
                          }
                        >
                          <span style={{
                            fontFamily: tokens.fontFamilyMonospace, fontSize: 11,
                            padding: '2px 6px', borderRadius: 4,
                            background: c.isLink ? tokens.colorPaletteBlueBackground2 : tokens.colorNeutralBackground3,
                            color: c.isLink ? tokens.colorPaletteBlueForeground2 : tokens.colorNeutralForeground1,
                            cursor: 'help',
                          }}>
                            {c.isLink ? `${c.linkAlias}.${c.attr}` : c.attr}
                          </span>
                        </Tooltip>
                      ))}
                    </div>
                    {selectedQuery.layoutColumns.some(c => c.isLink) && (
                      <Caption1 style={{ display: 'block', marginTop: 6, color: tokens.colorNeutralForeground3 }}>
                        Blue pills are joined columns (link-entity in FetchXml).
                      </Caption1>
                    )}
                  </div>
                )}
              </div>
            )}
          </div>
        )}
      </div>
    ); break;
    case 'conditions': pane = (
      <div>
        <PaneHead icon={Filter20Filled} title="Conditions" sub="Read-only summary — predefined queries embed their own FetchXml. Edit the saved query in the model-driven app to change conditions." />
        <MessageBar layout="multiline" intent="warning" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            <MessageBarTitle>$select / $filter / $orderby / $expand are locked</MessageBarTitle>
            With <code>?savedQuery=…</code> / <code>?userQuery=…</code>, the FetchXml in the query owns column projection, filter, sort and joins. Only <code>$top</code> and <code>Prefer</code> headers remain editable.
          </MessageBarBody>
        </MessageBar>
        {selectedQuery?.fetchXml ? (
          // Read-only Monaco — same theme + indent guides + line numbers as
          // the Code tab. FetchXml is generated by Dataverse so it's already
          // pretty-printed; we don't apply formatOnType (it's read-only).
          // `automaticLayout` lets Monaco adapt when the user resizes the
          // window or toggles the sidebar.
          <div style={{
            border: `1px solid ${tokens.colorNeutralStroke2}`,
            borderRadius: tokens.borderRadiusMedium,
            overflow: 'hidden',
            height: 480,
          }}>
            <Editor
              height="100%"
              language="xml"
              value={selectedQuery.fetchXml}
              theme={themeMode === 'dark' ? 'vs-dark' : 'light'}
              options={{
                readOnly: true,
                domReadOnly: true,
                lineNumbers: 'on',
                renderLineHighlight: 'line',
                minimap: { enabled: false },
                scrollBeyondLastLine: false,
                fontFamily: 'Cascadia Mono, Consolas, monospace',
                fontSize: 13,
                tabSize: 2,
                wordWrap: 'on',
                automaticLayout: true,
                padding: { top: 12, bottom: 12 },
                bracketPairColorization: { enabled: true },
                guides: { bracketPairs: true, indentation: true },
                smoothScrolling: true,
              }}
            />
          </div>
        ) : selectedQuery ? (
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            (No FetchXml available for this query — it may be a layout-only definition.)
          </Caption1>
        ) : (
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            Pick a query in Target to see its FetchXml.
          </Caption1>
        )}
      </div>
    ); break;
    case 'top':     pane = <TopEditor top={state.top} setTop={n => set('top', n, 'top')} maxPageSize={state.prefer.maxpagesize} />; break;
    case 'prefer':  pane = <PreferEditor spec={state.prefer} setSpec={p => set('prefer', p, 'prefer')} />; break;
    case 'headers': pane = <HeadersEditor items={state.headers} setItems={h => set('headers', h, 'headers')} />; break;
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
          activeNode={activeNode}
          onSelect={(id) => setActiveNode(id as ClauseId)}
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
        />
      }
    >
      <MainTabs tab={tab} onTabChange={setTab} resultCount={(result?.body as { value?: unknown[] } | null)?.value?.length ?? null}>
        {tab === 'builder' && pane}
        {tab === 'code'    && <CodeView themeMode={themeMode} inputs={codeInputs} />}
        {tab === 'results' && (
          <ResultsView
            result={result}
            mode="multi"
            table={state.table}
            // Pass the layoutxml column order so the results grid columns
            // match what Dynamics 365 / Sales Hub shows for the same view.
            // Without this the grid falls back to JSON-key order, which on
            // a `?savedQuery=…` request is whatever order Dataverse decided
            // to serialize the columns in (usually not the layout order).
            // We pass `fullName` (not `attr`) so linked-entity dotted keys
            // like `accountprimarycontactidcontactcontactid.emailaddress1`
            // stay intact.
            preferredColumnOrder={selectedQuery?.layoutColumns.map(c => c.fullName)}
          />
        )}
      </MainTabs>
    </ModeShell>
  );
}
