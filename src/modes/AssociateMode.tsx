// Associate — POST /<set>(<id>)/<nav>/$ref (collection) | PUT for single-valued.
//
// Sidebar layout:
//   ModeCard (relate · POST /accounts(id)/contact_customer_accounts/$ref)
//   * Target         → Source record (account)
//   * Relationship   → Navigation property picker
//                      (collection-valued or single, cardinality detected)
//   * Related        → One or more target records (multi-pick if N:N)
//   * Headers · 1 active
//   * Recent runs
//
// Cardinality is metadata-driven, not picker-driven. The URL bar verb flips
// automatically as the user picks a different nav property.
//
// Grounded in:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/associate-disassociate-entities-using-web-api

import { useMemo, useState } from 'react';
import {
  Table20Regular,
  Table20Filled,
  Link20Regular,
  Link20Filled,
  PersonAccounts20Regular,
  PersonAccounts20Filled,
  LineHorizontal320Regular,
  LineHorizontal320Filled,
  Code20Filled,
} from '@fluentui/react-icons';
import {
  Caption1,
  Badge,
  tokens,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  mergeClasses,
} from '@fluentui/react-components';
import { Sidebar } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { notify } from '../shell/dialogs';
import { useStudioStyles } from '../primitives/styles';
import { PaneHead } from '../editors/PaneHead';
import { NavPropertyPicker } from '../editors/NavPropertyPicker';
import { AssociateTargetsEditor, targetsTitle } from '../editors/AssociateTargetsEditor';
import { AllOrNothingOption } from '../editors/AllOrNothingOption';
import { TargetEditor } from '../editors/TargetEditor';
import { HeadersEditor, defaultWriteHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { findTable } from '../metadata/metadata';
import { findRequestType } from '../registry/requestTypes';
import {
  buildAssociate,
  buildAssociateBody,
  buildAssociateRequests,
  relateRequestHeaders,
} from '../engine/urlBuilder';
import { runtime, type ExecResult } from '../engine/runtime';
import { MAX_CHANGE_SET, relateAsChangeSet, relateDoneCount } from '../engine/dataverseExecutor';
import type { AssociateState } from '../state/relateState';
import { isSingleValuedNav } from '../state/relateState';
import type { RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useLiveTable } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import {
  serializeAssociate,
  deserializeAssociate,
  hashState,
  type SavedRequest,
  type SerializedAssociateState,
  recentRunAsSavedEntry,
} from '../state/savedRequests';
import { useModeDraft, useStashModeDraft } from '../state/modeDrafts';
import { usePublishSaveContext } from '../state/SaveContext';
import { useLatestCallback } from '../primitives/useLatestCallback';
import { invalidUrlPartReason } from '../primitives/urlParts';
import { disabledReasonFromAdvisories, firstBlocker } from '../primitives/advisories';
import { useRecentRuns } from '../state/runHistory';

// Empty initial state. The user picks a table → record → nav property →
// target(s) before the request can fire. Seeding with a mock account id
// confused PPTB users on first load ("why is this account selected?").
const initialState = (): AssociateState => ({
  table: '',
  sourceId: null,
  navProperty: null,
  targets: [],
  targetNames: {},
  headers: defaultWriteHeaders(),
  dirty: new Set(),
});

type RootClauseId = 'source' | 'nav' | 'targets' | 'preview' | 'headers';

export function AssociateMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('associate');
  const draft = useModeDraft('associate', deserializeAssociate, initialState);
  const [state, setState] = useState(draft.state);
  // Live-metadata subscription so child editors re-render when columns/relationships land.
  // The live table is a memo dependency below: build*() reads metadata via
  // findTable(), which is empty until the fetch lands — keying on `state`
  // alone left the URL / Code tab blank after picking a table (#78).
  const { table: liveTable } = useLiveTable(state.table || null);
  // Every table, not just the pickers' scope: a saved request's table only
  // has to exist (#44).
  const { allEntities: entities } = useScopedEntities();
  const [activePath, setActivePath] = useState<string>('source');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  // Kept across reloads, per environment (#226).
  const [recents, setRecents] = useRecentRuns('associate');

  // Save / Load tracking
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(draft.lastSavedId);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(draft.lastSavedHash);

  // Source-record primary-name cache. Populated when the user picks via
  // TargetEditor (the RecordPicker hands back the primary name); used for
  // the WriteResultCard success narrative.
  const [sourceName, setSourceName] = useState<string>('');

  // eslint-disable-next-line react-hooks/exhaustive-deps -- liveTable rebuilds the request once its metadata lands (#78)
  const built = useMemo(() => buildAssociate(state), [state, liveTable]);
  // eslint-disable-next-line react-hooks/exhaustive-deps -- liveTable rebuilds the request once its metadata lands (#78)
  const body = useMemo(() => buildAssociateBody(state), [state, liveTable]);
  // eslint-disable-next-line react-hooks/exhaustive-deps -- liveTable rebuilds the request once its metadata lands (#78)
  const requests = useMemo(() => buildAssociateRequests(state), [state, liveTable]);

  const tbl = findTable(state.table);
  const nav = state.navProperty
    ? tbl?.navigationProperties.find((n) => n.name === state.navProperty)
    : undefined;
  const singleValued = nav ? isSingleValuedNav(nav) : false;
  // Docs-preferred shape per cardinality:
  //   N:1 → PATCH /<source>(<id>)              (body owns the link)
  //   1:N / N:N → POST /<source>(<id>)/<nav>/$ref
  // PPTB constraint: PUT $ref isn't exposed on dataverseAPI, so PATCH is
  // also the only path that actually works in PPTB for single-valued.
  const method: 'POST' | 'PATCH' = singleValued ? 'PATCH' : 'POST';
  // All or nothing (#224): the targets go as one change set, as Execute sends them.
  const changeSet = !singleValued && relateAsChangeSet(state.allOrNothing, requests.length);

  const markDirty = (id: string) =>
    setState((s) => {
      const d = new Set(s.dirty);
      d.add(id);
      return { ...s, dirty: d };
    });
  const set = <K extends keyof AssociateState>(k: K, v: AssociateState[K], dirtyId?: string) => {
    setState((s) => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  // Ids and names that go into the URL as-is (#221).
  const urlPartProblem = invalidUrlPartReason(
    [
      ['Source record id', state.sourceId],
      ...state.targets.map((id) => ['Target record id', id] as const),
    ],
    [['Navigation property', state.navProperty]],
  );
  // Execute blockers, as advisories with a way to the fix (#294).
  const advisories = firstBlocker([
    { when: !tbl, id: 'assoc-table', title: 'Pick a source table first.', focusNode: 'source' },
    {
      when: !state.sourceId,
      id: 'assoc-source',
      title: 'Pick a source record.',
      focusNode: 'source',
    },
    {
      when: !state.navProperty,
      id: 'assoc-nav',
      title: 'Pick a navigation property.',
      focusNode: 'nav',
    },
    {
      when: state.targets.length === 0,
      id: 'assoc-targets',
      title: 'Pick at least one target record.',
      focusNode: 'targets',
    },
    {
      when: changeSet && requests.length > MAX_CHANGE_SET,
      id: 'assoc-change-set-cap',
      title: `All or nothing sends at most ${MAX_CHANGE_SET} requests in one change set; this has ${requests.length}. Remove targets or turn it off.`,
      focusNode: 'targets',
    },
    {
      when: state.headers.some((h) => h.enabled && !h.name),
      id: 'assoc-header-name',
      title: 'Fix empty header name.',
      focusNode: 'headers',
    },
    {
      when: urlPartProblem,
      id: 'assoc-url-part',
      title: urlPartProblem ?? '',
      focusNode: urlPartProblem?.startsWith('Target')
        ? 'targets'
        : urlPartProblem?.startsWith('Navigation')
          ? 'nav'
          : 'source',
    },
  ]);
  const disabledReason = disabledReasonFromAdvisories(advisories);

  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.associate(state, headerItemsToObject(state.headers, null));
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents((rs) =>
      [
        {
          id: `r-${Date.now()}`,
          snapshot: currentSerialized,
          modeId: 'associate',
          url: built.relativeUrl,
          method,
          ts: Date.now(),
          status: res.status,
          ms: res.ms,
          rowCount: relateDoneCount(res, requests.length) ?? 0,
          // Recent runs shows no row count for writes (#309); say how many links were done.
          summary: `${relateDoneCount(res, requests.length) ?? 0} of ${requests.length} linked`,
        },
        ...rs,
      ].slice(0, 8),
    );
    setState((s) => ({ ...s, dirty: new Set() }));
  };

  // ── Save / Load ──
  const currentSerialized = useMemo(() => serializeAssociate(state), [state]);
  useStashModeDraft('associate', currentSerialized, lastSavedId, lastSavedHash);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;

  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };

  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== 'associate') return;
    const snap = entry.state as SerializedAssociateState;
    if (entities.length > 0 && !entities.some((e) => e.logicalName === snap.table)) {
      notify(
        `Can't load "${entry.name}": entity \`${snap.table}\` ` +
          `isn't available in this environment. The solution may have been ` +
          `removed or you may be connected to a different org.`,
      );
      return;
    }
    setState(deserializeAssociate(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setSourceName('');
    setActivePath('source');
  };

  // Recent runs (#2) — restore a run's snapshot through the same loader as
  // the Saved library, but keep the current save baseline: a run isn't a
  // saved entry, so Save keeps targeting whatever it targeted before.
  const onRecentClick = (run: RecentRun) => {
    const entry = recentRunAsSavedEntry(run, 'associate');
    if (!entry) return;
    const keepId = lastSavedId;
    const keepHash = lastSavedHash;
    onLoadSaved(entry);
    setLastSavedId(keepId);
    setLastSavedHash(keepHash);
  };

  // The header holds on to these; the stable wrappers always run this
  // render's handlers, not the ones from when the context was published.
  const onSavedLatest = useLatestCallback(onSaved);
  const onLoadSavedLatest = useLatestCallback(onLoadSaved);
  usePublishSaveContext(
    useMemo(() => {
      if (!state.table) return null;
      return {
        state: currentSerialized,
        modeId: 'associate' as const,
        dirty: isDirty,
        lastSavedId,
        onSaved: onSavedLatest,
        onLoadSaved: onLoadSavedLatest,
      };
    }, [currentSerialized, isDirty, lastSavedId, state.table, onSavedLatest, onLoadSavedLatest]),
  );

  // ── Sidebar ──
  const sections = [
    {
      id: 'source',
      label: 'Target',
      meta: tbl?.displayName ?? 'Pick a table',
      items: [
        {
          id: 'source',
          icon: Table20Regular,
          iconFilled: Table20Filled,
          label: state.sourceId
            ? sourceName || `${tbl?.displayName ?? ''} (selected)`
            : 'Pick a record',
          badge: nav ? method : null,
          badgeAppearance: 'tint' as const,
          badgeColor: 'brand' as const,
          dirty: state.dirty.has('source'),
        },
      ],
    },
    {
      id: 'relationship',
      label: 'Relationship',
      meta: nav ? cardinalityShort(nav.cardinality) : 'pick one',
      items: [
        {
          id: 'nav',
          icon: Link20Regular,
          iconFilled: Link20Filled,
          label: nav ? nav.name : 'Navigation property',
          code: !!nav,
          badge: nav ? cardinalityShort(nav.cardinality) : null,
          badgeAppearance: 'ghost' as const,
          dirty: state.dirty.has('nav'),
        },
      ],
    },
    {
      id: 'targets',
      label: 'Related',
      meta: state.targets.length
        ? `${state.targets.length} target${state.targets.length === 1 ? '' : 's'}`
        : 'none',
      items: [
        {
          id: 'targets',
          icon: PersonAccounts20Regular,
          iconFilled: PersonAccounts20Filled,
          label: targetsTitle(singleValued, state.targets.length),
          badge: state.targets.length || null,
          badgeAppearance: 'tint' as const,
          badgeColor: state.targets.length > 0 ? ('success' as const) : ('subtle' as const),
          dirty: state.dirty.has('targets'),
        },
        ...(requests.length > 0
          ? [
              {
                id: 'preview',
                icon: Code20Filled,
                iconFilled: Code20Filled,
                label: `${requests.length} request${requests.length === 1 ? '' : 's'}${changeSet ? ' · all or nothing' : ''}`,
                badge: requests.length,
                badgeAppearance: 'ghost' as const,
              },
            ]
          : []),
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

  // ── Builder pane router ──
  let pane: React.ReactNode;
  const root = activePath as RootClauseId;
  switch (root) {
    case 'source':
      pane = (
        <TargetEditor
          table={state.table}
          onTableChange={(t) => {
            // Switching/clearing the source entity invalidates the source
            // record, the chosen nav property (belongs to the old entity's
            // relationship set), and any target ids picked under it.
            setState((s) => ({
              ...s,
              table: t,
              sourceId: null,
              navProperty: null,
              targets: [],
              targetNames: {},
              dirty: new Set(['source', 'nav', 'targets']),
            }));
            setSourceName('');
            setResult(null);
            // New table → fresh draft. Forget the last save so Save can't
            // overwrite an entry that belongs to the old table (#58).
            setLastSavedId(undefined);
            setLastSavedHash(null);
          }}
          recordId={state.sourceId}
          onRecordChange={(id, primary) => {
            set('sourceId', id, 'source');
            setSourceName(primary ?? '');
          }}
          group="relate"
          sub="Pick the source entity and the record that the relationship anchors on."
        />
      );
      break;
    case 'nav':
      pane = (
        <NavPropertyPicker
          table={state.table}
          value={state.navProperty}
          onChange={(n) => {
            // Switching nav → drop targets (target entity may change).
            // Also drop the name cache since those names belong to the
            // old target entity.
            setState((s) => ({ ...s, navProperty: n, targets: [], targetNames: {} }));
            markDirty('nav');
            markDirty('targets');
          }}
          group="relate"
          forOperation="associate"
          footer={
            <MessageBar layout="multiline" intent="info">
              <MessageBarBody>
                <MessageBarTitle>How cardinality drives the request</MessageBarTitle>
                <strong>Collection-valued</strong> (1:N / N:N) → <code>POST /…/$ref</code> with one
                request per target. <strong>Single-valued</strong> (N:1) →{' '}
                <code>PATCH /source(id)</code> with body{' '}
                <code>{'{ "<nav>@odata.bind": "<target-set>(<id>)" }'}</code> (one target). This is
                the docs-preferred shape — the URL bar verb updates automatically. The PATCH is sent
                with <code>If-Match: *</code>, so a wrong source id fails with 404 instead of
                creating a record; Power Platform ToolBox asks you to approve that header once.{' '}
                <a
                  href="https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/associate-disassociate-entities-using-web-api"
                  target="_blank"
                  rel="noopener noreferrer"
                  style={{ color: 'inherit' }}
                >
                  MS Learn ↗
                </a>
              </MessageBarBody>
            </MessageBar>
          }
        />
      );
      break;
    case 'targets':
      pane = (
        <AssociateTargetsEditor
          table={state.table}
          navProperty={state.navProperty}
          targets={state.targets}
          setTargets={(t) => set('targets', t, 'targets')}
          targetNames={state.targetNames}
          setTargetNames={(n) => set('targetNames', n)}
          group="relate"
          resolvedMethod={nav ? method : null}
          changeSet={changeSet}
          footer={
            nav && !singleValued ? (
              <AllOrNothingOption
                checked={!!state.allOrNothing}
                onChange={(v) => set('allOrNothing', v || undefined, 'targets')}
                targets={state.targets.length}
                verb="POST"
              />
            ) : undefined
          }
        />
      );
      break;
    case 'preview':
      pane = <RequestsPreview requests={requests} changeSet={changeSet} />;
      break;
    case 'headers':
      pane = (
        <HeadersEditor
          policy="relate"
          items={state.headers}
          setItems={(h) => set('headers', h, 'headers')}
          group="relate"
        />
      );
      break;
  }

  // Single-valued: the PATCH carries If-Match: *, as DRS sends it (#254).
  const headersMap = relateRequestHeaders(method, headerItemsToObject(state.headers, null));
  const codeInputs = {
    method,
    built,
    headers: headersMap,
    body,
    entityLogical: built.entityLogical,
    // When 2+ targets are queued, surface every request to the code generators
    multiRequests:
      requests.length > 1
        ? requests.map((r) => ({
            method: r.method,
            relativeUrl: r.relativeUrl,
            body: r.body as unknown as Record<string, unknown>,
            description: `target ${r.targetId}`,
          }))
        : undefined,
    changeSet,
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
          onRecentClick={onRecentClick}
        />
      }
      urlBar={
        <UrlBar
          method={method}
          url={built.relativeUrl}
          executeVerb={type.executeVerb}
          // What this request changes, for the write confirmation (#241).
          writeConfirm={{
            operation: nav ? `Associate (${nav.name})` : 'Associate',
            table: state.table,
            scope: { kind: 'records', count: Math.max(1, state.targets.length) },
            highImpact: state.targets.length > 1,
          }}
          disabledReason={disabledReason}
          loading={loading}
          onExecute={onExecute}
          advisories={advisories}
          onAdvisoryFocus={setActivePath}
        />
      }
    >
      <MainTabs
        tab={tab}
        onTabChange={setTab}
        resultCount={relateDoneCount(result, requests.length)}
      >
        {tab === 'builder' && pane}
        {tab === 'code' && <CodeView themeMode={themeMode} inputs={codeInputs} />}
        {tab === 'results' && (
          <ResultsView
            result={result}
            mode="single"
            table={state.table}
            writeContext={{
              operation: 'associate',
              table: state.table,
              recordId: state.sourceId,
              recordName: sourceName || null,
              navProperty: state.navProperty ?? undefined,
              targetCount: state.targets.length,
              targetNames: state.targets
                .map((id) => state.targetNames[id])
                .filter((n): n is string => !!n),
              targetNameById: state.targetNames,
            }}
          />
        )}
      </MainTabs>
    </ModeShell>
  );
}

// ──────────────────────────────────────────────────────────────
// RequestsPreview — JSON-style preview of the N requests that will fire
// ──────────────────────────────────────────────────────────────
function RequestsPreview({
  requests,
  changeSet,
}: {
  requests: import('../engine/urlBuilder').AssociateRequest[];
  changeSet: boolean;
}) {
  const s = useStudioStyles();
  return (
    <div>
      <PaneHead
        icon={Code20Filled}
        title={`Generated requests (${requests.length})`}
        sub={
          changeSet
            ? 'All or nothing: these requests go together as one $batch change set, so all succeed or none do.'
            : 'Each target queues a separate HTTP request. For large counts, prefer a $batch request.'
        }
        group="relate"
      />
      <div
        className={mergeClasses(s.inlineCard)}
        style={{ padding: 12, maxWidth: 980, fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}
      >
        {requests.map((r, i) => (
          <div
            key={r.targetId}
            style={{
              display: 'grid',
              gridTemplateColumns: '60px 1fr',
              gap: 8,
              padding: '8px 0',
              borderBottom:
                i < requests.length - 1 ? `1px solid ${tokens.colorNeutralStroke3}` : 'none',
            }}
          >
            <Badge
              appearance="filled"
              color={r.method === 'POST' ? 'success' : 'warning'}
              style={{ fontWeight: 700, justifySelf: 'start' }}
            >
              {r.method}
            </Badge>
            <div style={{ minWidth: 0 }}>
              <div style={{ wordBreak: 'break-all' }}>{r.relativeUrl}</div>
              <div style={{ marginTop: 4, color: tokens.colorNeutralForeground3 }}>
                {JSON.stringify(r.body)}
              </div>
            </div>
          </div>
        ))}
        {requests.length === 0 && (
          <Caption1
            style={{
              color: tokens.colorNeutralForeground3,
              fontStyle: 'italic',
              padding: '12px 4px',
              textAlign: 'center',
              display: 'block',
            }}
          >
            No requests queued — pick at least one target on the Related pane.
          </Caption1>
        )}
      </div>
    </div>
  );
}

function cardinalityShort(c: import('../metadata/metadata').NavProperty['cardinality']): string {
  switch (c) {
    case 'OneToMany':
      return '1:N';
    case 'ManyToOne':
      return 'N:1';
    case 'ManyToMany':
      return 'N:N';
  }
}
