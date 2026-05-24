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
  Table20Regular, Table20Filled,
  Link20Regular, Link20Filled,
  PersonAccounts20Regular, PersonAccounts20Filled,
  LineHorizontal320Regular, LineHorizontal320Filled,
  Code20Filled,
} from '@fluentui/react-icons';
import {
  Caption1, Badge, tokens, MessageBar, MessageBarBody, MessageBarTitle, mergeClasses,
} from '@fluentui/react-components';
import { Sidebar } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { useStudioStyles } from '../primitives/styles';
import { PaneHead } from '../editors/PaneHead';
import { NavPropertyPicker } from '../editors/NavPropertyPicker';
import { AssociateTargetsEditor } from '../editors/AssociateTargetsEditor';
import { TargetEditor } from '../editors/TargetEditor';
import { HeadersEditor, defaultWriteHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { findTable } from '../mock/metadata';
import { findRequestType } from '../registry/requestTypes';
import { buildAssociate, buildAssociateBody, buildAssociateRequests } from '../engine/urlBuilder';
import { runtime, type ExecResult } from '../engine/runtime';
import type { AssociateState } from '../state/relateState';
import { isSingleValuedNav } from '../state/relateState';
import type { RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useLiveTable } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import {
  serializeAssociate, deserializeAssociate, hashState,
  type SavedRequest, type SerializedAssociateState,
} from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';

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
  const [state, setState] = useState(initialState);
  // Live-metadata subscription so child editors re-render when columns/relationships land.
  useLiveTable((state).table || null);
  const { entities } = useScopedEntities();
  const [activePath, setActivePath] = useState<string>('source');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);

  // Save / Load tracking
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);

  // Source-record primary-name cache. Populated when the user picks via
  // TargetEditor (the RecordPicker hands back the primary name); used for
  // the WriteResultCard success narrative.
  const [sourceName, setSourceName] = useState<string>('');

  const built = useMemo(() => buildAssociate(state), [state]);
  const body = useMemo(() => buildAssociateBody(state), [state]);
  const requests = useMemo(() => buildAssociateRequests(state), [state]);

  const tbl = findTable(state.table);
  const nav = state.navProperty ? tbl?.navigationProperties.find(n => n.name === state.navProperty) : undefined;
  const singleValued = nav ? isSingleValuedNav(nav) : false;
  // Docs-preferred shape per cardinality:
  //   N:1 → PATCH /<source>(<id>)              (body owns the link)
  //   1:N / N:N → POST /<source>(<id>)/<nav>/$ref
  // PPTB constraint: PUT $ref isn't exposed on dataverseAPI, so PATCH is
  // also the only path that actually works in PPTB for single-valued.
  const method: 'POST' | 'PATCH' = singleValued ? 'PATCH' : 'POST';

  const markDirty = (id: string) => setState(s => { const d = new Set(s.dirty); d.add(id); return { ...s, dirty: d }; });
  const set = <K extends keyof AssociateState>(k: K, v: AssociateState[K], dirtyId?: string) => {
    setState(s => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  const disabledReason =
    !tbl ? 'Pick a source table first.' :
    !state.sourceId ? 'Pick a source record.' :
    !state.navProperty ? 'Pick a navigation property.' :
    state.targets.length === 0 ? 'Pick at least one target record.' :
    state.headers.some(h => h.enabled && !h.name) ? 'Fix empty header name.' :
    null;

  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.associate(state);
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents(rs => [{
      id: `r-${Date.now()}`, modeId: 'associate',
      url: built.relativeUrl, method, ts: Date.now(),
      status: res.status, ms: res.ms, rowCount: res.ok ? requests.length : 0,
    }, ...rs].slice(0, 8));
    setState(s => ({ ...s, dirty: new Set() }));
  };

  // ── Save / Load ──
  const currentSerialized = useMemo(() => serializeAssociate(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;

  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };

  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== 'associate') return;
    const snap = entry.state as SerializedAssociateState;
    if (entities.length > 0 && !entities.some(e => e.logicalName === snap.table)) {
      window.alert(
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

  usePublishSaveContext(useMemo(() => {
    if (!state.table) return null;
    return {
      state: currentSerialized,
      modeId: 'associate' as const,
      dirty: isDirty,
      lastSavedId,
      onSaved,
      onLoadSaved,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentSerialized, isDirty, lastSavedId, state.table]));

  // ── Sidebar ──
  const sections = [
    {
      id: 'source', label: 'Target',
      meta: tbl?.displayName ?? 'Pick a table',
      items: [{
        id: 'source',
        icon: Table20Regular, iconFilled: Table20Filled,
        label: state.sourceId ? (sourceName || `${tbl?.displayName ?? ''} (selected)`) : 'Pick a record',
        badge: nav ? method : null,
        badgeAppearance: 'tint' as const, badgeColor: 'brand' as const,
        dirty: state.dirty.has('source'),
      }],
    },
    {
      id: 'relationship', label: 'Relationship',
      meta: nav ? cardinalityShort(nav.cardinality) : 'pick one',
      items: [{
        id: 'nav',
        icon: Link20Regular, iconFilled: Link20Filled,
        label: nav ? nav.name : 'Navigation property',
        code: !!nav,
        badge: nav ? cardinalityShort(nav.cardinality) : null,
        badgeAppearance: 'ghost' as const,
        dirty: state.dirty.has('nav'),
      }],
    },
    {
      id: 'targets', label: 'Related',
      meta: state.targets.length ? `${state.targets.length} target${state.targets.length === 1 ? '' : 's'}` : 'none',
      items: [
        {
          id: 'targets',
          icon: PersonAccounts20Regular, iconFilled: PersonAccounts20Filled,
          label: singleValued ? 'Target record' : 'Target records',
          badge: state.targets.length || null,
          badgeAppearance: 'tint' as const,
          badgeColor: state.targets.length > 0 ? ('success' as const) : ('subtle' as const),
          dirty: state.dirty.has('targets'),
        },
        ...(requests.length > 0 ? [{
          id: 'preview',
          icon: Code20Filled, iconFilled: Code20Filled,
          label: `${requests.length} request${requests.length === 1 ? '' : 's'}`,
          badge: requests.length, badgeAppearance: 'ghost' as const,
        }] : []),
      ],
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

  // ── Builder pane router ──
  let pane: React.ReactNode;
  const root = activePath as RootClauseId;
  switch (root) {
    case 'source':
      pane = (
        <TargetEditor
          table={state.table}
          onTableChange={t => {
            // Switching/clearing the source entity invalidates the source
            // record, the chosen nav property (belongs to the old entity's
            // relationship set), and any target ids picked under it.
            setState(s => ({
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
            setState(s => ({ ...s, navProperty: n, targets: [], targetNames: {} }));
            markDirty('nav'); markDirty('targets');
          }}
          group="relate"
          forOperation="associate"
          footer={
            <MessageBar layout="multiline" intent="info">
              <MessageBarBody>
                <MessageBarTitle>How cardinality drives the request</MessageBarTitle>
                <strong>Collection-valued</strong> (1:N / N:N) → <code>POST /…/$ref</code>{' '}
                with one request per target.
                {' '}
                <strong>Single-valued</strong> (N:1) → <code>PATCH /source(id)</code>{' '}
                with body <code>{'{ "<nav>@odata.bind": "<target-set>(<id>)" }'}</code>{' '}
                (one target). This is the docs-preferred shape — the URL bar verb updates automatically.
                {' '}
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
        />
      );
      break;
    case 'preview':
      pane = <RequestsPreview requests={requests} />;
      break;
    case 'headers':
      pane = (
        <HeadersEditor
          items={state.headers}
          setItems={h => set('headers', h, 'headers')}
          group="relate"
        />
      );
      break;
  }

  const headersMap = headerItemsToObject(state.headers, null);
  const codeInputs = {
    method,
    built,
    headers: headersMap,
    body,
    entityLogical: built.entityLogical,
    // When 2+ targets are queued, surface every request to the code generators
    multiRequests: requests.length > 1
      ? requests.map(r => ({
          method: r.method,
          relativeUrl: r.relativeUrl,
          body: r.body as unknown as Record<string, unknown>,
          description: `target ${r.targetId}`,
        }))
      : undefined,
  };

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
          method={method}
          url={built.relativeUrl}
          executeVerb={type.executeVerb}
          disabledReason={disabledReason}
          loading={loading}
          onExecute={onExecute}
        />
      }
    >
      <MainTabs tab={tab} onTabChange={setTab} resultCount={result?.ok ? requests.length : null}>
        {tab === 'builder' && pane}
        {tab === 'code'    && <CodeView themeMode={themeMode} inputs={codeInputs} />}
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
                .map(id => state.targetNames[id])
                .filter((n): n is string => !!n),
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
function RequestsPreview({ requests }: { requests: import('../engine/urlBuilder').AssociateRequest[] }) {
  const s = useStudioStyles();
  return (
    <div>
      <PaneHead
        icon={Code20Filled}
        title={`Generated requests (${requests.length})`}
        sub="Each target queues a separate HTTP request. For large counts, prefer a $batch request."
        group="relate"
      />
      <div className={mergeClasses(s.inlineCard)} style={{ padding: 12, maxWidth: 980, fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>
        {requests.map((r, i) => (
          <div key={r.targetId} style={{
            display: 'grid', gridTemplateColumns: '60px 1fr', gap: 8, padding: '8px 0',
            borderBottom: i < requests.length - 1 ? `1px solid ${tokens.colorNeutralStroke3}` : 'none',
          }}>
            <Badge appearance="filled" color={r.method === 'POST' ? 'success' : 'warning'} style={{ fontWeight: 700, justifySelf: 'start' }}>
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
          <Caption1 style={{ color: tokens.colorNeutralForeground3, fontStyle: 'italic', padding: '12px 4px', textAlign: 'center', display: 'block' }}>
            No requests queued — pick at least one target on the Related pane.
          </Caption1>
        )}
      </div>
    </div>
  );
}

function cardinalityShort(c: import('../mock/metadata').NavProperty['cardinality']): string {
  switch (c) {
    case 'OneToMany':  return '1:N';
    case 'ManyToOne':  return 'N:1';
    case 'ManyToMany': return 'N:N';
  }
}
