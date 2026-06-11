// Disassociate — cardinality-dependent wire shape:
//
//   • Collection-valued (1:N / N:N):
//       DELETE /<set>(<id>)/<nav>(<targetId>)/$ref   (one per target)
//
//   • Single-valued (N:1):
//       PATCH  /<set>(<id>)   with body { "<nav>@odata.bind": null }
//       — clears the lookup column. Per docs this is the preferred shape;
//       the DELETE /…/$ref form (no target id) also works at the wire BUT
//       isn't reachable through PPTB's dataverseAPI (its .disassociate
//       requires a targetId), so PATCH is the only path that actually
//       executes inside PPTB.
//
// Grounded in:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/associate-disassociate-entities-using-web-api

import { useEffect, useMemo, useState } from 'react';
import {
  Table20Regular,
  Table20Filled,
  Link20Regular,
  Link20Filled,
  PersonAccounts20Regular,
  PersonAccounts20Filled,
  LineHorizontal320Regular,
  LineHorizontal320Filled,
  Warning20Filled,
} from '@fluentui/react-icons';
import {
  Caption1,
  Badge,
  tokens,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  mergeClasses,
  Persona,
  Spinner,
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
import { buildDisassociate, buildDisassociateRequests } from '../engine/urlBuilder';
import { runtime, type ExecResult } from '../engine/runtime';
import type { DisassociateState } from '../state/relateState';
import { isCollectionValuedNav, isSingleValuedNav } from '../state/relateState';
import type { RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useLiveTable } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import {
  serializeDisassociate,
  deserializeDisassociate,
  hashState,
  type SavedRequest,
  type SerializedDisassociateState,
} from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';

// Empty initial state — same pattern as MergeMode/AssociateMode. User picks
// the table → source record → nav property → targets (if collection-valued).
const initialState = (): DisassociateState => ({
  table: '',
  sourceId: null,
  navProperty: null,
  targetIds: [],
  targetNames: {},
  headers: defaultWriteHeaders(),
  dirty: new Set(),
});

type RootClauseId = 'source' | 'nav' | 'target' | 'headers';

export function DisassociateMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('disassociate');
  const [state, setState] = useState(initialState);
  // Live-metadata subscription so child editors re-render when columns/relationships land.
  useLiveTable(state.table || null);
  const { entities } = useScopedEntities();
  const [activePath, setActivePath] = useState<string>('source');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);

  // Save / Load tracking
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);

  // Source-record primary-name cache — populated when the user picks via
  // TargetEditor. Used for the WriteResultCard narrative + sidebar label.
  const [sourceName, setSourceName] = useState<string>('');

  // For single-valued: we live-fetch the source row's current lookup value so
  // the Target pane can preview "this is what'll be cleared".
  const [currentLookup, setCurrentLookup] = useState<{ id: string; name: string | null } | null>(
    null,
  );
  const [currentLookupLoading, setCurrentLookupLoading] = useState(false);

  const built = useMemo(() => buildDisassociate(state), [state]);
  const requests = useMemo(() => buildDisassociateRequests(state), [state]);
  const tbl = findTable(state.table);
  const nav = state.navProperty
    ? tbl?.navigationProperties.find((n) => n.name === state.navProperty)
    : undefined;
  const isCollection = nav ? isCollectionValuedNav(nav) : false;
  const isSingle = nav ? isSingleValuedNav(nav) : false;

  // Docs-preferred verb per cardinality (see file header).
  const method: 'DELETE' | 'PATCH' = isSingle ? 'PATCH' : 'DELETE';

  // ── Live source-lookup fetch (single-valued only) ──
  //
  // When single-valued nav is picked, we GET the source row to read the
  // current value of the lookup. The Target pane previews it so the user
  // sees "you're about to clear: Contoso, Ltd." rather than just
  // "PATCH … @odata.bind: null".
  useEffect(() => {
    let cancelled = false;
    if (!isSingle || !state.sourceId || !tbl || !nav) {
      setCurrentLookup(null);
      return;
    }
    if (typeof window === 'undefined' || !window.dataverseAPI?.queryData) return;
    setCurrentLookupLoading(true);
    // Project the lookup's raw value + formatted-value annotation. The
    // formatted-value gives us the related row's primary-name without a
    // second fetch.
    const url = `/${tbl.entitySetName}(${state.sourceId})?$select=_${nav.name}_value`;
    window.dataverseAPI
      .queryData(url)
      .then((res) => {
        if (cancelled) return;
        const r = res as Record<string, unknown> | null;
        const lookupId = r ? String(r[`_${nav.name}_value`] ?? '') : '';
        const lookupName = r
          ? String(r[`_${nav.name}_value@OData.Community.Display.V1.FormattedValue`] ?? '')
          : '';
        setCurrentLookup(lookupId ? { id: lookupId, name: lookupName || null } : null);
        setCurrentLookupLoading(false);
      })
      .catch(() => {
        if (cancelled) return;
        setCurrentLookup(null);
        setCurrentLookupLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [isSingle, state.sourceId, state.navProperty, tbl, nav]);

  const markDirty = (id: string) =>
    setState((s) => {
      const d = new Set(s.dirty);
      d.add(id);
      return { ...s, dirty: d };
    });
  const set = <K extends keyof DisassociateState>(
    k: K,
    v: DisassociateState[K],
    dirtyId?: string,
  ) => {
    setState((s) => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  const disabledReason = !tbl
    ? 'Pick a source table first.'
    : !state.sourceId
      ? 'Pick a source record.'
      : !state.navProperty
        ? 'Pick a navigation property.'
        : isCollection && state.targetIds.length === 0
          ? 'Pick at least one target record to disassociate.'
          : isSingle && currentLookup === null && !currentLookupLoading
            ? "Source's lookup is already empty — nothing to disassociate."
            : state.headers.some((h) => h.enabled && !h.name)
              ? 'Fix empty header name.'
              : null;

  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.disassociate(state);
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents((rs) =>
      [
        {
          id: `r-${Date.now()}`,
          modeId: 'disassociate',
          url: built.relativeUrl,
          method,
          ts: Date.now(),
          status: res.status,
          ms: res.ms,
          rowCount: res.ok ? requests.length : 0,
        },
        ...rs,
      ].slice(0, 8),
    );
    setState((s) => ({ ...s, dirty: new Set() }));
  };

  // ── Save / Load ──
  const currentSerialized = useMemo(() => serializeDisassociate(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;

  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };

  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== 'disassociate') return;
    const snap = entry.state as SerializedDisassociateState;
    if (entities.length > 0 && !entities.some((e) => e.logicalName === snap.table)) {
      window.alert(
        `Can't load "${entry.name}": entity \`${snap.table}\` ` +
          `isn't available in this environment. The solution may have been ` +
          `removed or you may be connected to a different org.`,
      );
      return;
    }
    setState(deserializeDisassociate(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setSourceName('');
    setActivePath('source');
  };

  usePublishSaveContext(
    useMemo(() => {
      if (!state.table) return null;
      return {
        state: currentSerialized,
        modeId: 'disassociate' as const,
        dirty: isDirty,
        lastSavedId,
        onSaved,
        onLoadSaved,
      };
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [currentSerialized, isDirty, lastSavedId, state.table]),
  );

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
      id: 'target',
      label: 'Related',
      meta: isSingle
        ? currentLookup
          ? `clear: ${currentLookup.name ?? currentLookup.id.slice(0, 8)}…`
          : 'lookup empty'
        : state.targetIds.length
          ? `${state.targetIds.length} target${state.targetIds.length === 1 ? '' : 's'}`
          : 'none',
      items: [
        {
          id: 'target',
          icon: PersonAccounts20Regular,
          iconFilled: PersonAccounts20Filled,
          label: isSingle
            ? 'Target — implicit'
            : state.targetIds.length > 1
              ? 'Target records'
              : 'Target record',
          badge: isSingle ? '∅' : state.targetIds.length || null,
          badgeAppearance: 'tint' as const,
          badgeColor: state.targetIds.length > 0 ? ('success' as const) : ('subtle' as const),
          dirty: state.dirty.has('target'),
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
  const root = activePath as RootClauseId;
  switch (root) {
    case 'source':
      pane = (
        <TargetEditor
          table={state.table}
          onTableChange={(t) => {
            // Switching/clearing the source entity invalidates the source
            // record + chosen nav property + target ids (all bound to the
            // old entity's relationship set).
            setState((s) => ({
              ...s,
              table: t,
              sourceId: null,
              navProperty: null,
              targetIds: [],
              targetNames: {},
              dirty: new Set(['source', 'nav', 'target']),
            }));
            setSourceName('');
            setCurrentLookup(null);
            setResult(null);
          }}
          recordId={state.sourceId}
          onRecordChange={(id, primary) => {
            set('sourceId', id, 'source');
            setSourceName(primary ?? '');
          }}
          group="relate"
          sub="Pick the source record holding the relationship link to remove."
        />
      );
      break;
    case 'nav':
      pane = (
        <div>
          <NavPropertyPicker
            table={state.table}
            value={state.navProperty}
            onChange={(n) => {
              setState((s) => ({ ...s, navProperty: n, targetIds: [], targetNames: {} }));
              markDirty('nav');
              markDirty('target');
            }}
            group="relate"
            forOperation="disassociate"
            footer={nav && <UrlShapeComparison nav={nav} />}
          />
        </div>
      );
      break;
    case 'target':
      pane = (
        <TargetPane
          state={state}
          setTargetIds={(ids) => set('targetIds', ids, 'target')}
          targetNames={state.targetNames}
          setTargetNames={(n) => set('targetNames', n)}
          currentLookup={currentLookup}
          currentLookupLoading={currentLookupLoading}
        />
      );
      break;
    case 'headers':
      pane = (
        <HeadersEditor
          items={state.headers}
          setItems={(h) => set('headers', h, 'headers')}
          group="relate"
        />
      );
      break;
  }

  const headersMap = headerItemsToObject(state.headers, null);
  // First request drives the URL bar + Code-tab "primary" snippet; the
  // multiRequests payload carries the rest for the collection-valued
  // multi-target case (one DELETE per target).
  const firstReq = requests[0];
  const codeInputs = {
    method,
    built,
    headers: headersMap,
    body: firstReq?.body,
    entityLogical: built.entityLogical,
    multiRequests:
      requests.length > 1
        ? requests.map((r) => ({
            method: r.method,
            relativeUrl: r.relativeUrl,
            body: r.body as unknown as Record<string, unknown> | undefined,
            description: r.targetId ? `target ${r.targetId}` : 'clear lookup (null)',
          }))
        : undefined,
  };

  // ── Build the WriteResultCard context. For single-valued, the "targets"
  //    are really "the previous lookup value we just cleared"; for
  //    collection-valued it's the picked target records.
  const writeCtxTargetNames = isSingle
    ? currentLookup?.name
      ? [currentLookup.name]
      : []
    : state.targetIds.map((id) => state.targetNames[id]).filter((n): n is string => !!n);
  const writeCtxTargetCount = isSingle ? (currentLookup ? 1 : 0) : state.targetIds.length;

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
        {tab === 'code' && <CodeView themeMode={themeMode} inputs={codeInputs} />}
        {tab === 'results' && (
          <ResultsView
            result={result}
            mode="single"
            table={state.table}
            writeContext={{
              operation: 'disassociate',
              table: state.table,
              recordId: state.sourceId,
              recordName: sourceName || null,
              navProperty: state.navProperty ?? undefined,
              targetCount: writeCtxTargetCount,
              targetNames: writeCtxTargetNames,
            }}
          />
        )}
      </MainTabs>
    </ModeShell>
  );
}

// ──────────────────────────────────────────────────────────────
// Target pane — record picker for collection, live-fetched info for single-valued
// ──────────────────────────────────────────────────────────────
function TargetPane({
  state,
  setTargetIds,
  targetNames,
  setTargetNames,
  currentLookup,
  currentLookupLoading,
}: {
  state: DisassociateState;
  setTargetIds: (ids: string[]) => void;
  targetNames: Record<string, string>;
  setTargetNames: (next: Record<string, string>) => void;
  currentLookup: { id: string; name: string | null } | null;
  currentLookupLoading: boolean;
}) {
  const tbl = findTable(state.table);
  const nav = state.navProperty
    ? tbl?.navigationProperties.find((n) => n.name === state.navProperty)
    : undefined;
  if (!nav) {
    return (
      <div>
        <PaneHead
          icon={PersonAccounts20Filled}
          title="Target record"
          sub="Pick a navigation property first."
          group="relate"
        />
        <MessageBar layout="multiline" intent="info" style={{ maxWidth: 720 }}>
          <MessageBarBody>
            Switch to the <strong>Relationship</strong> pane and pick a navigation property.
          </MessageBarBody>
        </MessageBar>
      </div>
    );
  }
  const targetTbl = findTable(nav.targetEntity);
  const isSingle = isSingleValuedNav(nav);

  if (isSingle) {
    return (
      <div>
        <PaneHead
          icon={PersonAccounts20Filled}
          title="Target — implicit (single-valued)"
          sub={
            <>
              Single-valued nav props don't take a target id. The <code>PATCH</code> body{' '}
              <code>{`{ "${nav.name}@odata.bind": null }`}</code> clears the lookup on the source
              row.
            </>
          }
          group="relate"
        />
        <MessageBar
          layout="multiline"
          intent="warning"
          icon={<Warning20Filled />}
          style={{ maxWidth: 880, marginBottom: 14 }}
        >
          <MessageBarBody>
            <MessageBarTitle>This clears the lookup.</MessageBarTitle>
            <code>{nav.name}</code> on the source row becomes <code>null</code>. The related row is{' '}
            <strong>not</strong> deleted — only the link is removed.
          </MessageBarBody>
        </MessageBar>
        <div
          style={{
            border: `1px solid ${tokens.colorNeutralStroke2}`,
            borderRadius: tokens.borderRadiusMedium,
            padding: 12,
            maxWidth: 560,
          }}
        >
          <Caption1
            style={{ color: tokens.colorNeutralForeground3, display: 'block', marginBottom: 8 }}
          >
            Current lookup value{currentLookupLoading ? ' (loading…)' : ''}
          </Caption1>
          {currentLookupLoading ? (
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <Spinner size="tiny" />
              <Caption1>Fetching the lookup's current value from Dataverse…</Caption1>
            </div>
          ) : currentLookup ? (
            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              <Persona
                size="small"
                name={currentLookup.name ?? '(unknown)'}
                avatar={{ color: 'colorful' }}
              />
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 12, fontWeight: 600 }}>
                  {currentLookup.name ?? '(name not available)'}
                </div>
                <div
                  style={{
                    fontFamily: tokens.fontFamilyMonospace,
                    fontSize: 10,
                    color: tokens.colorNeutralForeground3,
                  }}
                >
                  /{targetTbl?.entitySetName ?? nav.targetEntity}({currentLookup.id})
                </div>
              </div>
            </div>
          ) : (
            <Caption1 style={{ color: tokens.colorNeutralForeground3, fontStyle: 'italic' }}>
              Source's lookup is already empty — the PATCH would be a no-op.
            </Caption1>
          )}
        </div>
      </div>
    );
  }

  // Collection-valued: reuse the AssociateTargetsEditor for multi-target picking.
  // Same UX as Associate — live RecordPicker, dedupe, name cache.
  return (
    <div>
      <PaneHead
        icon={PersonAccounts20Filled}
        title={state.targetIds.length > 1 ? 'Target records' : 'Target record'}
        sub={
          <>
            Pick one or more <code>{targetTbl?.logicalName}</code> rows to remove from the
            collection. Each fires a separate DELETE request per docs — the target rows are{' '}
            <strong>not</strong> deleted, only the links.
          </>
        }
        group="relate"
      />
      <AssociateTargetsEditor
        table={state.table}
        navProperty={state.navProperty}
        targets={state.targetIds}
        setTargets={setTargetIds}
        targetNames={targetNames}
        setTargetNames={setTargetNames}
        group="relate"
        resolvedMethod={null}
      />
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// URL shape comparison card — surfaces both wire shapes per cardinality
// ──────────────────────────────────────────────────────────────
function UrlShapeComparison({ nav }: { nav: import('../mock/metadata').NavProperty }) {
  const s = useStudioStyles();
  const isSingle = isSingleValuedNav(nav);
  return (
    <div className={mergeClasses(s.inlineCard)} style={{ padding: 12 }}>
      <strong style={{ fontSize: 12, display: 'block', marginBottom: 8 }}>
        URL shape per cardinality
      </strong>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 10 }}>
        <UrlShapeCard
          label="Collection-valued (1:N / N:N)"
          syntax={'DELETE /<set>(<id>)/<nav>(<related-id>)/$ref'}
          note="One DELETE per target id."
          active={!isSingle}
        />
        <UrlShapeCard
          label="Single-valued (N:1)"
          syntax={`PATCH /<set>(<id>)\n{ "<nav>@odata.bind": null }`}
          note="Clears the lookup — docs-preferred over DELETE/$ref (and the only path PPTB exposes)."
          active={isSingle}
        />
      </div>
    </div>
  );
}

function UrlShapeCard({
  label,
  syntax,
  note,
  active,
}: {
  label: string;
  syntax: string;
  note: string;
  active: boolean;
}) {
  return (
    <div
      style={{
        padding: 10,
        border: `1px solid ${active ? tokens.colorBrandStroke1 : tokens.colorNeutralStroke2}`,
        background: active ? tokens.colorBrandBackground2 : 'transparent',
        borderRadius: tokens.borderRadiusMedium,
        opacity: active ? 1 : 0.55,
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 4 }}>
        <strong style={{ fontSize: 11 }}>{label}</strong>
        {active && (
          <Badge
            appearance="tint"
            color="brand"
            size="extra-small"
            style={{ marginLeft: 'auto', fontWeight: 700 }}
          >
            ACTIVE
          </Badge>
        )}
      </div>
      <code
        style={{
          fontFamily: tokens.fontFamilyMonospace,
          fontSize: 11,
          wordBreak: 'break-all',
          display: 'block',
          whiteSpace: 'pre-wrap',
        }}
      >
        {syntax}
      </code>
      <Caption1 style={{ display: 'block', marginTop: 4, color: tokens.colorNeutralForeground3 }}>
        {note}
      </Caption1>
    </div>
  );
}

function cardinalityShort(c: import('../mock/metadata').NavProperty['cardinality']): string {
  switch (c) {
    case 'OneToMany':
      return '1:N';
    case 'ManyToOne':
      return 'N:1';
    case 'ManyToMany':
      return 'N:N';
  }
}
