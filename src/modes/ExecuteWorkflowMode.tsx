// Execute Workflow — POST /workflows(<wf-id>)/Microsoft.Dynamics.CRM.ExecuteWorkflow
// Body: { "EntityId": "<target-record-guid>" }
//
// Per docs:
//   • On-demand workflows are defined in the workflow table and have
//     `primaryentity` = the entity they operate on.
//   • The body's `EntityId` must be a GUID of that entity.
//   • Response: 204 No Content (workflow runs async — check workflowlogs).
//
// Reference:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-web-api-functions
//   (workflows aren't actions/functions per se but invoke similarly)

import { useEffect, useMemo, useState } from 'react';
import {
  Flowchart20Regular, Flowchart20Filled,
  Table20Regular, Table20Filled,
  LineHorizontal320Regular, LineHorizontal320Filled,
} from '@fluentui/react-icons';
import {
  Field, Combobox, Option, Caption1, Badge, tokens, MessageBar, MessageBarBody, MessageBarTitle,
  Spinner,
} from '@fluentui/react-components';
import { Sidebar } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { PaneHead } from '../editors/PaneHead';
import { KvGrid } from '../primitives/KvGrid';
import { RecordPicker } from '../primitives/RecordPicker';
import { HeadersEditor, defaultWriteHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { findTable } from '../mock/metadata';
import { findRequestType } from '../registry/requestTypes';
import { buildExecuteWorkflow, buildExecuteWorkflowBody } from '../engine/urlBuilder';
import { runtime, type ExecResult } from '../engine/runtime';
import type { ExecuteWorkflowState } from '../state/executeState';
import type { RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useScopedEntities } from '../host/useScopedEntities';
import { useLiveTable } from '../host/useLiveMetadata';
import {
  serializeExecuteWorkflow, deserializeExecuteWorkflow, hashState,
  type SavedRequest, type SerializedExecuteWorkflowState,
} from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';

// Workflow shape used by the picker — the wire response fields are
// camelCased differently per the OData convention (workflowid not Id, etc).
// Normalized at fetch time so the picker doesn't have to branch.
interface WorkflowRow {
  id: string;
  name: string;
  uniqueName: string | null;
  primaryEntity: string;
  description: string | null;
}

// Empty initial state — user picks from the live workflows list once it loads.
const initialState = (): ExecuteWorkflowState => ({
  workflowId: null,
  entityId: null,
  headers: defaultWriteHeaders(),
  dirty: new Set(),
});

type RootClauseId = 'pick' | 'target' | 'headers';

export function ExecuteWorkflowMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('exec-workflow');
  const [state, setState] = useState(initialState);
  const [activePath, setActivePath] = useState<string>('pick');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);

  const built = useMemo(() => buildExecuteWorkflow(state), [state]);
  const body = useMemo(() => buildExecuteWorkflowBody(state), [state]);

  // Live workflow list — fetch from the `workflows` table. Filter:
  //   type eq 1        → Definition (not an Activation Copy)
  //   category eq 0    → Workflow (not Dialog / Business Process Flow / etc)
  //   statecode eq 1   → Activated (only activated workflows can be triggered)
  // Reference:
  //   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/workflow
  const [workflows, setWorkflows] = useState<WorkflowRow[]>([]);
  const [workflowsLoading, setWorkflowsLoading] = useState(false);
  const [workflowsError, setWorkflowsError] = useState<string | null>(null);
  useEffect(() => {
    let cancelled = false;
    if (typeof window === 'undefined' || !window.dataverseAPI?.queryData) {
      // Outside PPTB — no live host to query. Surface a clear error so the
      // picker doesn't silently sit empty.
      setWorkflows([]);
      setWorkflowsError('Workflows can only be loaded from a live Dataverse connection (PPTB).');
      return;
    }
    setWorkflowsLoading(true);
    setWorkflowsError(null);
    // Filter: activated (statecode=1), Definition (type=1, not Activation
    // Copies which are type=2), Workflow category (category=0, not Dialog/
    // BPF/etc), AND ondemand=true (excludes auto-triggered workflows that
    // can't be invoked via ExecuteWorkflow). Per the user's confirmed
    // filter; matches Dynamics 365 / Sales Hub's "on-demand workflow" list.
    const query =
      `workflows?$select=workflowid,name,uniquename,primaryentity,description` +
      `&$filter=(statecode eq 1 and type eq 1 and ondemand eq true and category eq 0)` +
      `&$orderby=name asc`;
    window.dataverseAPI.queryData(query).then(res => {
      if (cancelled) return;
      const rows = (res as { value?: Array<Record<string, unknown>> })?.value ?? [];
      setWorkflows(rows.map(r => ({
        id: String(r.workflowid ?? ''),
        name: String(r.name ?? '(unnamed)'),
        uniqueName: r.uniquename != null ? String(r.uniquename) : null,
        primaryEntity: String(r.primaryentity ?? ''),
        description: r.description != null ? String(r.description) : null,
      })));
      setWorkflowsLoading(false);
    }).catch(e => {
      if (cancelled) return;
      setWorkflowsError(e instanceof Error ? e.message : String(e));
      setWorkflowsLoading(false);
    });
    return () => { cancelled = true; };
  }, []);

  const wf = state.workflowId ? workflows.find(w => w.id === state.workflowId) : undefined;
  // Warm the workflow's primary-entity metadata as soon as a workflow is
  // picked. findTable() reads synchronously from the live registry; the
  // hook below kicks off the fetch when needed and triggers a re-render
  // once the metadata lands. Without this, the display name + entity-set
  // name stay as raw logical names indefinitely.
  useLiveTable(wf?.primaryEntity ?? null);
  const tbl = wf ? findTable(wf.primaryEntity) : undefined;

  // Save / Load
  const { entities } = useScopedEntities();
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);
  const [entityRecordName, setEntityRecordName] = useState<string>('');

  const markDirty = (id: string) => setState(s => { const d = new Set(s.dirty); d.add(id); return { ...s, dirty: d }; });
  const set = <K extends keyof ExecuteWorkflowState>(k: K, v: ExecuteWorkflowState[K], dirtyId?: string) => {
    setState(s => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  const disabledReason =
    !wf ? 'Pick a workflow.' :
    !state.entityId ? `Pick a ${wf.primaryEntity} record to run the workflow on.` :
    state.headers.some(h => h.enabled && !h.name) ? 'Fix empty header name.' :
    null;

  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.workflow(state);
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents(rs => [{
      id: `r-${Date.now()}`, modeId: 'exec-workflow',
      url: built.relativeUrl, method: 'POST', ts: Date.now(),
      status: res.status, ms: res.ms, rowCount: res.ok ? 1 : 0,
    }, ...rs].slice(0, 8));
    setState(s => ({ ...s, dirty: new Set() }));
  };

  // ── Save / Load ──
  const currentSerialized = useMemo(() => serializeExecuteWorkflow(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;

  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };
  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== 'exec-workflow') return;
    const snap = entry.state as SerializedExecuteWorkflowState;
    // Entity-existence guard on the workflow's primary entity (resolved
    // from the live workflows list after rehydration).
    const targetWf = workflows.find(w => w.id === snap.workflowId);
    if (targetWf && entities.length > 0 && !entities.some(e => e.logicalName === targetWf.primaryEntity)) {
      window.alert(
        `Can't load "${entry.name}": workflow's primary entity \`${targetWf.primaryEntity}\` ` +
        `isn't available in this environment.`,
      );
      return;
    }
    setState(deserializeExecuteWorkflow(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setEntityRecordName('');
    setActivePath('pick');
  };

  usePublishSaveContext(useMemo(() => {
    if (!state.workflowId) return null;
    return {
      state: currentSerialized,
      modeId: 'exec-workflow' as const,
      dirty: isDirty,
      lastSavedId,
      onSaved,
      onLoadSaved,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentSerialized, isDirty, lastSavedId, state.workflowId, workflows]));

  const sections = [
    {
      // The Workflow itself IS the target; the record context is "Bound to".
      id: 'workflow', label: 'Target',
      meta: wf?.name ?? 'pick one',
      items: [{
        id: 'pick',
        icon: Flowchart20Regular, iconFilled: Flowchart20Filled,
        label: wf?.name ?? 'Pick on-demand workflow',
        // Method already lives on the URL bar pill; no sidebar badge needed.
        dirty: state.dirty.has('pick'),
      }],
    },
    {
      // The record context is labeled "Bound to" (not "Target") — the workflow
      // itself is the target.
      id: 'target', label: 'Bound to',
      meta: wf ? `${tbl?.displayName ?? wf.primaryEntity} record` : '—',
      items: [{
        id: 'target',
        icon: Table20Regular, iconFilled: Table20Filled,
        label: state.entityId ? `${tbl?.displayName ?? ''} (selected)` : 'Pick a record',
        badge: state.entityId ? '✓' : null,
        badgeAppearance: 'ghost' as const,
        dirty: state.dirty.has('target'),
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
  const root = activePath as RootClauseId;
  switch (root) {
    case 'pick':
      pane = (
        <WorkflowPickerPane
          workflows={workflows}
          loading={workflowsLoading}
          error={workflowsError}
          workflowId={state.workflowId}
          setWorkflowId={(id) => {
            setState(s => ({ ...s, workflowId: id, entityId: null }));
            setEntityRecordName('');
            markDirty('pick'); markDirty('target');
          }}
        />
      );
      break;
    case 'target':
      // BUG FIX: the previous guard required `wf && tbl` where `tbl` comes
      // from the synchronous in-memory metadata registry (findTable). On
      // first-load the workflow's primaryEntity hasn't been warmed yet, so
      // tbl is undefined and the picker silently swaps to "Pick a workflow
      // first" — even though a workflow IS selected. New behavior: only
      // require `wf` (the user's actual selection); use the live entity's
      // primaryEntity directly with the RecordPicker (which warms metadata
      // on its own), and fall back to the logical name for display when
      // findTable hasn't returned yet.
      pane = wf ? (
        <div>
          <PaneHead
            icon={Table20Filled}
            title="Target record"
            sub={`The ${tbl?.displayName ?? wf.primaryEntity} record the workflow operates on. Slots into the request body as "EntityId".`}
            group="execute"
          />
          <div style={{ maxWidth: 480 }}>
            <Field label={`${tbl?.displayName ?? wf.primaryEntity} record`}>
              <RecordPicker
                table={wf.primaryEntity}
                selectedId={state.entityId}
                onPick={r => {
                  set('entityId', r?.id ?? null, 'target');
                  setEntityRecordName(r?.primary ?? '');
                }}
                placeholder={`Search ${tbl?.displayName ?? wf.primaryEntity} records…`}
              />
            </Field>
          </div>
        </div>
      ) : (
        <MessageBar layout="multiline" intent="info"><MessageBarBody>Pick a workflow first.</MessageBarBody></MessageBar>
      );
      break;
    case 'headers':
      pane = (
        <HeadersEditor
          items={state.headers}
          setItems={h => set('headers', h, 'headers')}
          group="execute"
        />
      );
      break;
  }

  const headersMap = headerItemsToObject(state.headers, null);
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
          onSelect={(id) => setActivePath(id)}
          recents={recents}
        />
      }
      urlBar={
        <UrlBar
          method="POST"
          url={built.relativeUrl}
          executeVerb={type.executeVerb}
          executeIcon={Flowchart20Filled}
          disabledReason={disabledReason}
          loading={loading}
          onExecute={onExecute}
        />
      }
    >
      <MainTabs tab={tab} onTabChange={setTab} resultCount={result?.ok ? 1 : null}>
        {tab === 'builder' && pane}
        {tab === 'code'    && <CodeView themeMode={themeMode} inputs={codeInputs} />}
        {tab === 'results' && (
          <ResultsView
            result={result}
            mode="single"
            table={wf?.primaryEntity ?? ''}
            writeContext={{
              operation: 'workflow',
              table: wf?.primaryEntity ?? '',
              recordId: state.entityId,
              recordName: entityRecordName || null,
              operationName: wf?.name ?? undefined,
            }}
          />
        )}
      </MainTabs>
    </ModeShell>
  );
}

function WorkflowPickerPane({
  workflows, loading, error, workflowId, setWorkflowId,
}: {
  workflows: WorkflowRow[];
  loading: boolean;
  error: string | null;
  workflowId: string | null;
  setWorkflowId: (id: string | null) => void;
}) {
  const wf = workflowId ? workflows.find(w => w.id === workflowId) : undefined;

  // Typed-input filter for the Combobox (freeform + clearable). The Combobox's
  // own input value tracks both the user's typed filter AND the picked
  // workflow's name — we sync on selection so the input reflects the
  // current state cleanly.
  const [search, setSearch] = useState('');
  // Sync the input text to the selected workflow's name (and reset on
  // table change). User typing overrides this via onChange.
  useEffect(() => {
    setSearch(wf?.name ?? '');
  }, [wf?.name]);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q || q === (wf?.name ?? '').toLowerCase()) return workflows;
    return workflows.filter(w =>
      w.name.toLowerCase().includes(q) ||
      (w.uniqueName?.toLowerCase().includes(q) ?? false) ||
      w.primaryEntity.toLowerCase().includes(q) ||
      (w.description?.toLowerCase().includes(q) ?? false));
  }, [workflows, search, wf]);
  return (
    <div>
      <PaneHead
        icon={Flowchart20Filled}
        title="On-demand workflow"
        sub="Pick an activated workflow (type=Definition, category=Workflow). The workflow's primary entity drives the Target picker."
        group="execute"
      >
        {loading
          ? <Badge appearance="ghost" icon={<Spinner size="extra-tiny" />}>loading…</Badge>
          : <Badge appearance="ghost">{workflows.length} workflow{workflows.length === 1 ? '' : 's'}</Badge>}
      </PaneHead>
      {error && (
        <MessageBar intent="error" layout="multiline" style={{ marginBottom: 12, maxWidth: 760 }}>
          <MessageBarBody>Couldn't load workflows: {error}</MessageBarBody>
        </MessageBar>
      )}
      <div style={{ maxWidth: 760 }}>
        <Field label="Workflow">
          <Combobox
            freeform
            clearable
            value={search}
            selectedOptions={workflowId ? [workflowId] : []}
            onChange={(e) => {
              const next = (e.target as HTMLInputElement).value;
              setSearch(next);
              // Clearing the input also clears the selection (matches the
              // native clearable X behavior).
              if (next === '' && workflowId) setWorkflowId(null);
            }}
            onOptionSelect={(_, d) => {
              const next = d.optionValue ?? null;
              setWorkflowId(next);
              if (next) {
                const picked = workflows.find(w => w.id === next);
                setSearch(picked?.name ?? '');
              } else {
                setSearch('');
              }
            }}
            placeholder={
              loading ? 'Loading workflows…' :
              workflows.length === 0 ? 'No activated workflows in this environment' :
              'Search workflows…'
            }
            disabled={loading || workflows.length === 0}
            listbox={{ style: { maxHeight: 420 } }}
          >
            {filtered.length === 0 && search.trim() && (
              <Option value="__none" text="" disabled>
                <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                  No workflows match "{search.trim()}"
                </Caption1>
              </Option>
            )}
            {filtered.map(w => (
              <Option key={w.id} value={w.id} text={w.name}>
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                  <span>{w.name}</span>
                  <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                    Primary entity: <span style={{ fontFamily: tokens.fontFamilyMonospace }}>{w.primaryEntity}</span>
                    {w.description ? ` · ${w.description}` : ''}
                  </Caption1>
                </div>
              </Option>
            ))}
          </Combobox>
        </Field>

        {wf && (
          <div style={{
            marginTop: 16, padding: 12,
            border: `1px solid ${tokens.colorNeutralStroke2}`,
            borderRadius: tokens.borderRadiusMedium,
            background: tokens.colorNeutralBackground1,
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
              <Flowchart20Filled style={{ color: tokens.colorBrandForeground1 }} />
              <strong style={{ fontSize: 13 }}>{wf.name}</strong>
              <Badge appearance="ghost">{wf.primaryEntity}</Badge>
              <span style={{ flexGrow: 1 }} />
              <Badge appearance="filled" color="success" style={{ fontWeight: 700 }}>POST</Badge>
            </div>
            <Caption1 style={{ color: tokens.colorNeutralForeground2, display: 'block', marginBottom: 8 }}>
              {wf.description}
            </Caption1>
            {/* Use the shared KvGrid primitive instead of a hand-rolled
                120/1fr grid — consistent typography across modes and a
                single source of truth for keyWidth. */}
            <KvGrid
              keyWidth={120}
              rows={[
                { k: 'Workflow id', v: wf.id, mono: true },
                { k: 'Primary entity', v: wf.primaryEntity, mono: true },
                { k: 'URL shape', v: `POST /workflows(${wf.id})/Microsoft.Dynamics.CRM.ExecuteWorkflow`, mono: true },
                { k: 'Body shape', v: '{ "EntityId": "<target-guid>" }', mono: true },
                { k: 'Response', v: '204 No Content — workflow runs async' },
              ]}
            />
            <MessageBar layout="multiline" intent="info" style={{ marginTop: 12 }}>
              <MessageBarBody>
                <MessageBarTitle>Workflow runs are asynchronous.</MessageBarTitle>
                The 204 only confirms the request was accepted. Check the <code>workflowlogs</code> / <code>asyncoperations</code> tables for completion + result.
              </MessageBarBody>
            </MessageBar>
          </div>
        )}
      </div>
    </div>
  );
}
