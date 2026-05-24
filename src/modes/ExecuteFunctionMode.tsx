// Execute Function — GET /<function>(p1=@p1)?@p1=value
//
// Per docs:
//   • Unbound:    GET /WhoAmI                or GET /F(p1=@p1)?@p1=...
//   • Bound:      GET /<set>(<id>)/Microsoft.Dynamics.CRM.<function>(...)
//   • Composable: append $select / $filter for collection-returning functions
//   • Prefer parameter aliases to avoid URL-length issues + DateTimeOffset bugs.
//
// Reference:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-web-api-functions

import { useEffect, useMemo, useState } from 'react';
import {
  Flash20Regular, Flash20Filled,
  DocumentSettings20Regular, DocumentSettings20Filled,
  Table20Regular, Table20Filled,
  LineHorizontal320Regular, LineHorizontal320Filled,
  Settings20Regular, Settings20Filled,
} from '@fluentui/react-icons';
import {
  Switch, Caption1, tokens, MessageBar, MessageBarBody, MessageBarTitle,
} from '@fluentui/react-components';
import { Sidebar } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { PaneHead } from '../editors/PaneHead';
import { ActionPicker } from '../editors/ActionPicker';
import { ActionParamForm } from '../editors/ActionParamForm';
import { FunctionUrlPreview } from '../editors/FunctionUrlPreview';
import { ResponsePropertiesCard } from '../editors/ResponsePropertiesCard';
import { TargetEditor } from '../editors/TargetEditor';
import { HeadersEditor, defaultReadHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { type CsdlAction } from '../mock/actionsCsdl';
import { actions as actionsProvider } from '../host/csdlProvider';
import { findRequestType } from '../registry/requestTypes';
import { buildExecuteFunction } from '../engine/urlBuilder';
import { runtime, type ExecResult } from '../engine/runtime';
import type { ExecuteFunctionState } from '../state/executeState';
import type { RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useScopedEntities } from '../host/useScopedEntities';
import {
  serializeExecuteFunction, deserializeExecuteFunction, hashState,
  type SavedRequest, type SerializedExecuteFunctionState,
} from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';

// Empty initial state — no pre-seeded function. User picks from the live
// CSDL once it loads.
const initialState = (): ExecuteFunctionState => ({
  functionName: null,
  boundRecordId: null,
  paramValues: {},
  useParamAliases: true,
  headers: defaultReadHeaders(),
  category: 'oob',
  dirty: new Set(),
});

type RootClauseId = 'pick' | 'target' | 'params' | 'options' | 'headers';

export function ExecuteFunctionMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('exec-function');
  const [state, setState] = useState(initialState);
  const [activePath, setActivePath] = useState<string>('pick');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);

  const built = useMemo(() => buildExecuteFunction(state), [state]);

  // Async live function lookup. actionsProvider caches the parsed CSDL.
  const [fn, setFn] = useState<CsdlAction | undefined>(undefined);
  useEffect(() => {
    let cancelled = false;
    if (!state.functionName) { setFn(undefined); return; }
    actionsProvider.find(state.functionName).then(a => {
      if (!cancelled) setFn(a);
    });
    return () => { cancelled = true; };
  }, [state.functionName]);

  // Save / Load
  const { entities } = useScopedEntities();
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);
  const [boundRecordName, setBoundRecordName] = useState<string>('');

  const markDirty = (id: string) => setState(s => { const d = new Set(s.dirty); d.add(id); return { ...s, dirty: d }; });
  const set = <K extends keyof ExecuteFunctionState>(k: K, v: ExecuteFunctionState[K], dirtyId?: string) => {
    setState(s => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  const requiresBoundRecord = fn?.binding.kind === 'entity';
  const boundEntity = fn?.binding.kind === 'entity' ? fn.binding.entityType : null;
  const missingRequired = (fn?.parameters ?? []).filter(p => p.required).filter(p => {
    const v = state.paramValues[p.name];
    return v == null || (typeof v === 'string' && v === '') || (Array.isArray(v) && v.length === 0);
  });

  const disabledReason =
    !fn ? 'Pick a function.' :
    requiresBoundRecord && !state.boundRecordId ? `Function '${fn.name}' is bound to ${boundEntity} — supply a source record.` :
    missingRequired.length > 0 ? `${missingRequired.length} required param${missingRequired.length === 1 ? '' : 's'} unset.` :
    state.headers.some(h => h.enabled && !h.name) ? 'Fix empty header name.' :
    null;

  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.function(state);
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents(rs => [{
      id: `r-${Date.now()}`, modeId: 'exec-function',
      url: built.relativeUrl, method: 'GET', ts: Date.now(),
      status: res.status, ms: res.ms, rowCount: res.ok ? 1 : 0,
    }, ...rs].slice(0, 8));
    setState(s => ({ ...s, dirty: new Set() }));
  };

  // ── Save / Load ──
  const currentSerialized = useMemo(() => serializeExecuteFunction(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;

  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };
  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== 'exec-function') return;
    const snap = entry.state as SerializedExecuteFunctionState;
    if (boundEntity && entities.length > 0 && !entities.some(e => e.logicalName === boundEntity)) {
      window.alert(
        `Can't load "${entry.name}": bound entity \`${boundEntity}\` ` +
        `isn't available in this environment.`,
      );
      return;
    }
    setState(deserializeExecuteFunction(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setBoundRecordName('');
    setActivePath('pick');
  };

  usePublishSaveContext(useMemo(() => {
    if (!state.functionName) return null;
    return {
      state: currentSerialized,
      modeId: 'exec-function' as const,
      dirty: isDirty,
      lastSavedId,
      onSaved,
      onLoadSaved,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentSerialized, isDirty, lastSavedId, state.functionName]));

  const sections = [
    {
      id: 'function', label: 'Function',
      meta: fn ? (fn.binding.kind === 'unbound' ? 'unbound' : `bound to ${boundEntity}`) : 'pick one',
      items: [{
        id: 'pick',
        icon: Flash20Regular, iconFilled: Flash20Filled,
        label: fn?.displayName ?? fn?.name ?? 'Pick function',
        code: !!fn,
        // Method already lives on the URL bar pill; no sidebar badge needed.
        dirty: state.dirty.has('pick'),
      }],
    },
    ...(requiresBoundRecord && boundEntity ? [{
      id: 'binding', label: `Bound to ${boundEntity}`,
      meta: state.boundRecordId ? '✓ record selected' : 'pick a record',
      items: [{
        id: 'target',
        icon: Table20Regular, iconFilled: Table20Filled,
        label: 'Source record',
        badge: state.boundRecordId ? '✓' : null,
        badgeAppearance: 'ghost' as const,
        dirty: state.dirty.has('target'),
      }],
    }] : []),
    ...(fn && fn.parameters.length > 0 ? [{
      id: 'params', label: 'Parameters',
      meta: `${Object.values(state.paramValues).filter(v => v != null && v !== '').length} of ${fn.parameters.length} set`,
      items: [{
        id: 'params',
        icon: DocumentSettings20Regular, iconFilled: DocumentSettings20Filled,
        label: 'Function parameters',
        badge: missingRequired.length > 0 ? `${missingRequired.length} req unset` : null,
        badgeAppearance: 'tint' as const,
        badgeColor: missingRequired.length > 0 ? ('danger' as const) : ('subtle' as const),
        dirty: state.dirty.has('params'),
        // Sidebar surfaces ONLY REQUIRED params; optional live in the
        // accordion in the main Parameters pane.
        children: fn?.parameters.filter(p => p.required).map(p => {
          const v = state.paramValues[p.name];
          const isSet = !(v == null || v === '' || (Array.isArray(v) && v.length === 0));
          return {
            id: `param:${p.name}`,
            icon: DocumentSettings20Regular,
            label: p.name,
            code: true,
            badge: isSet ? '✓' : 'req',
            badgeAppearance: 'ghost' as const,
            badgeColor: !isSet ? ('danger' as const) : ('subtle' as const),
          };
        }) ?? [],
      }],
    }] : []),
    {
      // Section label is "Aliasing" — choice between Aliased and Inline encoding.
      id: 'options', label: 'Aliasing',
      meta: state.useParamAliases ? 'param aliases' : 'inline literals',
      items: [{
        id: 'options',
        icon: Settings20Regular, iconFilled: Settings20Filled,
        label: 'Parameter encoding',
        // The user-facing state label is the encoding *style*, not the
        // raw OData prefix. Mirror the editor's prose copy
        // ("aliased" / "inline") rather than the literal "@aliases" token.
        badge: state.useParamAliases ? 'aliased' : 'inline',
        badgeAppearance: 'ghost' as const,
        dirty: state.dirty.has('options'),
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
  const root = (activePath.startsWith('param:') ? 'params' : activePath) as RootClauseId;
  switch (root) {
    case 'pick':
      pane = (
        <ActionPicker
          // Exclude query functions (Last7Days, Between, EqualUserId,
          // Contains, etc.) — those are $filter operators, not callable
          // standalone. csdlProvider tags them via the IsBound + PropertyName
          // + Edm.Boolean signature. Keeps the picker focused on actual
          // invokable functions (WhoAmI, RetrieveTotalRecordCount, etc.).
          filter={(a) => a.kind === 'Function' && !a.isQueryFunction}
          title="Browse OData functions"
          sub="Functions are read-only operations exposed via GET. Composable functions accept $select/$filter for column scoping."
          value={state.functionName}
          onChange={(name) => {
            setState(s => ({ ...s, functionName: name, paramValues: {}, boundRecordId: null }));
            markDirty('pick'); markDirty('params'); markDirty('target');
          }}
          group="execute"
        />
      );
      break;
    case 'target':
      pane = boundEntity ? (
        <TargetEditor
          table={boundEntity}
          onTableChange={() => undefined}
          recordId={state.boundRecordId}
          onRecordChange={(id, primary) => {
            set('boundRecordId', id, 'target');
            setBoundRecordName(primary ?? '');
          }}
          group="execute"
          sub={`Pick the ${boundEntity} record. The id slots into the URL: /${boundEntity}s(<id>)/Microsoft.Dynamics.CRM.${fn?.name}.`}
        />
      ) : null;
      break;
    case 'params':
      pane = fn ? (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
          <ActionParamForm
            action={fn}
            values={state.paramValues}
            setValue={(name, v) => set('paramValues', { ...state.paramValues, [name]: v }, 'params')}
            setValues={(next) => set('paramValues', next, 'params')}
            group="execute"
            themeMode={themeMode}
          />
          {/* URL preview with colorized alias substitutions */}
          <FunctionUrlPreview built={built} useParamAliases={state.useParamAliases} />
          {/* Response Properties card — function return type contract */}
          <ResponsePropertiesCard action={fn} />
        </div>
      ) : null;
      break;
    case 'options':
      pane = (
        <div>
          <PaneHead
            icon={Settings20Filled}
            title="Parameter encoding"
            sub="Controls how function parameters are serialized into the URL."
            group="execute"
          />
          <div style={{ maxWidth: 760, display: 'flex', flexDirection: 'column', gap: 14 }}>
            <div style={{ padding: 12, border: `1px solid ${tokens.colorNeutralStroke2}`, borderRadius: tokens.borderRadiusMedium }}>
              <Switch
                checked={state.useParamAliases}
                onChange={(_, d) => set('useParamAliases', d.checked, 'options')}
                label={
                  <span>
                    <strong>Use parameter aliases</strong>{' '}
                    <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11, color: tokens.colorBrandForeground2 }}>(p1=@p1)?@p1=value</code>
                    <Caption1 style={{ display: 'block', color: tokens.colorNeutralForeground3, marginTop: 2 }}>
                      Recommended by docs — avoids the 400 Bad Request thrown on long URLs and the DateTimeOffset inline-encoding bug.
                    </Caption1>
                  </span>
                }
              />
            </div>
            <MessageBar layout="multiline" intent={state.useParamAliases ? 'info' : 'warning'}>
              <MessageBarBody>
                <MessageBarTitle>{state.useParamAliases ? 'Param aliases are on.' : 'Inline literals are on.'}</MessageBarTitle>
                {state.useParamAliases
                  ? <>Each param becomes <code>p=@pN</code> in the URL path, with values in the query string. Strings stay single-quoted; numbers/booleans/GUIDs go bare.</>
                  : <>Params are written inline like <code>(p='val',q=42)</code>. Faster to read but breaks for long URLs and DateTimeOffset.</>}
              </MessageBarBody>
            </MessageBar>
          </div>
        </div>
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
    method: 'GET',
    built,
    headers: headersMap,
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
          method="GET"
          url={built.relativeUrl}
          executeVerb={type.executeVerb}
          executeIcon={Flash20Filled}
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
            table={boundEntity ?? ''}
            writeContext={{
              operation: 'function',
              table: boundEntity ?? '',
              recordId: state.boundRecordId,
              recordName: boundRecordName || null,
              operationName: fn?.name ?? state.functionName ?? undefined,
              boundEntity: boundEntity ?? undefined,
            }}
          />
        )}
      </MainTabs>
    </ModeShell>
  );
}
