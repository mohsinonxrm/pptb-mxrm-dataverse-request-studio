// Execute Function — GET /<function>(p1=@p1)?@p1=value
//
// Per docs:
//   • Unbound:    GET /WhoAmI                or GET /F(p1=@p1)?@p1=...
//   • Bound:      GET /<set>(<id>)/Microsoft.Dynamics.CRM.<function>(...)
//   • Collection: GET /<set>/Microsoft.Dynamics.CRM.<function>(...)
//   • Composable: append $select / $filter for collection-returning functions
//   • Prefer parameter aliases to avoid URL-length issues + DateTimeOffset bugs.
//
// Reference:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-web-api-functions

import { useEffect, useMemo, useState } from 'react';
import {
  Flash20Regular,
  Flash20Filled,
  DocumentSettings20Regular,
  DocumentSettings20Filled,
  Table20Regular,
  Table20Filled,
  LineHorizontal320Regular,
  LineHorizontal320Filled,
  Settings20Regular,
  Settings20Filled,
} from '@fluentui/react-icons';
import {
  Switch,
  Caption1,
  tokens,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
} from '@fluentui/react-components';
import { Sidebar } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { notify } from '../shell/dialogs';
import { PaneHead } from '../editors/PaneHead';
import { ActionPicker } from '../editors/ActionPicker';
import { ActionParamForm } from '../editors/ActionParamForm';
import { FunctionUrlPreview } from '../editors/FunctionUrlPreview';
import { ResponsePropertiesCard } from '../editors/ResponsePropertiesCard';
import { TargetEditor } from '../editors/TargetEditor';
import { HeadersEditor, defaultReadHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { type CsdlAction } from '../metadata/actionsCsdl';
import { actions as actionsProvider } from '../host/csdlProvider';
import { findRequestType } from '../registry/requestTypes';
import { buildExecuteFunction } from '../engine/urlBuilder';
import { runtime, type ExecResult } from '../engine/runtime';
import type { ExecuteFunctionState } from '../state/executeState';
import type { RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useScopedEntities } from '../host/useScopedEntities';
import { useLiveTable } from '../host/useLiveMetadata';
import {
  serializeExecuteFunction,
  deserializeExecuteFunction,
  hashState,
  type SavedRequest,
  type SerializedExecuteFunctionState,
  recentRunAsSavedEntry,
} from '../state/savedRequests';
import { useModeDraft, useStashModeDraft } from '../state/modeDrafts';
import { usePublishSaveContext } from '../state/SaveContext';
import { useLatestCallback } from '../primitives/useLatestCallback';
import { invalidUrlPartReason } from '../primitives/urlParts';
import { disabledReasonFromAdvisories, firstBlocker } from '../primitives/advisories';
import { invalidFunctionParamReason } from '../engine/executeBuilders';
import { addRun, useRecentRuns } from '../state/runHistory';

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
  const draft = useModeDraft('exec-function', deserializeExecuteFunction, initialState);
  const [state, setState] = useState(draft.state);
  const [activePath, setActivePath] = useState<string>('pick');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  // Kept across reloads, per environment (#226).
  const [recents, setRecents] = useRecentRuns('exec-function');

  // Async live function lookup. actionsProvider caches the parsed CSDL.
  const [fn, setFn] = useState<CsdlAction | undefined>(undefined);
  useEffect(() => {
    let cancelled = false;
    if (!state.functionName) {
      setFn(undefined);
      return;
    }
    actionsProvider.find(state.functionName).then((a) => {
      if (!cancelled) setFn(a);
    });
    return () => {
      cancelled = true;
    };
  }, [state.functionName]);

  // Bound functions (entity OR collection) need the binding table's metadata
  // for the real entity set name. Load it and rebuild the URL when it (or
  // the function) resolves. The executor sends this same URL (#55).
  const bindingEntity = fn && fn.binding.kind !== 'unbound' ? fn.binding.entityType : null;
  const { table: bindingTable } = useLiveTable(bindingEntity);
  const built = useMemo(
    () => buildExecuteFunction(state),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [state, fn, bindingTable],
  );

  // Save / Load
  // Every table, not just the pickers' scope: a saved request's table only
  // has to exist (#44).
  const { allEntities: entities } = useScopedEntities();
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(draft.lastSavedId);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(draft.lastSavedHash);
  const [boundRecordName, setBoundRecordName] = useState<string>('');

  const markDirty = (id: string) =>
    setState((s) => {
      const d = new Set(s.dirty);
      d.add(id);
      return { ...s, dirty: d };
    });
  const set = <K extends keyof ExecuteFunctionState>(
    k: K,
    v: ExecuteFunctionState[K],
    dirtyId?: string,
  ) => {
    setState((s) => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  const requiresBoundRecord = fn?.binding.kind === 'entity';
  const boundEntity = fn?.binding.kind === 'entity' ? fn.binding.entityType : null;
  const missingRequired = (fn?.parameters ?? [])
    .filter((p) => p.required)
    .filter((p) => {
      const v = state.paramValues[p.name];
      return (
        v == null || (typeof v === 'string' && v === '') || (Array.isArray(v) && v.length === 0)
      );
    });

  // Ids, names and parameter values that go into the URL (#221).
  const urlPartProblem =
    invalidUrlPartReason(requiresBoundRecord ? [['Bound record id', state.boundRecordId]] : []) ??
    invalidFunctionParamReason(fn?.parameters ?? [], state.paramValues);
  // Execute blockers, as advisories with a way to the fix (#294).
  const advisories = firstBlocker([
    { when: !fn, id: 'function-pick', title: 'Pick a function.', focusNode: 'pick' },
    {
      when: requiresBoundRecord && !state.boundRecordId,
      id: 'function-bound-record',
      title: `Function '${fn?.name}' is bound to ${boundEntity} — supply a source record.`,
      focusNode: 'target',
    },
    {
      when: bindingEntity && !bindingTable,
      id: 'function-binding-metadata',
      title: `Can't build the URL yet — ${bindingEntity} table metadata isn't loaded.`,
      // The source record when one is needed; otherwise the action picker.
      focusNode: requiresBoundRecord ? 'target' : 'pick',
    },
    {
      when: missingRequired.length > 0,
      id: 'function-required',
      title: `${missingRequired.length} required param${missingRequired.length === 1 ? '' : 's'} unset.`,
      focusNode: missingRequired[0] ? `param:${missingRequired[0].name}` : 'params',
    },
    {
      when: state.headers.some((h) => h.enabled && !h.name),
      id: 'function-header-name',
      title: 'Fix empty header name.',
      focusNode: 'headers',
    },
    {
      when: urlPartProblem,
      id: 'function-url-part',
      title: urlPartProblem ?? '',
      // 'Parameter <name>: …' names the param; each has its own node.
      focusNode: urlPartProblem?.startsWith('Bound')
        ? 'target'
        : `param:${/^Parameter (\S+):/.exec(urlPartProblem ?? '')?.[1] ?? ''}`,
    },
  ]);
  const disabledReason = disabledReasonFromAdvisories(advisories);

  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.function(state, headerItemsToObject(state.headers, null));
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents(
      addRun({
        snapshot: currentSerialized,
        modeId: 'exec-function',
        url: built.relativeUrl,
        method: 'GET',
        status: res.status,
        ms: res.ms,
        rowCount: res.ok ? 1 : 0,
      }),
    );
    setState((s) => ({ ...s, dirty: new Set() }));
  };

  // ── Save / Load ──
  const currentSerialized = useMemo(() => serializeExecuteFunction(state), [state]);
  useStashModeDraft('exec-function', currentSerialized, lastSavedId, lastSavedHash);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;

  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };
  const onLoadSaved = async (entry: SavedRequest) => {
    if (entry.modeId !== 'exec-function') return;
    const snap = entry.state as SerializedExecuteFunctionState;
    // Check the SNAPSHOT's function binding, not the currently selected one
    // (#56). A CSDL load failure skips the guard; the picker surfaces it.
    const snapFn = snap.functionName
      ? await actionsProvider.find(snap.functionName).catch(() => undefined)
      : undefined;
    const snapEntity =
      snapFn && snapFn.binding.kind !== 'unbound' ? snapFn.binding.entityType : null;
    if (snapEntity && entities.length > 0 && !entities.some((e) => e.logicalName === snapEntity)) {
      notify(
        `Can't load "${entry.name}": bound entity \`${snapEntity}\` ` +
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

  // Recent runs (#2) — restore a run's snapshot through the same loader as
  // the Saved library, but keep the current save baseline: a run isn't a
  // saved entry, so Save keeps targeting whatever it targeted before.
  const onRecentClick = (run: RecentRun) => {
    const entry = recentRunAsSavedEntry(run, 'exec-function');
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
      if (!state.functionName) return null;
      return {
        state: currentSerialized,
        modeId: 'exec-function' as const,
        dirty: isDirty,
        lastSavedId,
        onSaved: onSavedLatest,
        onLoadSaved: onLoadSavedLatest,
      };
    }, [
      currentSerialized,
      isDirty,
      lastSavedId,
      state.functionName,
      onSavedLatest,
      onLoadSavedLatest,
    ]),
  );

  const sections = [
    {
      id: 'function',
      label: 'Function',
      meta: fn
        ? fn.binding.kind === 'unbound'
          ? 'unbound'
          : fn.binding.kind === 'collection'
            ? `bound to ${bindingEntity} collection`
            : `bound to ${boundEntity}`
        : 'pick one',
      items: [
        {
          id: 'pick',
          icon: Flash20Regular,
          iconFilled: Flash20Filled,
          label: fn?.displayName ?? fn?.name ?? 'Pick function',
          code: !!fn,
          // Method already lives on the URL bar pill; no sidebar badge needed.
          dirty: state.dirty.has('pick'),
        },
      ],
    },
    ...(requiresBoundRecord && boundEntity
      ? [
          {
            id: 'binding',
            label: `Bound to ${boundEntity}`,
            meta: state.boundRecordId ? '✓ record selected' : 'pick a record',
            items: [
              {
                id: 'target',
                icon: Table20Regular,
                iconFilled: Table20Filled,
                label: 'Source record',
                badge: state.boundRecordId ? '✓' : null,
                badgeAppearance: 'ghost' as const,
                dirty: state.dirty.has('target'),
              },
            ],
          },
        ]
      : []),
    ...(fn && fn.parameters.length > 0
      ? [
          {
            id: 'params',
            label: 'Parameters',
            meta: `${Object.values(state.paramValues).filter((v) => v != null && v !== '').length} of ${fn.parameters.length} set`,
            items: [
              {
                id: 'params',
                icon: DocumentSettings20Regular,
                iconFilled: DocumentSettings20Filled,
                label: 'Function parameters',
                badge: missingRequired.length > 0 ? `${missingRequired.length} req unset` : null,
                badgeAppearance: 'tint' as const,
                badgeColor: missingRequired.length > 0 ? ('danger' as const) : ('subtle' as const),
                dirty: state.dirty.has('params'),
                // Sidebar surfaces ONLY REQUIRED params; optional live in the
                // accordion in the main Parameters pane.
                children:
                  fn?.parameters
                    .filter((p) => p.required)
                    .map((p) => {
                      const v = state.paramValues[p.name];
                      const isSet = !(
                        v == null ||
                        v === '' ||
                        (Array.isArray(v) && v.length === 0)
                      );
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
              },
            ],
          },
        ]
      : []),
    {
      // Section label is "Aliasing" — choice between Aliased and Inline encoding.
      id: 'options',
      label: 'Aliasing',
      meta: state.useParamAliases ? 'param aliases' : 'inline literals',
      items: [
        {
          id: 'options',
          icon: Settings20Regular,
          iconFilled: Settings20Filled,
          label: 'Parameter encoding',
          // The user-facing state label is the encoding *style*, not the
          // raw OData prefix. Mirror the editor's prose copy
          // ("aliased" / "inline") rather than the literal "@aliases" token.
          badge: state.useParamAliases ? 'aliased' : 'inline',
          badgeAppearance: 'ghost' as const,
          dirty: state.dirty.has('options'),
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
            setState((s) => ({ ...s, functionName: name, paramValues: {}, boundRecordId: null }));
            markDirty('pick');
            markDirty('params');
            markDirty('target');
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
          sub={`Pick the ${boundEntity} record. The id slots into the URL: /${bindingTable?.entitySetName ?? `${boundEntity}s`}(<id>)/Microsoft.Dynamics.CRM.${fn?.name}.`}
        />
      ) : null;
      break;
    case 'params':
      pane = fn ? (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
          <ActionParamForm
            action={fn}
            values={state.paramValues}
            setValue={(name, v) =>
              set('paramValues', { ...state.paramValues, [name]: v }, 'params')
            }
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
            <div
              style={{
                padding: 12,
                border: `1px solid ${tokens.colorNeutralStroke2}`,
                borderRadius: tokens.borderRadiusMedium,
              }}
            >
              <Switch
                checked={state.useParamAliases}
                onChange={(_, d) => set('useParamAliases', d.checked, 'options')}
                label={
                  <span>
                    <strong>Use parameter aliases</strong>{' '}
                    <code
                      style={{
                        fontFamily: tokens.fontFamilyMonospace,
                        fontSize: 11,
                        color: tokens.colorBrandForeground2,
                      }}
                    >
                      (p1=@p1)?@p1=value
                    </code>
                    <Caption1
                      style={{
                        display: 'block',
                        color: tokens.colorNeutralForeground3,
                        marginTop: 2,
                      }}
                    >
                      Recommended by docs — avoids the 400 Bad Request thrown on long URLs and the
                      DateTimeOffset inline-encoding bug.
                    </Caption1>
                  </span>
                }
              />
            </div>
            <MessageBar layout="multiline" intent={state.useParamAliases ? 'info' : 'warning'}>
              <MessageBarBody>
                <MessageBarTitle>
                  {state.useParamAliases ? 'Param aliases are on.' : 'Inline literals are on.'}
                </MessageBarTitle>
                {state.useParamAliases ? (
                  <>
                    Each param becomes <code>p=@pN</code> in the URL path, with values in the query
                    string. Strings stay single-quoted; numbers/booleans/GUIDs/date-times go bare;
                    collections are JSON arrays; enums are{' '}
                    <code>Namespace.Type&apos;Member&apos;</code>.
                  </>
                ) : (
                  <>
                    Params are written inline like <code>(p='val',q=42)</code>. Faster to read but
                    breaks for long URLs, and Dataverse rejects inline values containing{' '}
                    <code>/ &lt; &gt; * % &amp; : \ ? +</code> — which includes every
                    DateTimeOffset. Use aliases for those.
                  </>
                )}
              </MessageBarBody>
            </MessageBar>
          </div>
        </div>
      );
      break;
    case 'headers':
      pane = (
        <HeadersEditor
          policy="impersonation"
          items={state.headers}
          setItems={(h) => set('headers', h, 'headers')}
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
          executeIcon={Flash20Filled}
          disabledReason={disabledReason}
          loading={loading}
          onExecute={onExecute}
          advisories={advisories}
          onAdvisoryFocus={setActivePath}
        />
      }
    >
      <MainTabs tab={tab} onTabChange={setTab} resultCount={result?.ok ? 1 : null}>
        {tab === 'builder' && pane}
        {tab === 'code' && <CodeView themeMode={themeMode} inputs={codeInputs} />}
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
