// Execute Action — one component, three modes:
//   • exec-action       — OOB actions (WinOpportunity, AddToQueue, GrantAccess, …)
//   • exec-customapi    — Custom APIs (modern message replacement)
//   • exec-customaction — Custom process actions (legacy)
//
// All three share the same OData wire shape (POST /<name> or POST /<set>(<id>)/
// Microsoft.Dynamics.CRM.<name>) and the same metadata-driven param form. The
// only thing that varies is which subset of the CSDL the picker lists.
//
// Reference:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-web-api-actions

import { useEffect, useMemo, useState } from 'react';
import {
  Flash20Regular,
  Flash20Filled,
  Settings20Regular,
  Settings20Filled,
  Table20Regular,
  Table20Filled,
  LineHorizontal320Regular,
  LineHorizontal320Filled,
} from '@fluentui/react-icons';
import { MessageBar, MessageBarBody, MessageBarTitle } from '@fluentui/react-components';
import { Sidebar } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { ActionPicker } from '../editors/ActionPicker';
import { ActionParamForm } from '../editors/ActionParamForm';
import { RequestBodyPreview } from '../editors/RequestBodyPreview';
import { ResponsePropertiesCard } from '../editors/ResponsePropertiesCard';
import { TargetEditor } from '../editors/TargetEditor';
import { HeadersEditor, defaultWriteHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { type CsdlAction } from '../mock/actionsCsdl';
import { actions as actionsProvider } from '../host/csdlProvider';
import { findRequestType } from '../registry/requestTypes';
import { buildExecuteAction, buildExecuteActionBody } from '../engine/urlBuilder';
import { runtime, type ExecResult } from '../engine/runtime';
import type { ExecuteActionState } from '../state/executeState';
import type { RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useScopedEntities } from '../host/useScopedEntities';
import {
  serializeExecuteAction,
  deserializeExecuteAction,
  hashState,
  type SavedRequest,
  type SerializedExecuteActionState,
} from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';

export type ActionCategory = 'oob' | 'custom-api' | 'custom-action';

interface ModeConfig {
  registryId: string;
  picker: { title: string; sub: string };
  filter: (a: CsdlAction) => boolean;
}

const CONFIGS: Record<ActionCategory, ModeConfig> = {
  oob: {
    registryId: 'exec-action',
    picker: {
      title: 'Browse OOB actions',
      sub: 'Out-of-the-box Dataverse actions exposed via the Microsoft.Dynamics.CRM namespace.',
    },
    filter: (a) => a.kind === 'Action' && a.source === 'oob',
  },
  'custom-api': {
    registryId: 'exec-customapi',
    picker: {
      title: 'Browse Custom APIs',
      sub: 'Custom API messages — modern replacement for custom process actions. Definition lives in the customapi table.',
    },
    filter: (a) => a.kind === 'Action' && a.source === 'custom-api',
  },
  'custom-action': {
    registryId: 'exec-customaction',
    picker: {
      title: 'Browse Custom Process Actions',
      sub: 'Legacy custom process actions — defined in Workflow designer with Category = Action.',
    },
    filter: (a) => a.kind === 'Action' && a.source === 'custom-action',
  },
};

// Empty initial state — no pre-seeded action. User picks from the live CSDL
// once it loads. Matches the empty-start pattern used across every mode.
const initialState = (cat: ActionCategory): ExecuteActionState => ({
  actionName: null,
  boundRecordId: null,
  paramValues: {},
  headers: defaultWriteHeaders(),
  category: cat,
  dirty: new Set(),
});

type RootClauseId = 'pick' | 'target' | 'params' | 'headers';

export interface ExecuteActionModeProps {
  themeMode: ThemeMode;
  category: ActionCategory;
}

export function ExecuteActionMode({ themeMode, category }: ExecuteActionModeProps) {
  const config = CONFIGS[category];
  const type = findRequestType(config.registryId);
  const [state, setState] = useState(() => initialState(category));
  const [activePath, setActivePath] = useState<string>('pick');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);

  const built = useMemo(() => buildExecuteAction(state), [state]);
  const body = useMemo(() => buildExecuteActionBody(state), [state]);

  // Async live action lookup. The actionsProvider caches the parsed CSDL for
  // the session, so this is one fetch on first action pick — every subsequent
  // pick is synchronous-fast. We resolve to `undefined` while loading; the
  // sidebar / URL bar handle the placeholder state.
  const [action, setAction] = useState<CsdlAction | undefined>(undefined);
  useEffect(() => {
    let cancelled = false;
    if (!state.actionName) {
      setAction(undefined);
      return;
    }
    actionsProvider.find(state.actionName).then((a) => {
      if (!cancelled) setAction(a);
    });
    return () => {
      cancelled = true;
    };
  }, [state.actionName]);

  // Save / Load tracking
  const { entities } = useScopedEntities();
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);
  // Bound-record name cache — populated when the user picks via TargetEditor.
  const [boundRecordName, setBoundRecordName] = useState<string>('');

  const markDirty = (id: string) =>
    setState((s) => {
      const d = new Set(s.dirty);
      d.add(id);
      return { ...s, dirty: d };
    });
  const set = <K extends keyof ExecuteActionState>(
    k: K,
    v: ExecuteActionState[K],
    dirtyId?: string,
  ) => {
    setState((s) => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  // Bound action requires a record id
  const requiresBoundRecord = action?.binding.kind === 'entity';
  const boundEntity = action?.binding.kind === 'entity' ? action.binding.entityType : null;

  const paramCount = action?.parameters.length ?? 0;
  const setParamCount = Object.values(state.paramValues).filter((v) => {
    if (v == null) return false;
    if (typeof v === 'string' && v === '') return false;
    if (Array.isArray(v) && v.length === 0) return false;
    return true;
  }).length;
  const missingRequired = (action?.parameters ?? [])
    .filter((p) => p.required)
    .filter((p) => {
      const v = state.paramValues[p.name];
      return (
        v == null || (typeof v === 'string' && v === '') || (Array.isArray(v) && v.length === 0)
      );
    });

  const disabledReason = !action
    ? 'Pick an action.'
    : requiresBoundRecord && !state.boundRecordId
      ? `Action '${action.name}' is bound to ${boundEntity} — supply a source record.`
      : missingRequired.length > 0
        ? `${missingRequired.length} required param${missingRequired.length === 1 ? '' : 's'} unset: ${missingRequired
            .slice(0, 3)
            .map((p) => p.name)
            .join(', ')}${missingRequired.length > 3 ? '…' : ''}`
        : state.headers.some((h) => h.enabled && !h.name)
          ? 'Fix empty header name.'
          : null;

  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.action(state);
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents((rs) =>
      [
        {
          id: `r-${Date.now()}`,
          modeId: config.registryId,
          url: built.relativeUrl,
          method: 'POST',
          ts: Date.now(),
          status: res.status,
          ms: res.ms,
          rowCount: res.ok ? 1 : 0,
        },
        ...rs,
      ].slice(0, 8),
    );
    setState((s) => ({ ...s, dirty: new Set() }));
  };

  // ── Save / Load ──
  // Execute-Action / Custom API / Custom Action all serialize under the same
  // `exec-action` saved-mode id (the inner `category` field distinguishes
  // which sub-mode they belong to). Library entries can be loaded into any
  // of the three so the loader doesn't gatekeep on category.
  const currentSerialized = useMemo(() => serializeExecuteAction(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;

  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };
  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== 'exec-action') return;
    const snap = entry.state as SerializedExecuteActionState;
    // Entity-existence guard only matters for bound actions; unbound is fine
    // even on orgs that don't have the source's entity provisioned (rare).
    if (
      boundEntity &&
      entities.length > 0 &&
      !entities.some((e) => e.logicalName === boundEntity)
    ) {
      window.alert(
        `Can't load "${entry.name}": bound entity \`${boundEntity}\` ` +
          `isn't available in this environment.`,
      );
      return;
    }
    setState(deserializeExecuteAction(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setBoundRecordName('');
    setActivePath('pick');
  };

  usePublishSaveContext(
    useMemo(() => {
      if (!state.actionName) return null;
      return {
        state: currentSerialized,
        modeId: 'exec-action' as const,
        dirty: isDirty,
        lastSavedId,
        onSaved,
        onLoadSaved,
      };
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [currentSerialized, isDirty, lastSavedId, state.actionName]),
  );

  const sections = [
    {
      id: 'action',
      label: 'Action',
      meta: action
        ? action.binding.kind === 'unbound'
          ? 'unbound'
          : `bound to ${boundEntity}`
        : 'pick one',
      items: [
        {
          id: 'pick',
          icon: Flash20Regular,
          iconFilled: Flash20Filled,
          label: action?.displayName ?? action?.name ?? `Pick ${category} action`,
          code: !!action,
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
    {
      // Execute group uses "Parameters" as the section label, not "Request body".
      id: 'request',
      label: 'Parameters',
      meta: `${setParamCount} of ${paramCount} param${paramCount === 1 ? '' : 's'}`,
      items: [
        {
          id: 'params',
          icon: Settings20Regular,
          iconFilled: Settings20Filled,
          label: 'Parameters',
          badge:
            missingRequired.length > 0
              ? `${missingRequired.length} req unset`
              : setParamCount || null,
          badgeAppearance: 'tint' as const,
          badgeColor: missingRequired.length > 0 ? ('danger' as const) : ('success' as const),
          dirty: state.dirty.has('params'),
          // Sidebar surfaces ONLY REQUIRED params as sub-items so the user's
          // attention goes to what's blocking the request. Optional params
          // live in the accordion inside the main Parameters pane.
          children:
            action?.parameters
              .filter((p) => p.required)
              .map((p) => {
                const v = state.paramValues[p.name];
                const isSet = !(v == null || v === '' || (Array.isArray(v) && v.length === 0));
                return {
                  id: `param:${p.name}`,
                  icon: Settings20Regular,
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
  // Param sub-items (param:<name>) all route to the same Parameters pane.
  let pane: React.ReactNode;
  const root = (activePath.startsWith('param:') ? 'params' : activePath) as RootClauseId;
  switch (root) {
    case 'pick':
      pane = (
        <ActionPicker
          filter={config.filter}
          title={config.picker.title}
          sub={config.picker.sub}
          value={state.actionName}
          onChange={(name) => {
            setState((s) => ({
              ...s,
              actionName: name ?? '',
              paramValues: {},
              boundRecordId: null,
            }));
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
          onTableChange={() => undefined /* fixed by action binding */}
          recordId={state.boundRecordId}
          onRecordChange={(id, primary) => {
            set('boundRecordId', id, 'target');
            setBoundRecordName(primary ?? '');
          }}
          group="execute"
          sub={`Pick the ${boundEntity} record this action operates on. The id slots into the URL: /${boundEntity}s(<id>)/Microsoft.Dynamics.CRM.${action?.name}.`}
        />
      ) : (
        <MessageBar layout="multiline" intent="info">
          <MessageBarBody>
            The selected action is unbound — no source record required.
          </MessageBarBody>
        </MessageBar>
      );
      break;
    case 'params':
      pane = action ? (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
          <ActionParamForm
            action={action}
            values={state.paramValues}
            setValue={(name, v) =>
              set('paramValues', { ...state.paramValues, [name]: v }, 'params')
            }
            setValues={(next) => set('paramValues', next, 'params')}
            group="execute"
            themeMode={themeMode}
          />
          {/* Body preview card under the form */}
          <RequestBodyPreview body={body} pillText="from $metadata" />
          {/* Response properties card for context */}
          <ResponsePropertiesCard action={action} />
        </div>
      ) : (
        <MessageBar layout="multiline" intent="info">
          <MessageBarBody>
            <MessageBarTitle>Pick an action first.</MessageBarTitle>
          </MessageBarBody>
        </MessageBar>
      );
      break;
    case 'headers':
      pane = (
        <HeadersEditor
          items={state.headers}
          setItems={(h) => set('headers', h, 'headers')}
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
          onSelect={(id) => {
            setActivePath(id);
            setTab('builder');
          }}
          recents={recents}
        />
      }
      urlBar={
        <UrlBar
          method="POST"
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
        {tab === 'code' && <CodeView themeMode={themeMode} inputs={codeInputs} />}
        {tab === 'results' && (
          <ResultsView
            result={result}
            mode="single"
            table={boundEntity ?? ''}
            writeContext={{
              operation: 'action',
              table: boundEntity ?? '',
              recordId: state.boundRecordId,
              recordName: boundRecordName || null,
              operationName: action?.name ?? state.actionName ?? undefined,
              boundEntity: boundEntity ?? undefined,
            }}
          />
        )}
      </MainTabs>
    </ModeShell>
  );
}
