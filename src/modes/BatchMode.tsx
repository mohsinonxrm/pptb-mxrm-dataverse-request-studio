// $batch Request (#9) — compose a Dataverse Web API `$batch`: requests
// grouped into change sets (atomic) and independent groups, with
// continue-on-error, a raw multipart preview, and per-request results.
//
// Grounded in:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/execute-batch-operations-using-web-api
//
// Sidebar:
//   • Batch — options (continue-on-error), add from Saved / last request
//   • one section per group — the group's settings + one node per request
//
// Execution goes through PPTB 1.2.7's `dataverseAPI.executeBatch` /
// `executeTransaction` (see engine/batch/batchExecutor.ts). On older hosts
// and standalone the batch can be composed and copied but not sent.
//
// `variant="execute-multiple"` (#11) is the same engine with the SDK's
// ExecuteMultiple semantics: one flat list of independent requests,
// ContinueOnError → `Prefer: odata.continue-on-error`, ReturnResponses →
// results list faults only, and an ExecuteMultipleRequest (C#) in the Code
// tab. ExecuteMultiple isn't a Web API action; `$batch` is its Web API
// equivalent, so that's what gets sent.

import { useCallback, useMemo, useRef, useState } from 'react';
import type { WriteConfirmSpec } from '../engine/writeConfirm';
import {
  Document20Regular,
  Document20Filled,
  LockClosed20Regular,
  LockClosed20Filled,
  Options20Regular,
  Options20Filled,
  TextBulletListSquare20Regular,
  TextBulletListSquare20Filled,
} from '@fluentui/react-icons';
import { Sidebar, type SidebarSection } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { findRequestType } from '../registry/requestTypes';
import type { ThemeMode } from '../theme/theme';
import { ENV } from '../env/environment';
import { adv, disabledReasonFromAdvisories, type Advisory } from '../primitives/advisories';
import { hashState, type SavedRequest } from '../state/savedRequests';
import { lastPublishedRequest, usePublishSaveContext } from '../state/SaveContext';
import { useLatestCallback } from '../primitives/useLatestCallback';
import { useModeDraft, useStashModeDraft } from '../state/modeDrafts';
import {
  API_BASE,
  compileBatch,
  deserializeBatch,
  initialBatchState,
  initialExecuteMultipleState,
  itemSummary,
  newGroup,
  newItem,
  planDispatch,
  serializeBatch,
  type BatchGroup,
  type BatchItem,
  type BatchOpKind,
  type BatchState,
  type SerializedBatchState,
} from '../engine/batch/batchModel';
import {
  batchUnavailableReason,
  executeBatchPlan,
  summarizeBatchRun,
  type BatchRunResult,
} from '../engine/batch/batchExecutor';
import { convertSnapshot, type RequestSnapshot } from '../engine/batch/savedToBatch';
import {
  BATCH_CODE_LABELS,
  generateBatchCode,
  toEncodableGroups,
  type BatchCodeFormat,
} from '../engine/batch/batchCode';
import { encodeBatch, makeBoundary } from '../engine/batch/multipart';
import { generateExecuteMultipleCSharp } from '../engine/batch/executeMultipleCode';
import { useScopedEntities } from '../host/useScopedEntities';
import { findTable } from '../metadata/metadata';
import { BatchItemEditor } from './batch/BatchItemEditor';
import { BatchGroupPane, BatchOverviewPane } from './batch/BatchPanes';
import { BatchResultsView, type ResultRow } from './batch/BatchResultsView';
import { BatchCodeView } from './batch/BatchCodeView';
import { AddFromSavedDialog } from './batch/AddFromSavedDialog';
import { addRun, useRecentRuns } from '../state/runHistory';
import { useImportedRequest } from '../state/modeHandoff';
import { importBatch } from '../engine/curlImport/applyImport';

const CODE_FORMATS: BatchCodeFormat[] = ['http', 'fetch', 'python', 'pptb'];
type EmCodeFormat = 'csharp' | BatchCodeFormat;
const EM_CODE_FORMATS: EmCodeFormat[] = ['csharp', 'http', 'python', 'pptb'];
const EM_CODE_LABELS: Record<EmCodeFormat, string> = {
  ...BATCH_CODE_LABELS,
  csharp: 'ExecuteMultipleRequest (C#)',
  http: 'Web API equivalent ($batch)',
  python: 'Web API equivalent (Python)',
};

/** Job / composite modes "Add last request" doesn't offer (#198). */
const NOT_BATCHABLE_MODES = new Set(['bulk-delete', 'bulk-write', 'batch', 'execute-multiple']);

export type BatchVariant = 'batch' | 'execute-multiple';

export function BatchMode({
  themeMode: _themeMode,
  variant = 'batch',
}: {
  themeMode: ThemeMode;
  variant?: BatchVariant;
}) {
  const em = variant === 'execute-multiple';
  const modeId = em ? ('execute-multiple' as const) : ('batch' as const);
  const type = findRequestType(modeId);
  // Keep the composed batch across mode switches (#101): building a batch
  // means visiting other modes to run requests and "Add last request". Each
  // variant keeps its own draft.
  const draft = useModeDraft(
    modeId,
    deserializeBatch,
    em ? initialExecuteMultipleState : initialBatchState,
  );
  const [state, setState] = useState<BatchState>(draft.state);
  const returnResponses = state.returnResponses ?? true;
  // Every table, not just the pickers' scope: a saved request's table only
  // has to exist (#44).
  const { allEntities: entities } = useScopedEntities();
  const [activePath, setActivePath] = useState('overview');
  const [tab, setTab] = useState<MainTab>('builder');
  const [run, setRun] = useState<BatchRunResult | null>(null);
  const [loading, setLoading] = useState(false);
  // Kept across reloads, per environment (#226).
  const [recents, setRecents] = useRecentRuns(modeId);
  const [savedOpen, setSavedOpen] = useState(false);
  const [notice, setNotice] = useState<string | null>(null);
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(draft.lastSavedId);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(draft.lastSavedHash);

  const orgUrl = ENV.host ? `https://${ENV.host}` : '';
  const compiled = useMemo(() => compileBatch(state.groups, orgUrl), [state.groups, orgUrl]);
  const plan = useMemo(() => planDispatch(compiled.items), [compiled]);
  const unavailable = batchUnavailableReason();
  // A batch with writes always asks before it's sent (#241); a read-only one doesn't.
  const batchWriteConfirm = useMemo((): WriteConfirmSpec | null => {
    const total = compiled.items.length;
    const writes = compiled.items.filter((c) => c.wire && c.wire.method !== 'GET').length;
    if (writes === 0) return null;
    return {
      operation: type.name,
      scope: {
        kind: 'text',
        text: `${total} request${total === 1 ? '' : 's'}, ${writes} of them change${writes === 1 ? 's' : ''} data`,
        // More than 100 writes needs the word typed, as for other bulk writes.
        count: writes,
      },
      highImpact: true,
      confirmWord: type.id === 'execute-multiple' ? 'ExecuteMultiple' : 'batch',
    };
  }, [compiled, type.id, type.name]);

  // ── State helpers ──
  const touch = (s: BatchState, id: string): BatchState => {
    const d = new Set(s.dirty);
    d.add(id);
    return { ...s, dirty: d };
  };
  const setGroups = useCallback(
    (fn: (groups: BatchGroup[]) => BatchGroup[], dirtyId = 'batch') =>
      setState((s) => touch({ ...s, groups: fn(s.groups) }, dirtyId)),
    [],
  );
  const updateItem = (itemId: string, patch: Partial<BatchItem>) =>
    setGroups(
      (gs) =>
        gs.map((g) => ({
          ...g,
          items: g.items.map((i) => (i.id === itemId ? { ...i, ...patch } : i)),
        })),
      `item:${itemId}`,
    );
  const addItems = (groupId: string | null, items: BatchItem[], select = true) => {
    if (items.length === 0) return;
    setGroups((gs) => {
      let target = gs.find((g) => g.id === groupId);
      // A GET can't go in a change set: route reads to an independent group.
      const anyGet = items.some((i) => itemSummary(i, orgUrl).method === 'GET');
      if (!target || (target.atomic && anyGet)) {
        target = [...gs].reverse().find((g) => (anyGet ? !g.atomic : true));
      }
      if (!target) {
        const g = newGroup(!anyGet, items);
        return [...gs, g];
      }
      const id = target.id;
      return gs.map((g) => (g.id === id ? { ...g, items: [...g.items, ...items] } : g));
    });
    if (select) {
      setActivePath(`item:${items[0].id}`);
      setTab('builder');
    }
  };
  const findItem = (itemId: string) => {
    for (const g of state.groups) {
      const it = g.items.find((i) => i.id === itemId);
      if (it) return { group: g, item: it };
    }
    return null;
  };
  const moveItem = (itemId: string, delta: -1 | 1) =>
    setGroups((gs) => {
      // Flatten, swap with the neighbour, and let a move past a group edge
      // hop into the adjacent group.
      const gi = gs.findIndex((g) => g.items.some((i) => i.id === itemId));
      if (gi < 0) return gs;
      const items = gs[gi].items;
      const ii = items.findIndex((i) => i.id === itemId);
      const next = gs.map((g) => ({ ...g, items: [...g.items] }));
      const j = ii + delta;
      if (j >= 0 && j < items.length) {
        [next[gi].items[ii], next[gi].items[j]] = [next[gi].items[j], next[gi].items[ii]];
      } else if (gs[gi + delta]) {
        const [it] = next[gi].items.splice(ii, 1);
        if (delta < 0) next[gi - 1].items.push(it);
        else next[gi + 1].items.unshift(it);
      }
      return next;
    });
  const moveItemToGroup = (itemId: string, groupId: string) =>
    setGroups((gs) => {
      const hit = gs.flatMap((g) => g.items).find((i) => i.id === itemId);
      if (!hit) return gs;
      return gs.map((g) => {
        const items = g.items.filter((i) => i.id !== itemId);
        return g.id === groupId ? { ...g, items: [...items, hit] } : { ...g, items };
      });
    });

  // ── Add from Saved / last request ──
  // Job and composite modes can't go inside a $batch (#198).
  const published = lastPublishedRequest();
  const last = published && !NOT_BATCHABLE_MODES.has(published.modeId) ? published : null;
  const lastLabel = last ? findRequestType(last.modeId).name : null;
  const addSnapshots = async (snaps: RequestSnapshot[]) => {
    const added: BatchItem[] = [];
    const skipped: string[] = [];
    for (const s of snaps) {
      const r = await convertSnapshot(s);
      if (r.ok) added.push(...r.items);
      else skipped.push(`“${s.name}” (${r.reason})`);
    }
    const selected = activePath.startsWith('group:') ? activePath.slice(6) : null;
    addItems(selected, added, false);
    setNotice(
      [
        added.length ? `Added ${added.length} request${added.length === 1 ? '' : 's'}.` : '',
        skipped.length ? `Skipped ${skipped.join(', ')}.` : '',
      ]
        .filter(Boolean)
        .join(' ') || null,
    );
  };

  // ── Validation → advisories ──
  const advisories = useMemo<Advisory[]>(() => {
    const out: Advisory[] = compiled.issues.map((i, n) => {
      const focus = i.target
        ? i.target.startsWith('g')
          ? `group:${i.target}`
          : `item:${i.target}`
        : 'overview';
      const label = compiled.items.find((c) => c.item.id === i.target);
      const title = label ? `$${label.contentId}: ${i.message}` : i.message;
      return i.severity === 'error'
        ? adv.err(`batch-${n}`, 'validation', title, undefined, focus)
        : adv.warn(`batch-${n}`, 'validation', title, undefined, focus);
    });
    const deletes = compiled.items.filter((c) => c.wire?.method === 'DELETE').length;
    if (deletes > 0) {
      out.push(
        adv.warn(
          'batch-deletes',
          'validation',
          `This batch deletes ${deletes} record${deletes === 1 ? '' : 's'}`,
          'The batch’s confirmation counts writes, not the records each delete removes, and there’s no typed record name as in the Delete mode. Check the ids first.',
          'overview',
        ),
      );
    }
    // No host that can send a batch: a blocker too, so it's listed with the
    // others and Open goes to the overview that explains it (#294).
    if (unavailable) {
      out.unshift(adv.err('batch-unavailable', 'validation', unavailable, undefined, 'overview'));
    }
    return out;
  }, [compiled, unavailable]);
  // Without a host that can send a batch, the tooltip says so first, as before.
  const disabledReason = unavailable ?? disabledReasonFromAdvisories(advisories);
  const needsConsent =
    compiled.items.some((c) => c.wire && Object.keys(c.wire.headers).length > 0) ||
    (state.continueOnError && plan.some((s) => !s.transaction));

  // ── Execute ──
  const onExecute = async () => {
    setLoading(true);
    try {
      const res = await executeBatchPlan(plan, { continueOnError: state.continueOnError });
      setRun(res);
      // The batch's own status plus a per-part summary — never one part's
      // error status (#204).
      const { status, summary, succeeded } = summarizeBatchRun(res);
      setRecents(
        addRun({
          modeId,
          url: `${API_BASE}/$batch`,
          method: 'BATCH',
          status,
          ms: res.ms,
          rowCount: succeeded,
          summary,
        }),
      );
      setState((s) => ({ ...s, dirty: new Set() }));
    } finally {
      setLoading(false);
      setTab('results');
    }
  };

  const resultRows = useMemo<ResultRow[]>(() => {
    if (!run) return [];
    // ExecuteMultiple with ReturnResponses = false returns faults only.
    const shown =
      em && !returnResponses
        ? run.items.filter((r) => r.outcome !== 'done' || (r.status ?? 0) >= 400)
        : run.items;
    return shown.map((r) => {
      const c = compiled.items.find((x) => x.item.id === r.itemId);
      return {
        key: r.itemId,
        contentId: r.contentId,
        method: c?.wire?.method,
        path: c?.wire?.path,
        label: c?.item.label,
        outcome: r.outcome,
        status: r.status,
        statusText: r.statusText,
        headers: r.headers,
        body: r.body,
      };
    });
  }, [run, compiled, em, returnResponses]);

  // ── Code tab ──
  // Boundaries stay fixed for the session so the preview doesn't churn on every keystroke.
  const boundaryRef = useRef<{ batch: string; cs: string[] }>({
    batch: makeBoundary('batch'),
    cs: [],
  });
  const boundaries = useMemo(
    () => ({
      batch: boundaryRef.current.batch,
      changeset: (i: number) => (boundaryRef.current.cs[i] ??= makeBoundary('changeset')),
    }),
    [],
  );
  // Size of the multipart body for the URL bar's byte counter (#198).
  const bodyBytes = useMemo(
    () =>
      new TextEncoder().encode(encodeBatch(toEncodableGroups(compiled.items), boundaries).body)
        .length,
    [compiled, boundaries],
  );
  const generate = (fmt: BatchCodeFormat) =>
    generateBatchCode(fmt, {
      items: compiled.items,
      plan,
      continueOnError: state.continueOnError,
      orgUrl,
      boundaries,
    });
  const generateEm = (fmt: EmCodeFormat) =>
    fmt === 'csharp'
      ? generateExecuteMultipleCSharp(
          compiled.items,
          { continueOnError: state.continueOnError, returnResponses },
          {
            logicalOf: (set) => entities.find((e) => e.entitySetName === set)?.logicalName ?? null,
            attrType: (logical, attr) =>
              findTable(logical)?.columns.find((col) => col.logicalName === attr)?.attributeType ??
              null,
          },
        )
      : generate(fmt);

  // ── Save / Load ──
  const currentSerialized = useMemo(() => serializeBatch(state), [state]);
  useStashModeDraft(modeId, currentSerialized, lastSavedId, lastSavedHash);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;
  const total = compiled.items.length;
  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };
  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== modeId) return;
    const snap = entry.state as SerializedBatchState;
    setState(deserializeBatch(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setRun(null);
    setActivePath('overview');
  };
  // Always published (like Retrieve Multiple) so the Saved library is
  // reachable from an empty batch.
  // The header holds on to these; the stable wrappers always run this
  // render's handlers, not the ones from when the context was published.
  const onSavedLatest = useLatestCallback(onSaved);
  // cURL import (#271): a command pasted anywhere opens here. Like a recent
  // run, it isn't a saved entry, so the save baseline stays as it was.
  useImportedRequest(modeId, (request) => {
    if (request.modeId !== modeId) return;
    setState(importBatch(initialBatchState(), request));
    setRun(null);
    setTab('builder');
    setActivePath('overview');
  });
  const onLoadSavedLatest = useLatestCallback(onLoadSaved);
  usePublishSaveContext(
    useMemo(() => {
      return {
        state: currentSerialized,
        modeId,
        dirty: isDirty,
        lastSavedId,
        onSaved: onSavedLatest,
        onLoadSaved: onLoadSavedLatest,
      };
    }, [currentSerialized, modeId, isDirty, lastSavedId, onSavedLatest, onLoadSavedLatest]),
  );

  // ── Sidebar ──
  const issuesFor = (id: string) => compiled.issues.filter((i) => i.target === id);
  const resultFor = (id: string) => run?.items.find((r) => r.itemId === id);
  const sections: SidebarSection[] = [
    {
      id: 'batch',
      label: em ? 'ExecuteMultiple' : 'Batch',
      meta: `${total} request${total === 1 ? '' : 's'}`,
      items: [
        {
          id: 'overview',
          icon: Options20Regular,
          iconFilled: Options20Filled,
          label: em ? 'Settings & add' : 'Options & add',
          badge: state.continueOnError ? 'continue' : null,
          badgeAppearance: 'tint',
          dirty: state.dirty.has('overview'),
        },
      ],
    },
    ...state.groups.map(
      (g, gi): SidebarSection => ({
        id: g.id,
        label: em ? 'Requests' : g.atomic ? `Change set ${gi + 1}` : `Independent ${gi + 1}`,
        meta: `${g.items.length}`,
        items: [
          {
            id: `group:${g.id}`,
            icon: g.atomic ? LockClosed20Regular : TextBulletListSquare20Regular,
            iconFilled: g.atomic ? LockClosed20Filled : TextBulletListSquare20Filled,
            label: em ? 'Add request' : g.atomic ? 'Atomic group' : 'Independent requests',
            badge: issuesFor(g.id).length || null,
            badgeColor: 'danger',
          },
          ...g.items.map((it) => {
            const c = compiled.items.find((x) => x.item.id === it.id);
            const sum = itemSummary(it, orgUrl);
            const errs = issuesFor(it.id).filter((i) => i.severity === 'error').length;
            const r = resultFor(it.id);
            const badge = errs
              ? `${errs} issue${errs === 1 ? '' : 's'}`
              : r
                ? r.status !== undefined
                  ? r.status
                  : r.outcome === 'rolled-back'
                    ? 'rolled back'
                    : r.outcome === 'unknown'
                      ? 'no result'
                      : 'not run'
                : null;
            return {
              id: `item:${it.id}`,
              icon: Document20Regular,
              iconFilled: Document20Filled,
              label: `$${c?.contentId ?? '?'} ${sum.method} ${it.label || sum.path || '…'}`,
              code: true,
              badge,
              badgeAppearance: 'tint' as const,
              badgeColor:
                errs || (r && (r.outcome !== 'done' || (r.status ?? 0) >= 400))
                  ? ('danger' as const)
                  : ('success' as const),
              dirty: state.dirty.has(`item:${it.id}`),
            };
          }),
        ],
      }),
    ),
  ];

  // ── Builder pane ──
  let pane: React.ReactNode = null;
  if (activePath.startsWith('item:')) {
    const hit = findItem(activePath.slice(5));
    const c = hit && compiled.items.find((x) => x.item.id === hit.item.id);
    if (hit && c) {
      pane = (
        <BatchItemEditor
          key={hit.item.id}
          item={hit.item}
          contentId={c.contentId}
          group={hit.group}
          groups={state.groups}
          wire={c.wire}
          issues={issuesFor(hit.item.id)}
          onChange={(patch) => updateItem(hit.item.id, patch)}
          onRemove={() => {
            setGroups((gs) =>
              gs.map((g) => ({ ...g, items: g.items.filter((i) => i.id !== hit.item.id) })),
            );
            setActivePath(`group:${hit.group.id}`);
          }}
          onMove={(d) => moveItem(hit.item.id, d)}
          onMoveToGroup={(gid) => moveItemToGroup(hit.item.id, gid)}
        />
      );
    }
  } else if (activePath.startsWith('group:')) {
    const gi = state.groups.findIndex((g) => g.id === activePath.slice(6));
    const g = state.groups[gi];
    if (g) {
      pane = (
        <BatchGroupPane
          group={g}
          index={gi}
          count={state.groups.length}
          variant={variant}
          orgUrl={orgUrl}
          contentIdOf={(id) => compiled.items.find((x) => x.item.id === id)?.contentId ?? '?'}
          onSetAtomic={(atomic) =>
            setGroups((gs) => gs.map((x) => (x.id === g.id ? { ...x, atomic } : x)))
          }
          onAdd={(op: BatchOpKind) => addItems(g.id, [newItem(op)])}
          onRemoveGroup={() => {
            setGroups((gs) => gs.filter((x) => x.id !== g.id));
            setActivePath('overview');
          }}
          onMoveGroup={(d) =>
            setGroups((gs) => {
              const next = [...gs];
              const j = gi + d;
              if (j < 0 || j >= next.length) return gs;
              [next[gi], next[j]] = [next[j], next[gi]];
              return next;
            })
          }
          onSelect={setActivePath}
        />
      );
    }
  }
  if (!pane) {
    pane = (
      <BatchOverviewPane
        variant={variant}
        returnResponses={returnResponses}
        setReturnResponses={(v) => setState((s) => touch({ ...s, returnResponses: v }, 'overview'))}
        groups={state.groups}
        continueOnError={state.continueOnError}
        setContinueOnError={(v) => setState((s) => touch({ ...s, continueOnError: v }, 'overview'))}
        plan={plan}
        unavailable={unavailable}
        needsConsent={needsConsent}
        orgUrl={orgUrl}
        onAddGroup={(atomic) => {
          const g = newGroup(atomic);
          setGroups((gs) => [...gs, g]);
          setActivePath(`group:${g.id}`);
        }}
        onOpenSaved={() => setSavedOpen(true)}
        onAddLast={() => {
          if (last) void addSnapshots([{ ...last, name: lastLabel ?? last.modeId }]);
        }}
        lastLabel={lastLabel}
        onSelect={setActivePath}
        notice={notice}
      />
    );
  }

  return (
    <ModeShell
      sidebar={
        <Sidebar
          type={type}
          urlPreview={`/$batch · ${total} request${total === 1 ? '' : 's'}${em ? ' · ExecuteMultiple' : ''}`}
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
          method="BATCH"
          url={`${API_BASE}/$batch`}
          bytes={bodyBytes}
          executeVerb={type.executeVerb}
          // What this request changes, for the write confirmation (#241).
          writeConfirm={batchWriteConfirm}
          disabledReason={disabledReason}
          loading={loading}
          onExecute={onExecute}
          advisories={advisories}
          onAdvisoryFocus={(id) => {
            setActivePath(id);
            setTab('builder');
          }}
        />
      }
    >
      <MainTabs tab={tab} onTabChange={setTab} resultCount={run ? run.items.length : null}>
        {tab === 'builder' && pane}
        {tab === 'code' && em && (
          <BatchCodeView
            formats={EM_CODE_FORMATS}
            labels={EM_CODE_LABELS}
            generate={generateEm}
            note="DRS sends these requests as a Web API $batch; the C# is the SDK form."
          />
        )}
        {tab === 'code' && !em && (
          <BatchCodeView
            formats={CODE_FORMATS}
            labels={BATCH_CODE_LABELS}
            generate={generate}
            note={
              compiled.issues.some((i) => i.severity === 'error')
                ? 'Requests with errors are left out until they’re fixed.'
                : undefined
            }
          />
        )}
        {tab === 'results' && (
          <BatchResultsView
            run={run}
            rows={resultRows}
            extraTitle={
              em && !returnResponses && run
                ? 'ReturnResponses = false: only faults are listed.'
                : undefined
            }
          />
        )}
      </MainTabs>
      <AddFromSavedDialog
        open={savedOpen}
        onClose={() => setSavedOpen(false)}
        onAdd={(entries) =>
          void addSnapshots(
            entries.map((e) => ({ modeId: e.modeId, state: e.state, name: e.name })),
          )
        }
      />
    </ModeShell>
  );
}
