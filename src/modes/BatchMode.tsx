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

import { useCallback, useMemo, useRef, useState } from 'react';
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
import type { RecentRun } from '../state/readState';
import { ENV } from '../mock/environment';
import { adv, disabledReasonFromAdvisories, type Advisory } from '../primitives/advisories';
import { hashState, type SavedRequest } from '../state/savedRequests';
import { lastPublishedRequest, usePublishSaveContext } from '../state/SaveContext';
import {
  API_BASE,
  compileBatch,
  deserializeBatch,
  initialBatchState,
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
  type BatchRunResult,
} from '../engine/batch/batchExecutor';
import { convertSnapshot, type RequestSnapshot } from '../engine/batch/savedToBatch';
import {
  BATCH_CODE_LABELS,
  generateBatchCode,
  type BatchCodeFormat,
} from '../engine/batch/batchCode';
import { makeBoundary } from '../engine/batch/multipart';
import { BatchItemEditor } from './batch/BatchItemEditor';
import { BatchGroupPane, BatchOverviewPane } from './batch/BatchPanes';
import { BatchResultsView, type ResultRow } from './batch/BatchResultsView';
import { BatchCodeView } from './batch/BatchCodeView';
import { AddFromSavedDialog } from './batch/AddFromSavedDialog';

const CODE_FORMATS: BatchCodeFormat[] = ['http', 'fetch', 'pptb'];

export function BatchMode({ themeMode: _themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('batch');
  const [state, setState] = useState<BatchState>(initialBatchState);
  const [activePath, setActivePath] = useState('overview');
  const [tab, setTab] = useState<MainTab>('builder');
  const [run, setRun] = useState<BatchRunResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);
  const [savedOpen, setSavedOpen] = useState(false);
  const [notice, setNotice] = useState<string | null>(null);
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);

  const orgUrl = ENV.host ? `https://${ENV.host}` : '';
  const compiled = useMemo(() => compileBatch(state.groups, orgUrl), [state.groups, orgUrl]);
  const plan = useMemo(() => planDispatch(compiled.items), [compiled]);
  const unavailable = batchUnavailableReason();

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
  const last = lastPublishedRequest();
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
          'Deletes in a batch run without the Delete mode’s confirmation. Check the ids first.',
          'overview',
        ),
      );
    }
    return out;
  }, [compiled]);
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
      setRecents((rs) =>
        [
          {
            id: `r-${Date.now()}`,
            modeId: 'batch',
            url: `${API_BASE}/$batch`,
            method: 'BATCH',
            ts: Date.now(),
            status: res.ok ? 200 : 400,
            ms: res.ms,
            rowCount: res.items.filter((i) => i.outcome === 'done').length,
          },
          ...rs,
        ].slice(0, 8),
      );
      setState((s) => ({ ...s, dirty: new Set() }));
    } finally {
      setLoading(false);
      setTab('results');
    }
  };

  const resultRows = useMemo<ResultRow[]>(() => {
    if (!run) return [];
    return run.items.map((r) => {
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
  }, [run, compiled]);

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
  const generate = (fmt: BatchCodeFormat) =>
    generateBatchCode(fmt, {
      items: compiled.items,
      plan,
      continueOnError: state.continueOnError,
      orgUrl,
      boundaries,
    });

  // ── Save / Load ──
  const currentSerialized = useMemo(() => serializeBatch(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;
  const total = compiled.items.length;
  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };
  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== 'batch') return;
    const snap = entry.state as SerializedBatchState;
    setState(deserializeBatch(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setRun(null);
    setActivePath('overview');
  };
  // Always published (like Retrieve Multiple) so the Saved library is
  // reachable from an empty batch.
  usePublishSaveContext(
    useMemo(() => {
      return {
        state: currentSerialized,
        modeId: 'batch' as const,
        dirty: isDirty,
        lastSavedId,
        onSaved,
        onLoadSaved,
      };
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [currentSerialized, isDirty, lastSavedId]),
  );

  // ── Sidebar ──
  const issuesFor = (id: string) => compiled.issues.filter((i) => i.target === id);
  const resultFor = (id: string) => run?.items.find((r) => r.itemId === id);
  const sections: SidebarSection[] = [
    {
      id: 'batch',
      label: 'Batch',
      meta: `${total} request${total === 1 ? '' : 's'}`,
      items: [
        {
          id: 'overview',
          icon: Options20Regular,
          iconFilled: Options20Filled,
          label: 'Options & add',
          badge: state.continueOnError ? 'continue' : null,
          badgeAppearance: 'tint',
          dirty: state.dirty.has('overview'),
        },
      ],
    },
    ...state.groups.map(
      (g, gi): SidebarSection => ({
        id: g.id,
        label: g.atomic ? `Change set ${gi + 1}` : `Independent ${gi + 1}`,
        meta: `${g.items.length}`,
        items: [
          {
            id: `group:${g.id}`,
            icon: g.atomic ? LockClosed20Regular : TextBulletListSquare20Regular,
            iconFilled: g.atomic ? LockClosed20Filled : TextBulletListSquare20Filled,
            label: g.atomic ? 'Atomic group' : 'Independent requests',
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
          urlPreview={`/$batch · ${total} request${total === 1 ? '' : 's'}`}
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
          executeVerb={type.executeVerb}
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
        {tab === 'code' && (
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
        {tab === 'results' && <BatchResultsView run={run} rows={resultRows} />}
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
