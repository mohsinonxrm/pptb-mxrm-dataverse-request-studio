// Bulk Delete — submit a BulkDelete system job (#10).
//
// Sidebar:
//   • Target        — the table
//   • Query         — the Retrieve Multiple $filter builder; mapped to a
//                     QueryExpression for the job (limits in engine/bulkDelete)
//   • Job           — name, start time, recurrence; email notification
//   • Headers       — HTTP headers (Code tab)
//
// Send never deletes straight away: it opens an in-app confirmation that
// counts the matching rows first ($count on the same filter) and requires
// the table's logical name to be typed. The job's JobId is then polled on
// asyncoperations / bulkdeleteoperations in the Results tab.
//
// https://learn.microsoft.com/en-us/power-apps/developer/data-platform/delete-data-bulk

import { useEffect, useMemo, useState } from 'react';
import {
  Button,
  Caption1,
  Checkbox,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Field,
  Input,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Radio,
  RadioGroup,
  Spinner,
  SpinButton,
  Switch,
  Text,
  Textarea,
  tokens,
} from '@fluentui/react-components';
import {
  Table20Regular,
  Table20Filled,
  Filter20Regular,
  Filter20Filled,
  CalendarClock20Regular,
  Clock20Filled,
  Mail20Regular,
  Mail20Filled,
  LineHorizontal320Regular,
  LineHorizontal320Filled,
  Person20Regular,
  ArrowSync20Regular,
  Copy20Regular,
  Open20Regular,
  DeleteLines20Regular,
} from '@fluentui/react-icons';
import { Sidebar } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { PaneHead } from '../editors/PaneHead';
import { TargetEditor } from '../editors/TargetEditor';
import { FilterEditor } from '../editors/filter/FilterEditor';
import { countRules, emptyTree } from '../editors/filter/filterTree';
import { HeadersEditor, defaultWriteHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { StatusPill } from '../primitives/StatusPill';
import { findTable, type TableMeta } from '../mock/metadata';
import { ENV } from '../mock/environment';
import { findRequestType } from '../registry/requestTypes';
import type { ExecResult } from '../engine/runtime';
import type { RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useLiveTable } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import { useHostSession } from '../host/HostContext';
import { EnvironmentBadge } from '../shell/confirmWrite';
import { hashState, type SavedRequest } from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';
import {
  ASYNC_STATE,
  ASYNC_STATUS,
  COUNT_CAP,
  buildBulkDelete,
  buildBulkDeleteBody,
  defaultJobName,
  deserializeBulkDelete,
  recurrencePattern,
  serializeBulkDelete,
  type BulkDeleteState,
  type SerializedBulkDeleteState,
} from '../engine/bulkDelete';
import {
  countMatching,
  currentUserId,
  executeBulkDelete,
  fetchJobStatus,
  jobIdFromResult,
  type JobStatus,
  type MatchCount,
} from '../engine/bulkDeleteExecutor';
import { copyText } from '../shell/copy';

const MODE_ID = 'bulk-delete' as const;
const GROUP = 'bulk' as const;
const POLL_MS = 5000;
/** Stop auto-refresh after ~10 minutes; Refresh still works. */
const POLL_LIMIT = 120;

const defaults = (): Omit<BulkDeleteState, 'dirty'> => ({
  table: '',
  filter: emptyTree(),
  jobName: '',
  sendEmail: false,
  toRecipients: '',
  ccRecipients: '',
  startMode: 'now',
  startAt: '',
  recurrenceDays: null,
  headers: defaultWriteHeaders(),
});

export function BulkDeleteMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType(MODE_ID);
  const [state, setState] = useState<BulkDeleteState>(() => ({ ...defaults(), dirty: new Set() }));
  const { table: liveTable } = useLiveTable(state.table || null);
  // Every table, not just the pickers' scope: a saved request's table only
  // has to exist (#44).
  const { allEntities: entities } = useScopedEntities();
  const [activePath, setActivePath] = useState('target');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [jobId, setJobId] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [confirmOpen, setConfirmOpen] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);
  // Re-render once a minute so a "now" StartDateTime in the preview stays current.
  const [now, setNow] = useState(() => new Date());
  useEffect(() => {
    const t = window.setInterval(() => setNow(new Date()), 60_000);
    return () => window.clearInterval(t);
  }, []);

  const tbl = findTable(state.table);
  const built = useMemo(() => buildBulkDelete(state), [state, liveTable]);
  const { body, problems, deletesAll } = useMemo(
    () => buildBulkDeleteBody(state, now),
    [state, liveTable, now],
  );
  const ruleCount = useMemo(() => countRules(state.filter), [state.filter]);

  const set = <K extends keyof BulkDeleteState>(k: K, v: BulkDeleteState[K], dirtyId: string) =>
    setState((s) => {
      const d = new Set(s.dirty);
      d.add(dirtyId);
      return { ...s, [k]: v, dirty: d };
    });

  const disabledReason = !state.table
    ? 'Pick a target table first.'
    : !tbl
      ? 'Loading table metadata…'
      : problems.length > 0
        ? problems[0]
        : state.headers.some((h) => h.enabled && !h.name)
          ? 'Fix empty header name.'
          : null;

  // ── Send → confirmation → BulkDelete ──
  const onConfirmed = async () => {
    setConfirmOpen(false);
    setLoading(true);
    const res = await executeBulkDelete(state, headerItemsToObject(state.headers, null));
    setResult(res);
    setJobId(res.ok ? jobIdFromResult(res.body) : null);
    setLoading(false);
    setTab('results');
    setRecents((rs) =>
      [
        {
          id: `r-${Date.now()}`,
          modeId: MODE_ID,
          url: built.relativeUrl,
          method: 'POST',
          ts: Date.now(),
          status: res.status,
          ms: res.ms,
        },
        ...rs,
      ].slice(0, 8),
    );
    setState((s) => ({ ...s, dirty: new Set() }));
  };

  // ── Save / Load ──
  const currentSerialized = useMemo(() => serializeBulkDelete(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;
  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };
  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== MODE_ID) return;
    const snap = entry.state as SerializedBulkDeleteState;
    if (entities.length > 0 && !entities.some((e) => e.logicalName === snap.table)) {
      setLoadError(
        `Can't load "${entry.name}": table '${snap.table}' isn't available in this environment.`,
      );
      return;
    }
    setLoadError(null);
    setState(deserializeBulkDelete(snap, defaults()));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setJobId(null);
    setActivePath('filter');
  };
  usePublishSaveContext(
    useMemo(() => {
      if (!state.table) return null;
      return {
        state: currentSerialized,
        modeId: MODE_ID,
        dirty: isDirty,
        lastSavedId,
        onSaved,
        onLoadSaved,
      };
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [currentSerialized, isDirty, lastSavedId, state.table]),
  );

  // ── Sidebar ──
  const sections = [
    {
      id: 'target',
      label: 'Target',
      meta: tbl ? `${tbl.displayName} · ${tbl.entitySetName}` : 'Pick a table',
      items: [
        {
          id: 'target',
          icon: Table20Regular,
          iconFilled: Table20Filled,
          label: tbl?.displayName ?? 'Pick a table',
          dirty: state.dirty.has('target'),
        },
      ],
    },
    {
      id: 'query',
      label: 'Query',
      meta: ruleCount ? `${ruleCount} condition${ruleCount === 1 ? '' : 's'}` : 'all rows',
      items: [
        {
          id: 'filter',
          icon: Filter20Regular,
          iconFilled: Filter20Filled,
          label: 'Criteria',
          code: true,
          badge: deletesAll ? 'ALL' : ruleCount || null,
          badgeAppearance: 'tint' as const,
          badgeColor: deletesAll ? ('danger' as const) : ('brand' as const),
          dirty: state.dirty.has('filter'),
        },
      ],
    },
    {
      id: 'job',
      label: 'Job',
      meta: state.recurrenceDays ? `every ${state.recurrenceDays} d` : 'run once',
      items: [
        {
          id: 'job',
          icon: CalendarClock20Regular,
          iconFilled: Clock20Filled,
          label: 'Name & schedule',
          badge: state.startMode === 'at' ? 'scheduled' : null,
          dirty: state.dirty.has('job'),
        },
        {
          id: 'notify',
          icon: Mail20Regular,
          iconFilled: Mail20Filled,
          label: 'Email notification',
          badge: state.sendEmail ? 'on' : null,
          dirty: state.dirty.has('notify'),
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

  // ── Builder panes ──
  let pane: React.ReactNode = null;
  switch (activePath) {
    case 'target':
      pane = (
        <TargetEditor
          table={state.table}
          onTableChange={(t) => {
            setState((s) => ({
              ...s,
              table: t,
              filter: emptyTree(),
              // Keep a name the user typed; otherwise suggest one for the new table.
              jobName:
                s.jobName && !s.jobName.startsWith('DRS bulk delete')
                  ? s.jobName
                  : t
                    ? defaultJobName(t, new Date())
                    : '',
              dirty: new Set(['target', 'filter']),
            }));
            setResult(null);
            setJobId(null);
          }}
          group={GROUP}
          sub="Rows of this table that match the query are deleted by an asynchronous system job. Switching the target clears the query."
        />
      );
      break;
    case 'filter':
      pane = (
        <div>
          <QueryNotes deletesAll={deletesAll} problems={problems} tbl={tbl} />
          <FilterEditor
            table={state.table}
            tree={state.filter}
            setTree={(t) => set('filter', t, 'filter')}
            group={GROUP}
          />
        </div>
      );
      break;
    case 'job':
      pane = (
        <JobPane
          state={state}
          now={now}
          onChange={(patch) =>
            setState((s) => ({ ...s, ...patch, dirty: new Set([...s.dirty, 'job']) }))
          }
        />
      );
      break;
    case 'notify':
      pane = (
        <NotifyPane
          state={state}
          onChange={(patch) =>
            setState((s) => ({ ...s, ...patch, dirty: new Set([...s.dirty, 'notify']) }))
          }
        />
      );
      break;
    case 'headers':
      pane = (
        <HeadersEditor
          policy="impersonation"
          items={state.headers}
          setItems={(h) => set('headers', h, 'headers')}
          group={GROUP}
        />
      );
      break;
  }

  const codeInputs = {
    method: 'POST',
    built,
    headers: headerItemsToObject(state.headers, null),
    body: (body ?? {}) as Record<string, unknown>,
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
          disabledReason={disabledReason}
          loading={loading}
          // Send only opens the confirmation; the job is submitted from there.
          onExecute={async () => setConfirmOpen(true)}
        />
      }
    >
      <MainTabs tab={tab} onTabChange={setTab} resultCount={null}>
        {tab === 'builder' && (
          <>
            {loadError && (
              <MessageBar intent="error" style={{ marginBottom: 12 }}>
                <MessageBarBody>{loadError}</MessageBarBody>
              </MessageBar>
            )}
            {pane}
          </>
        )}
        {tab === 'code' && <CodeView themeMode={themeMode} inputs={codeInputs} group={GROUP} />}
        {tab === 'results' &&
          (result?.ok && jobId ? (
            <JobStatusPanel jobId={jobId} result={result} />
          ) : (
            <ResultsView result={result} mode="single" table={state.table} />
          ))}
      </MainTabs>
      {confirmOpen && tbl && (
        <ConfirmDialog
          state={state}
          tbl={tbl}
          deletesAll={deletesAll}
          onCancel={() => setConfirmOpen(false)}
          onConfirm={onConfirmed}
        />
      )}
    </ModeShell>
  );
}

// ──────────────────────────────────────────────────────────────
// Query notes (mapping limits)
// ──────────────────────────────────────────────────────────────

function QueryNotes({
  deletesAll,
  problems,
  tbl,
}: {
  deletesAll: boolean;
  problems: string[];
  tbl: TableMeta | undefined;
}) {
  if (!tbl) return null;
  const mapping = problems.filter((p) => /bulk delete query|mapping/.test(p));
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 12 }}>
      {deletesAll && (
        <MessageBar intent="error" layout="multiline">
          <MessageBarBody>
            <MessageBarTitle>No conditions — every {tbl.displayName} row matches</MessageBarTitle>
            Add at least one condition unless you really mean to empty the table.
          </MessageBarBody>
        </MessageBar>
      )}
      {mapping.length > 0 && (
        <MessageBar intent="error" layout="multiline">
          <MessageBarBody>
            <MessageBarTitle>Can't convert to a QueryExpression</MessageBarTitle>
            <ul style={{ margin: '4px 0 0', paddingLeft: 18 }}>
              {mapping.map((p) => (
                <li key={p}>{p}</li>
              ))}
            </ul>
          </MessageBarBody>
        </MessageBar>
      )}
      <MessageBar intent="info" layout="multiline">
        <MessageBarBody>
          BulkDelete takes a <code>QueryExpression</code>, so this $filter is converted:
          comparisons, null checks, contains / starts / ends with (as <code>Like</code>), and the
          Dataverse query functions map 1:1. Related-table paths, any/all, column-to-column
          comparisons and negated groups or functions aren't supported. The confirmation counts
          matches with the same $filter before anything is submitted.
        </MessageBarBody>
      </MessageBar>
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Job: name, start, recurrence
// ──────────────────────────────────────────────────────────────

function JobPane({
  state,
  now,
  onChange,
}: {
  state: BulkDeleteState;
  now: Date;
  onChange: (patch: Partial<BulkDeleteState>) => void;
}) {
  return (
    <div style={{ maxWidth: 640, display: 'flex', flexDirection: 'column', gap: 14 }}>
      <PaneHead
        icon={Clock20Filled}
        title="Job"
        sub="The bulk delete runs as a system job; it only deletes rows created before it starts."
        group={GROUP}
      />
      <Field label="Job name" required>
        <Input
          value={state.jobName}
          placeholder={defaultJobName(state.table, now)}
          onChange={(_, d) => onChange({ jobName: d.value })}
        />
      </Field>
      <Field label="Start">
        <RadioGroup
          value={state.startMode}
          onChange={(_, d) => onChange({ startMode: d.value as 'now' | 'at' })}
        >
          <Radio value="now" label="Immediately (StartDateTime = when you send)" />
          <Radio value="at" label="At a specific time" />
        </RadioGroup>
      </Field>
      {state.startMode === 'at' && (
        <Field label="Start date and time (your local time)">
          <Input
            type="datetime-local"
            value={state.startAt}
            onChange={(_, d) => onChange({ startAt: d.value })}
          />
        </Field>
      )}
      <Switch
        checked={state.recurrenceDays !== null}
        onChange={(_, d) => onChange({ recurrenceDays: d.checked ? 7 : null })}
        label="Repeat the job"
      />
      {state.recurrenceDays !== null && (
        <Field
          label="Every N days"
          hint={
            <>
              RecurrencePattern <code>{recurrencePattern(state.recurrenceDays) || '—'}</code>. A
              recurring job keeps deleting new matching rows until you cancel it in System Jobs.
            </>
          }
        >
          <SpinButton
            min={1}
            max={365}
            value={state.recurrenceDays}
            onChange={(_, d) => {
              const v = d.value ?? Number(d.displayValue);
              if (Number.isFinite(v)) onChange({ recurrenceDays: Math.max(1, Math.floor(v!)) });
            }}
          />
        </Field>
      )}
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Email notification
// ──────────────────────────────────────────────────────────────

function NotifyPane({
  state,
  onChange,
}: {
  state: BulkDeleteState;
  onChange: (patch: Partial<BulkDeleteState>) => void;
}) {
  const [meError, setMeError] = useState<string | null>(null);
  const addMe = async () => {
    try {
      const id = await currentUserId();
      if (!state.toRecipients.toLowerCase().includes(id)) {
        onChange({ toRecipients: [state.toRecipients.trim(), id].filter(Boolean).join(', ') });
      }
      setMeError(null);
    } catch (e) {
      setMeError(e instanceof Error ? e.message : String(e));
    }
  };
  return (
    <div style={{ maxWidth: 640, display: 'flex', flexDirection: 'column', gap: 14 }}>
      <PaneHead
        icon={Mail20Filled}
        title="Email notification"
        sub="Email users when the job finishes. Recipients are systemuser ids with a valid email address."
        group={GROUP}
      />
      <Switch
        checked={state.sendEmail}
        onChange={(_, d) => onChange({ sendEmail: d.checked })}
        label="Send an email when the job completes"
      />
      {state.sendEmail && (
        <>
          <Field label="To (systemuser ids)" required hint="Comma, space or new-line separated.">
            <Textarea
              value={state.toRecipients}
              onChange={(_, d) => onChange({ toRecipients: d.value })}
              rows={2}
            />
          </Field>
          <div>
            <Button icon={<Person20Regular />} onClick={() => void addMe()}>
              Add me
            </Button>
            {meError && (
              <Caption1 style={{ color: tokens.colorPaletteRedForeground1, marginLeft: 8 }}>
                {meError}
              </Caption1>
            )}
          </div>
          <Field label="Cc (systemuser ids)">
            <Textarea
              value={state.ccRecipients}
              onChange={(_, d) => onChange({ ccRecipients: d.value })}
              rows={2}
            />
          </Field>
        </>
      )}
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Confirmation (in-app; no native confirm — #111)
// ──────────────────────────────────────────────────────────────

function ConfirmDialog({
  state,
  tbl,
  deletesAll,
  onCancel,
  onConfirm,
}: {
  state: BulkDeleteState;
  tbl: TableMeta;
  deletesAll: boolean;
  onCancel: () => void;
  onConfirm: () => void;
}) {
  const [count, setCount] = useState<MatchCount | null>(null);
  const [typed, setTyped] = useState('');
  const [ackAll, setAckAll] = useState(false);
  const host = useHostSession();
  const envHost = host.environmentUrl.replace(/^https?:\/\//, '').replace(/\/.*$/, '');
  useEffect(() => {
    let cancelled = false;
    void countMatching(state).then((c) => {
      if (!cancelled) setCount(c);
    });
    return () => {
      cancelled = true;
    };
    // Count once per opening; the request can't change while the dialog is up.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const nameOk = typed.trim() === tbl.logicalName;
  const zero = count?.ok && count.count === 0;
  const canDelete =
    nameOk && count !== null && (!deletesAll || ackAll) && !(zero && !state.recurrenceDays);
  const when =
    state.startMode === 'at' && state.startAt
      ? new Date(state.startAt).toLocaleString()
      : 'as soon as the system job runs';

  return (
    <Dialog open modalType="alert" onOpenChange={(_, d) => !d.open && onCancel()}>
      <DialogSurface style={{ maxWidth: 560 }}>
        <DialogBody>
          <DialogTitle>Delete {tbl.displayName} rows in bulk?</DialogTitle>
          <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            <MessageBar intent="error" layout="multiline">
              <MessageBarBody>
                <MessageBarTitle>This permanently deletes data</MessageBarTitle>
                Deleted rows can't be restored, related rows follow the table's cascade rules,
                plug-ins and flows on Delete run for every row, and a failed job doesn't roll back
                what it already deleted.
              </MessageBarBody>
            </MessageBar>
            {/* The environment the job runs in (#241). */}
            {host.connected && (
              <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 8 }}>
                <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Environment</Caption1>
                <Text weight="semibold">{host.connectionName || envHost || '(no connection)'}</Text>
                <EnvironmentBadge environment={host.environment} />
                {host.connectionName && envHost && (
                  <Caption1
                    style={{ color: tokens.colorNeutralForeground3, wordBreak: 'break-all' }}
                  >
                    {envHost}
                  </Caption1>
                )}
              </div>
            )}
            <div>
              {count === null ? (
                <Spinner size="tiny" label="Counting matching rows…" labelPosition="after" />
              ) : count.ok ? (
                <Text weight="semibold" size={400}>
                  {count.capped
                    ? `${COUNT_CAP.toLocaleString()} or more`
                    : count.count.toLocaleString()}{' '}
                  {tbl.displayName} row{count.count === 1 ? '' : 's'} match right now.
                </Text>
              ) : (
                <Caption1 style={{ color: tokens.colorPaletteRedForeground1 }}>
                  Couldn't count matching rows: {count.error}
                </Caption1>
              )}
            </div>
            {zero && !state.recurrenceDays && (
              <Caption1>Nothing matches, so there's nothing to delete.</Caption1>
            )}
            <Caption1>
              Job <strong>{state.jobName || '(unnamed)'}</strong> starts {when}
              {state.recurrenceDays
                ? ` and repeats every ${state.recurrenceDays} day${state.recurrenceDays === 1 ? '' : 's'}, deleting rows that match at each run`
                : ''}
              .
            </Caption1>
            {deletesAll && (
              <Checkbox
                checked={ackAll}
                onChange={(_, d) => setAckAll(!!d.checked)}
                label={`The query has no conditions — delete every ${tbl.displayName} row.`}
              />
            )}
            <Field
              label={
                <>
                  Type the table name <code>{tbl.logicalName}</code> to confirm
                </>
              }
            >
              <Input
                value={typed}
                onChange={(_, d) => setTyped(d.value)}
                autoComplete="off"
                aria-label="Table logical name"
              />
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onCancel}>
              Cancel
            </Button>
            <Button
              appearance="primary"
              icon={<DeleteLines20Regular />}
              disabled={!canDelete}
              onClick={onConfirm}
              style={
                canDelete
                  ? {
                      backgroundColor: tokens.colorPaletteRedBackground3,
                      color: tokens.colorNeutralForegroundOnBrand,
                    }
                  : undefined
              }
            >
              Submit bulk delete job
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

// ──────────────────────────────────────────────────────────────
// Job status — polls asyncoperations(<JobId>)
// ──────────────────────────────────────────────────────────────

function num(v: unknown): number | undefined {
  return typeof v === 'number' ? v : undefined;
}
function fmt(row: Record<string, unknown> | null, key: string, labels?: Record<number, string>) {
  if (!row) return '—';
  const f = row[`${key}@OData.Community.Display.V1.FormattedValue`];
  if (typeof f === 'string') return f;
  const v = row[key];
  if (labels && typeof v === 'number') return labels[v] ?? String(v);
  return v == null || v === '' ? '—' : String(v);
}

function JobStatusPanel({ jobId, result }: { jobId: string; result: ExecResult }) {
  const [status, setStatus] = useState<JobStatus | null>(null);
  const [polls, setPolls] = useState(0);
  const [auto, setAuto] = useState(true);
  const [copied, setCopied] = useState(false);
  const done = num(status?.job?.statecode) === 3;

  const refresh = async () => {
    const s = await fetchJobStatus(jobId);
    setStatus(s);
    setPolls((p) => p + 1);
  };
  useEffect(() => {
    void refresh();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [jobId]);
  useEffect(() => {
    if (!auto || done || polls >= POLL_LIMIT) return;
    const t = window.setTimeout(() => void refresh(), POLL_MS);
    return () => window.clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [auto, done, polls]);

  const job = status?.job ?? null;
  const op = status?.operation ?? null;
  const statusCode = num(job?.statuscode);
  const kind =
    statusCode === 30 ? 'success' : statusCode === 31 || statusCode === 32 ? 'danger' : 'info';
  const jobUrl = ENV.host
    ? `https://${ENV.host}/main.aspx?pagetype=entityrecord&etn=asyncoperation&id=${jobId}`
    : '';
  const cell = { padding: '4px 12px 4px 0', verticalAlign: 'top' as const };
  const rows: [string, React.ReactNode][] = [
    ['Job id', <code key="id">{jobId}</code>],
    ['Name', fmt(job, 'name')],
    ['State', fmt(job, 'statecode', ASYNC_STATE)],
    ['Status', fmt(job, 'statuscode', ASYNC_STATUS)],
    ['Deleted', op ? fmt(op, 'successcount') : '—'],
    ['Failed', op ? fmt(op, 'failurecount') : '—'],
    ['Created', fmt(job, 'createdon')],
    ['Started', fmt(job, 'startedon')],
    ['Completed', fmt(job, 'completedon')],
    ['Recurrence', fmt(job, 'recurrencepattern')],
    [
      'Message',
      fmt(job, 'friendlymessage') !== '—' ? fmt(job, 'friendlymessage') : fmt(job, 'message'),
    ],
  ];

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12, maxWidth: 760 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
        <StatusPill status="success" code={result.status} ms={result.ms} />
        <Text weight="semibold">Bulk delete job submitted</Text>
        {status && <StatusPill status={kind} code={fmt(job, 'statuscode', ASYNC_STATUS)} />}
      </div>
      <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
        <Button icon={<ArrowSync20Regular />} onClick={() => void refresh()}>
          Refresh
        </Button>
        <Switch
          checked={auto && !done && polls < POLL_LIMIT}
          disabled={done || polls >= POLL_LIMIT}
          onChange={(_, d) => setAuto(d.checked)}
          label={done ? 'Job completed' : `Auto-refresh every ${POLL_MS / 1000} s`}
        />
        <Button
          icon={<Copy20Regular />}
          onClick={() =>
            void copyText(jobId, 'the job ID', { toastOnSuccess: false }).then(
              (ok) => ok && setCopied(true),
            )
          }
        >
          {copied ? 'Copied' : 'Copy job id'}
        </Button>
        {jobUrl && (
          <Button
            as="a"
            href={jobUrl}
            target="_blank"
            rel="noreferrer"
            icon={<Open20Regular />}
            appearance="subtle"
          >
            Open system job
          </Button>
        )}
      </div>
      {status?.error && (
        <MessageBar intent="warning">
          <MessageBarBody>Couldn't read the job: {status.error}</MessageBarBody>
        </MessageBar>
      )}
      <table style={{ borderCollapse: 'collapse', fontSize: 13 }}>
        <tbody>
          {rows.map(([k, v]) => (
            <tr key={k}>
              <td style={{ ...cell, color: tokens.colorNeutralForeground3 }}>{k}</td>
              <td style={cell}>{v}</td>
            </tr>
          ))}
        </tbody>
      </table>
      <Caption1>
        Status comes from <code>asyncoperations({jobId})</code>; deleted / failed counts from the
        matching <code>bulkdeleteoperations</code> row once the job has run.
      </Caption1>
    </div>
  );
}
