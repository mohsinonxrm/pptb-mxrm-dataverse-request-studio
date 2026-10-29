// Bulk Write — CreateMultiple / UpdateMultiple / UpsertMultiple / DeleteMultiple (#8).
//
// Sidebar:
//   • Target    — the table (one table per request; every Target shares it)
//   • Operation — which bulk message, plus its availability on the table
//   • Rows      — overview (add / duplicate / remove / import CSV or JSON)
//                 and one entry per row that opens the shared field-set editor
//   • Headers   — HTTP headers + bypass options (Code tab)
//
// Wire shape and rules live in engine/bulkWrite.ts; execution (and the
// sdkmessagefilters availability check) in engine/bulkWriteExecutor.ts.
//
// https://learn.microsoft.com/en-us/power-apps/developer/data-platform/bulk-operations

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
  Text,
  Textarea,
  Tooltip,
  tokens,
} from '@fluentui/react-components';
import {
  Table20Regular,
  Table20Filled,
  Stack20Regular,
  Stack20Filled,
  Flash20Regular,
  Flash20Filled,
  FormNew20Regular,
  FormNew20Filled,
  LineHorizontal320Regular,
  LineHorizontal320Filled,
  Warning20Filled,
  Add20Regular,
  ArrowImport20Regular,
  Copy20Regular,
  Delete20Regular,
  Edit20Regular,
  ArrowLeft20Regular,
  ChevronLeft20Regular,
  ChevronRight20Regular,
} from '@fluentui/react-icons';
import { Sidebar, type SidebarClauseItem } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { PaneHead } from '../editors/PaneHead';
import { TargetEditor } from '../editors/TargetEditor';
import { FieldSetEditor, seedRequiredFieldValues } from '../editors/FieldSetEditor';
import { HeadersEditor, defaultWriteHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { BypassEditor, summarize as summarizeBypass } from '../editors/BypassEditor';
import { applyBypassToHeaders } from '../engine/bypassHeaders';
import { detectBypassAdvisories } from '../engine/bypassAdvisories';
import { detectUnsentHeaderAdvisories } from '../engine/hostHeaderAdvisories';
import { disabledReasonFromAdvisories } from '../primitives/advisories';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { StatusPill } from '../primitives/StatusPill';
import { findTable, type TableMeta } from '../mock/metadata';
import { findRequestType } from '../registry/requestTypes';
import type { ExecResult } from '../engine/runtime';
import { defaultBypassOptions, type CreateFieldValue } from '../state/writeState';
import type { RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useLiveTable } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import { copyToClipboardSafe, isEmbedded } from '../host/pptbBridge';
import { hashState, type SavedRequest } from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';
import {
  BULK_ELASTIC_SUGGESTED,
  BULK_OPERATIONS,
  BULK_WARN_THRESHOLD,
  OPERATION_INFO,
  buildBulkWrite,
  buildBulkWriteBody,
  bulkDisabledReason,
  deserializeBulkWrite,
  duplicateBulkRow,
  newBulkRow,
  parseBulkRows,
  serializeBulkWrite,
  type BulkOperation,
  type BulkRow,
  type BulkSupport,
  type BulkWriteState,
  type SerializedBulkWriteState,
} from '../engine/bulkWrite';
import { checkBulkSupport, executeBulkWrite, idsFromResult } from '../engine/bulkWriteExecutor';

const MODE_ID = 'bulk-write' as const;
const GROUP = 'bulk' as const;
const PAGE_SIZE = 100;
/** Rows listed individually in the sidebar; the Rows pane lists all of them. */
const SIDEBAR_ROW_LIMIT = 25;

const defaults = (): Omit<BulkWriteState, 'dirty'> => ({
  table: '',
  operation: 'CreateMultiple',
  rows: [],
  headers: defaultWriteHeaders(),
  bypass: defaultBypassOptions(),
});

const initialState = (): BulkWriteState => ({ ...defaults(), dirty: new Set() });

const rowPath = (id: string) => `row:${id}`;

export function BulkWriteMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType(MODE_ID);
  const [state, setState] = useState(initialState);
  const { table: liveTable } = useLiveTable(state.table || null);
  const { entities } = useScopedEntities();
  const [activePath, setActivePath] = useState<string>('target');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [resultOp, setResultOp] = useState<BulkOperation>('CreateMultiple');
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);
  const [support, setSupport] = useState<BulkSupport>({ state: 'idle' });
  const [importOpen, setImportOpen] = useState(false);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);

  const tbl = findTable(state.table);
  const built = useMemo(() => buildBulkWrite(state), [state, liveTable]);
  const { body, issues } = useMemo(() => buildBulkWriteBody(state), [state, liveTable]);
  const info = OPERATION_INFO[state.operation];

  // ── Availability (sdkmessagefilters) ──
  useEffect(() => {
    if (!state.table) {
      setSupport({ state: 'idle' });
      return;
    }
    if (!isEmbedded()) {
      setSupport({
        state: 'unknown',
        message: 'Not connected to Power Platform ToolBox — availability not checked.',
      });
      return;
    }
    let cancelled = false;
    setSupport({ state: 'checking' });
    void checkBulkSupport(state.table, state.operation).then((r) => {
      if (!cancelled) setSupport(r);
    });
    return () => {
      cancelled = true;
    };
  }, [state.table, state.operation]);

  // ── State helpers ──
  const markDirty = (s: BulkWriteState, id: string): Set<string> => {
    const d = new Set(s.dirty);
    d.add(id);
    return d;
  };
  const set = <K extends keyof BulkWriteState>(k: K, v: BulkWriteState[K], dirtyId: string) =>
    setState((s) => ({ ...s, [k]: v, dirty: markDirty(s, dirtyId) }));
  const setRows = (fn: (rows: BulkRow[]) => BulkRow[]) =>
    setState((s) => ({ ...s, rows: fn(s.rows), dirty: markDirty(s, 'rows') }));
  const patchRow = (id: string, patch: Partial<BulkRow>) =>
    setRows((rows) => rows.map((r) => (r.id === id ? { ...r, ...patch } : r)));

  const seedRow = (): BulkRow =>
    newBulkRow({
      fieldValues: tbl && state.operation === 'CreateMultiple' ? seedRequiredFieldValues(tbl) : {},
    });
  const addRow = (open: boolean) => {
    const row = seedRow();
    setRows((rows) => [...rows, row]);
    if (open) setActivePath(rowPath(row.id));
  };
  const duplicateRow = (id: string) => {
    const src = state.rows.find((r) => r.id === id);
    if (!src) return;
    const copy = duplicateBulkRow(src);
    setRows((rows) => {
      const i = rows.findIndex((r) => r.id === id);
      const next = [...rows];
      next.splice(i + 1, 0, copy);
      return next;
    });
  };
  const removeRow = (id: string) => {
    setRows((rows) => rows.filter((r) => r.id !== id));
    if (activePath === rowPath(id)) setActivePath('rows');
  };

  // ── Headers / advisories ──
  const headers = useMemo(
    () => applyBypassToHeaders(state.headers, state.bypass),
    [state.headers, state.bypass],
  );
  const advisories = useMemo(
    () => [
      ...detectBypassAdvisories(state.bypass, 'bypass'),
      ...detectUnsentHeaderAdvisories({ op: 'create', bypass: state.bypass }),
    ],
    [state.bypass],
  );
  const bypassBlocker = disabledReasonFromAdvisories(advisories);
  const disabledReason =
    bulkDisabledReason(state, issues, support) ??
    (support.state === 'checking' ? `Checking ${state.operation} availability…` : null);

  // ── Execute ──
  const onExecute = async () => {
    setLoading(true);
    const op = state.operation;
    // The same headers the Code tab shows, bypass included (#257).
    const res = await executeBulkWrite(state, headerItemsToObject(headers, null));
    setResult(res);
    setResultOp(op);
    setLoading(false);
    setTab('results');
    const ids = idsFromResult(res.body);
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
          rowCount: res.ok ? ids.length || state.rows.length : 0,
          // Recent runs shows no row count for writes (#309).
          summary: res.ok
            ? `${ids.length || state.rows.length} ${(ids.length || state.rows.length) === 1 ? 'row' : 'rows'} written`
            : undefined,
        },
        ...rs,
      ].slice(0, 8),
    );
    setState((s) => ({ ...s, dirty: new Set() }));
  };

  // ── Save / Load ──
  const currentSerialized = useMemo(() => serializeBulkWrite(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;

  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };
  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== MODE_ID) return;
    const snap = entry.state as SerializedBulkWriteState;
    if (entities.length > 0 && !entities.some((e) => e.logicalName === snap.table)) {
      setLoadError(
        `Can't load "${entry.name}": table '${snap.table}' isn't available in this environment.`,
      );
      return;
    }
    setLoadError(null);
    setState(deserializeBulkWrite(snap, defaults()));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setActivePath('rows');
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
  const issueByRow = useMemo(() => {
    const m = new Map<string, string>();
    for (const i of issues) if (!m.has(i.rowId)) m.set(i.rowId, i.message);
    return m;
  }, [issues]);
  const rowLabel = (row: BulkRow, i: number) => {
    const name = tbl ? row.fieldValues[tbl.primaryName] : undefined;
    const tail = row.key || (typeof name === 'string' && name) || '';
    return tail ? `#${i + 1} · ${tail}` : `#${i + 1}`;
  };
  const rowItems: SidebarClauseItem[] = state.rows.slice(0, SIDEBAR_ROW_LIMIT).map((row, i) => {
    const issue = issueByRow.get(row.id);
    const n = Object.keys(row.fieldValues).length + row.nullFields.length;
    return {
      id: rowPath(row.id),
      icon: FormNew20Regular,
      iconFilled: FormNew20Filled,
      label: rowLabel(row, i),
      badge: issue ? '!' : n || null,
      badgeAppearance: 'tint' as const,
      badgeColor: issue ? ('danger' as const) : ('success' as const),
    };
  });
  const supportBadge =
    support.state === 'supported'
      ? { badge: 'supported', color: 'success' as const }
      : support.state === 'unsupported'
        ? { badge: 'unsupported', color: 'danger' as const }
        : support.state === 'checking'
          ? { badge: '…', color: 'subtle' as const }
          : support.state === 'unknown'
            ? { badge: '?', color: 'warning' as const }
            : { badge: null, color: 'subtle' as const };
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
      id: 'operation',
      label: 'Operation',
      meta: state.operation,
      items: [
        {
          id: 'operation',
          icon: Flash20Regular,
          iconFilled: Flash20Filled,
          label: info.label,
          badge: supportBadge.badge,
          badgeAppearance: 'tint' as const,
          badgeColor: supportBadge.color,
          dirty: state.dirty.has('operation'),
        },
      ],
    },
    {
      id: 'rows',
      label: 'Rows',
      meta: `${state.rows.length} row${state.rows.length === 1 ? '' : 's'}`,
      items: [
        {
          id: 'rows',
          icon: Stack20Regular,
          iconFilled: Stack20Filled,
          label: 'All rows',
          // Count only — the sidebar badge column is narrow; the Rows pane spells it out.
          badge: issues.length || null,
          badgeAppearance: 'tint' as const,
          badgeColor: 'danger' as const,
          dirty: state.dirty.has('rows'),
          children: rowItems,
        },
      ],
    },
    {
      id: 'headers',
      label: 'Headers',
      meta: `${headers.filter((h) => h.enabled).length} active`,
      items: [
        {
          id: 'headers',
          icon: LineHorizontal320Regular,
          iconFilled: LineHorizontal320Filled,
          label: 'HTTP headers',
          badge: headers.filter((h) => h.enabled).length || null,
          dirty: state.dirty.has('headers'),
        },
        {
          id: 'bypass',
          icon: Warning20Filled,
          iconFilled: Warning20Filled,
          label: 'Bypass logic',
          badge: summarizeBypass(state.bypass),
          badgeAppearance:
            state.bypass.businessLogic !== 'none' || state.bypass.suppressFlows
              ? ('tint' as const)
              : ('ghost' as const),
          badgeColor:
            state.bypass.businessLogic !== 'none' || state.bypass.suppressFlows
              ? ('warning' as const)
              : ('subtle' as const),
          dirty: state.dirty.has('bypass'),
        },
      ],
    },
  ];

  // ── Builder pane ──
  let pane: React.ReactNode = null;
  if (activePath.startsWith('row:')) {
    const idx = state.rows.findIndex((r) => rowPath(r.id) === activePath);
    const row = state.rows[idx];
    pane = row ? (
      <RowEditor
        key={row.id}
        row={row}
        index={idx}
        count={state.rows.length}
        table={state.table}
        tbl={tbl}
        operation={state.operation}
        issue={issueByRow.get(row.id)}
        themeMode={themeMode}
        onPatch={(p) => patchRow(row.id, p)}
        onNavigate={(i) => setActivePath(rowPath(state.rows[i].id))}
        onBack={() => setActivePath('rows')}
        onDuplicate={() => duplicateRow(row.id)}
        onRemove={() => removeRow(row.id)}
      />
    ) : (
      <Caption1>That row was removed.</Caption1>
    );
  } else {
    switch (activePath) {
      case 'target':
        pane = (
          <TargetEditor
            table={state.table}
            onTableChange={(t) => {
              // Rows are column-bound — switching the table clears them.
              setState((s) => ({
                ...s,
                table: t,
                rows: [],
                dirty: new Set(['target', 'rows']),
              }));
              setResult(null);
            }}
            group={GROUP}
            sub="Every row in a bulk request targets the same table. Switching the target clears the rows."
          />
        );
        break;
      case 'operation':
        pane = (
          <OperationPane
            operation={state.operation}
            setOperation={(op) => set('operation', op, 'operation')}
            support={support}
            table={state.table}
          />
        );
        break;
      case 'rows':
        pane = (
          <RowsPane
            state={state}
            tbl={tbl}
            issueByRow={issueByRow}
            issueCount={issues.length}
            rowLabel={rowLabel}
            onAdd={() => addRow(false)}
            onImport={() => setImportOpen(true)}
            onClear={() => setRows(() => [])}
            onEdit={(id) => setActivePath(rowPath(id))}
            onDuplicate={duplicateRow}
            onRemove={removeRow}
            onKeyChange={(id, key) => patchRow(id, { key })}
          />
        );
        break;
      case 'headers':
        pane = (
          <HeadersEditor
            items={state.headers}
            setItems={(h) => set('headers', h, 'headers')}
            group={GROUP}
            sendsHeaders
          />
        );
        break;
      case 'bypass':
        pane = (
          <BypassEditor
            value={state.bypass}
            onChange={(b) => set('bypass', b, 'bypass')}
            group={GROUP}
          />
        );
        break;
    }
  }

  const codeInputs = {
    method: 'POST',
    built,
    headers: headerItemsToObject(headers, null),
    body: body as Record<string, unknown>,
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
          // What this request changes, for the write confirmation (#241).
          writeConfirm={{
            operation: `Bulk Write · ${state.operation}`,
            table: state.table,
            scope: { kind: 'records', count: state.rows.length },
            highImpact: true,
          }}
          disabledReason={disabledReason ?? bypassBlocker}
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
        resultCount={result?.ok ? idsFromResult(result.body).length || null : null}
      >
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
          (result && result.ok ? (
            <BulkResult result={result} operation={resultOp} />
          ) : (
            <ResultsView result={result} mode="single" table={state.table} />
          ))}
      </MainTabs>
      <ImportDialog
        open={importOpen}
        tbl={tbl}
        onClose={() => setImportOpen(false)}
        onImport={(rows, replace) => {
          setRows((cur) => (replace ? rows : [...cur, ...rows]));
          setImportOpen(false);
          setActivePath('rows');
        }}
      />
    </ModeShell>
  );
}

// ──────────────────────────────────────────────────────────────
// Operation pane
// ──────────────────────────────────────────────────────────────

function OperationPane({
  operation,
  setOperation,
  support,
  table,
}: {
  operation: BulkOperation;
  setOperation: (op: BulkOperation) => void;
  support: BulkSupport;
  table: string;
}) {
  return (
    <div style={{ maxWidth: 760 }}>
      <PaneHead
        icon={Flash20Filled}
        title="Operation"
        sub="One request, many rows of the same table. On standard tables the whole request rolls back if any row fails."
        group={GROUP}
      />
      <RadioGroup
        value={operation}
        onChange={(_, d) => setOperation(d.value as BulkOperation)}
        aria-label="Bulk operation"
      >
        {BULK_OPERATIONS.map((op) => (
          <Radio
            key={op}
            value={op}
            label={
              <span>
                <code>{op}</code>
                <br />
                <Caption1>{OPERATION_INFO[op].description}</Caption1>
              </span>
            }
          />
        ))}
      </RadioGroup>

      <div style={{ marginTop: 16, display: 'flex', flexDirection: 'column', gap: 8 }}>
        {!table ? (
          <MessageBar intent="info">
            <MessageBarBody>Pick a table to check whether it supports {operation}.</MessageBarBody>
          </MessageBar>
        ) : support.state === 'checking' ? (
          <Spinner size="tiny" label={`Checking ${operation} availability for ${table}…`} />
        ) : support.state === 'supported' ? (
          <MessageBar intent="success">
            <MessageBarBody>
              <code>{table}</code> supports {operation}.
            </MessageBarBody>
          </MessageBar>
        ) : support.state === 'unsupported' ? (
          <MessageBar intent="error" layout="multiline">
            <MessageBarBody>
              <MessageBarTitle>Not available</MessageBarTitle>
              {support.message}
            </MessageBarBody>
          </MessageBar>
        ) : support.state === 'unknown' ? (
          <MessageBar intent="warning" layout="multiline">
            <MessageBarBody>{support.message}</MessageBarBody>
          </MessageBar>
        ) : null}
        <MessageBar intent="info" layout="multiline">
          <MessageBarBody>
            Availability comes from the table's <code>sdkmessagefilters</code>.{' '}
            <code>UpsertMultiple</code> has no filter of its own — it's available where both{' '}
            <code>CreateMultiple</code> and <code>UpdateMultiple</code> are. Send 100–1,000 rows per
            request on standard tables ({BULK_ELASTIC_SUGGESTED} on elastic tables); there's no
            fixed cap, but large payloads hit size and time limits.{' '}
            <a
              href="https://learn.microsoft.com/en-us/power-apps/developer/data-platform/bulk-operations"
              target="_blank"
              rel="noreferrer"
            >
              Use bulk operation messages
            </a>
          </MessageBarBody>
        </MessageBar>
      </div>
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Rows overview
// ──────────────────────────────────────────────────────────────

function keyPlaceholder(op: BulkOperation, tbl: TableMeta | undefined): string {
  const pk = tbl?.primaryKey ?? 'primary key';
  switch (op) {
    case 'CreateMultiple':
      return `optional ${pk} GUID`;
    case 'UpsertMultiple':
      return `${pk} GUID or key='value'`;
    default:
      return `${pk} GUID`;
  }
}

function RowsPane({
  state,
  tbl,
  issueByRow,
  issueCount,
  rowLabel,
  onAdd,
  onImport,
  onClear,
  onEdit,
  onDuplicate,
  onRemove,
  onKeyChange,
}: {
  state: BulkWriteState;
  tbl: TableMeta | undefined;
  issueByRow: Map<string, string>;
  issueCount: number;
  rowLabel: (row: BulkRow, i: number) => string;
  onAdd: () => void;
  onImport: () => void;
  onClear: () => void;
  onEdit: (id: string) => void;
  onDuplicate: (id: string) => void;
  onRemove: (id: string) => void;
  onKeyChange: (id: string, key: string) => void;
}) {
  const [page, setPage] = useState(0);
  const pages = Math.max(1, Math.ceil(state.rows.length / PAGE_SIZE));
  const p = Math.min(page, pages - 1);
  const visible = state.rows.slice(p * PAGE_SIZE, (p + 1) * PAGE_SIZE);
  const info = OPERATION_INFO[state.operation];
  const cell = { padding: '4px 8px', borderBottom: `1px solid ${tokens.colorNeutralStroke2}` };

  return (
    <div>
      <PaneHead
        icon={Stack20Filled}
        title="Rows"
        sub={`Each row becomes one item in Targets. ${info.keyRequired ? 'Every row needs a key.' : 'A key is optional.'}`}
        group={GROUP}
      >
        <Button icon={<Add20Regular />} onClick={onAdd} disabled={!tbl}>
          Add row
        </Button>
        <Button icon={<ArrowImport20Regular />} onClick={onImport} disabled={!tbl}>
          Import CSV / JSON…
        </Button>
        <Button
          icon={<Delete20Regular />}
          appearance="subtle"
          onClick={onClear}
          disabled={state.rows.length === 0}
        >
          Clear
        </Button>
      </PaneHead>

      {!tbl && (
        <MessageBar intent="info">
          <MessageBarBody>Pick a target table first.</MessageBarBody>
        </MessageBar>
      )}
      {state.rows.length > BULK_WARN_THRESHOLD && (
        <MessageBar intent="warning" layout="multiline" style={{ marginBottom: 8 }}>
          <MessageBarBody>
            <MessageBarTitle>{state.rows.length.toLocaleString()} rows</MessageBarTitle>
            Microsoft recommends 100–1,000 rows per request. Larger requests are more likely to hit
            message-size or time limits, and on standard tables a failure rolls back every row.
            Consider splitting the rows into several requests.
          </MessageBarBody>
        </MessageBar>
      )}
      {issueCount > 0 && (
        <MessageBar intent="error" style={{ marginBottom: 8 }}>
          <MessageBarBody>
            {issueCount} row issue{issueCount === 1 ? '' : 's'} — fix them before sending.
          </MessageBarBody>
        </MessageBar>
      )}

      {tbl && state.rows.length === 0 && (
        <Caption1>
          No rows yet. Add rows one by one, or import them from CSV (header row of column names —
          Excel copy-paste works) or a JSON array of objects.
        </Caption1>
      )}

      {state.rows.length > 0 && (
        <table style={{ borderCollapse: 'collapse', width: '100%', fontSize: 13 }}>
          <thead>
            <tr style={{ textAlign: 'left' }}>
              <th style={cell}>#</th>
              <th style={cell}>Key</th>
              <th style={cell}>Columns</th>
              <th style={cell}>Status</th>
              <th style={cell} aria-label="Actions" />
            </tr>
          </thead>
          <tbody>
            {visible.map((row, vi) => {
              const i = p * PAGE_SIZE + vi;
              const cols = [...Object.keys(row.fieldValues), ...row.nullFields];
              const issue = issueByRow.get(row.id);
              return (
                <tr key={row.id}>
                  <td style={cell}>{i + 1}</td>
                  <td style={{ ...cell, minWidth: 220 }}>
                    <Input
                      size="small"
                      value={row.key}
                      placeholder={keyPlaceholder(state.operation, tbl)}
                      onChange={(_, d) => onKeyChange(row.id, d.value)}
                      aria-label={`Row ${i + 1} key`}
                      style={{ width: '100%' }}
                    />
                  </td>
                  <td style={cell}>
                    <Caption1 title={cols.join(', ')}>
                      {cols.length === 0
                        ? '—'
                        : `${cols.length}: ${cols.slice(0, 3).join(', ')}${cols.length > 3 ? '…' : ''}`}
                    </Caption1>
                  </td>
                  <td style={cell}>
                    {issue ? (
                      <Caption1 style={{ color: tokens.colorPaletteRedForeground1 }}>
                        {issue}
                      </Caption1>
                    ) : (
                      <Caption1 style={{ color: tokens.colorPaletteGreenForeground1 }}>
                        Ready
                      </Caption1>
                    )}
                  </td>
                  <td style={{ ...cell, whiteSpace: 'nowrap' }}>
                    <Tooltip content={`Edit ${rowLabel(row, i)}`} relationship="label">
                      <Button
                        size="small"
                        appearance="subtle"
                        icon={<Edit20Regular />}
                        onClick={() => onEdit(row.id)}
                      />
                    </Tooltip>
                    <Tooltip content="Duplicate row" relationship="label">
                      <Button
                        size="small"
                        appearance="subtle"
                        icon={<Copy20Regular />}
                        onClick={() => onDuplicate(row.id)}
                      />
                    </Tooltip>
                    <Tooltip content="Remove row" relationship="label">
                      <Button
                        size="small"
                        appearance="subtle"
                        icon={<Delete20Regular />}
                        onClick={() => onRemove(row.id)}
                      />
                    </Tooltip>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      )}
      {pages > 1 && (
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginTop: 8 }}>
          <Button
            size="small"
            icon={<ChevronLeft20Regular />}
            disabled={p === 0}
            onClick={() => setPage(p - 1)}
            aria-label="Previous page"
          />
          <Caption1>
            Rows {p * PAGE_SIZE + 1}–{Math.min((p + 1) * PAGE_SIZE, state.rows.length)} of{' '}
            {state.rows.length}
          </Caption1>
          <Button
            size="small"
            icon={<ChevronRight20Regular />}
            disabled={p >= pages - 1}
            onClick={() => setPage(p + 1)}
            aria-label="Next page"
          />
        </div>
      )}
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// One row — key + the shared field-set editor
// ──────────────────────────────────────────────────────────────

function RowEditor({
  row,
  index,
  count,
  table,
  tbl,
  operation,
  issue,
  themeMode,
  onPatch,
  onNavigate,
  onBack,
  onDuplicate,
  onRemove,
}: {
  row: BulkRow;
  index: number;
  count: number;
  table: string;
  tbl: TableMeta | undefined;
  operation: BulkOperation;
  issue?: string;
  themeMode: ThemeMode;
  onPatch: (p: Partial<BulkRow>) => void;
  onNavigate: (index: number) => void;
  onBack: () => void;
  onDuplicate: () => void;
  onRemove: () => void;
}) {
  const info = OPERATION_INFO[operation];
  return (
    <div>
      <PaneHead
        icon={FormNew20Filled}
        title={`Row ${index + 1} of ${count}`}
        sub={
          operation === 'DeleteMultiple'
            ? 'Only the key (and partitionid on partitioned elastic tables) is needed.'
            : undefined
        }
        group={GROUP}
      >
        <Button icon={<ArrowLeft20Regular />} appearance="subtle" onClick={onBack}>
          All rows
        </Button>
        <Button
          icon={<ChevronLeft20Regular />}
          disabled={index === 0}
          onClick={() => onNavigate(index - 1)}
          aria-label="Previous row"
        />
        <Button
          icon={<ChevronRight20Regular />}
          disabled={index >= count - 1}
          onClick={() => onNavigate(index + 1)}
          aria-label="Next row"
        />
        <Button icon={<Copy20Regular />} onClick={onDuplicate}>
          Duplicate
        </Button>
        <Button icon={<Delete20Regular />} appearance="subtle" onClick={onRemove}>
          Remove
        </Button>
      </PaneHead>
      <Field
        label={info.keyRequired ? 'Key (required)' : 'Key (optional)'}
        hint={
          operation === 'UpsertMultiple'
            ? `A ${tbl?.primaryKey ?? 'primary key'} GUID, or an alternate key segment such as name='Contoso' — sent as @odata.id.`
            : operation === 'CreateMultiple'
              ? `Leave empty to let Dataverse assign the id, or set a ${tbl?.primaryKey ?? 'primary key'} GUID.`
              : `The ${tbl?.primaryKey ?? 'primary key'} GUID of the row.`
        }
        validationState={issue ? 'error' : 'none'}
        validationMessage={issue}
        style={{ maxWidth: 520, marginBottom: 12 }}
      >
        <Input value={row.key} onChange={(_, d) => onPatch({ key: d.value })} />
      </Field>
      {info.usesFields && (
        <FieldSetEditor
          table={table}
          values={row.fieldValues}
          setValues={(next) => onPatch({ fieldValues: next as Record<string, CreateFieldValue> })}
          nullFields={row.nullFields}
          setNullFields={(n) => onPatch({ nullFields: n })}
          group={GROUP}
          themeMode={themeMode}
          purpose={info.fieldPurpose}
        />
      )}
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Import dialog (in-app; no native prompts — #111)
// ──────────────────────────────────────────────────────────────

function ImportDialog({
  open,
  tbl,
  onClose,
  onImport,
}: {
  open: boolean;
  tbl: TableMeta | undefined;
  onClose: () => void;
  onImport: (rows: BulkRow[], replace: boolean) => void;
}) {
  const [text, setText] = useState('');
  const [replace, setReplace] = useState(false);
  const parsed = useMemo(
    () => (tbl ? parseBulkRows(text, tbl) : { rows: [], skippedColumns: [], errors: [] }),
    [text, tbl],
  );
  const onFile = async (file: File | undefined) => {
    if (!file) return;
    setText(await file.text());
  };
  return (
    <Dialog open={open} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface style={{ maxWidth: 760 }}>
        <DialogBody>
          <DialogTitle>Import rows{tbl ? ` into ${tbl.displayName}` : ''}</DialogTitle>
          <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            <Caption1>
              Paste CSV / tab-separated text with a header row of column logical (or display) names,
              or a JSON array of objects. A <code>{tbl?.primaryKey ?? 'primary key'}</code>,{' '}
              <code>id</code> or <code>@key</code> column fills the row key. Lookups take a GUID
              (first target) or <code>table:guid</code>; choices take the value or the label.
            </Caption1>
            <Textarea
              value={text}
              onChange={(_, d) => setText(d.value)}
              rows={12}
              aria-label="Rows to import"
              placeholder={'name,numberofemployees\nContoso,10\nFabrikam,25'}
              textarea={{ style: { fontFamily: tokens.fontFamilyMonospace, fontSize: 12 } }}
            />
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
              <input
                type="file"
                accept=".csv,.tsv,.txt,.json"
                aria-label="Load rows from a file"
                onChange={(e) => void onFile(e.target.files?.[0])}
              />
              <Checkbox
                checked={replace}
                onChange={(_, d) => setReplace(!!d.checked)}
                label="Replace existing rows"
              />
            </div>
            {text.trim() && (
              <MessageBar
                intent={parsed.errors.length ? 'warning' : parsed.rows.length ? 'success' : 'info'}
                layout="multiline"
              >
                <MessageBarBody>
                  {parsed.rows.length} row{parsed.rows.length === 1 ? '' : 's'} parsed.
                  {parsed.skippedColumns.length > 0 && (
                    <> Skipped columns (not on this table): {parsed.skippedColumns.join(', ')}.</>
                  )}
                  {parsed.errors.length > 0 && (
                    <ul style={{ margin: '4px 0 0', paddingLeft: 18 }}>
                      {parsed.errors.slice(0, 6).map((e) => (
                        <li key={e}>{e}</li>
                      ))}
                      {parsed.errors.length > 6 && <li>…and {parsed.errors.length - 6} more</li>}
                    </ul>
                  )}
                </MessageBarBody>
              </MessageBar>
            )}
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>
              Cancel
            </Button>
            <Button
              appearance="primary"
              disabled={parsed.rows.length === 0}
              onClick={() => {
                onImport(parsed.rows, replace);
                setText('');
              }}
            >
              Import {parsed.rows.length || ''} row{parsed.rows.length === 1 ? '' : 's'}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

// ──────────────────────────────────────────────────────────────
// Results — returned Ids (CreateMultiple) or the 204 confirmation
// ──────────────────────────────────────────────────────────────

function BulkResult({ result, operation }: { result: ExecResult; operation: BulkOperation }) {
  const ids = idsFromResult(result.body);
  const [copied, setCopied] = useState(false);
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12, maxWidth: 760 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
        <StatusPill status="success" code={result.status} ms={result.ms} />
        <Text weight="semibold">
          {operation}{' '}
          {ids.length > 0 ? `created ${ids.length} row${ids.length === 1 ? '' : 's'}` : 'completed'}
        </Text>
      </div>
      {ids.length > 0 ? (
        <>
          <div>
            <Button
              icon={<Copy20Regular />}
              onClick={() => void copyToClipboardSafe(ids.join('\n')).then(() => setCopied(true))}
            >
              {copied ? 'Copied' : 'Copy ids'}
            </Button>
          </div>
          <ol
            style={{
              margin: 0,
              paddingLeft: 32,
              fontFamily: tokens.fontFamilyMonospace,
              fontSize: 12,
              maxHeight: 480,
              overflow: 'auto',
            }}
          >
            {ids.map((id) => (
              <li key={id}>{id}</li>
            ))}
          </ol>
        </>
      ) : (
        <Caption1>
          {operation} returns 204 No Content — Dataverse doesn't report per-row results for it.
          Every row in the request succeeded (a failure would have rolled back the request on a
          standard table).
        </Caption1>
      )}
    </div>
  );
}
