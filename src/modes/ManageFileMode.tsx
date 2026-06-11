// Manage File Data — File column upload / download / delete pipelines.
//
// Layout:
//   Sidebar:
//     • Operation     — Upload | Download | Delete (radio cards)
//     • Target        — table + record picker
//     • File column   — File-typed column picker w/ metadata summary
//     • Method        — Single | Chunked | Dataverse messages (upload)
//                       Single | Range | Messages | SAS URL (download)
//     • Pipeline      — visual tree showing each request in the pipeline
//   Main:
//     • Banner explaining the operation context + size threshold
//     • Source card (file picker / URL / base64) — upload only
//     • Pipeline card with all queued HTTP requests
//
// Reference docs:
//   files-images-overview · file-column-data · attachment-annotation-files · getfilesasurl

import { useMemo, useState } from 'react';
import { Warning20Filled } from '@fluentui/react-icons';
import {
  Table20Regular,
  Table20Filled,
  Document20Regular,
  Document20Filled,
  Settings20Regular,
  Settings20Filled,
  LineHorizontal320Regular,
  LineHorizontal320Filled,
  ArrowSwap20Regular,
  ArrowSwap20Filled,
  Code20Regular,
  Code20Filled,
  ArrowUpload20Filled,
  ArrowDownload20Filled,
  Delete20Filled,
  Checkmark16Filled,
} from '@fluentui/react-icons';
import {
  Field,
  RadioGroup,
  Radio,
  Caption1,
  Badge,
  tokens,
  mergeClasses,
  SpinButton,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
} from '@fluentui/react-components';
import { Sidebar } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { useStudioStyles } from '../primitives/styles';
import { PaneHead } from '../editors/PaneHead';
import { TargetEditor } from '../editors/TargetEditor';
import { FileColumnPicker } from '../editors/BinaryColumnPickers';
import { BinarySourceCard } from '../editors/BinarySourceCard';
import { BinaryPipelineCard } from '../editors/BinaryPipelineCard';
import { HeadersEditor, defaultWriteHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { findTable, type FileColumnMeta } from '../mock/metadata';
import { findRequestType } from '../registry/requestTypes';
import { buildManageFile, manageFilePipeline, formatSize } from '../engine/urlBuilder';
import { runtime, type ExecResult } from '../engine/runtime';
import {
  checkUploadSize,
  type ManageFileState,
  type FileOperation,
  type FileUploadMethod,
  type FileDownloadMethod,
} from '../state/binaryState';
import type { RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useLiveTable } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import { isEmbedded } from '../host/pptbBridge';
import { adv, type Advisory } from '../primitives/advisories';
import {
  serializeManageFile,
  deserializeManageFile,
  hashState,
  type SavedRequest,
  type SerializedManageFileState,
} from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';

const initialState = (): ManageFileState => ({
  // Empty initial state — user picks table + record + file column from live
  // metadata; no fixture seed.
  table: '',
  recordId: null,
  fileColumn: '',
  operation: 'upload',
  uploadMethod: { kind: 'chunked-patch', chunkSize: 2 * 1024 * 1024 },
  downloadMethod: { kind: 'single-request' },
  source: { kind: 'file' },
  fileName: '',
  fileSize: 0,
  mimeType: 'application/octet-stream',
  bodyBase64: '',
  bodyUrl: '',
  headers: defaultWriteHeaders(),
  dirty: new Set(),
});

type RootClauseId =
  | 'operation'
  | 'target'
  | 'column'
  | 'method'
  | 'source'
  | 'pipeline'
  | 'headers';

export function ManageFileMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('manage-file');
  const [state, setState] = useState(initialState);
  // Live-metadata subscription so child editors re-render when columns/relationships land.
  useLiveTable(state.table || null);
  const [activePath, setActivePath] = useState<string>('pipeline');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);

  const built = useMemo(() => buildManageFile(state), [state]);
  const pipeline = useMemo(() => manageFilePipeline(state), [state]);
  const tbl = findTable(state.table);
  const fileCol = useMemo(
    () =>
      (tbl?.columns ?? []).find(
        (c): c is FileColumnMeta =>
          c.attributeType === 'File' && c.logicalName === state.fileColumn,
      ),
    [tbl, state.fileColumn],
  );

  const markDirty = (id: string) =>
    setState((s) => {
      const d = new Set(s.dirty);
      d.add(id);
      return { ...s, dirty: d };
    });
  const set = <K extends keyof ManageFileState>(k: K, v: ManageFileState[K], dirtyId?: string) => {
    setState((s) => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  const firstStepMethod = pipeline[0]?.method ?? 'PATCH';

  const disabledReason = !tbl
    ? 'Pick a target table first.'
    : !state.recordId
      ? 'Pick a source record.'
      : !state.fileColumn
        ? 'Pick a file column.'
        : state.operation === 'upload' &&
            state.source.kind === 'file' &&
            !state.bodyBase64 &&
            state.source.kind === 'file' &&
            !state.fileName
          ? 'Choose a file to upload.'
          : state.headers.some((h) => h.enabled && !h.name)
            ? 'Fix empty header name.'
            : null;

  // Bound-record name capture for the WriteResultCard narrative.
  const [recordName, setRecordName] = useState<string>('');

  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.manageFile(state);
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents((rs) =>
      [
        {
          id: `r-${Date.now()}`,
          modeId: 'manage-file',
          url: built.relativeUrl,
          method: firstStepMethod,
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
  const { entities } = useScopedEntities();
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);
  const currentSerialized = useMemo(() => serializeManageFile(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;
  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };
  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== 'manage-file') return;
    const snap = entry.state as SerializedManageFileState;
    if (entities.length > 0 && snap.table && !entities.some((e) => e.logicalName === snap.table)) {
      window.alert(
        `Can't load "${entry.name}": entity \`${snap.table}\` ` +
          `isn't available in this environment.`,
      );
      return;
    }
    setState(deserializeManageFile(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setRecordName('');
    setActivePath('pipeline');
  };
  usePublishSaveContext(
    useMemo(() => {
      if (!state.table) return null;
      return {
        state: currentSerialized,
        modeId: 'manage-file' as const,
        dirty: isDirty,
        lastSavedId,
        onSaved,
        onLoadSaved,
      };
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [currentSerialized, isDirty, lastSavedId, state.table]),
  );

  // ── PPTB host-limitation advisory ──
  //
  // The host's dataverseAPI exposes no file/image endpoints (no upload,
  // chunked PATCH, ranged GET, or message-based pipelines). Executing
  // inside DRS always returns 501 Not Implemented — we surface a clear
  // warning advisory so the user knows up-front to copy the snippet from
  // the Code tab and run it externally.
  const advisories = useMemo<Advisory[]>(() => {
    if (!isEmbedded()) return [];
    return [
      adv.warn(
        'binary-host-unsupported',
        'header',
        'File operations run externally',
        <span>
          PPTB's <code>dataverseAPI</code> doesn't expose file column endpoints. Build the request
          here, then copy a snippet from the <strong>Code</strong> tab (fetch / curl / C# /
          PowerShell) and run it from there. Execute will return <code>501 Not Implemented</code>.
        </span>,
        'pipeline',
        'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/file-column-data',
      ),
    ];
  }, []);

  const opBadge =
    state.operation === 'upload'
      ? 'UPLOAD'
      : state.operation === 'download'
        ? 'DOWNLOAD'
        : 'DELETE';
  const opBadgeColor = state.operation === 'delete' ? ('danger' as const) : ('brand' as const);
  const methodLabel =
    state.operation === 'upload'
      ? state.uploadMethod.kind
      : state.operation === 'download'
        ? state.downloadMethod.kind
        : 'simple';

  const sections = [
    {
      id: 'operation',
      label: 'Operation',
      meta: opBadge.toLowerCase(),
      items: [
        {
          id: 'operation',
          icon: ArrowSwap20Regular,
          iconFilled: ArrowSwap20Filled,
          label: opBadge.charAt(0) + opBadge.slice(1).toLowerCase(),
          badge: opBadge,
          badgeAppearance: 'tint' as const,
          badgeColor: opBadgeColor,
          dirty: state.dirty.has('operation'),
        },
      ],
    },
    {
      id: 'target',
      label: 'Target',
      meta: tbl?.displayName ?? '?',
      items: [
        {
          id: 'target',
          icon: Table20Regular,
          iconFilled: Table20Filled,
          label: state.recordId ? `${tbl?.displayName ?? ''} (selected)` : 'Pick a record',
          badge: state.recordId ? '✓' : null,
          badgeAppearance: 'ghost' as const,
          dirty: state.dirty.has('target'),
        },
        {
          id: 'column',
          icon: Document20Regular,
          iconFilled: Document20Filled,
          label: state.fileColumn || 'Pick a file column',
          code: !!state.fileColumn,
          badge: fileCol ? formatSize((fileCol.maxSizeInKB ?? 32768) * 1024) : null,
          badgeAppearance: 'ghost' as const,
          dirty: state.dirty.has('column'),
        },
      ],
    },
    ...(state.operation !== 'delete'
      ? [
          {
            id: 'method',
            label: 'Method',
            meta: methodLabel,
            items: [
              {
                id: 'method',
                icon: Settings20Regular,
                iconFilled: Settings20Filled,
                label: state.operation === 'upload' ? 'Upload method' : 'Download method',
                badge: methodLabel,
                badgeAppearance: 'ghost' as const,
                dirty: state.dirty.has('method'),
              },
            ],
          },
        ]
      : []),
    ...(state.operation === 'upload'
      ? [
          {
            id: 'payload',
            label: 'Payload',
            meta: state.fileName || '—',
            items: [
              {
                id: 'source',
                icon: Document20Regular,
                iconFilled: Document20Filled,
                label: 'Source',
                badge: state.source.kind,
                badgeAppearance: 'ghost' as const,
                dirty: state.dirty.has('source'),
              },
            ],
          },
        ]
      : []),
    {
      id: 'pipe',
      label: 'Pipeline',
      meta: `${pipeline.length} request${pipeline.length === 1 ? '' : 's'}`,
      items: [
        {
          id: 'pipeline',
          icon: Code20Regular,
          iconFilled: Code20Filled,
          label: 'Generated requests',
          badge: pipeline.length || null,
          badgeAppearance: 'tint' as const,
          badgeColor: 'brand' as const,
          dirty: state.dirty.has('pipeline'),
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
  let pane: React.ReactNode;
  const root = activePath as RootClauseId;
  switch (root) {
    case 'operation':
      pane = (
        <OperationPicker
          value={state.operation}
          onChange={(op) => set('operation', op, 'operation')}
        />
      );
      break;
    case 'target':
      pane = (
        <TargetEditor
          table={state.table}
          onTableChange={(t) => {
            // Switching/clearing the target invalidates the record + file
            // column (those File-typed columns belong to the old entity).
            setState((s) => ({
              ...s,
              table: t,
              recordId: null,
              fileColumn: null,
              dirty: new Set(['target', 'column']),
            }));
            setResult(null);
          }}
          recordId={state.recordId}
          onRecordChange={(id, primary) => {
            set('recordId', id, 'target');
            setRecordName(primary ?? '');
          }}
          group="binary"
          sub="Pick the entity + record that holds the file column."
        />
      );
      break;
    case 'column':
      pane = (
        <FileColumnPicker
          table={state.table}
          value={state.fileColumn}
          onChange={(c) => set('fileColumn', c, 'column')}
        />
      );
      break;
    case 'method':
      pane =
        state.operation === 'upload' ? (
          <UploadMethodPicker
            value={state.uploadMethod}
            onChange={(m) => set('uploadMethod', m, 'method')}
          />
        ) : (
          <DownloadMethodPicker
            value={state.downloadMethod}
            onChange={(m) => set('downloadMethod', m, 'method')}
          />
        );
      break;
    case 'source':
      pane = (
        <BinarySourceCard
          source={state.source}
          setSource={(s) => set('source', s, 'source')}
          fileName={state.fileName}
          setFileName={(n) => set('fileName', n, 'source')}
          fileSize={state.fileSize}
          setFileSize={(n) => set('fileSize', n, 'source')}
          mimeType={state.mimeType}
          setMimeType={(m) => set('mimeType', m, 'source')}
          bodyBase64={state.bodyBase64}
          setBodyBase64={(b) => set('bodyBase64', b, 'source')}
          bodyUrl={state.bodyUrl}
          setBodyUrl={(u) => set('bodyUrl', u, 'source')}
          group="binary"
        />
      );
      break;
    case 'pipeline':
      pane = (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
          <OperationBanner state={state} fileCol={fileCol} />
          <BinaryPipelineCard
            steps={pipeline}
            title={`${opBadge} pipeline — ${methodLabel}`}
            sub={pipeline.length > 1 ? 'Each request fires sequentially.' : undefined}
            sizeAdvisory={
              state.operation === 'upload'
                ? checkUploadSize(
                    state.fileSize,
                    fileCol?.maxSizeInKB ? fileCol.maxSizeInKB * 1024 : undefined,
                    'file',
                  )
                : null
            }
          />
        </div>
      );
      break;
    case 'headers':
      pane = (
        <HeadersEditor
          items={state.headers}
          setItems={(h) => set('headers', h, 'headers')}
          group="binary"
        />
      );
      break;
  }

  const headersMap = headerItemsToObject(state.headers, null);
  const codeInputs = {
    method: firstStepMethod,
    built,
    headers: headersMap,
    entityLogical: built.entityLogical,
    multiRequests:
      pipeline.length > 1
        ? pipeline.map((step) => ({
            method: step.method,
            relativeUrl: step.relativeUrl.startsWith('<')
              ? step.relativeUrl
              : step.relativeUrl.replace(/^https?:\/\/[^/]+/, ''),
            body: typeof step.body === 'object' ? step.body : undefined,
            description: step.title,
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
          onSelect={(id) => {
            setActivePath(id);
            setTab('builder');
          }}
          recents={recents}
        />
      }
      urlBar={
        <UrlBar
          method={firstStepMethod}
          url={built.relativeUrl}
          // Manage File/Image button verb mirrors the active operation:
          // Upload / Download / Delete.
          executeVerb={
            state.operation === 'upload'
              ? 'Upload'
              : state.operation === 'download'
                ? 'Download'
                : 'Delete'
          }
          executeIcon={
            state.operation === 'upload'
              ? ArrowUpload20Filled
              : state.operation === 'download'
                ? ArrowDownload20Filled
                : Delete20Filled
          }
          disabledReason={disabledReason}
          loading={loading}
          onExecute={onExecute}
          advisories={advisories}
          onAdvisoryFocus={setActivePath}
        />
      }
    >
      <MainTabs tab={tab} onTabChange={setTab} resultCount={result?.ok ? pipeline.length : null}>
        {tab === 'builder' && pane}
        {tab === 'code' && <CodeView themeMode={themeMode} inputs={codeInputs} />}
        {tab === 'results' && (
          <ResultsView
            result={result}
            mode="single"
            table={state.table}
            writeContext={{
              operation: 'manage-file',
              table: state.table,
              recordId: state.recordId,
              recordName: recordName || null,
              fileOperation: state.operation,
              columnName: state.fileColumn ?? undefined,
              fileName: state.fileName || undefined,
            }}
          />
        )}
      </MainTabs>
    </ModeShell>
  );
}

// ──────────────────────────────────────────────────────────────
// Operation picker — 3 radio cards
// ──────────────────────────────────────────────────────────────
function OperationPicker({
  value,
  onChange,
}: {
  value: FileOperation;
  onChange: (op: FileOperation) => void;
}) {
  const s = useStudioStyles();
  const cards: { id: FileOperation; label: string; sub: string; tone: 'brand' | 'danger' }[] = [
    {
      id: 'upload',
      label: 'Upload',
      sub: 'PATCH or message-based upload pipeline.',
      tone: 'brand',
    },
    {
      id: 'download',
      label: 'Download',
      sub: 'GET /$value, ranged GET, message pipeline, or SAS URL.',
      tone: 'brand',
    },
    {
      id: 'delete',
      label: 'Delete',
      sub: 'DELETE /<set>(<id>)/<col> · clears the column value.',
      tone: 'danger',
    },
  ];
  return (
    <div>
      <PaneHead
        icon={ArrowSwap20Filled}
        title="Operation"
        sub="What to do with the file column — drives the URL bar method + pipeline shape."
        group="binary"
      />
      <div
        style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 12, maxWidth: 960 }}
      >
        {cards.map((c) => {
          const selected = value === c.id;
          return (
            <button
              key={c.id}
              type="button"
              onClick={() => onChange(c.id)}
              className={mergeClasses(s.inlineCard)}
              style={{
                textAlign: 'left',
                padding: 14,
                cursor: 'pointer',
                background: selected
                  ? c.tone === 'danger'
                    ? tokens.colorPaletteRedBackground2
                    : tokens.colorBrandBackground2
                  : tokens.colorNeutralBackground1,
                border: `1px solid ${selected ? (c.tone === 'danger' ? tokens.colorPaletteRedBorderActive : tokens.colorBrandStroke1) : tokens.colorNeutralStroke2}`,
                borderRadius: tokens.borderRadiusMedium,
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <strong style={{ fontSize: 13 }}>{c.label}</strong>
                {/* Use an iconic checkmark, not a "SELECTED" caps label.
                    The selection state is already encoded twice (bg + border);
                    a text label is redundant chrome. */}
                {selected && (
                  <Checkmark16Filled
                    style={{
                      marginLeft: 'auto',
                      color:
                        c.tone === 'danger'
                          ? tokens.colorPaletteRedForeground1
                          : tokens.colorBrandForeground1,
                    }}
                    aria-label="Selected"
                  />
                )}
              </div>
              <Caption1
                style={{
                  color: tokens.colorNeutralForeground2,
                  lineHeight: 1.4,
                  marginTop: 6,
                  display: 'block',
                }}
              >
                {c.sub}
              </Caption1>
            </button>
          );
        })}
      </div>
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Upload method picker — single / chunked / messages
// ──────────────────────────────────────────────────────────────
function UploadMethodPicker({
  value,
  onChange,
}: {
  value: FileUploadMethod;
  onChange: (m: FileUploadMethod) => void;
}) {
  return (
    <div>
      <PaneHead
        icon={Settings20Filled}
        title="Upload method"
        sub="Per file-column-data — three supported approaches."
        group="binary"
      />
      <div style={{ maxWidth: 760, display: 'flex', flexDirection: 'column', gap: 14 }}>
        <Field label="Approach">
          <RadioGroup
            value={value.kind}
            onChange={(_, d) => {
              switch (d.value) {
                case 'single-request':
                  onChange({ kind: 'single-request' });
                  break;
                case 'chunked-patch':
                  onChange({
                    kind: 'chunked-patch',
                    chunkSize: value.kind === 'chunked-patch' ? value.chunkSize : 2 * 1024 * 1024,
                  });
                  break;
                case 'dataverse-messages':
                  onChange({
                    kind: 'dataverse-messages',
                    blockSize:
                      value.kind === 'dataverse-messages' ? value.blockSize : 4 * 1024 * 1024,
                  });
                  break;
              }
            }}
          >
            <Radio
              value="single-request"
              label={
                <span>
                  <strong>Single-request PATCH</strong>{' '}
                  <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>
                    PATCH /(set)(id)/col
                  </code>
                  <Caption1
                    style={{
                      display: 'block',
                      color: tokens.colorNeutralForeground3,
                      marginTop: 2,
                    }}
                  >
                    Fastest for files up to ~128 MB. Body = raw bytes; x-ms-file-name header. No
                    init/commit ceremony.
                  </Caption1>
                </span>
              }
            />
            <Radio
              value="chunked-patch"
              label={
                <span>
                  <strong>Chunked PATCH</strong>{' '}
                  <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>
                    x-ms-transfer-mode: chunked
                  </code>
                  <Caption1
                    style={{
                      display: 'block',
                      color: tokens.colorNeutralForeground3,
                      marginTop: 2,
                    }}
                  >
                    Init PATCH returns Location + x-ms-chunk-size. Each chunk = PATCH with
                    Content-Range. Server commits on final chunk.
                  </Caption1>
                </span>
              }
            />
            <Radio
              value="dataverse-messages"
              label={
                <span>
                  <strong>Dataverse messages</strong>{' '}
                  <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>
                    InitializeFileBlocksUpload / UploadBlock / CommitFileBlocksUpload
                  </code>
                  <Caption1
                    style={{
                      display: 'block',
                      color: tokens.colorNeutralForeground3,
                      marginTop: 2,
                    }}
                  >
                    Three OData actions. Up to 10 GB. Returns FileId on commit. Use for files &gt;
                    128 MB or when you need parallel block uploads.
                  </Caption1>
                </span>
              }
            />
          </RadioGroup>
        </Field>

        {value.kind === 'chunked-patch' && (
          <Field
            label="Chunk size (bytes)"
            hint="Defaults to 2 MB. Server may suggest a different size via x-ms-chunk-size on init."
          >
            <SpinButton
              value={value.chunkSize}
              min={64 * 1024}
              max={64 * 1024 * 1024}
              step={1024 * 1024}
              onChange={(_, d) =>
                onChange({
                  kind: 'chunked-patch',
                  chunkSize: Number(d.value ?? d.displayValue ?? value.chunkSize),
                })
              }
            />
            <Caption1
              style={{ marginTop: 4, display: 'block', color: tokens.colorNeutralForeground3 }}
            >
              ≈ {formatSize(value.chunkSize)} per chunk.
            </Caption1>
          </Field>
        )}

        {value.kind === 'dataverse-messages' && (
          <Field
            label="Block size (bytes)"
            hint="Default = 4 MB. The docs cap each UploadBlock at 4 MB; smaller is fine."
          >
            <SpinButton
              value={value.blockSize}
              min={64 * 1024}
              max={4 * 1024 * 1024}
              step={1024 * 1024}
              onChange={(_, d) =>
                onChange({
                  kind: 'dataverse-messages',
                  blockSize: Number(d.value ?? d.displayValue ?? value.blockSize),
                })
              }
            />
            <Caption1
              style={{ marginTop: 4, display: 'block', color: tokens.colorNeutralForeground3 }}
            >
              ≈ {formatSize(value.blockSize)} per block.
            </Caption1>
          </Field>
        )}
      </div>
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Download method picker — single / ranged / messages / SAS URL
// ──────────────────────────────────────────────────────────────
function DownloadMethodPicker({
  value,
  onChange,
}: {
  value: FileDownloadMethod;
  onChange: (m: FileDownloadMethod) => void;
}) {
  return (
    <div>
      <PaneHead
        icon={Settings20Filled}
        title="Download method"
        sub="Per file-column-data + getfilesasurl — four supported approaches."
        group="binary"
      />
      <div style={{ maxWidth: 760, display: 'flex', flexDirection: 'column', gap: 14 }}>
        <Field label="Approach">
          <RadioGroup
            value={value.kind}
            onChange={(_, d) => {
              switch (d.value) {
                case 'single-request':
                  onChange({ kind: 'single-request' });
                  break;
                case 'ranged':
                  onChange({ kind: 'ranged', rangeStart: 0, rangeEnd: 4 * 1024 * 1024 - 1 });
                  break;
                case 'dataverse-messages':
                  onChange({ kind: 'dataverse-messages', blockSize: 4 * 1024 * 1024 });
                  break;
                case 'sas-url':
                  onChange({ kind: 'sas-url' });
                  break;
              }
            }}
          >
            <Radio
              value="single-request"
              label={
                <span>
                  <strong>GET /$value</strong>
                  <Caption1
                    style={{
                      display: 'block',
                      color: tokens.colorNeutralForeground3,
                      marginTop: 2,
                    }}
                  >
                    Stream the entire file in one request. Returns raw bytes. Easy but not great for
                    huge files.
                  </Caption1>
                </span>
              }
            />
            <Radio
              value="ranged"
              label={
                <span>
                  <strong>Ranged GET</strong>{' '}
                  <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>
                    Range: bytes=A-B
                  </code>
                  <Caption1
                    style={{
                      display: 'block',
                      color: tokens.colorNeutralForeground3,
                      marginTop: 2,
                    }}
                  >
                    Partial content — server returns 206. Use for resume / parallel download.
                  </Caption1>
                </span>
              }
            />
            <Radio
              value="dataverse-messages"
              label={
                <span>
                  <strong>Dataverse messages</strong>{' '}
                  <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>
                    InitializeFileBlocksDownload + DownloadBlock
                  </code>
                  <Caption1
                    style={{
                      display: 'block',
                      color: tokens.colorNeutralForeground3,
                      marginTop: 2,
                    }}
                  >
                    Block-by-block download. Returns FileName + FileSizeInBytes +
                    FileContinuationToken from init.
                  </Caption1>
                </span>
              }
            />
            <Radio
              value="sas-url"
              label={
                <span>
                  <strong>GetFileSasUrl</strong>{' '}
                  <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>
                    GET /GetFileSasUrl(…)
                  </code>
                  <Caption1
                    style={{
                      display: 'block',
                      color: tokens.colorNeutralForeground3,
                      marginTop: 2,
                    }}
                  >
                    Returns a 1-hour SAS URL that anyone can use to download directly from Azure
                    Storage. Skips Web API roundtrip on the actual download.
                  </Caption1>
                </span>
              }
            />
          </RadioGroup>
        </Field>

        {value.kind === 'ranged' && (
          <div style={{ display: 'flex', gap: 10, maxWidth: 480 }}>
            <Field label="Range start (bytes)">
              <SpinButton
                value={value.rangeStart}
                min={0}
                onChange={(_, d) =>
                  onChange({
                    kind: 'ranged',
                    rangeStart: Number(d.value ?? d.displayValue ?? 0),
                    rangeEnd: value.rangeEnd,
                  })
                }
              />
            </Field>
            <Field label="Range end (bytes)">
              <SpinButton
                value={value.rangeEnd}
                min={0}
                onChange={(_, d) =>
                  onChange({
                    kind: 'ranged',
                    rangeStart: value.rangeStart,
                    rangeEnd: Number(d.value ?? d.displayValue ?? 0),
                  })
                }
              />
            </Field>
          </div>
        )}

        {value.kind === 'dataverse-messages' && (
          <Field label="Block size (bytes)">
            <SpinButton
              value={value.blockSize}
              min={64 * 1024}
              max={4 * 1024 * 1024}
              step={1024 * 1024}
              onChange={(_, d) =>
                onChange({
                  kind: 'dataverse-messages',
                  blockSize: Number(d.value ?? d.displayValue ?? value.blockSize),
                })
              }
            />
          </Field>
        )}

        {value.kind === 'sas-url' && (
          <MessageBar layout="multiline" intent="info">
            <MessageBarBody>
              <MessageBarTitle>SAS URL caveats</MessageBarTitle>
              <ul style={{ margin: '4px 0 0', paddingLeft: 18, lineHeight: 1.6 }}>
                <li>Only works for image columns that support full-size images.</li>
                <li>SAS URL valid for 1 hour from generation.</li>
                <li>
                  Does NOT work for environments using BYOK encryption — migrate to CMK first.
                </li>
                <li>
                  Caller must have access to the record; the resulting URL grants anonymous access
                  for 1 hour.
                </li>
              </ul>
            </MessageBarBody>
          </MessageBar>
        )}
      </div>
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Operation banner — contextual intro at the top of the pipeline pane
// ──────────────────────────────────────────────────────────────
function OperationBanner({
  state,
  fileCol,
}: {
  state: ManageFileState;
  fileCol: FileColumnMeta | undefined;
}) {
  const maxBytes = (fileCol?.maxSizeInKB ?? 32768) * 1024;
  if (state.operation === 'upload') {
    const above4mb = state.fileSize > 4 * 1024 * 1024;
    const above128mb = state.fileSize > 128 * 1024 * 1024;
    return (
      <MessageBar layout="multiline" intent={above128mb ? 'warning' : 'info'}>
        <MessageBarBody>
          <MessageBarTitle>Upload</MessageBarTitle>
          {above128mb ? (
            <>
              File is over <strong>128 MB</strong> — single-request PATCH won't work. Use the
              Dataverse messages path which supports up to 10 GB.
            </>
          ) : above4mb ? (
            <>
              File is over <strong>4 MB</strong> — chunked PATCH (or Dataverse messages) is
              recommended; single-request still works up to 128 MB.
            </>
          ) : (
            <>Small enough for single-request PATCH. Column max: {formatSize(maxBytes)}.</>
          )}
        </MessageBarBody>
      </MessageBar>
    );
  }
  if (state.operation === 'download') {
    return (
      <MessageBar layout="multiline" intent="info">
        <MessageBarBody>
          <MessageBarTitle>Download</MessageBarTitle>
          Retrieving a file column returns the <strong>file id</strong> (not bytes) — use the
          operations below to fetch actual content. GET /$value is the simplest; SAS URL skips the
          Web API roundtrip when an anonymous browser-side download is fine.
        </MessageBarBody>
      </MessageBar>
    );
  }
  return (
    <MessageBar layout="multiline" intent="warning">
      <MessageBarBody>
        <MessageBarTitle>Delete</MessageBarTitle>
        Clears the file column value + the <code>_name</code> companion. The underlying file is
        removed by Dataverse. Alternative equivalent paths: PATCH the column to null, or call the{' '}
        <code>DeleteFile</code> action with the column's current FileId.
      </MessageBarBody>
    </MessageBar>
  );
}
