// Manage Image Data — Image column upload / download / delete.
//
// Layout:
//   Sidebar:
//     • Operation     — Upload | Retrieve | Delete (radio cards)
//     • Target        — table + record picker
//     • Image column  — entityimage (primary) vs custom image column,
//                       with isPrimary / canStoreFullImage badges
//     • Derivatives   — companion columns tree (<col>id / _Timestamp / _URL)
//   Main:
//     • Banner explaining entityimage vs custom (thumbnail-only vs full)
//     • Preview tile (144×144) with source picker (file / URL / base64)
//     • Generated request card
//
// Reference docs:
//   image-column-data · files-images-overview

import { useMemo, useState } from 'react';
import {
  Table20Regular,
  Table20Filled,
  Image20Regular,
  Image20Filled,
  ArrowSwap20Regular,
  ArrowSwap20Filled,
  LineHorizontal320Regular,
  LineHorizontal320Filled,
  Settings20Regular,
  Settings20Filled,
  Code20Regular,
  Code20Filled,
  ArrowUpload20Filled,
  ArrowDownload20Filled,
  Delete20Filled,
  Warning20Filled,
} from '@fluentui/react-icons';
import {
  RadioGroup,
  Radio,
  Caption1,
  Badge,
  Field,
  tokens,
  mergeClasses,
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
import { ImageColumnPicker } from '../editors/BinaryColumnPickers';
import { BinarySourceCard } from '../editors/BinarySourceCard';
import { BinaryPipelineCard } from '../editors/BinaryPipelineCard';
import { HeadersEditor, defaultWriteHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { findTable, type ImageColumnMeta } from '../mock/metadata';
import { findRequestType } from '../registry/requestTypes';
import { buildManageImage, manageImagePipeline } from '../engine/urlBuilder';
import { runtime, type ExecResult } from '../engine/runtime';
import { checkUploadSize, type ManageImageState, type FileOperation } from '../state/binaryState';

/**
 * Approximate decoded byte count for a base 64 payload. base64 grows the
 * source by ~33% — for the pre-flight advisory a rough estimate (with padding
 * accounted for) is enough; we don't need to actually decode.
 */
function approxBytesFromBase64(b64: string): number {
  if (!b64) return 0;
  const padding = b64.endsWith('==') ? 2 : b64.endsWith('=') ? 1 : 0;
  return Math.max(0, Math.floor((b64.length * 3) / 4) - padding);
}
import type { RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useLiveTable } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import { isEmbedded } from '../host/pptbBridge';
import { adv, type Advisory } from '../primitives/advisories';
import {
  serializeManageImage,
  deserializeManageImage,
  hashState,
  type SavedRequest,
  type SerializedManageImageState,
} from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';

const initialState = (): ManageImageState => ({
  // Empty initial state — user picks everything from live metadata.
  table: '',
  recordId: null,
  imageColumn: '',
  operation: 'upload',
  downloadSize: 'thumbnail',
  source: { kind: 'file' },
  fileName: '',
  bodyBase64: '',
  bodyUrl: '',
  mimeType: 'image/png',
  headers: defaultWriteHeaders(),
  dirty: new Set(),
});

type RootClauseId =
  | 'operation'
  | 'target'
  | 'column'
  | 'source'
  | 'preview'
  | 'options'
  | 'pipeline'
  | 'headers';

export function ManageImageMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('manage-image');
  const [state, setState] = useState(initialState);
  // Live-metadata subscription so child editors re-render when columns/relationships land.
  useLiveTable(state.table || null);
  const [activePath, setActivePath] = useState<string>('preview');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);

  const built = useMemo(() => buildManageImage(state), [state]);
  const pipeline = useMemo(() => manageImagePipeline(state), [state]);
  const tbl = findTable(state.table);
  const imageCol = useMemo(
    () =>
      (tbl?.columns ?? []).find(
        (c): c is ImageColumnMeta =>
          c.attributeType === 'Image' && c.logicalName === state.imageColumn,
      ),
    [tbl, state.imageColumn],
  );

  const markDirty = (id: string) =>
    setState((s) => {
      const d = new Set(s.dirty);
      d.add(id);
      return { ...s, dirty: d };
    });
  const set = <K extends keyof ManageImageState>(
    k: K,
    v: ManageImageState[K],
    dirtyId?: string,
  ) => {
    setState((s) => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  const firstStepMethod = pipeline[0]?.method ?? 'PATCH';

  const disabledReason = !tbl
    ? 'Pick a target table first.'
    : !state.recordId
      ? 'Pick a source record.'
      : !state.imageColumn
        ? 'Pick an image column.'
        : state.operation === 'upload' && !state.bodyBase64 && state.source.kind === 'file'
          ? 'Pick an image file.'
          : state.operation === 'download' &&
              state.downloadSize === 'full' &&
              !imageCol?.canStoreFullImage
            ? `${state.imageColumn} doesn't store full-size images (canStoreFullImage = false).`
            : state.headers.some((h) => h.enabled && !h.name)
              ? 'Fix empty header name.'
              : null;

  // Bound-record name capture for the WriteResultCard narrative.
  const [recordName, setRecordName] = useState<string>('');

  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.manageImage(state);
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents((rs) =>
      [
        {
          id: `r-${Date.now()}`,
          modeId: 'manage-image',
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
  const currentSerialized = useMemo(() => serializeManageImage(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;
  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };
  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== 'manage-image') return;
    const snap = entry.state as SerializedManageImageState;
    if (entities.length > 0 && snap.table && !entities.some((e) => e.logicalName === snap.table)) {
      window.alert(
        `Can't load "${entry.name}": entity \`${snap.table}\` ` +
          `isn't available in this environment.`,
      );
      return;
    }
    setState(deserializeManageImage(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setRecordName('');
    setActivePath('preview');
  };
  usePublishSaveContext(
    useMemo(() => {
      if (!state.table) return null;
      return {
        state: currentSerialized,
        modeId: 'manage-image' as const,
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
  // dataverseAPI exposes no image endpoints (no upload, no ?size=full
  // download, no entityimage_url derivative reads). Executing inside DRS
  // always returns 501. Surface a warning so the user knows to copy a
  // snippet from the Code tab and run it externally.
  const advisories = useMemo<Advisory[]>(() => {
    if (!isEmbedded()) return [];
    return [
      adv.warn(
        'binary-host-unsupported',
        'header',
        'Image operations run externally',
        <span>
          PPTB's <code>dataverseAPI</code> doesn't expose image column endpoints. Build the request
          here, then copy a snippet from the <strong>Code</strong> tab (fetch / curl / C# /
          PowerShell) and run it from there. Execute will return <code>501 Not Implemented</code>.
        </span>,
        'preview',
        'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/image-column-data',
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

  // Initials placeholder for the avatar preview tile. The richer "fetch the
  // primary name to derive initials" path lived against the mock records
  // catalog; with that gone we just render a question mark until the user
  // wires up a live record-name lookup (TODO if/when this mode lands).
  const initials = '?';

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
          icon: Image20Regular,
          iconFilled: Image20Filled,
          label: state.imageColumn || 'Pick an image column',
          code: !!state.imageColumn,
          badge: imageCol?.isPrimaryImage
            ? 'primary'
            : imageCol?.canStoreFullImage
              ? 'full'
              : 'thumb',
          badgeAppearance: 'tint' as const,
          badgeColor: imageCol?.canStoreFullImage ? ('success' as const) : ('brand' as const),
          dirty: state.dirty.has('column'),
        },
      ],
    },
    ...(imageCol
      ? [
          {
            id: 'deriv',
            label: 'Derivatives',
            meta: '3 companion columns',
            items: [
              {
                id: 'd1',
                icon: Image20Regular,
                label: `${imageCol.logicalName}id`,
                code: true,
                badge: 'guid',
                badgeAppearance: 'ghost' as const,
              },
              {
                id: 'd2',
                icon: Image20Regular,
                label: `${imageCol.logicalName}_Timestamp`,
                code: true,
                badge: 'bigint',
                badgeAppearance: 'ghost' as const,
              },
              {
                id: 'd3',
                icon: Image20Regular,
                label: `${imageCol.logicalName}_URL`,
                code: true,
                badge: 'url',
                badgeAppearance: 'ghost' as const,
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
                icon: Image20Regular,
                iconFilled: Image20Filled,
                label: 'Source',
                badge: state.source.kind,
                badgeAppearance: 'ghost' as const,
                dirty: state.dirty.has('source'),
              },
            ],
          },
        ]
      : []),
    ...(state.operation === 'download'
      ? [
          {
            id: 'options',
            label: 'Options',
            meta: state.downloadSize,
            items: [
              {
                id: 'options',
                icon: Settings20Regular,
                iconFilled: Settings20Filled,
                label: 'Download size',
                badge: state.downloadSize,
                badgeAppearance: 'tint' as const,
                badgeColor: 'brand' as const,
                dirty: state.dirty.has('options'),
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
            // Switching/clearing the target invalidates the record +
            // image column (those Image-typed columns belong to the old entity).
            setState((s) => ({
              ...s,
              table: t,
              recordId: null,
              imageColumn: null,
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
          sub="Pick the entity + record that holds the image column."
        />
      );
      break;
    case 'column':
      pane = (
        <ImageColumnPicker
          table={state.table}
          value={state.imageColumn}
          onChange={(c) => set('imageColumn', c, 'column')}
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
          fileSize={0}
          setFileSize={() => undefined}
          mimeType={state.mimeType}
          setMimeType={(m) => set('mimeType', m as ManageImageState['mimeType'], 'source')}
          bodyBase64={state.bodyBase64}
          setBodyBase64={(b) => set('bodyBase64', b, 'source')}
          bodyUrl={state.bodyUrl}
          setBodyUrl={(u) => set('bodyUrl', u, 'source')}
          accept="image/png,image/jpeg,image/gif,image/bmp,image/tiff"
          group="binary"
        />
      );
      break;
    case 'options':
      pane = (
        <DownloadOptionsPane
          downloadSize={state.downloadSize}
          setDownloadSize={(s) => set('downloadSize', s, 'options')}
          canStoreFullImage={imageCol?.canStoreFullImage ?? false}
        />
      );
      break;
    case 'preview':
      pane = (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
          <OperationBanner state={state} imageCol={imageCol} />
          <PreviewCard
            state={state}
            imageCol={imageCol}
            initials={initials}
            setSource={(s) => set('source', s, 'source')}
            setBodyBase64={(b) => set('bodyBase64', b, 'source')}
            setFileName={(n) => set('fileName', n, 'source')}
            setMimeType={(m) => set('mimeType', m as ManageImageState['mimeType'], 'source')}
          />
          <BinaryPipelineCard
            steps={pipeline}
            title={`${opBadge} pipeline`}
            sizeAdvisory={
              state.operation === 'upload'
                ? checkUploadSize(
                    approxBytesFromBase64(state.bodyBase64),
                    imageCol?.maxSizeInKB ? imageCol.maxSizeInKB * 1024 : undefined,
                    'image',
                  )
                : null
            }
          />
        </div>
      );
      break;
    case 'pipeline':
      pane = (
        <BinaryPipelineCard
          steps={pipeline}
          title={`${opBadge} pipeline`}
          sizeAdvisory={
            state.operation === 'upload'
              ? checkUploadSize(
                  approxBytesFromBase64(state.bodyBase64),
                  imageCol?.maxSizeInKB ? imageCol.maxSizeInKB * 1024 : undefined,
                  'image',
                )
              : null
          }
        />
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
    body:
      pipeline[0] && typeof pipeline[0].body === 'object'
        ? (pipeline[0].body as Record<string, unknown>)
        : undefined,
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
          method={firstStepMethod}
          url={built.relativeUrl}
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
      <MainTabs tab={tab} onTabChange={setTab} resultCount={result?.ok ? 1 : null}>
        {tab === 'builder' && pane}
        {tab === 'code' && <CodeView themeMode={themeMode} inputs={codeInputs} />}
        {tab === 'results' && (
          <ResultsView
            result={result}
            mode="single"
            table={state.table}
            writeContext={{
              operation: 'manage-image',
              table: state.table,
              recordId: state.recordId,
              recordName: recordName || null,
              fileOperation: state.operation,
              columnName: state.imageColumn ?? undefined,
              fileName: state.fileName || undefined,
            }}
          />
        )}
      </MainTabs>
    </ModeShell>
  );
}

// ──────────────────────────────────────────────────────────────
// Preview card with 144×144 tile (per v2.2 design)
// ──────────────────────────────────────────────────────────────
function PreviewCard({
  state,
  imageCol,
  initials,
  setSource,
  setBodyBase64,
  setFileName,
  setMimeType,
}: {
  state: ManageImageState;
  imageCol: ImageColumnMeta | undefined;
  initials: string;
  setSource: (s: ManageImageState['source']) => void;
  setBodyBase64: (b: string) => void;
  setFileName: (n: string) => void;
  setMimeType: (m: ManageImageState['mimeType']) => void;
}) {
  const s = useStudioStyles();
  const isThumb = imageCol?.isPrimaryImage && !imageCol?.canStoreFullImage;

  return (
    <div className={mergeClasses(s.inlineCard)} style={{ padding: 14, maxWidth: 880 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12 }}>
        <Image20Filled style={{ color: tokens.colorBrandForeground1 }} />
        <strong style={{ fontSize: 13 }}>Preview · Source</strong>
        <Badge appearance="tint" color="brand">
          {imageCol?.isPrimaryImage ? 'primary image' : 'custom image column'}
        </Badge>
        {isThumb && <Badge appearance="ghost">144×144 only</Badge>}
        {imageCol?.canStoreFullImage && (
          <Badge appearance="tint" color="success">
            full-size enabled
          </Badge>
        )}
      </div>

      <div style={{ display: 'flex', gap: 20, alignItems: 'flex-start' }}>
        {/* 144×144 preview tile */}
        <div>
          <Caption1
            style={{ color: tokens.colorNeutralForeground3, display: 'block', marginBottom: 4 }}
          >
            {state.operation === 'upload'
              ? 'Preview'
              : state.downloadSize === 'full'
                ? 'Full-size'
                : 'Thumbnail'}{' '}
            · 144×144
          </Caption1>
          {/* Empty-state tile uses a flat neutral background — a colored
              gradient here would compete with the group-driven brand color
              and break the no-hardcoded-hex rule used elsewhere. */}
          <div
            style={{
              width: 144,
              height: 144,
              background: state.bodyBase64
                ? `url(data:${state.mimeType};base64,${state.bodyBase64}) center/cover`
                : tokens.colorNeutralBackground3,
              borderRadius: tokens.borderRadiusMedium,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              color: state.bodyBase64 ? '#fff' : tokens.colorNeutralForeground3,
              fontWeight: 700,
              fontSize: 40,
              border: `1px solid ${tokens.colorNeutralStroke2}`,
            }}
          >
            {!state.bodyBase64 && initials}
          </div>
          {state.bodyBase64 && (
            <Caption1
              style={{ color: tokens.colorNeutralForeground3, display: 'block', marginTop: 4 }}
            >
              {state.fileName} · {state.mimeType}
            </Caption1>
          )}
        </div>

        {/* Source controls */}
        {state.operation === 'upload' && (
          <div style={{ flex: 1 }}>
            <BinarySourceCard
              source={state.source}
              setSource={setSource}
              fileName={state.fileName}
              setFileName={setFileName}
              fileSize={0}
              setFileSize={() => undefined}
              mimeType={state.mimeType}
              setMimeType={(m) => setMimeType(m as ManageImageState['mimeType'])}
              bodyBase64={state.bodyBase64}
              setBodyBase64={setBodyBase64}
              bodyUrl={state.bodyUrl}
              setBodyUrl={() => undefined}
              accept="image/png,image/jpeg,image/gif,image/bmp,image/tiff"
              group="binary"
            />
          </div>
        )}
        {state.operation !== 'upload' && (
          <div
            style={{
              flex: 1,
              fontSize: 11,
              color: tokens.colorNeutralForeground3,
              lineHeight: 1.6,
            }}
          >
            {state.operation === 'download' && (
              <>
                <strong style={{ color: tokens.colorNeutralForeground2 }}>Download:</strong> server
                returns the{' '}
                {state.downloadSize === 'full'
                  ? 'original full-size'
                  : 'auto-cropped 144×144 thumbnail'}{' '}
                as raw bytes. The Content-Type matches the stored image's MIME type (one of
                image/png · image/jpeg · image/gif · image/bmp · image/tiff).
              </>
            )}
            {state.operation === 'delete' && (
              <>
                <strong style={{ color: tokens.colorNeutralForeground2 }}>Delete:</strong> clears
                the image data + companion columns (
                <code style={{ fontFamily: tokens.fontFamilyMonospace }}>
                  {imageCol?.logicalName}id
                </code>
                , <code style={{ fontFamily: tokens.fontFamilyMonospace }}>_Timestamp</code>,{' '}
                <code style={{ fontFamily: tokens.fontFamilyMonospace }}>_URL</code>). Equivalent
                forms: PATCH set null, PUT /col with body <code>{'{"value":null}'}</code>, or DELETE
                /col (shown).
              </>
            )}
          </div>
        )}
      </div>
    </div>
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
      sub: 'PATCH /<set>(<id>) with body {col: <base64>}.',
      tone: 'brand',
    },
    {
      id: 'download',
      label: 'Download',
      sub: 'GET /<set>(<id>)/<col>/$value (?size=full for full).',
      tone: 'brand',
    },
    {
      id: 'delete',
      label: 'Delete',
      sub: 'DELETE / PATCH null / PUT /col body=null — three equivalents.',
      tone: 'danger',
    },
  ];
  return (
    <div>
      <PaneHead icon={ArrowSwap20Filled} title="Operation" group="binary" />
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
              <strong style={{ fontSize: 13, display: 'block' }}>{c.label}</strong>
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
// Download options — thumbnail vs full
// ──────────────────────────────────────────────────────────────
function DownloadOptionsPane({
  downloadSize,
  setDownloadSize,
  canStoreFullImage,
}: {
  downloadSize: 'thumbnail' | 'full';
  setDownloadSize: (s: 'thumbnail' | 'full') => void;
  canStoreFullImage: boolean;
}) {
  return (
    <div>
      <PaneHead
        icon={Settings20Filled}
        title="Download options"
        sub="Image columns can store both thumbnail + full-size when CanStoreFullImage is true."
        group="binary"
      />
      <div style={{ maxWidth: 720 }}>
        <Field label="Size">
          <RadioGroup
            value={downloadSize}
            onChange={(_, d) => setDownloadSize(d.value as 'thumbnail' | 'full')}
          >
            <Radio
              value="thumbnail"
              label={
                <span>
                  <strong>Thumbnail (144×144)</strong>
                  <Caption1
                    style={{
                      display: 'block',
                      color: tokens.colorNeutralForeground3,
                      marginTop: 2,
                    }}
                  >
                    Default for all image columns. Center-cropped square.
                  </Caption1>
                </span>
              }
            />
            <Radio
              value="full"
              disabled={!canStoreFullImage}
              label={
                <span>
                  <strong>Full-size</strong>{' '}
                  <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>
                    ?size=full
                  </code>
                  <Caption1
                    style={{
                      display: 'block',
                      color: tokens.colorNeutralForeground3,
                      marginTop: 2,
                    }}
                  >
                    Only available when the column has CanStoreFullImage = true. Server returns 204
                    if not configured.
                    {!canStoreFullImage && (
                      <>
                        {' '}
                        ·{' '}
                        <span style={{ color: tokens.colorPaletteRedForeground1 }}>
                          this column doesn't support full-size
                        </span>
                      </>
                    )}
                  </Caption1>
                </span>
              }
            />
          </RadioGroup>
        </Field>
      </div>
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Operation banner
// ──────────────────────────────────────────────────────────────
function OperationBanner({
  state,
  imageCol,
}: {
  state: ManageImageState;
  imageCol: ImageColumnMeta | undefined;
}) {
  if (!imageCol) {
    return (
      <MessageBar layout="multiline" intent="info">
        <MessageBarBody>
          Pick an image column on the left to see operation-specific guidance.
        </MessageBarBody>
      </MessageBar>
    );
  }
  if (state.operation === 'upload') {
    return (
      <MessageBar layout="multiline" intent="info">
        <MessageBarBody>
          <MessageBarTitle>
            {imageCol.isPrimaryImage ? 'Primary image — thumbnail only' : 'Custom image column'}
          </MessageBarTitle>
          {imageCol.isPrimaryImage ? (
            <>
              <code>{imageCol.logicalName}</code> is the table's primary image (auto-cropped
              144×144). Full-size data is discarded. To store full-resolution images, create a
              custom image column with <code>CanStoreFullImage = true</code>.
            </>
          ) : imageCol.canStoreFullImage ? (
            <>
              <code>{imageCol.logicalName}</code> supports full-size (max{' '}
              {((imageCol.maxSizeInKB ?? 10240) / 1024).toFixed(0)} MB). Upload via PATCH; for files
              &gt; 4 MB use the file-column chunked upload path on the underlying file storage.
            </>
          ) : (
            <>
              <code>{imageCol.logicalName}</code> stores thumbnail only (CanStoreFullImage = false).
              Full-size won't be retained.
            </>
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
          GET /$value returns the{' '}
          {state.downloadSize === 'full' ? 'full-size original' : '144×144 thumbnail'}. Image
          companion columns (<code>{imageCol.logicalName}_URL</code> etc.) let you compose a
          download URL for browser-side use.
        </MessageBarBody>
      </MessageBar>
    );
  }
  return (
    <MessageBar layout="multiline" intent="warning">
      <MessageBarBody>
        <MessageBarTitle>Delete</MessageBarTitle>
        Clears <code>{imageCol.logicalName}</code> + its companion columns (id / _Timestamp / _URL).
        Three equivalent paths per docs: DELETE /col (shown), PATCH set null, or PUT /col with body{' '}
        <code>{'{"value":null}'}</code>.
      </MessageBarBody>
    </MessageBar>
  );
}
