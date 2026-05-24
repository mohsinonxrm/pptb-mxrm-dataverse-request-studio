// Manage Attachment / Note — ActivityMimeAttachment.body and Annotation.documentbody.
//
// Per attachment-annotation-files doc the wire shape is:
//   Upload:
//     (a) inline base64 — create/update record with body column populated
//     (b) Initialize{Attachment|Annotation}BlocksUpload → UploadBlock × N → Commit*
//   Download:
//     (a) GET /<set>(<id>)/{body|documentbody}/$value (text/plain base64)
//     (b) Initialize{Attachment|Annotation}BlocksDownload → DownloadBlock × N
//   Delete:
//     DELETE /<set>(<id>) (removes the entire record — body lives on the record).
//
// Differences from ManageFileMode:
//   • No ranged downloads, no SAS URL — those are file-column-only.
//   • No "file column" picker — the body column is fixed by the target type.
//   • Attachment requires `objectid_email@odata.bind` to the parent email
//     activity (caller-supplied parentActivityId).
//   • Annotation requires a caller-generated GUID for `annotationid` on
//     Initialize/Commit (the server will NOT auto-assign for these messages).
//
// Pre-flight cap is Organization.MaxUploadFileSize (default 5 MB; max 128 MB).
// We surface the default cap when no live value is known — the advisory has a
// "size limit unknown" path for that case.

import { useMemo, useState } from 'react';
import {
  Table20Regular, Table20Filled,
  Document20Regular, Document20Filled,
  Settings20Regular, Settings20Filled,
  LineHorizontal320Regular, LineHorizontal320Filled,
  ArrowSwap20Regular, ArrowSwap20Filled,
  Code20Regular, Code20Filled,
  ArrowUpload20Filled, ArrowDownload20Filled, Delete20Filled,
  Checkmark16Filled, Attach20Regular,
} from '@fluentui/react-icons';
import {
  Field, RadioGroup, Radio, Caption1, Input, Textarea, tokens, mergeClasses, SpinButton,
} from '@fluentui/react-components';
import { Sidebar } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { useStudioStyles } from '../primitives/styles';
import { PaneHead } from '../editors/PaneHead';
import { BinarySourceCard } from '../editors/BinarySourceCard';
import { BinaryPipelineCard } from '../editors/BinaryPipelineCard';
import { HeadersEditor, defaultWriteHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { findRequestType } from '../registry/requestTypes';
import {
  buildManageAttachment, manageAttachmentPipeline, attachmentTargetLabel, formatSize,
} from '../engine/binaryBuilders';
import { runtime, type ExecResult } from '../engine/runtime';
import {
  ATTACHMENT_TARGET_INFO, checkUploadSize,
  type ManageAttachmentState, type AttachmentTarget, type FileOperation,
  type AttachmentUploadMethod, type AttachmentDownloadMethod,
} from '../state/binaryState';
import type { RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useScopedEntities } from '../host/useScopedEntities';
import { isEmbedded } from '../host/pptbBridge';
import { adv, type Advisory } from '../primitives/advisories';
import {
  serializeManageAttachment, deserializeManageAttachment, hashState,
  type SavedRequest, type SerializedManageAttachmentState,
} from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';

// Default cap = Organization.MaxUploadFileSize default per doc (5 MB).
const DEFAULT_MAX_UPLOAD_BYTES = 5 * 1024 * 1024;

const initialState = (): ManageAttachmentState => ({
  target: 'attachment',
  recordId: null,
  parentActivityId: null,
  parentEntityLogical: null,
  parentEntityId: null,
  operation: 'upload',
  uploadMethod: { kind: 'inline-base64' },
  downloadMethod: { kind: 'single-request' },
  source: { kind: 'file' },
  fileName: '',
  fileSize: 0,
  mimeType: 'application/octet-stream',
  bodyBase64: '',
  bodyUrl: '',
  subject: '',
  noteText: '',
  headers: defaultWriteHeaders(),
  dirty: new Set(),
});

type RootClauseId = 'operation' | 'target' | 'parent' | 'method' | 'source' | 'pipeline' | 'headers';

export function ManageAttachmentMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('manage-attachment');
  const [state, setState] = useState(initialState);
  const [activePath, setActivePath] = useState<string>('pipeline');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);

  const built = useMemo(() => buildManageAttachment(state), [state]);
  const pipeline = useMemo(() => manageAttachmentPipeline(state), [state]);
  const info = ATTACHMENT_TARGET_INFO[state.target];

  const markDirty = (id: string) => setState(s => { const d = new Set(s.dirty); d.add(id); return { ...s, dirty: d }; });
  const set = <K extends keyof ManageAttachmentState>(k: K, v: ManageAttachmentState[K], dirtyId?: string) => {
    setState(s => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  const firstStepMethod = pipeline[0]?.method ?? 'POST';

  const disabledReason =
    state.operation === 'upload' && state.target === 'attachment' && !state.parentActivityId
      ? 'Pick the parent email activity for the attachment.' :
    state.operation === 'upload' && state.target === 'annotation' && !state.recordId
      ? 'Provide a caller-generated annotationid (Annotations require an explicit GUID).' :
    state.operation !== 'upload' && !state.recordId
      ? `Pick the existing ${state.target} record.` :
    state.operation === 'upload' && state.source.kind === 'file' && !state.bodyBase64 && !state.fileName
      ? 'Choose a file to upload.' :
    state.headers.some(h => h.enabled && !h.name)
      ? 'Fix empty header name.' :
    null;

  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.manageAttachment(state);
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents(rs => [{
      id: `r-${Date.now()}`, modeId: 'manage-attachment',
      url: built.relativeUrl, method: firstStepMethod, ts: Date.now(),
      status: res.status, ms: res.ms, rowCount: res.ok ? 1 : 0,
    }, ...rs].slice(0, 8));
    setState(s => ({ ...s, dirty: new Set() }));
  };

  // ── Save / Load ──
  const { entities } = useScopedEntities();
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);
  const currentSerialized = useMemo(() => serializeManageAttachment(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;
  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };
  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== 'manage-attachment') return;
    const snap = entry.state as SerializedManageAttachmentState;
    // We don't gate by entity availability — both 'activitymimeattachment' and
    // 'annotation' are system tables present in every Dataverse environment.
    void entities; // intentionally unused — see comment above
    setState(deserializeManageAttachment(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setActivePath('pipeline');
  };
  usePublishSaveContext(useMemo(() => ({
    state: currentSerialized,
    modeId: 'manage-attachment' as const,
    dirty: isDirty,
    lastSavedId,
    onSaved,
    onLoadSaved,
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }), [currentSerialized, isDirty, lastSavedId]));

  // ── PPTB host-limitation advisory ──
  const advisories = useMemo<Advisory[]>(() => {
    if (!isEmbedded()) return [];
    return [adv.warn(
      'binary-host-unsupported',
      'header',
      'Attachment / Note operations run externally',
      <span>
        PPTB's <code>dataverseAPI</code> doesn't expose the attachment / annotation
        message endpoints. Build the request here, then copy a snippet from the{' '}
        <strong>Code</strong> tab (fetch / curl / C# / PowerShell) and run it
        from there. Execute will return <code>501 Not Implemented</code>.
      </span>,
      'pipeline',
      'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/attachment-annotation-files',
    )];
  }, []);

  const opBadge = state.operation === 'upload' ? 'UPLOAD'
                : state.operation === 'download' ? 'DOWNLOAD'
                : 'DELETE';
  const opBadgeColor = state.operation === 'delete' ? ('danger' as const) : ('brand' as const);
  const methodLabel = state.operation === 'upload'
    ? state.uploadMethod.kind
    : state.operation === 'download'
      ? state.downloadMethod.kind
      : 'simple';

  const sections = [
    {
      id: 'operation', label: 'Operation',
      meta: opBadge.toLowerCase(),
      items: [{
        id: 'operation',
        icon: ArrowSwap20Regular, iconFilled: ArrowSwap20Filled,
        label: opBadge.charAt(0) + opBadge.slice(1).toLowerCase(),
        badge: opBadge, badgeAppearance: 'tint' as const, badgeColor: opBadgeColor,
        dirty: state.dirty.has('operation'),
      }],
    },
    {
      id: 'target', label: 'Target',
      meta: state.target,
      items: [{
        id: 'target',
        icon: Table20Regular, iconFilled: Table20Filled,
        label: attachmentTargetLabel(state.target),
        badge: state.target, badgeAppearance: 'ghost' as const,
        dirty: state.dirty.has('target'),
      }],
    },
    {
      id: 'parent', label: state.target === 'attachment' ? 'Parent email' : 'Annotation id',
      meta: state.target === 'attachment'
        ? (state.parentActivityId ? state.parentActivityId.slice(0, 8) : 'unset')
        : (state.recordId ? state.recordId.slice(0, 8) : 'unset'),
      items: [{
        id: 'parent',
        icon: Document20Regular, iconFilled: Document20Filled,
        label: state.target === 'attachment' ? 'Email + record reference' : 'Annotation id + optional parent',
        badge: (state.target === 'attachment' ? !!state.parentActivityId : !!state.recordId) ? '✓' : null,
        badgeAppearance: 'ghost' as const,
        dirty: state.dirty.has('parent'),
      }],
    },
    ...(state.operation !== 'delete' ? [{
      id: 'method', label: 'Method',
      meta: methodLabel,
      items: [{
        id: 'method',
        icon: Settings20Regular, iconFilled: Settings20Filled,
        label: state.operation === 'upload' ? 'Upload method' : 'Download method',
        badge: methodLabel, badgeAppearance: 'ghost' as const,
        dirty: state.dirty.has('method'),
      }],
    }] : []),
    ...(state.operation === 'upload' ? [{
      id: 'payload', label: 'Payload',
      meta: state.fileName || '—',
      items: [{
        id: 'source',
        icon: Document20Regular, iconFilled: Document20Filled,
        label: 'Source',
        badge: state.source.kind, badgeAppearance: 'ghost' as const,
        dirty: state.dirty.has('source'),
      }],
    }] : []),
    {
      id: 'pipe', label: 'Pipeline',
      meta: `${pipeline.length} request${pipeline.length === 1 ? '' : 's'}`,
      items: [{
        id: 'pipeline',
        icon: Code20Regular, iconFilled: Code20Filled,
        label: 'Generated requests',
        badge: pipeline.length || null,
        badgeAppearance: 'tint' as const, badgeColor: 'brand' as const,
        dirty: state.dirty.has('pipeline'),
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
        <TargetPicker
          value={state.target}
          onChange={(t) => {
            // Switching target invalidates the record + parent refs (the
            // identifiers are scoped to one entity type).
            setState(s => ({
              ...s,
              target: t,
              recordId: null,
              parentActivityId: null,
              parentEntityLogical: null,
              parentEntityId: null,
              dirty: new Set(['target', 'parent']),
            }));
            setResult(null);
          }}
        />
      );
      break;
    case 'parent':
      pane = (
        <ParentEditor
          target={state.target}
          recordId={state.recordId}
          parentActivityId={state.parentActivityId}
          parentEntityLogical={state.parentEntityLogical}
          parentEntityId={state.parentEntityId}
          subject={state.subject}
          noteText={state.noteText}
          onRecordIdChange={(v) => set('recordId', v, 'parent')}
          onParentActivityIdChange={(v) => set('parentActivityId', v, 'parent')}
          onParentEntityLogicalChange={(v) => set('parentEntityLogical', v, 'parent')}
          onParentEntityIdChange={(v) => set('parentEntityId', v, 'parent')}
          onSubjectChange={(v) => set('subject', v, 'parent')}
          onNoteTextChange={(v) => set('noteText', v, 'parent')}
        />
      );
      break;
    case 'method':
      pane = state.operation === 'upload'
        ? <UploadMethodPicker value={state.uploadMethod} onChange={(m) => set('uploadMethod', m, 'method')} />
        : <DownloadMethodPicker value={state.downloadMethod} onChange={(m) => set('downloadMethod', m, 'method')} />;
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
          <BinaryPipelineCard
            steps={pipeline}
            title={`${opBadge} pipeline — ${methodLabel}`}
            sub={pipeline.length > 1 ? 'Each request fires sequentially.' : undefined}
            sizeAdvisory={state.operation === 'upload'
              ? checkUploadSize(state.fileSize, DEFAULT_MAX_UPLOAD_BYTES, state.target)
              : null}
          />
        </div>
      );
      break;
    case 'headers':
      pane = (
        <HeadersEditor
          items={state.headers}
          setItems={h => set('headers', h, 'headers')}
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
    multiRequests: pipeline.length > 1
      ? pipeline.map(step => ({
          method: step.method,
          relativeUrl: step.relativeUrl.startsWith('<') ? step.relativeUrl : step.relativeUrl.replace(/^https?:\/\/[^/]+/, ''),
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
          onSelect={(id) => setActivePath(id)}
          recents={recents}
        />
      }
      urlBar={
        <UrlBar
          method={firstStepMethod}
          url={built.relativeUrl}
          executeVerb={state.operation === 'upload' ? 'Upload' : state.operation === 'download' ? 'Download' : 'Delete'}
          executeIcon={
            state.operation === 'upload'   ? ArrowUpload20Filled :
            state.operation === 'download' ? ArrowDownload20Filled :
                                              Delete20Filled
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
        {tab === 'code'    && <CodeView themeMode={themeMode} inputs={codeInputs} />}
        {tab === 'results' && (
          <ResultsView
            result={result}
            mode="single"
            table={info.entityLogical}
            writeContext={{
              operation: 'manage-attachment',
              table: info.entityLogical,
              recordId: state.recordId,
              recordName: state.fileName || null,
              fileOperation: state.operation,
              columnName: info.bodyColumn,
              fileName: state.fileName || undefined,
            }}
          />
        )}
      </MainTabs>
    </ModeShell>
  );
}

// ──────────────────────────────────────────────────────────────
// Operation picker
// ──────────────────────────────────────────────────────────────
function OperationPicker({ value, onChange }: { value: FileOperation; onChange: (op: FileOperation) => void }) {
  const s = useStudioStyles();
  const cards: { id: FileOperation; label: string; sub: string; tone: 'brand' | 'danger' }[] = [
    { id: 'upload',   label: 'Upload',   sub: 'Inline base 64 (≲4 MB) or Initialize/UploadBlock/Commit messages.', tone: 'brand' },
    { id: 'download', label: 'Download', sub: 'GET /<set>(<id>)/{body|documentbody}/$value or Initialize/DownloadBlock messages.', tone: 'brand' },
    { id: 'delete',   label: 'Delete',   sub: 'DELETE /<set>(<id>) — removes the entire attachment / note record.', tone: 'danger' },
  ];
  return (
    <div>
      <PaneHead icon={ArrowSwap20Filled} title="Operation" sub="What to do with the attachment or annotation." group="binary" />
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 12, maxWidth: 960 }}>
        {cards.map(c => {
          const selected = value === c.id;
          return (
            <button
              key={c.id}
              type="button"
              onClick={() => onChange(c.id)}
              className={mergeClasses(s.inlineCard)}
              style={{
                textAlign: 'left', padding: 14, cursor: 'pointer',
                background: selected ? (c.tone === 'danger' ? tokens.colorPaletteRedBackground2 : tokens.colorBrandBackground2) : tokens.colorNeutralBackground1,
                border: `1px solid ${selected ? (c.tone === 'danger' ? tokens.colorPaletteRedBorderActive : tokens.colorBrandStroke1) : tokens.colorNeutralStroke2}`,
                borderRadius: tokens.borderRadiusMedium,
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <strong style={{ fontSize: 13 }}>{c.label}</strong>
                {selected && (
                  <Checkmark16Filled
                    style={{
                      marginLeft: 'auto',
                      color: c.tone === 'danger' ? tokens.colorPaletteRedForeground1 : tokens.colorBrandForeground1,
                    }}
                    aria-label="Selected"
                  />
                )}
              </div>
              <Caption1 style={{ color: tokens.colorNeutralForeground2, lineHeight: 1.4, marginTop: 6, display: 'block' }}>
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
// Target picker (attachment vs annotation)
// ──────────────────────────────────────────────────────────────
function TargetPicker({ value, onChange }: { value: AttachmentTarget; onChange: (t: AttachmentTarget) => void }) {
  return (
    <div>
      <PaneHead icon={Attach20Regular} title="Target type" sub="Pick the record type — drives the body column, message names, and parent reference shape." group="binary" />
      <div style={{ maxWidth: 760 }}>
        <Field label="Type">
          <RadioGroup value={value} onChange={(_, d) => onChange(d.value as AttachmentTarget)}>
            <Radio value="attachment" label={
              <span>
                <strong>Attachment</strong>{' '}
                <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>activitymimeattachment.body</code>
                <Caption1 style={{ display: 'block', color: tokens.colorNeutralForeground3, marginTop: 2 }}>
                  Hangs off an email activity. Upload requires <code>objectid_email@odata.bind</code>. Server generates the activitymimeattachmentid.
                </Caption1>
              </span>
            }/>
            <Radio value="annotation" label={
              <span>
                <strong>Note (Annotation)</strong>{' '}
                <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>annotation.documentbody</code>
                <Caption1 style={{ display: 'block', color: tokens.colorNeutralForeground3, marginTop: 2 }}>
                  Hangs off any HasNotes=true entity. Caller MUST generate the annotationid GUID for Initialize/Commit messages.
                </Caption1>
              </span>
            }/>
          </RadioGroup>
        </Field>
      </div>
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Parent / record-id editor
// ──────────────────────────────────────────────────────────────
interface ParentEditorProps {
  target: AttachmentTarget;
  recordId: string | null;
  parentActivityId: string | null;
  parentEntityLogical: string | null;
  parentEntityId: string | null;
  subject: string;
  noteText: string;
  onRecordIdChange: (v: string | null) => void;
  onParentActivityIdChange: (v: string | null) => void;
  onParentEntityLogicalChange: (v: string | null) => void;
  onParentEntityIdChange: (v: string | null) => void;
  onSubjectChange: (v: string) => void;
  onNoteTextChange: (v: string) => void;
}
function ParentEditor(p: ParentEditorProps) {
  return (
    <div>
      <PaneHead
        icon={Document20Filled}
        title={p.target === 'attachment' ? 'Parent email & record reference' : 'Annotation id & parent reference'}
        sub={p.target === 'attachment'
          ? 'Attachment must reference its parent email activity. The attachment id is server-assigned on create.'
          : 'Annotation messages require a caller-generated GUID. Parent reference is optional but recommended for navigation.'}
        group="binary"
      />
      <div style={{ maxWidth: 760, display: 'flex', flexDirection: 'column', gap: 14 }}>
        {p.target === 'attachment' ? (
          <>
            <Field label="Parent email activityid (GUID)" hint="Bound via objectid_email@odata.bind: emails(<id>) on the Initialize/Commit Target.">
              <Input value={p.parentActivityId ?? ''} onChange={(_, d) => p.onParentActivityIdChange(d.value || null)} placeholder="00000000-0000-0000-0000-000000000000" />
            </Field>
            <Field label="Existing activitymimeattachmentid (update/download/delete only)">
              <Input value={p.recordId ?? ''} onChange={(_, d) => p.onRecordIdChange(d.value || null)} placeholder="leave empty for create" />
            </Field>
            <Field label="Subject" hint="Optional. Stored on the attachment record alongside the binary.">
              <Input value={p.subject} onChange={(_, d) => p.onSubjectChange(d.value)} placeholder="e.g. Sample attached 25mb.pdf" />
            </Field>
          </>
        ) : (
          <>
            <Field label="Annotation id (GUID)" hint="Caller-generated. Initialize / Commit will NOT auto-assign one for these messages.">
              <Input value={p.recordId ?? ''} onChange={(_, d) => p.onRecordIdChange(d.value || null)} placeholder="00000000-0000-0000-0000-000000000000" />
            </Field>
            <Field label="Parent entity logical name (optional)" hint="Annotations can hang off any HasNotes=true entity. Leave blank for unparented notes.">
              <Input value={p.parentEntityLogical ?? ''} onChange={(_, d) => p.onParentEntityLogicalChange(d.value || null)} placeholder="account / contact / opportunity / …" />
            </Field>
            <Field label="Parent entity id (GUID)">
              <Input value={p.parentEntityId ?? ''} onChange={(_, d) => p.onParentEntityIdChange(d.value || null)} placeholder="00000000-0000-0000-0000-000000000000" />
            </Field>
            <Field label="Notetext" hint="Free-text body of the note — separate from documentbody (which carries the binary).">
              <Textarea value={p.noteText} onChange={(_, d) => p.onNoteTextChange(d.value)} placeholder="Please see attached…" />
            </Field>
          </>
        )}
      </div>
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Upload method picker
// ──────────────────────────────────────────────────────────────
function UploadMethodPicker({ value, onChange }: { value: AttachmentUploadMethod; onChange: (m: AttachmentUploadMethod) => void }) {
  return (
    <div>
      <PaneHead icon={Settings20Filled} title="Upload method" sub="Per attachment-annotation-files — two supported approaches." group="binary" />
      <div style={{ maxWidth: 760, display: 'flex', flexDirection: 'column', gap: 14 }}>
        <Field label="Approach">
          <RadioGroup
            value={value.kind}
            onChange={(_, d) => {
              if (d.value === 'inline-base64') onChange({ kind: 'inline-base64' });
              else onChange({ kind: 'dataverse-messages', blockSize: value.kind === 'dataverse-messages' ? value.blockSize : 4 * 1024 * 1024 });
            }}
          >
            <Radio value="inline-base64" label={
              <span>
                <strong>Inline base 64</strong>{' '}
                <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>POST / PATCH with body column set</code>
                <Caption1 style={{ display: 'block', color: tokens.colorNeutralForeground3, marginTop: 2 }}>
                  Works for files ≲ 4 MB (Organization.MaxUploadFileSize default 5 MB). One round-trip. Body column = base 64 string in the JSON body.
                </Caption1>
              </span>
            }/>
            <Radio value="dataverse-messages" label={
              <span>
                <strong>Dataverse messages</strong>{' '}
                <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>Initialize* / UploadBlock / Commit*</code>
                <Caption1 style={{ display: 'block', color: tokens.colorNeutralForeground3, marginTop: 2 }}>
                  Three OData actions. Up to 128 MB. UploadBlock is shared with file columns. Commit returns FileSizeInBytes + record id.
                </Caption1>
              </span>
            }/>
          </RadioGroup>
        </Field>

        {value.kind === 'dataverse-messages' && (
          <Field label="Block size (bytes)" hint="Defaults to 4 MB. BlockData must be ≤ 4 MB before base 64 encoding.">
            <SpinButton
              value={value.blockSize}
              min={64 * 1024}
              max={4 * 1024 * 1024}
              step={1024 * 1024}
              onChange={(_, d) => onChange({ kind: 'dataverse-messages', blockSize: Number(d.value ?? d.displayValue ?? value.blockSize) })}
            />
            <Caption1 style={{ marginTop: 4, display: 'block', color: tokens.colorNeutralForeground3 }}>
              ≈ {formatSize(value.blockSize)} per block.
            </Caption1>
          </Field>
        )}
      </div>
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Download method picker
// ──────────────────────────────────────────────────────────────
function DownloadMethodPicker({ value, onChange }: { value: AttachmentDownloadMethod; onChange: (m: AttachmentDownloadMethod) => void }) {
  return (
    <div>
      <PaneHead icon={Settings20Filled} title="Download method" sub="Per attachment-annotation-files — two supported approaches. No ranged download / SAS URL (those are file-column-only)." group="binary" />
      <div style={{ maxWidth: 760, display: 'flex', flexDirection: 'column', gap: 14 }}>
        <Field label="Approach">
          <RadioGroup
            value={value.kind}
            onChange={(_, d) => {
              if (d.value === 'single-request') onChange({ kind: 'single-request' });
              else onChange({ kind: 'dataverse-messages', blockSize: value.kind === 'dataverse-messages' ? value.blockSize : 4 * 1024 * 1024 });
            }}
          >
            <Radio value="single-request" label={
              <span>
                <strong>Single request</strong>{' '}
                <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>GET …/(body|documentbody)/$value</code>
                <Caption1 style={{ display: 'block', color: tokens.colorNeutralForeground3, marginTop: 2 }}>
                  Returns the binary as base 64 text/plain. No size / filename / mimetype headers — query those via $select on filename / mimetype.
                </Caption1>
              </span>
            }/>
            <Radio value="dataverse-messages" label={
              <span>
                <strong>Dataverse messages</strong>{' '}
                <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>Initialize* / DownloadBlock</code>
                <Caption1 style={{ display: 'block', color: tokens.colorNeutralForeground3, marginTop: 2 }}>
                  Returns FileName + FileSizeInBytes + FileContinuationToken on Initialize. Then loop DownloadBlock with Offset stepping by BlockLength.
                </Caption1>
              </span>
            }/>
          </RadioGroup>
        </Field>

        {value.kind === 'dataverse-messages' && (
          <Field label="Block size (bytes)" hint="Defaults to 4 MB. Can stay constant — final block may be smaller and only those bytes are returned.">
            <SpinButton
              value={value.blockSize}
              min={64 * 1024}
              max={4 * 1024 * 1024}
              step={1024 * 1024}
              onChange={(_, d) => onChange({ kind: 'dataverse-messages', blockSize: Number(d.value ?? d.displayValue ?? value.blockSize) })}
            />
            <Caption1 style={{ marginTop: 4, display: 'block', color: tokens.colorNeutralForeground3 }}>
              ≈ {formatSize(value.blockSize)} per block.
            </Caption1>
          </Field>
        )}
      </div>
    </div>
  );
}
