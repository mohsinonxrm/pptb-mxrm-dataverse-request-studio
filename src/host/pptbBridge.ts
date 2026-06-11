// Power Platform ToolBox (PPTB) integration bridge.
//
// PPTB exposes two preload-injected globals (typed via `@pptb/types`):
//
//   window.toolboxAPI    — host services (theme, active connection, event bus,
//                          file system, settings, terminal, events, utils,
//                          notifications)
//   window.dataverseAPI  — Dataverse operations (queryData, fetchXmlQuery,
//                          create/retrieve/update/delete, execute,
//                          associate/disassociate, metadata family, solution
//                          deploy, batch create/update)
//
// Detection: `@pptb/types` declares these globals as non-optional on `Window`,
// so we use `typeof === 'undefined'` (not `!!`) because the latter is always-
// true under the official types — but the runtime value is genuinely absent
// when the studio runs standalone in a regular browser tab.
//
// This bridge owns:
//   - `isEmbedded()`         — whether we have a PPTB host
//   - `getSession()`         — current { embedded, connected, theme, env… }
//   - `subscribe(listener)`  — fires on initial mount + every host event
//
// Nothing else in the codebase touches `window.toolboxAPI` / `window.dataverseAPI`
// directly — go through this module.

// Pull in the global `Window.toolboxAPI` / `Window.dataverseAPI` augmentations
// from @pptb/types. We use a triple-slash type reference (NOT a side-effect
// `import '@pptb/types'`) because the package ships .d.ts only — Vite's bundler
// would otherwise try to resolve `./dataverseAPI` as a runtime module and fail
// at build time. The reference triggers the type augmentation at compile time
// without emitting any runtime import.
/// <reference types="@pptb/types" />

export interface HostSession {
  /** True when running inside PPTB (window.toolboxAPI is present). */
  embedded: boolean;
  /** True when the host has an active Dataverse connection. */
  connected: boolean;
  /** Current theme — pushed by host via getCurrentTheme() + settings:updated event. */
  theme: 'light' | 'dark';
  /** The active environment URL (e.g. "https://contoso.crm.dynamics.com"). Empty in standalone. */
  environmentUrl: string;
  /** Environment label from the host — "Dev" / "Test" / "UAT" / "Production" or empty. */
  environment: string;
}

type Listener = (session: HostSession) => void;

// ──────────────────────────────────────────────────────────────
// Detection
// ──────────────────────────────────────────────────────────────
function detectEmbedded(): boolean {
  return typeof window !== 'undefined' && !!window.toolboxAPI;
}
const EMBEDDED = detectEmbedded();

// ──────────────────────────────────────────────────────────────
// Mutable session + subscribers
// ──────────────────────────────────────────────────────────────
let session: HostSession = {
  embedded: EMBEDDED,
  connected: false,
  theme: readBootTheme(),
  environmentUrl: '',
  environment: '',
};

function readBootTheme(): 'light' | 'dark' {
  // Used both standalone (where index.html may set ?theme=dark for iframe demos)
  // and during the very first paint inside PPTB before getCurrentTheme() resolves.
  if (typeof window === 'undefined') return 'light';
  const p = new URLSearchParams(window.location.search).get('theme');
  return p === 'dark' ? 'dark' : 'light';
}

const listeners = new Set<Listener>();

function emit(next: HostSession) {
  session = next;
  for (const l of listeners) {
    try {
      l(session);
    } catch (e) {
      console.error('[pptbBridge] listener threw', e);
    }
  }
}

// ──────────────────────────────────────────────────────────────
// Initial fetch + event subscription (only in embedded mode)
// ──────────────────────────────────────────────────────────────
async function fetchInitialTheme(
  api: NonNullable<Window['toolboxAPI']>,
): Promise<'light' | 'dark'> {
  try {
    const t = await api.utils?.getCurrentTheme?.();
    return String(t ?? 'light').toLowerCase() === 'dark' ? 'dark' : 'light';
  } catch (e) {
    console.warn('[pptbBridge] getCurrentTheme failed', e);
    return 'light';
  }
}

async function fetchInitialConnection(api: NonNullable<Window['toolboxAPI']>): Promise<{
  connected: boolean;
  environmentUrl: string;
  environment: string;
}> {
  try {
    const c = await api.connections?.getActiveConnection?.();
    if (!c) return { connected: false, environmentUrl: '', environment: '' };
    return { connected: true, environmentUrl: c.url ?? '', environment: c.environment ?? '' };
  } catch (e) {
    console.warn('[pptbBridge] getActiveConnection failed', e);
    return { connected: false, environmentUrl: '', environment: '' };
  }
}

function hookEvents(api: NonNullable<Window['toolboxAPI']>): void {
  if (!api.events?.on) {
    console.warn('[pptbBridge] toolboxAPI.events.on missing — running without live updates');
    return;
  }

  // Per @pptb/types: events.on takes `(event: any, payload: ToolBoxEventPayload) => void`.
  // The first arg is the raw IPC event we don't read; the second is the typed payload.
  const handler = (_e: unknown, payload: ToolBoxAPI.ToolBoxEventPayload) => {
    if (!payload || typeof payload.event !== 'string') return;
    switch (payload.event) {
      case 'settings:updated': {
        // The settings panel changed something. Theme is the most common —
        // payload.data.theme is set when theme changed. Otherwise refetch.
        if (payload.data && typeof payload.data === 'object' && 'theme' in payload.data) {
          const t = String((payload.data as { theme: string }).theme ?? '').toLowerCase();
          emit({ ...session, theme: t === 'dark' ? 'dark' : 'light' });
        } else {
          void refetchTheme(api);
        }
        break;
      }
      case 'connection:updated':
      case 'connection:created':
        void refetchConnection(api);
        break;
      case 'connection:deleted':
        emit({ ...session, connected: false, environmentUrl: '', environment: '' });
        break;
    }
  };

  api.events.on(handler);
}

async function refetchTheme(api: NonNullable<Window['toolboxAPI']>): Promise<void> {
  const theme = await fetchInitialTheme(api);
  emit({ ...session, theme });
}
async function refetchConnection(api: NonNullable<Window['toolboxAPI']>): Promise<void> {
  const conn = await fetchInitialConnection(api);
  emit({ ...session, ...conn });
}

// Kick off the handshake on module load.
if (EMBEDDED && typeof window !== 'undefined' && window.toolboxAPI) {
  const api = window.toolboxAPI;
  // Theme first (controls boot paint), then connection.
  void (async () => {
    const theme = await fetchInitialTheme(api);
    const conn = await fetchInitialConnection(api);
    emit({ ...session, theme, ...conn });
    hookEvents(api);
  })();
}

// ──────────────────────────────────────────────────────────────
// Public API
// ──────────────────────────────────────────────────────────────
export const isEmbedded = (): boolean => EMBEDDED;
export const getSession = (): HostSession => session;

export function subscribe(listener: Listener): () => void {
  listeners.add(listener);
  // Fire immediately so React effects get current state without a race.
  listener(session);
  return () => listeners.delete(listener);
}

/**
 * Run an OData query against the active connection. Resolves to the
 * Dataverse response envelope (`{ value, '@odata...': … }`). Throws when
 * PPTB isn't embedded — callers must check `isEmbedded()` first OR fall
 * back to the mock executor.
 */
// Return type from @pptb/types' queryData. We name it locally so callers can
// use it without importing the namespace.
export type DataverseQueryResponse = { value: Record<string, unknown>[] };

export async function dataverseQuery(odataQuery: string): Promise<DataverseQueryResponse> {
  if (!EMBEDDED || typeof window === 'undefined' || !window.dataverseAPI) {
    throw new Error('PPTB Dataverse API not available — running standalone.');
  }
  return window.dataverseAPI.queryData(odataQuery);
}

// ════════════════════════════════════════════════════════════════════════
// Full host surface — typed wrappers over window.dataverseAPI / window.toolboxAPI
// ════════════════════════════════════════════════════════════════════════
//
// Every consumer in the studio goes through these helpers. They:
//   - throw `HostNotAvailableError` when called outside PPTB
//   - preserve the exact dataverseAPI signature (no re-arranging args)
//   - never inject custom headers (the host strips most of them anyway;
//     studio-composed headers like `MSCRM.*` bypass and `Prefer` are
//     surfaced via the Code tab for external execution)

export class HostNotAvailableError extends Error {
  constructor(method: string) {
    super(`PPTB host method '${method}' is not available — running standalone.`);
    this.name = 'HostNotAvailableError';
  }
}

function dv(): DataverseAPI.API {
  if (!EMBEDDED || typeof window === 'undefined' || !window.dataverseAPI) {
    throw new HostNotAvailableError('window.dataverseAPI');
  }
  return window.dataverseAPI;
}

function tb(): ToolBoxAPI.API {
  if (!EMBEDDED || typeof window === 'undefined' || !window.toolboxAPI) {
    throw new HostNotAvailableError('window.toolboxAPI');
  }
  return window.toolboxAPI;
}

// ── Dataverse host surface ───────────────────────────────────────────────
export const dvHost = {
  // CRUD
  create: (entity: string, body: Record<string, unknown>) => dv().create(entity, body),
  retrieve: (entity: string, id: string, columns?: string[]) => dv().retrieve(entity, id, columns),
  update: (entity: string, id: string, body: Record<string, unknown>) =>
    dv().update(entity, id, body),
  delete: (entity: string, id: string) => dv().delete(entity, id),

  // Batch
  createMultiple: (entity: string, records: Record<string, unknown>[]) =>
    dv().createMultiple(entity, records),
  updateMultiple: (entity: string, records: Record<string, unknown>[]) =>
    dv().updateMultiple(entity, records),

  // Queries
  queryData: (odata: string) => dv().queryData(odata),
  fetchXmlQuery: (xml: string) => dv().fetchXmlQuery(xml),

  // Actions / Functions
  execute: (req: {
    operationName: string;
    operationType: 'action' | 'function';
    entityName?: string;
    entityId?: string;
    parameters?: Record<string, unknown>;
  }) => dv().execute(req),

  // Relationships
  associate: (primary: string, id: string, rel: string, target: string, targetId: string) =>
    dv().associate(primary, id, rel, target, targetId),
  disassociate: (primary: string, id: string, rel: string, targetId: string) =>
    dv().disassociate(primary, id, rel, targetId),

  // ── Metadata ────────────────────────────────────────────────────────
  // We expose three flavors:
  //   getEntity        — full entity metadata (used sparingly; can be heavy)
  //   getAllEntities   — list-only; for entity pickers
  //   getRelatedMeta   — type-safe path-based reader (Attributes / Keys / Relationships)
  //   queryData(meta)  — escape hatch for type-cast OData metadata queries
  //                      (e.g. PicklistAttributeMetadata + $expand=OptionSet)
  metadata: {
    getEntity: (logicalName: string, properties?: string[]) =>
      dv().getEntityMetadata(logicalName, true, properties),
    getAllEntities: (properties?: string[]) => dv().getAllEntitiesMetadata(properties),
    /**
     * Strongly-typed related metadata reader. The return type narrows to
     * either a single record or a collection based on the path shape per
     * @pptb/types' EntityRelatedMetadataResponse<P> type.
     */
    getRelated: <P extends DataverseAPI.EntityRelatedMetadataPath>(
      logicalName: string,
      path: P,
      properties?: string[],
    ) => dv().getEntityRelatedMetadata(logicalName, path, properties),
    /**
     * For type-cast OData queries that need $expand or $select-with-cast
     * (e.g. PicklistAttributeMetadata + OptionSet). Use sparingly — most
     * needs are covered by getRelated.
     */
    queryRaw: (odata: string) => dv().queryData(odata),
    getSolutions: (cols: string[]) => dv().getSolutions(cols),
    getEntitySetName: (logicalName: string) => dv().getEntitySetName(logicalName),
    /**
     * The full `$metadata` CSDL/EDMX document as raw XML. ~1-5 MB; cache it.
     * Used by `csdlProvider` to enumerate live actions/functions/custom APIs
     * with their typed parameter lists.
     */
    getCSDLDocument: () => dv().getCSDLDocument(),
  },
};

// ── ToolBox host surface ─────────────────────────────────────────────────
// File system, clipboard, notifications, persistent settings, terminal.
// Most callers only need fileSystem.saveFile + utils.copyToClipboard +
// settings.{get,set}.
export const tbHost = {
  fileSystem: {
    /** Native OS save dialog. `content` is `string | Buffer | Uint8Array`. */
    saveFile: (
      defaultPath: string,
      content: string | Uint8Array,
      filters?: ToolBoxAPI.FileDialogFilter[],
    ) => tb().fileSystem.saveFile(defaultPath, content, filters),
    readText: (path: string) => tb().fileSystem.readText(path),
    readBinary: (path: string) => tb().fileSystem.readBinary(path),
    exists: (path: string) => tb().fileSystem.exists(path),
    selectPath: (opts?: ToolBoxAPI.SelectPathOptions) => tb().fileSystem.selectPath(opts),
  },
  utils: {
    copyToClipboard: (text: string) => tb().utils.copyToClipboard(text),
    showNotification: (opts: ToolBoxAPI.NotificationOptions) => tb().utils.showNotification(opts),
  },
  settings: {
    /** Strongly-typed get. T is the caller's expected shape; falls back to undefined. */
    get: <T>(key: string) => tb().settings.get(key) as Promise<T | undefined>,
    set: <T>(key: string, value: T) => tb().settings.set(key, value),
    getAll: <T extends Record<string, unknown>>() => tb().settings.getAll() as Promise<T>,
  },
};

// ── Standalone-safe helpers ──────────────────────────────────────────────
// These work in both modes — when not embedded they fall back to browser
// equivalents. Callers should prefer these over the raw `tbHost.*` accessors
// when there's a sensible fallback.

/**
 * Copy text to the clipboard. Uses the host's clipboard API inside PPTB
 * (which goes through Electron's IPC and avoids the focus requirement that
 * navigator.clipboard sometimes hits in iframes), falls back to the browser
 * clipboard when standalone.
 */
export async function copyToClipboardSafe(text: string): Promise<void> {
  if (EMBEDDED && typeof window !== 'undefined' && window.toolboxAPI) {
    try {
      await window.toolboxAPI.utils.copyToClipboard(text);
      return;
    } catch (e) {
      console.warn(
        '[pptbBridge] toolboxAPI.utils.copyToClipboard failed, falling back to navigator.clipboard',
        e,
      );
    }
  }
  if (typeof navigator !== 'undefined' && navigator.clipboard) {
    await navigator.clipboard.writeText(text);
  }
}

/**
 * Save bytes to a file. Uses the host's native save dialog inside PPTB,
 * falls back to a Blob + anchor download when standalone. Returns the
 * resolved path on PPTB, or `null` when the browser-fallback download is
 * triggered (the browser doesn't expose the chosen path).
 */
export async function saveFileSafe(
  defaultPath: string,
  content: string | Uint8Array,
  filters?: ToolBoxAPI.FileDialogFilter[],
): Promise<string | null> {
  if (EMBEDDED && typeof window !== 'undefined' && window.toolboxAPI) {
    try {
      return await window.toolboxAPI.fileSystem.saveFile(defaultPath, content, filters);
    } catch (e) {
      console.warn(
        '[pptbBridge] toolboxAPI.fileSystem.saveFile failed, falling back to browser download',
        e,
      );
    }
  }
  // Browser fallback — Blob + anchor download
  if (typeof document !== 'undefined') {
    const mimeFromExt = (path: string): string => {
      const ext = path.toLowerCase().split('.').pop() ?? '';
      switch (ext) {
        case 'xlsx':
          return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        case 'csv':
          return 'text/csv';
        case 'json':
          return 'application/json';
        case 'xml':
          return 'application/xml';
        case 'txt':
          return 'text/plain';
        default:
          return 'application/octet-stream';
      }
    };
    // Cast Uint8Array to BlobPart — BlobPart's typing is overly narrow in
    // some TS lib versions but accepts ArrayBufferView at runtime.
    const blob = new Blob([content as BlobPart], { type: mimeFromExt(defaultPath) });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = defaultPath;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }
  return null;
}
