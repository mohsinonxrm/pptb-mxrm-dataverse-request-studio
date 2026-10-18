// Saved request → collection item (#27).
//
// Rebuilds each saved entry's HTTP request(s) with the same builders the
// modes use (URL, body) and composes the headers the way each mode does at
// execute time: user headers + Prefer + concurrency (If-Match /
// If-None-Match) + the bypass family + the duplicate-detection switches.
// Multi-request entries (several Associate / Disassociate targets, binary
// pipelines) become a folder with one item per step, in order.
//
// The builders read table metadata from the live registry, so the caller
// loads each entry's table first (see `loadTablesFor`).

import { ENV } from '../mock/environment';
import { metadata } from '../host/metadataProvider';
import {
  buildRetrieveMultiple,
  buildRetrieveSingle,
  buildPredefinedQuery,
  buildCreate,
  buildCreateBody,
  buildUpdate,
  buildUpdateBody,
  buildUpsert,
  buildUpsertBody,
  buildDelete,
  buildMerge,
  buildMergeBody,
  buildAssociateRequests,
  buildDisassociateRequests,
  buildExecuteAction,
  buildExecuteActionBody,
  buildExecuteFunction,
  buildExecuteWorkflow,
  buildExecuteWorkflowBody,
  manageFilePipeline,
  manageImagePipeline,
  manageAttachmentPipeline,
  type BinaryPipelineStep,
} from './urlBuilder';
import { applyBypassToHeaders } from './bypassHeaders';
import { headerItemsToObject, type HeaderItem } from '../editors/HeadersEditor';
import { preferToHeaderString, type PreferSpec } from '../editors/PreferEditor';
import { preconditionToHeader } from '../editors/PreconditionEditor';
import type { BypassOptions, ConcurrencyMode } from '../state/writeState';
import {
  deserializeRetrieveMultiple,
  deserializeRetrieveSingle,
  deserializePredefinedQuery,
  deserializeCreate,
  deserializeUpdate,
  deserializeUpsert,
  deserializeDelete,
  deserializeMerge,
  deserializeAssociate,
  deserializeDisassociate,
  deserializeExecuteAction,
  deserializeExecuteFunction,
  deserializeExecuteWorkflow,
  deserializeManageFile,
  deserializeManageImage,
  deserializeManageAttachment,
  tableNameFromState,
  type SavedRequest,
  type SerializedRetrieveMultipleState,
  type SerializedRetrieveSingleState,
  type SerializedPredefinedQueryState,
  type SerializedCreateState,
  type SerializedUpdateState,
  type SerializedUpsertState,
  type SerializedDeleteState,
  type SerializedMergeState,
  type SerializedAssociateState,
  type SerializedDisassociateState,
  type SerializedExecuteActionState,
  type SerializedExecuteFunctionState,
  type SerializedExecuteWorkflowState,
  type SerializedManageFileState,
  type SerializedManageImageState,
  type SerializedManageAttachmentState,
} from '../state/savedRequests';
import { UPLOAD_URL_VAR, type CollectionItem, type CollectionRequest } from './collectionFormats';

/** The header-related fields a mode state may carry. */
interface HeaderBits {
  headers?: HeaderItem[];
  prefer?: PreferSpec;
  concurrency?: ConcurrencyMode;
  bypass?: BypassOptions;
  /** Delete's legacy "bypass custom plugins" switch. */
  bypassCustomPlugins?: boolean;
  /** Create: duplicate detection on → MSCRM.SuppressDuplicateDetection: false. */
  duplicateDetection?: boolean;
  /** Merge: skip duplicate detection on UpdateContent. */
  suppressDuplicateDetection?: boolean;
}

function synth(name: string, value: string): HeaderItem {
  return { id: `__export__${name}`, name, value, enabled: true, builtin: true };
}

/** Execute-time headers for a mode state, mirroring each mode's composition. */
export function composeHeaders(s: HeaderBits): Record<string, string> {
  let h: HeaderItem[] = (s.headers ?? []).map((it) =>
    it.name === 'MSCRM.SuppressDuplicateDetection' && s.duplicateDetection !== undefined
      ? { ...it, enabled: s.duplicateDetection, value: 'false' }
      : it,
  );
  const cc = s.concurrency ? preconditionToHeader(s.concurrency) : null;
  if (cc) h.push(synth(cc.name, cc.value));
  if (s.suppressDuplicateDetection) h.push(synth('MSCRM.SuppressDuplicateDetection', 'true'));
  if (s.bypass) {
    const b =
      s.bypassCustomPlugins && s.bypass.businessLogic === 'none'
        ? { ...s.bypass, businessLogic: 'sync' as const, useLegacyHeader: true }
        : s.bypass;
    h = applyBypassToHeaders(h, b);
  }
  return headerItemsToObject(h, s.prefer ? preferToHeaderString(s.prefer) : null);
}

/** Path + query without the host, for builder URLs that may be absolute. */
function pathOnly(url: string): string {
  return url.replace(/^https?:\/\/[^/]+/i, '');
}

function pipelineSteps(
  steps: BinaryPipelineStep[],
  headers: Record<string, string>,
): CollectionRequest[] {
  return steps.map((st) => {
    // Chunk uploads go to the URL the init step returns in its Location
    // header: export it as an `uploadUrl` variable the user fills in.
    const fromLocation = st.relativeUrl.startsWith('<');
    const notes = [
      st.detail,
      fromLocation ? "Set uploadUrl to the previous response's Location header." : '',
      typeof st.body === 'string' ? 'Replace the body with the file bytes before sending.' : '',
    ].filter(Boolean);
    return {
      name: `${st.n}. ${st.title}`,
      method: st.method,
      path: fromLocation ? UPLOAD_URL_VAR : pathOnly(st.relativeUrl),
      headers: { ...headers, ...st.headers },
      body: st.body,
      description: notes.join(' ') || undefined,
    };
  });
}

export type SavedConversion =
  | { ok: true; item: CollectionItem }
  | { ok: false; name: string; reason: string };

function single(
  entry: SavedRequest,
  method: string,
  url: string,
  headers: Record<string, string>,
  body?: Record<string, unknown>,
): SavedConversion {
  if (!url) {
    return {
      ok: false,
      name: entry.name,
      reason: `table "${tableNameFromState(entry.state)}" isn't available on this connection`,
    };
  }
  return {
    ok: true,
    item: { kind: 'request', name: entry.name, method, path: url, headers, body },
  };
}

function multi(entry: SavedRequest, steps: CollectionRequest[]): SavedConversion {
  if (steps.length === 0 || steps.some((s) => !s.path)) {
    return { ok: false, name: entry.name, reason: 'the request is incomplete' };
  }
  if (steps.length === 1) {
    return { ok: true, item: { kind: 'request', ...steps[0], name: entry.name } };
  }
  return {
    ok: true,
    item: {
      kind: 'folder',
      name: entry.name,
      description: `${steps.length} requests, run in order.`,
      steps,
    },
  };
}

/** Convert one saved entry into a collection request (or folder of steps). */
export function savedToCollectionItem(entry: SavedRequest): SavedConversion {
  try {
    switch (entry.modeId) {
      case 'retrieve-multiple': {
        const s = deserializeRetrieveMultiple(entry.state as SerializedRetrieveMultipleState);
        return single(entry, 'GET', buildRetrieveMultiple(s).relativeUrl, composeHeaders(s));
      }
      case 'retrieve-single': {
        const s = deserializeRetrieveSingle(entry.state as SerializedRetrieveSingleState);
        return single(entry, 'GET', buildRetrieveSingle(s).relativeUrl, composeHeaders(s));
      }
      case 'predefined-query': {
        const s = deserializePredefinedQuery(entry.state as SerializedPredefinedQueryState);
        return single(entry, 'GET', buildPredefinedQuery(s).relativeUrl, composeHeaders(s));
      }
      case 'create': {
        const s = deserializeCreate(entry.state as SerializedCreateState);
        return single(
          entry,
          'POST',
          buildCreate(s).relativeUrl,
          composeHeaders(s),
          buildCreateBody(s),
        );
      }
      case 'update': {
        const s = deserializeUpdate(entry.state as SerializedUpdateState);
        return single(
          entry,
          s.method,
          buildUpdate(s).relativeUrl,
          composeHeaders(s),
          buildUpdateBody(s),
        );
      }
      case 'upsert': {
        const s = deserializeUpsert(entry.state as SerializedUpsertState);
        return single(
          entry,
          'PATCH',
          buildUpsert(s).relativeUrl,
          composeHeaders(s),
          buildUpsertBody(s),
        );
      }
      case 'delete': {
        const s = deserializeDelete(entry.state as SerializedDeleteState);
        return single(entry, 'DELETE', buildDelete(s).relativeUrl, composeHeaders(s));
      }
      case 'merge': {
        const s = deserializeMerge(entry.state as SerializedMergeState);
        return single(
          entry,
          'POST',
          buildMerge(s).relativeUrl,
          composeHeaders(s),
          buildMergeBody(s),
        );
      }
      case 'associate': {
        const s = deserializeAssociate(entry.state as SerializedAssociateState);
        const headers = composeHeaders(s);
        return multi(
          entry,
          buildAssociateRequests(s).map((r) => ({
            name: `Associate ${r.targetId}`,
            method: r.method,
            path: r.relativeUrl,
            headers,
            body: r.body,
          })),
        );
      }
      case 'disassociate': {
        const s = deserializeDisassociate(entry.state as SerializedDisassociateState);
        const headers = composeHeaders(s);
        return multi(
          entry,
          buildDisassociateRequests(s).map((r) => ({
            name: r.targetId ? `Disassociate ${r.targetId}` : 'Disassociate',
            method: r.method,
            path: r.relativeUrl,
            headers,
            body: r.body,
          })),
        );
      }
      case 'exec-action': {
        const s = deserializeExecuteAction(entry.state as SerializedExecuteActionState);
        return single(
          entry,
          'POST',
          buildExecuteAction(s).relativeUrl,
          composeHeaders(s),
          buildExecuteActionBody(s),
        );
      }
      case 'exec-function': {
        const s = deserializeExecuteFunction(entry.state as SerializedExecuteFunctionState);
        return single(entry, 'GET', buildExecuteFunction(s).relativeUrl, composeHeaders(s));
      }
      case 'exec-workflow': {
        const s = deserializeExecuteWorkflow(entry.state as SerializedExecuteWorkflowState);
        return single(
          entry,
          'POST',
          buildExecuteWorkflow(s).relativeUrl,
          composeHeaders(s),
          buildExecuteWorkflowBody(s),
        );
      }
      case 'manage-file': {
        const s = deserializeManageFile(entry.state as SerializedManageFileState);
        return multi(entry, pipelineSteps(manageFilePipeline(s), composeHeaders(s)));
      }
      case 'manage-image': {
        const s = deserializeManageImage(entry.state as SerializedManageImageState);
        return multi(entry, pipelineSteps(manageImagePipeline(s), composeHeaders(s)));
      }
      case 'manage-attachment': {
        const s = deserializeManageAttachment(entry.state as SerializedManageAttachmentState);
        return multi(entry, pipelineSteps(manageAttachmentPipeline(s), composeHeaders(s)));
      }
      default:
        // Modes added later (e.g. bulk or batch modes) aren't exportable yet;
        // skip them with a reason instead of failing the whole export.
        return {
          ok: false,
          name: entry.name,
          reason: `${entry.modeId} requests can't be exported as a collection yet`,
        };
    }
  } catch (e) {
    return { ok: false, name: entry.name, reason: e instanceof Error ? e.message : String(e) };
  }
}

/**
 * Load the metadata of every table the entries target, so the builders can
 * resolve entity sets, lookup columns and nav properties. Best effort: a
 * table that fails to load surfaces later as a skipped entry.
 */
export async function loadTablesFor(entries: SavedRequest[]): Promise<void> {
  const tables = new Set(
    entries.map((e) => tableNameFromState(e.state)).filter((t) => t && t !== '—'),
  );
  await Promise.all([...tables].map((t) => metadata.getTable(t).catch(() => undefined)));
}

/** `https://<current org host>` — the `baseUrl` collection variable. */
export function currentBaseUrl(): string {
  return `https://${ENV.host}`;
}
