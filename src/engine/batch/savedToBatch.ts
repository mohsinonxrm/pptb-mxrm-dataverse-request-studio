// Saved request / current request → $batch items (#9).
//
// Rebuilds the request(s) a mode would send with the same builders the
// modes use, composes its headers the way the mode does at execute time
// (user headers + Prefer + precondition + bypass + duplicate detection),
// and snapshots the result as raw "custom" batch items the user can edit.
// Associate / Disassociate with several targets become one item per target.
// Binary modes (file / image / attachment) run multi-step pipelines with
// binary bodies, so they aren't offered.

import { metadata } from '../../host/metadataProvider';
import { findTable } from '../../mock/metadata';
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
} from '../urlBuilder';
import { applyBypassToHeaders } from '../bypassHeaders';
import { headerItemsToObject, type HeaderItem } from '../../editors/HeadersEditor';
import { preferToHeaderString, type PreferSpec } from '../../editors/PreferEditor';
import { preconditionToHeader } from '../../editors/PreconditionEditor';
import type { BypassOptions, ConcurrencyMode } from '../../state/writeState';
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
  type SavedModeId,
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
} from '../../state/savedRequests';
import { newItem, normalizePath, prettyJson, type BatchItem, type BatchMethod } from './batchModel';

/** A mode's request to convert: a saved entry, or the last request a mode published. */
export interface RequestSnapshot {
  modeId: SavedModeId;
  state: unknown;
  /** Display name — saved entry name or mode label. */
  name: string;
}

/** Modes whose requests can go into a batch. */
export const BATCHABLE_MODES: ReadonlySet<SavedModeId> = new Set<SavedModeId>([
  'retrieve-multiple',
  'retrieve-single',
  'predefined-query',
  'create',
  'update',
  'upsert',
  'delete',
  'merge',
  'associate',
  'disassociate',
  'exec-action',
  'exec-function',
  'exec-workflow',
]);

interface HeaderBits {
  headers?: HeaderItem[];
  prefer?: PreferSpec;
  concurrency?: ConcurrencyMode;
  bypass?: BypassOptions;
  bypassCustomPlugins?: boolean;
  duplicateDetection?: boolean;
  suppressDuplicateDetection?: boolean;
}

// Inside a batch part these add nothing (the host sets the outer request's
// Accept / OData-* and the part's Content-Type) but would trigger the host's
// header consent prompt, so they're left out.
const PART_DROPPED = new Set([
  'authorization',
  'accept',
  'content-type',
  'content-length',
  'odata-version',
  'odata-maxversion',
]);

function synth(name: string, value: string): HeaderItem {
  return { id: `__batch__${name}`, name, value, enabled: true, builtin: true };
}

/** The headers a mode sends at execute time, minus the ones a batch part doesn't need. */
export function composePartHeaders(s: HeaderBits): Record<string, string> {
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
  const all = headerItemsToObject(h, s.prefer ? preferToHeaderString(s.prefer) : null);
  return Object.fromEntries(
    Object.entries(all).filter(([name]) => !PART_DROPPED.has(name.toLowerCase())),
  );
}

function toItem(
  label: string,
  source: string,
  method: string,
  url: string,
  headers: Record<string, string>,
  body?: Record<string, unknown>,
): BatchItem {
  return newItem('custom', {
    label,
    source,
    method: method as BatchMethod,
    url: normalizePath(url),
    body: body && Object.keys(body).length > 0 ? prettyJson(body) : body ? '{}' : '',
    headers: Object.entries(headers).map(([name, value]) => ({ name, value, enabled: true })),
  });
}

export type SnapshotConversion = { ok: true; items: BatchItem[] } | { ok: false; reason: string };

/** Convert synchronously — the request's table metadata must already be loaded. */
export function snapshotToItems(snap: RequestSnapshot): SnapshotConversion {
  const { name } = snap;
  const one = (
    method: string,
    url: string,
    headers: Record<string, string>,
    body?: Record<string, unknown>,
  ): SnapshotConversion =>
    url
      ? { ok: true, items: [toItem(name, name, method, url, headers, body)] }
      : { ok: false, reason: 'its table isn’t available on this connection' };
  try {
    switch (snap.modeId) {
      case 'retrieve-multiple': {
        const s = deserializeRetrieveMultiple(snap.state as SerializedRetrieveMultipleState);
        return one('GET', buildRetrieveMultiple(s).relativeUrl, composePartHeaders(s));
      }
      case 'retrieve-single': {
        const s = deserializeRetrieveSingle(snap.state as SerializedRetrieveSingleState);
        return one('GET', buildRetrieveSingle(s).relativeUrl, composePartHeaders(s));
      }
      case 'predefined-query': {
        const s = deserializePredefinedQuery(snap.state as SerializedPredefinedQueryState);
        return one('GET', buildPredefinedQuery(s).relativeUrl, composePartHeaders(s));
      }
      case 'create': {
        const s = deserializeCreate(snap.state as SerializedCreateState);
        return one('POST', buildCreate(s).relativeUrl, composePartHeaders(s), buildCreateBody(s));
      }
      case 'update': {
        const s = deserializeUpdate(snap.state as SerializedUpdateState);
        return one(s.method, buildUpdate(s).relativeUrl, composePartHeaders(s), buildUpdateBody(s));
      }
      case 'upsert': {
        const s = deserializeUpsert(snap.state as SerializedUpsertState);
        return one('PATCH', buildUpsert(s).relativeUrl, composePartHeaders(s), buildUpsertBody(s));
      }
      case 'delete': {
        const s = deserializeDelete(snap.state as SerializedDeleteState);
        return one('DELETE', buildDelete(s).relativeUrl, composePartHeaders(s));
      }
      case 'merge': {
        const s = deserializeMerge(snap.state as SerializedMergeState);
        return one('POST', buildMerge(s).relativeUrl, composePartHeaders(s), buildMergeBody(s));
      }
      case 'associate': {
        const s = deserializeAssociate(snap.state as SerializedAssociateState);
        const headers = composePartHeaders(s);
        const reqs = buildAssociateRequests(s);
        if (reqs.length === 0) return { ok: false, reason: 'the request is incomplete' };
        return {
          ok: true,
          items: reqs.map((r) =>
            toItem(`${name} → ${r.targetId}`, name, r.method, r.relativeUrl, headers, r.body),
          ),
        };
      }
      case 'disassociate': {
        const s = deserializeDisassociate(snap.state as SerializedDisassociateState);
        const headers = composePartHeaders(s);
        const reqs = buildDisassociateRequests(s);
        if (reqs.length === 0) return { ok: false, reason: 'the request is incomplete' };
        return {
          ok: true,
          items: reqs.map((r) =>
            toItem(
              r.targetId ? `${name} × ${r.targetId}` : name,
              name,
              r.method,
              r.relativeUrl,
              headers,
              r.body,
            ),
          ),
        };
      }
      case 'exec-action': {
        const s = deserializeExecuteAction(snap.state as SerializedExecuteActionState);
        return one(
          'POST',
          buildExecuteAction(s).relativeUrl,
          composePartHeaders(s),
          buildExecuteActionBody(s),
        );
      }
      case 'exec-function': {
        const s = deserializeExecuteFunction(snap.state as SerializedExecuteFunctionState);
        return one('GET', buildExecuteFunction(s).relativeUrl, composePartHeaders(s));
      }
      case 'exec-workflow': {
        const s = deserializeExecuteWorkflow(snap.state as SerializedExecuteWorkflowState);
        return one(
          'POST',
          buildExecuteWorkflow(s).relativeUrl,
          composePartHeaders(s),
          buildExecuteWorkflowBody(s),
        );
      }
      default:
        return {
          ok: false,
          reason: 'file, image and attachment requests can’t be sent in a batch',
        };
    }
  } catch (e) {
    return { ok: false, reason: e instanceof Error ? e.message : String(e) };
  }
}

function tableOf(state: unknown): string {
  if (state && typeof state === 'object' && 'table' in state) {
    const t = (state as { table?: unknown }).table;
    if (typeof t === 'string') return t;
  }
  return '';
}

/** Load the snapshot's table metadata (so the builders resolve), then convert. */
export async function convertSnapshot(snap: RequestSnapshot): Promise<SnapshotConversion> {
  const table = tableOf(snap.state);
  if (table) await metadata.getTable(table).catch(() => undefined);
  // Associate / Disassociate also resolve the navigation target's entity set.
  const nav = (snap.state as { navProperty?: unknown } | null)?.navProperty;
  if (table && typeof nav === 'string' && nav) {
    const target = findTable(table)?.navigationProperties.find((n) => n.name === nav)?.targetEntity;
    if (target) await metadata.getTable(target).catch(() => undefined);
  }
  return snapshotToItems(snap);
}
