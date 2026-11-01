// $batch composer model (#9) — pure state + helpers, no React, no host.
//
// A batch is an ordered list of groups. A group is either a change set
// (atomic: every request succeeds or all roll back) or a run of independent
// requests. Content-IDs are numbered 1..n across the whole batch in order,
// so `$1`, `$2`, … can reference a request created earlier in the same
// change set.
//
// https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/execute-batch-operations-using-web-api
//
// How it reaches Dataverse: PPTB 1.2.7 exposes `dataverseAPI.executeBatch`
// (one $batch of independent requests) and `executeTransaction` (one $batch
// holding a single change set). It has no call for a mixed envelope, so
// `planDispatch` splits the batch into consecutive host calls — each change
// set its own `executeTransaction`, each run of independent groups one
// `executeBatch` — sent in order. A batch with one group is one call.

import { isCredentialHeader } from '../sensitiveHeaders';
import { API_BASE } from '../../env/environment';

export type BatchMethod = 'GET' | 'POST' | 'PATCH' | 'PUT' | 'DELETE';
export type BatchOpKind = 'create' | 'update' | 'delete' | 'retrieve' | 'associate' | 'custom';

export const MAX_BATCH_REQUESTS = 1000;
export { API_BASE };

export interface BatchHeaderRow {
  name: string;
  value: string;
  enabled: boolean;
}

export interface BatchItem {
  id: string;
  op: BatchOpKind;
  /** Optional user label; the sidebar falls back to METHOD + path. */
  label: string;
  /** Entity set name (create / update / delete / retrieve / associate source). */
  entitySet: string;
  /** Record id — a GUID, an alternate key (`key='v'`) or a `$n` reference. */
  recordId: string;
  /** Retrieve only: query options without the leading `?`. */
  query: string;
  /** Associate only: collection-valued navigation property. */
  navProperty: string;
  targetSet: string;
  /** Associate target id — a GUID or a `$n` reference. */
  targetId: string;
  /** Custom only. */
  method: BatchMethod;
  /** Custom only: path relative to the API base, e.g. `WhoAmI` or `accounts?$top=1`. */
  url: string;
  /** JSON body text (create / update / custom). */
  body: string;
  headers: BatchHeaderRow[];
  /** Where the item came from (saved request name, mode), for display. */
  source?: string;
}

export interface BatchGroup {
  id: string;
  /** True → change set (atomic). False → independent requests. */
  atomic: boolean;
  items: BatchItem[];
}

export interface BatchState {
  groups: BatchGroup[];
  /** `Prefer: odata.continue-on-error` — keep going after a failed request. */
  continueOnError: boolean;
  /**
   * ExecuteMultiple-style mode (#11) only: SDK `ReturnResponses`. False →
   * the results list only the faults, as ExecuteMultipleResponse does.
   */
  returnResponses?: boolean;
  dirty: Set<string>;
}

let seq = 1;
export const newBatchId = (prefix: string) =>
  `${prefix}${Date.now().toString(36)}${(seq++).toString(36)}`;

export function newItem(op: BatchOpKind = 'create', patch: Partial<BatchItem> = {}): BatchItem {
  return {
    id: newBatchId('i'),
    op,
    label: '',
    entitySet: '',
    recordId: '',
    query: '',
    navProperty: '',
    targetSet: '',
    targetId: '',
    method: 'POST',
    url: '',
    body: op === 'create' || op === 'update' ? '{\n  \n}' : '',
    headers: [],
    ...patch,
  };
}

export function newGroup(atomic: boolean, items: BatchItem[] = []): BatchGroup {
  return { id: newBatchId('g'), atomic, items };
}

export const initialBatchState = (): BatchState => ({
  groups: [newGroup(true)],
  continueOnError: false,
  dirty: new Set(),
});

/** Every item with its 1-based Content-ID and owning group, in batch order. */
export interface NumberedItem {
  item: BatchItem;
  group: BatchGroup;
  groupIndex: number;
  contentId: string;
}

export function numberItems(groups: BatchGroup[]): NumberedItem[] {
  const out: NumberedItem[] = [];
  let n = 0;
  groups.forEach((group, groupIndex) => {
    for (const item of group.items) {
      n += 1;
      out.push({ item, group, groupIndex, contentId: String(n) });
    }
  });
  return out;
}

export const methodOf = (item: BatchItem): BatchMethod => {
  switch (item.op) {
    case 'create':
    case 'associate':
      return 'POST';
    case 'update':
      return 'PATCH';
    case 'delete':
      return 'DELETE';
    case 'retrieve':
      return 'GET';
    case 'custom':
      return item.method;
  }
};

const REF_RE = /^\$\d+$/;
/** `$3` → true. A Content-ID reference stands in for a created entity's URI. */
export const isContentRef = (s: string) => REF_RE.test(s.trim());

/** Strip a leading `/` and `api/data/v9.x/` so paths are stored relative to the API base. */
export function normalizePath(path: string): string {
  return path
    .trim()
    .replace(/^https?:\/\/[^/]+/i, '')
    .replace(/^\/+/, '')
    .replace(/^api\/data\/v\d+(\.\d+)?\//i, '');
}

function recordSegment(set: string, id: string): string {
  const t = id.trim();
  if (isContentRef(t)) return t;
  return `${set.trim()}(${t.replace(/^\{|\}$/g, '')})`;
}

/** The request as it goes on the wire, relative to the API base. */
export interface WireRequest {
  method: BatchMethod;
  /** Relative to the API base, no leading slash (`accounts(…)`, `$1/…`). */
  path: string;
  headers: Record<string, string>;
  /** Canonical JSON text, or undefined when the request has no body. */
  body?: string;
}

export interface ItemProblem {
  severity: 'error' | 'warning';
  message: string;
}

/**
 * Build the wire request for one item. `orgUrl` (`https://contoso.crm…`)
 * makes the associate `@odata.id` absolute, as the $ref docs require; when
 * empty the reference stays relative.
 */
export function itemToWire(
  item: BatchItem,
  orgUrl = '',
): { wire: WireRequest | null; problems: ItemProblem[] } {
  const problems: ItemProblem[] = [];
  const err = (message: string) => problems.push({ severity: 'error', message });
  const method = methodOf(item);
  const set = item.entitySet.trim();
  const id = item.recordId.trim();
  let path = '';
  let body: string | undefined;

  const parseBody = (required: boolean): string | undefined => {
    const text = item.body.trim();
    if (!text) {
      if (required) err('Body is empty.');
      return undefined;
    }
    try {
      const parsed = JSON.parse(text) as unknown;
      if (required && (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed))) {
        err('Body must be a JSON object.');
      }
      return prettyJson(parsed);
    } catch (e) {
      err(`Body isn't valid JSON: ${e instanceof Error ? e.message : String(e)}`);
      return undefined;
    }
  };

  switch (item.op) {
    case 'create':
      if (!set) err('Pick a table (entity set).');
      path = set;
      body = parseBody(true);
      break;
    case 'update':
      if (!set && !isContentRef(id)) err('Pick a table (entity set).');
      if (!id) err('Enter the record id (or a $n reference).');
      path = recordSegment(set, id);
      body = parseBody(true);
      break;
    case 'delete':
      if (!set && !isContentRef(id)) err('Pick a table (entity set).');
      if (!id) err('Enter the record id (or a $n reference).');
      path = recordSegment(set, id);
      break;
    case 'retrieve': {
      if (!set) err('Pick a table (entity set).');
      const q = item.query.trim().replace(/^\?/, '');
      path = `${id ? recordSegment(set, id) : set}${q ? `?${q}` : ''}`;
      break;
    }
    case 'associate': {
      if (!set && !isContentRef(id)) err('Pick the source table (entity set).');
      if (!id) err('Enter the source record id.');
      if (!item.navProperty.trim()) err('Enter the collection-valued navigation property.');
      const tid = item.targetId.trim();
      if (!tid) err('Enter the target record id.');
      if (!item.targetSet.trim() && !isContentRef(tid)) err('Pick the target table.');
      path = `${recordSegment(set, id)}/${item.navProperty.trim()}/$ref`;
      const ref = isContentRef(tid)
        ? tid
        : `${orgUrl ? `${orgUrl.replace(/\/+$/, '')}${API_BASE}/` : ''}${recordSegment(item.targetSet, tid)}`;
      body = prettyJson({ '@odata.id': ref });
      break;
    }
    case 'custom':
      path = normalizePath(item.url);
      if (!path) err('Enter the request URL.');
      if (/\$batch\b/i.test(path)) err('A batch cannot contain another $batch request.');
      body = parseBody(false);
      if (method === 'GET' && body !== undefined) err('A GET request cannot have a body.');
      break;
  }

  const headers: Record<string, string> = {};
  for (const h of item.headers) {
    if (!h.enabled) continue;
    const name = h.name.trim();
    if (!name && !h.value) continue;
    if (!HEADER_NAME_RE.test(name)) {
      err(`Invalid header name "${name}".`);
      continue;
    }
    // eslint-disable-next-line no-control-regex
    if (/[\x00-\x1f\x7f]/.test(h.value)) {
      err(`Header ${name} has a control character in its value.`);
      continue;
    }
    // Never sent, saved or exported (#217): the host authenticates the batch.
    if (isCredentialHeader(name)) {
      problems.push({
        severity: 'warning',
        message: `${name} isn't sent: Power Platform ToolBox authenticates the batch.`,
      });
      continue;
    }
    headers[name] = h.value;
  }

  if (problems.some((p) => p.severity === 'error')) return { wire: null, problems };
  return { wire: { method, path, headers, body }, problems };
}

/** Indented JSON with CRLF line breaks — a $batch payload must be CRLF throughout. */
export const prettyJson = (v: unknown): string => JSON.stringify(v, null, 2).replace(/\n/g, '\r\n');

const HEADER_NAME_RE = /^[!#$%&'*+\-.^_`|~0-9A-Za-z]+$/;

/** `$n` references used in a path or body. */
export function contentRefsIn(wire: WireRequest): { inUrl: string[]; inBody: string[] } {
  const inUrl = [...wire.path.matchAll(/(?:^|\/)\$(\d+)(?=\/|$|\?)/g)].map((m) => m[1]);
  const inBody = wire.body ? [...wire.body.matchAll(/"\$(\d+)"/g)].map((m) => m[1]) : [];
  return { inUrl, inBody };
}

export interface BatchIssue {
  severity: 'error' | 'warning';
  /** Item id, group id, or '' for the whole batch. */
  target: string;
  message: string;
}

export interface CompiledItem extends NumberedItem {
  wire: WireRequest | null;
}

export interface CompiledBatch {
  items: CompiledItem[];
  issues: BatchIssue[];
}

/** Turn the state into wire requests and collect every blocking or advisory problem. */
export function compileBatch(groups: BatchGroup[], orgUrl = ''): CompiledBatch {
  const numbered = numberItems(groups);
  const issues: BatchIssue[] = [];
  const items: CompiledItem[] = numbered.map((n) => {
    const { wire, problems } = itemToWire(n.item, orgUrl);
    for (const p of problems) issues.push({ ...p, target: n.item.id });
    return { ...n, wire };
  });

  if (items.length === 0) issues.push({ severity: 'error', target: '', message: 'Add a request.' });
  if (items.length > MAX_BATCH_REQUESTS) {
    issues.push({
      severity: 'error',
      target: '',
      message: `A batch can hold at most ${MAX_BATCH_REQUESTS} requests (this one has ${items.length}).`,
    });
  }

  const contentIdGroup = new Map(items.map((c) => [c.contentId, c]));
  for (const c of items) {
    if (!c.wire) continue;
    if (c.group.atomic && c.wire.method === 'GET') {
      issues.push({
        severity: 'error',
        target: c.item.id,
        message:
          'GET requests aren’t allowed inside a change set. Move it to an independent group.',
      });
    }
    const { inUrl, inBody } = contentRefsIn(c.wire);
    for (const ref of [...inUrl, ...inBody]) {
      const target = contentIdGroup.get(ref);
      if (!c.group.atomic) {
        issues.push({
          severity: 'error',
          target: c.item.id,
          message: `$${ref} can only be referenced inside a change set.`,
        });
      } else if (!target || target.group !== c.group || Number(ref) >= Number(c.contentId)) {
        issues.push({
          severity: 'error',
          target: c.item.id,
          message: `$${ref} must be an earlier request in the same change set.`,
        });
      }
    }
    if (inUrl.length > 0) {
      issues.push({
        severity: 'warning',
        target: c.item.id,
        message:
          `The URL starts with a Content-ID reference ($${inUrl[0]}). Power Platform ToolBox ` +
          `prefixes every batch URL with ${API_BASE}/, so Dataverse may not resolve it. ` +
          `References in the body (@odata.bind / @odata.id) aren’t affected.`,
      });
    }
  }
  return { items, issues };
}

/** One host call: a change set (`executeTransaction`) or independent requests (`executeBatch`). */
export interface DispatchSegment {
  transaction: boolean;
  items: CompiledItem[];
}

/** Split the batch into the host calls that send it, in order. Empty groups are skipped. */
export function planDispatch(items: CompiledItem[]): DispatchSegment[] {
  const segments: DispatchSegment[] = [];
  let lastGroup: BatchGroup | null = null;
  for (const c of items) {
    const tail = segments[segments.length - 1];
    if (c.group.atomic) {
      if (tail && tail.transaction && lastGroup === c.group) tail.items.push(c);
      else segments.push({ transaction: true, items: [c] });
    } else if (tail && !tail.transaction) {
      tail.items.push(c);
    } else {
      segments.push({ transaction: false, items: [c] });
    }
    lastGroup = c.group;
  }
  return segments;
}

/** Short display path: `POST accounts`, `PATCH $1`, … */
export function itemSummary(item: BatchItem, orgUrl = ''): { method: BatchMethod; path: string } {
  const { wire } = itemToWire(item, orgUrl);
  if (wire) return { method: wire.method, path: wire.path };
  return {
    method: methodOf(item),
    path: item.op === 'custom' ? normalizePath(item.url) : item.entitySet,
  };
}

// ── Save / load ─────────────────────────────────────────────────────────

export type SerializedBatchState = Omit<BatchState, 'dirty'>;

export function serializeBatch(state: BatchState): SerializedBatchState {
  const { dirty: _d, ...rest } = state;
  return rest;
}

export function deserializeBatch(snap: SerializedBatchState): BatchState {
  const groups = Array.isArray(snap.groups) ? snap.groups : [];
  return {
    groups: groups.map((g) => ({
      id: g.id || newBatchId('g'),
      atomic: !!g.atomic,
      items: (Array.isArray(g.items) ? g.items : []).map((i) => ({ ...newItem(i.op), ...i })),
    })),
    continueOnError: !!snap.continueOnError,
    ...(snap.returnResponses !== undefined ? { returnResponses: !!snap.returnResponses } : {}),
    dirty: new Set(),
  };
}

export function autoSuggestBatch(snap: SerializedBatchState): string {
  const n = (snap.groups ?? []).reduce((acc, g) => acc + (g.items?.length ?? 0), 0);
  const cs = (snap.groups ?? []).filter((g) => g.atomic && g.items?.length).length;
  const parts = ['batch', `${n} request${n === 1 ? '' : 's'}`];
  if (cs) parts.push(`${cs} change set${cs === 1 ? '' : 's'}`);
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}

// ── ExecuteMultiple-style mode (#11) ────────────────────────────────────
// One flat list of independent requests: ExecuteMultiple has no change
// sets (ExecuteTransaction is the atomic message — a change set here).

export const initialExecuteMultipleState = (): BatchState => ({
  groups: [newGroup(false)],
  continueOnError: false,
  returnResponses: true,
  dirty: new Set(),
});

export function autoSuggestExecuteMultiple(snap: SerializedBatchState): string {
  const n = (snap.groups ?? []).reduce((acc, g) => acc + (g.items?.length ?? 0), 0);
  const parts = ['ExecuteMultiple', `${n} request${n === 1 ? '' : 's'}`];
  if (snap.continueOnError) parts.push('continue on error');
  parts.push(new Date().toISOString().slice(0, 10));
  return parts.join(' · ');
}
