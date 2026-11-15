// SDK code generation (#232): the current request as Dataverse SDK calls
// instead of a Web API HTTP call.
//
//   • C# SDK — ServiceClient (Microsoft.PowerPlatform.Dataverse.Client):
//     QueryExpression for queries, Entity + Create / Update / Upsert /
//     Delete for writes, Associate / Disassociate, and OrganizationRequest
//     for actions and functions.
//   • PowerShell (Xrm) — the Microsoft.Xrm.Data.PowerShell module:
//     Get-CrmRecordsByFetch with FetchXML, New-CrmRecord, Set-CrmRecord, …,
//     and SDK request objects through $conn.Execute where no cmdlet fits.
//
// Both render one shared plan (`buildPlan`). OData values become typed SDK
// values: a lookup an EntityReference, a choice an OptionSetValue, money a
// Money, using the table's loaded metadata. A request with no faithful SDK
// form (an $expand, a lambda, deep insert, a file's bytes, …) becomes a
// `webapi` / `binary` plan: the snippet says why, and the C# one sends the
// Web API request through ServiceClient.ExecuteWebRequest instead of
// generating SDK code that would mean something else.

import type { CodegenInputs } from './codeGenerators';
import { ENV } from '../env/environment';
import {
  findColumn,
  findTable,
  findTableBySet,
  type AttributeTypeCode,
  type ColumnMeta,
} from '../metadata/metadata';
import { parseFilterForCodegen, topLevelSplit } from './odataParser';
import type { FilterGroup, FilterNode } from '../editors/filter/filterTree';
import { findOperator } from '../editors/filter/operators';
import { findActionSync } from '../host/csdlProvider';
import type { ActionParam } from '../metadata/actionsCsdl';
import { isCredentialHeader } from './sensitiveHeaders';
import { encodeSpaces } from '../primitives/urlParts';
import { csharpString, csharpVerbatimEscape, psQuote } from './codegenStrings';

// ────────────────────────────────────────────────────────────
// Plan model
// ────────────────────────────────────────────────────────────

/** A typed SDK value. Numbers keep their source text so nothing is rounded. */
export type SdkValue =
  | { t: 'null' }
  | { t: 'string'; v: string }
  | { t: 'int' | 'long' | 'decimal' | 'double'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'guid'; v: string }
  | { t: 'datetime'; v: string }
  | { t: 'option'; v: string }
  | { t: 'options'; v: string[] }
  | { t: 'money'; v: string }
  | { t: 'ref'; entity: string; id: string }
  | { t: 'bytes'; v: string }
  | { t: 'array'; elem: 'string' | 'int' | 'guid'; v: string[] };

interface Attr {
  name: string;
  value: SdkValue;
}

type Target = { kind: 'id'; id: string } | { kind: 'key'; pairs: [string, SdkValue][] };

/** A filter as QueryExpression / FetchXML see it. `op` is a ConditionOperator name. */
type Cond =
  | { type: 'group'; or: boolean; children: Cond[] }
  | { type: 'cond'; attr: string; op: string; values: SdkValue[]; column?: string };

interface QueryModel {
  entity: string;
  /** null = all columns (no $select). */
  columns: string[] | null;
  filter: Cond | null;
  orders: { attr: string; desc: boolean }[];
  top?: number;
  count: boolean;
  /** `$top=1&$count=true` + the limit annotation: Count only (#272). */
  countOnly: boolean;
  pageSize?: number;
}

type Plan =
  | { k: 'query'; q: QueryModel }
  | { k: 'retrieve'; entity: string; target: Target; columns: string[] | null }
  | { k: 'create'; entity: string; attrs: Attr[]; id?: string }
  | { k: 'update'; entity: string; target: Target; attrs: Attr[]; rowVersion?: string }
  | { k: 'upsert'; entity: string; target: Target; attrs: Attr[] }
  | { k: 'delete'; entity: string; target: Target; rowVersion?: string }
  | {
      k: 'associate' | 'disassociate';
      entity: string;
      id: string;
      relationship: string;
      targets: Extract<SdkValue, { t: 'ref' }>[];
      /** One request for every target (all or nothing), not one per target. */
      atomic: boolean;
      /** The source row's role in a self-referencing relationship. */
      selfRole?: 'Referenced' | 'Referencing';
    }
  | {
      k: 'execute';
      name: string;
      params: [string, SdkValue][];
      target?: Extract<SdkValue, { t: 'ref' }>;
    }
  | { k: 'webapi'; reason: string }
  | { k: 'binary'; reason: string };

/** Header-driven request options. */
interface Opts {
  /** Optional parameters (bypass, duplicate detection, …) — need a request class. */
  params: [string, SdkValue][];
  callerId?: string;
  callerObjectId?: string;
  ifMatch?: string;
  ifNoneMatch?: string;
  pageSize?: number;
  annotations: boolean;
  returnRepresentation: boolean;
  /** Headers with no SDK equivalent. */
  dropped: [string, string][];
}

interface Built {
  plan: Plan;
  opts: Opts;
  /** Comment lines shown above the code (checks the user should make). */
  notes: string[];
}

class Unsupported extends Error {}
const no = (reason: string): never => {
  throw new Unsupported(reason);
};

// ────────────────────────────────────────────────────────────
// Small helpers
// ────────────────────────────────────────────────────────────

const NAME_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;
const GUID_RE = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;
const INT_RE = /^-?\d+$/;
const NUM_RE = /^-?\d+(\.\d+)?([eE][+-]?\d+)?$/;
const DATE_RE = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:\d{2})?)?$/;
const INT32_MAX = 2147483647;

const safeDecode = (v: string): string => {
  try {
    return decodeURIComponent(v);
  } catch {
    return v;
  }
};

const checkName = (n: string, what: string): string =>
  NAME_RE.test(n) ? n : no(`${what} '${n}' isn't a plain column or table name.`);

/** `_parentaccountid_value` → `parentaccountid`: the SDK uses the attribute name. */
const bareAttr = (n: string): string =>
  /^_[A-Za-z0-9_]+_value$/.test(n) ? n.slice(1, -'_value'.length) : n;

/** Text safe for a `//` or `#` comment: no line breaks or control characters. */
export function lineComment(s: string): string {
  return (
    s
      // eslint-disable-next-line no-control-regex
      .replace(/[\u0000-\u001f\u007f\u0085\u2028\u2029]+/g, ' ')
      .replace(/[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/g, '\ufffd')
  );
}

function colMeta(entity: string, attr: string): ColumnMeta | undefined {
  const t = findTable(entity);
  return t ? findColumn(t, attr) : undefined;
}

/** Logical name for an entity set: from metadata, else a guess the snippet flags. */
function logicalOfSet(set: string, notes: string[]): string {
  const t = findTableBySet(set);
  if (t) return t.logicalName;
  const guess = /ies$/.test(set)
    ? `${set.slice(0, -3)}y`
    : /(ss|x|ch|sh)es$/.test(set)
      ? set.slice(0, -2)
      : set.replace(/s$/, '');
  notes.push(
    `Check: '${guess}' is a guess at the table behind '${set}' (its metadata isn't loaded).`,
  );
  return guess;
}

/** The record a bind / @odata.id URL points at: `/contacts(<id>)`, `contacts(<id>)` or absolute. */
function refFromUrl(url: string, notes: string[]): Extract<SdkValue, { t: 'ref' }> {
  const path = url
    .trim()
    .replace(/^https?:\/\/[^/]+/, '')
    .replace(/^\/?api\/data\/v\d+(\.\d+)?\//, '')
    .replace(/^\//, '');
  const m = /^([A-Za-z_][A-Za-z0-9_]*)\(([^)]*)\)$/.exec(path);
  if (!m) return no(`'${url}' isn't a record URL like /contacts(<id>).`);
  if (!GUID_RE.test(m[2])) return no(`'${url}' addresses the record by alternate key.`);
  return { t: 'ref', entity: logicalOfSet(m[1], notes), id: m[2] };
}

// ────────────────────────────────────────────────────────────
// URL path
// ────────────────────────────────────────────────────────────

interface ParsedPath {
  set: string;
  /** Raw text between the key parentheses. */
  key?: string;
  /** Segments after the key, e.g. ['primarycontactid', '$ref']. */
  segs: string[];
}

function parsePath(noBase: string): ParsedPath {
  const path = safeDecode(noBase.split('?')[0]).replace(/^\//, '');
  const m = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(path);
  if (!m) return no('The URL has no table or message name.');
  const set = m[0];
  let rest = path.slice(set.length);
  let key: string | undefined;
  if (rest.startsWith('(')) {
    // Scan to the matching ')' — a quoted key value may contain one.
    let inStr = false;
    let end = -1;
    for (let p = 1; p < rest.length; p++) {
      const c = rest[p];
      if (inStr) {
        if (c === "'") {
          if (rest[p + 1] === "'") p++;
          else inStr = false;
        }
      } else if (c === "'") inStr = true;
      else if (c === ')') {
        end = p;
        break;
      }
    }
    if (end < 0) return no('The URL key has no closing parenthesis.');
    key = rest.slice(1, end);
    rest = rest.slice(end + 1);
  }
  return { set, key, segs: rest.split('/').filter(Boolean) };
}

/** A key literal (`'A-1'`, `42`, a GUID) typed for the key column. */
function keyValue(entity: string, col: string, raw: string): SdkValue {
  const r = raw.trim();
  if (r.length >= 2 && r.startsWith("'") && r.endsWith("'")) {
    return { t: 'string', v: r.slice(1, -1).replace(/''/g, "'") };
  }
  return literal(colMeta(entity, col)?.attributeType, r, false);
}

function targetOf(entity: string, key: string | undefined): Target {
  if (!key) return no('The URL has no record key.');
  if (GUID_RE.test(key.trim())) return { kind: 'id', id: key.trim() };
  const pairs = topLevelSplit(key, ',').map((p): [string, SdkValue] => {
    const eq = p.indexOf('=');
    if (eq < 0) return no(`The key '${key}' isn't a GUID or a column=value list.`);
    const col = checkName(p.slice(0, eq).trim(), 'Key column');
    if (bareAttr(col) !== col) {
      return no(
        `The alternate key uses the lookup '${col}': its SDK key value isn't a plain value.`,
      );
    }
    return [col, keyValue(entity, col, p.slice(eq + 1))];
  });
  return { kind: 'key', pairs };
}

// ────────────────────────────────────────────────────────────
// Values
// ────────────────────────────────────────────────────────────

const INT_TYPES = new Set<AttributeTypeCode>(['Integer', 'Picklist', 'State', 'Status']);
const GUID_TYPES = new Set<AttributeTypeCode>(['Uniqueidentifier', 'Lookup', 'Customer', 'Owner']);

/** A number typed for its column (int when unknown and whole, else decimal). */
function numberFor(type: AttributeTypeCode | undefined, v: string): SdkValue {
  if (!NUM_RE.test(v)) return no(`'${v}' isn't a number.`);
  const whole = INT_RE.test(v);
  const long = whole && fitsLong(v);
  switch (type) {
    case 'BigInt':
      return long ? { t: 'long', v } : no(`'${v}' isn't a 64-bit whole number.`);
    case 'Decimal':
    case 'Money':
      return { t: 'decimal', v };
    case 'Double':
      return { t: 'double', v };
    default:
      if (whole && Math.abs(Number(v)) <= INT32_MAX) return { t: 'int', v };
      return long ? { t: 'long', v } : { t: 'decimal', v };
  }
}

/** A whole number's text fits a C# long. */
function fitsLong(v: string): boolean {
  const n = BigInt(v);
  return n >= -(2n ** 63n) && n < 2n ** 63n;
}

/**
 * An OData literal as an SDK condition value. `quoted` = it was a quoted
 * string; otherwise its shape (true/false, GUID, date, number) decides,
 * with the column type choosing the number type.
 */
function literal(type: AttributeTypeCode | undefined, raw: string, quoted: boolean): SdkValue {
  if (quoted) return { t: 'string', v: raw };
  const lower = raw.toLowerCase();
  if (lower === 'true' || lower === 'false') return { t: 'bool', v: lower === 'true' };
  if (GUID_RE.test(raw)) return { t: 'guid', v: raw };
  if (DATE_RE.test(raw) && !NUM_RE.test(raw)) return { t: 'datetime', v: raw };
  return numberFor(type, raw);
}

/**
 * A query-function value (always text in the URL, quoted or not), typed
 * from the column's metadata; without it, a GUID stays a Guid and a whole
 * number an int.
 */
function looseValue(type: AttributeTypeCode | undefined, raw: string): SdkValue {
  if (type) {
    if (INT_TYPES.has(type) || type === 'MultiSelectPicklist' || type === 'EntityName') {
      return INT_RE.test(raw) ? { t: 'int', v: raw } : { t: 'string', v: raw };
    }
    if (type === 'BigInt' || type === 'Decimal' || type === 'Money' || type === 'Double') {
      return numberFor(type, raw);
    }
    if (type === 'DateTime') return { t: 'datetime', v: raw };
    if (GUID_TYPES.has(type) && GUID_RE.test(raw)) return { t: 'guid', v: raw };
    if (type === 'Boolean' && /^(true|false)$/i.test(raw)) {
      return { t: 'bool', v: raw.toLowerCase() === 'true' };
    }
    return { t: 'string', v: raw };
  }
  if (GUID_RE.test(raw)) return { t: 'guid', v: raw };
  // A leading zero ('00123') means text: an int would drop it.
  if (/^-?(0|[1-9]\d*)$/.test(raw) && Math.abs(Number(raw)) <= INT32_MAX) {
    return { t: 'int', v: raw };
  }
  return { t: 'string', v: raw };
}

/** A JSON body value for a column, typed by its metadata when loaded. */
function columnValue(col: ColumnMeta | undefined, attr: string, v: unknown): SdkValue {
  if (v === null) return { t: 'null' };
  const json = jsonValue(v, `Column '${attr}'`);
  if (!col) return json;
  const type = col.attributeType;
  if (type === 'Lookup' || type === 'Customer' || type === 'Owner') {
    return no(`'${attr}' is a lookup: the Web API sets it with ${attr}@odata.bind.`);
  }
  if (INT_TYPES.has(type) && type !== 'Integer') {
    if (json.t === 'int') return { t: 'option', v: json.v };
    if (json.t === 'string' && INT_RE.test(json.v)) return { t: 'option', v: json.v };
  }
  if (type === 'MultiSelectPicklist') {
    const list = typeof v === 'string' ? v.split(',').map((s) => s.trim()) : null;
    if (list && list.every((s) => INT_RE.test(s))) return { t: 'options', v: list };
    if (Array.isArray(v) && v.every((n) => Number.isInteger(n))) {
      return { t: 'options', v: v.map(String) };
    }
  }
  if (typeof v === 'number') {
    if (type === 'Money') return { t: 'money', v: String(v) };
    if (type === 'Decimal' || type === 'Double' || type === 'BigInt' || type === 'Integer') {
      return numberFor(type, String(v));
    }
  }
  if (typeof v === 'string') {
    if (type === 'DateTime' && DATE_RE.test(v)) return { t: 'datetime', v };
    if (type === 'Uniqueidentifier' && GUID_RE.test(v)) return { t: 'guid', v };
    if (type === 'Image') return { t: 'bytes', v };
    if (type === 'File') return no(`'${attr}' is a file column: its content isn't a column value.`);
  }
  return json;
}

/** A JSON primitive as an SDK value. Objects and mixed arrays have no faithful form. */
function jsonValue(v: unknown, what: string): SdkValue {
  if (v === null || v === undefined) return { t: 'null' };
  if (typeof v === 'boolean') return { t: 'bool', v };
  if (typeof v === 'number') {
    if (!Number.isFinite(v)) return no(`${what} isn't a finite number.`);
    return numberFor(undefined, String(v));
  }
  if (typeof v === 'string') return { t: 'string', v };
  if (Array.isArray(v)) {
    if (v.every((x) => typeof x === 'string')) return { t: 'array', elem: 'string', v };
    if (v.every((x) => Number.isInteger(x))) return { t: 'array', elem: 'int', v: v.map(String) };
  }
  return no(`${what} is an object or a list: its SDK type isn't known from the JSON.`);
}

// ────────────────────────────────────────────────────────────
// Headers → SDK options
// ────────────────────────────────────────────────────────────

const BOOL_PARAMS: Record<string, string> = {
  'mscrm.bypasscustompluginexecution': 'BypassCustomPluginExecution',
  'mscrm.suppressduplicatedetection': 'SuppressDuplicateDetection',
  'mscrm.suppresscallbackregistrationexpanderjob': 'SuppressCallbackRegistrationExpanderJob',
};
const STRING_PARAMS: Record<string, string> = {
  'mscrm.bypassbusinesslogicexecution': 'BypassBusinessLogicExecution',
  'mscrm.bypassbusinesslogicexecutionstepids': 'BypassBusinessLogicExecutionStepIds',
  'mscrm.solutionuniquename': 'SolutionUniqueName',
};
/** Protocol headers the SDK sets itself. */
const IGNORED = new Set(['accept', 'content-type', 'odata-version', 'odata-maxversion']);

function readOpts(headers: Record<string, string>): Opts {
  const o: Opts = {
    params: [],
    dropped: [],
    annotations: false,
    returnRepresentation: false,
  };
  for (const [k, raw] of Object.entries(headers)) {
    if (!raw || isCredentialHeader(k)) continue;
    const key = k.toLowerCase();
    const v = raw.trim();
    if (IGNORED.has(key)) continue;
    if (key in BOOL_PARAMS && /^(true|false)$/i.test(v)) {
      o.params.push([BOOL_PARAMS[key], { t: 'bool', v: v.toLowerCase() === 'true' }]);
    } else if (key in STRING_PARAMS) {
      o.params.push([STRING_PARAMS[key], { t: 'string', v }]);
    } else if (key === 'mscrmcallerid' && GUID_RE.test(v)) {
      o.callerId = v;
    } else if (key === 'callerobjectid' && GUID_RE.test(v)) {
      o.callerObjectId = v;
    } else if (key === 'if-match') {
      o.ifMatch = v;
    } else if (key === 'if-none-match') {
      o.ifNoneMatch = v;
    } else if (key === 'prefer') {
      const rest: string[] = [];
      for (const part of topLevelSplit(v, ',').map((s) => s.trim())) {
        const size = /^odata\.maxpagesize\s*=\s*(\d+)$/i.exec(part);
        if (size) o.pageSize = Number(size[1]);
        else if (/^odata\.include-annotations\s*=/i.test(part)) o.annotations = true;
        else if (/^return\s*=\s*representation$/i.test(part)) o.returnRepresentation = true;
        else if (part) rest.push(part);
      }
      if (rest.length) o.dropped.push([k, rest.join(',')]);
    } else {
      o.dropped.push([k, v]);
    }
  }
  return o;
}

/** `W/"123"` / `"123"` → `123` (the SDK's RowVersion). `*` → undefined. */
function rowVersionOf(etag: string | undefined): string | undefined {
  if (!etag || etag === '*') return undefined;
  return etag.replace(/^W\//, '').replace(/^"|"$/g, '');
}

// ────────────────────────────────────────────────────────────
// Plan builder
// ────────────────────────────────────────────────────────────

function buildPlan(i: CodegenInputs): Built {
  const opts = readOpts(i.headers);
  const notes: string[] = [];
  try {
    return { plan: planFor(i, opts, notes), opts, notes };
  } catch (e) {
    if (!(e instanceof Unsupported)) throw e;
    return {
      plan: hasBinary(i) ? { k: 'binary', reason: e.message } : { k: 'webapi', reason: e.message },
      opts,
      notes: [],
    };
  }
}

function hasBinary(i: CodegenInputs): boolean {
  if (i.rawBody || i.multiRequests?.some((r) => r.rawBody)) return true;
  // `/<column>/$value` is raw bytes for a file or image column; for a plain
  // column it's just the value, which Retrieve reads.
  const value = /\/([A-Za-z_][A-Za-z0-9_]*)\/\$value$/.exec(i.built.relativeUrl.split('?')[0]);
  if (!value) return false;
  const type = colMeta(i.built.entityLogical, value[1])?.attributeType;
  return !type || type === 'File' || type === 'Image';
}

function planFor(i: CodegenInputs, opts: Opts, notes: string[]): Plan {
  if (!i.built.relativeUrl && !i.isNextLink) return no("The request isn't complete yet.");
  if (hasBinary(i)) {
    return no(
      'File and image bytes: the SDK moves them with the InitializeFileBlocksUpload / UploadBlock / CommitFileBlocksUpload (or …Download / DownloadBlock) messages, not a binary body.',
    );
  }
  if (i.isNextLink) {
    return no(
      'A nextLink is a Web API paging URL. With the SDK, page with PagingInfo.PagingCookie instead.',
    );
  }
  if (i.multiRequests && i.multiRequests.length > 1) return relatePlan(i, notes);

  const path = parsePath(i.built.relativeNoBase);
  const nsSeg = path.segs.find((s) => s.startsWith('Microsoft.Dynamics.CRM.'));
  // An unbound message: no table, or a path that isn't the table's (Bulk
  // Delete and Merge carry their table in `entitySet` but post to
  // /BulkDelete, /Merge).
  const unbound = !i.built.entitySet || path.set !== i.built.entitySet;
  const tableOf = () => checkName(i.built.entityLogical || logicalOfSet(path.set, notes), 'Table');

  switch (i.method) {
    case 'GET': {
      if (nsSeg || unbound) return executePlan(i, path, notes, true);
      const entity = tableOf();
      if (!path.key) {
        if (path.segs.length) return no(`'/${path.segs.join('/')}' after the table name.`);
        return { k: 'query', q: queryModel(i, entity, opts, notes) };
      }
      return retrievePlan(i, entity, path, notes);
    }
    case 'POST': {
      if (nsSeg || unbound) return executePlan(i, path, notes, false);
      const entity = tableOf();
      if (path.segs[path.segs.length - 1] === '$ref' && path.segs.length === 2) {
        return relateOne(i, entity, path, 'associate', i.body?.['@odata.id'], notes);
      }
      if (path.key || path.segs.length) {
        return no('A POST to a related collection creates the row through a navigation property.');
      }
      if (rejectQuery(i)) return no(rejectQuery(i)!);
      return { k: 'create', entity, attrs: bodyAttrs(entity, i.body ?? {}, notes) };
    }
    case 'PATCH': {
      const entity = tableOf();
      if (path.segs.length) return no(`A PATCH to '/${path.segs.join('/')}'.`);
      if (rejectQuery(i)) return no(rejectQuery(i)!);
      const target = targetOf(entity, path.key);
      const attrs = bodyAttrs(entity, i.body ?? {}, notes);
      if (opts.ifNoneMatch === '*') {
        if (target.kind !== 'id') {
          return no(
            'Create-only by alternate key: the SDK creates by id, and Upsert also updates.',
          );
        }
        return { k: 'create', entity, attrs, id: target.id };
      }
      if (opts.ifNoneMatch) return no(`If-None-Match: ${opts.ifNoneMatch} on a write.`);
      if (opts.ifMatch) {
        return { k: 'update', entity, target, attrs, rowVersion: rowVersionOf(opts.ifMatch) };
      }
      return { k: 'upsert', entity, target, attrs };
    }
    case 'PUT': {
      const entity = tableOf();
      const target = targetOf(entity, path.key);
      // Single-valued $ref: PUT /<set>(<id>)/<nav>/$ref { @odata.id }.
      if (path.segs.length === 2 && path.segs[1] === '$ref') {
        const nav = lookupNav(entity, path.segs[0], notes);
        const ref = refFromUrl(String(i.body?.['@odata.id'] ?? ''), notes);
        return update(entity, target, [{ name: nav, value: ref }], opts);
      }
      if (path.segs.length === 1) {
        const col = checkName(path.segs[0], 'Column');
        const value = columnValue(colMeta(entity, col), col, i.body?.value ?? null);
        return update(entity, target, [{ name: col, value }], opts);
      }
      return no("A PUT that isn't a single column or a lookup $ref.");
    }
    case 'DELETE': {
      const entity = tableOf();
      if (path.segs[path.segs.length - 1] === '$ref') {
        if (path.segs.length === 2 && !/\(/.test(path.segs[0])) {
          const idParam = i.built.queryParts.find((p) => p.key === '$id');
          if (idParam) {
            return relateOne(i, entity, path, 'disassociate', safeDecode(idParam.value), notes);
          }
          // Single-valued: clear the lookup.
          const nav = lookupNav(entity, path.segs[0], notes);
          return update(
            entity,
            targetOf(entity, path.key),
            [{ name: nav, value: { t: 'null' } }],
            opts,
          );
        }
        return relateOne(i, entity, path, 'disassociate', undefined, notes);
      }
      const target = targetOf(entity, path.key);
      if (path.segs.length === 1) {
        // DELETE /<set>(<id>)/<column> clears one column. File and image
        // columns have their own messages, so only a known plain column maps.
        const col = checkName(path.segs[0], 'Column');
        if (!plainColumn(entity, col)) {
          return no(
            `Clearing '/${col}': a file, image or lookup segment (or one whose metadata isn't loaded) isn't a plain column value.`,
          );
        }
        return update(entity, target, [{ name: col, value: { t: 'null' } }], opts);
      }
      if (path.segs.length) return no(`A DELETE of '/${path.segs.join('/')}'.`);
      return { k: 'delete', entity, target, rowVersion: rowVersionOf(opts.ifMatch) };
    }
  }
  return no(`The ${i.method} method.`);
}

/**
 * A plain column segment (`/<set>(<id>)/<column>`). A lookup's segment is
 * its navigation property, which reads or changes the related row, and file
 * and image columns have their own messages.
 */
function plainColumn(entity: string, col: string): boolean {
  const meta = colMeta(entity, col);
  if (!meta) return false;
  if (['File', 'Image', 'Lookup', 'Customer', 'Owner'].includes(meta.attributeType)) return false;
  return !findTable(entity)?.navigationProperties.some((n) => n.name === col);
}

function update(entity: string, target: Target, attrs: Attr[], opts: Opts): Plan {
  return { k: 'update', entity, target, attrs, rowVersion: rowVersionOf(opts.ifMatch) };
}

/** A query string on a write (e.g. `?$select=` with return=representation). */
function rejectQuery(i: CodegenInputs): string | null {
  const keys = i.built.queryParts.map((p) => p.key).filter((k) => k !== '$select');
  return keys.length ? `The query option ${keys.join(', ')} on a write.` : null;
}

/** Attribute name behind a single-valued navigation property. */
function lookupNav(entity: string, nav: string, notes: string[]): string {
  checkName(nav, 'Navigation property');
  const meta = findTable(entity)?.navigationProperties.find(
    (n) => n.name === nav && n.cardinality === 'ManyToOne',
  );
  if (meta?.referencingAttribute) return meta.referencingAttribute;
  if (!findTable(entity)) {
    notes.push(
      `Check: '${nav.toLowerCase()}' is taken as the lookup column behind '${nav}' (${entity} metadata isn't loaded).`,
    );
  }
  return nav.toLowerCase();
}

/** The body's columns as typed attributes. */
function bodyAttrs(entity: string, body: Record<string, unknown>, notes: string[]): Attr[] {
  const attrs: Attr[] = [];
  const tbl = findTable(entity);
  let untyped = false;
  for (const [key, v] of Object.entries(body)) {
    if (v === undefined) continue;
    const bind = /^([A-Za-z_][A-Za-z0-9_]*)@odata\.bind$/.exec(key);
    if (bind) {
      const name = lookupNav(entity, bind[1], notes);
      if (v === null) attrs.push({ name, value: { t: 'null' } });
      else if (typeof v === 'string') attrs.push({ name, value: refFromUrl(v, notes) });
      else no(`'${key}' binds a list of rows (a collection-valued navigation property).`);
      continue;
    }
    if (key.includes('@')) return no(`The annotation '${key}' in the body.`);
    checkName(key, 'Column');
    if (v !== null && typeof v === 'object' && !Array.isArray(v)) {
      return no(`'${key}' is a nested row: deep insert creates related rows in one request.`);
    }
    if (Array.isArray(v) && v.some((x) => x !== null && typeof x === 'object')) {
      return no(`'${key}' is a list of rows (deep insert or activity parties).`);
    }
    const meta = tbl ? findColumn(tbl, key) : undefined;
    if (!meta) untyped = true;
    attrs.push({ name: key, value: columnValue(meta, key, v) });
  }
  if (untyped) {
    notes.push(
      `Some column types aren't loaded, so their values follow the JSON: wrap a choice in OptionSetValue and money in Money.`,
    );
  }
  return attrs;
}

// ── Retrieve ────────────────────────────────────────────────

function selectColumns(i: CodegenInputs): string[] | null {
  const sel = i.built.queryParts.find((p) => p.key === '$select');
  if (!sel) return null;
  return safeDecode(sel.value)
    .split(',')
    .map((c) => c.trim())
    .filter(Boolean)
    .map((c) => checkName(bareAttr(c), 'Column'));
}

function retrievePlan(i: CodegenInputs, entity: string, path: ParsedPath, notes: string[]): Plan {
  const target = targetOf(entity, path.key);
  const extra = i.built.queryParts.map((p) => p.key).filter((k) => k !== '$select');
  if (extra.length) {
    return no(
      extra.includes('$expand')
        ? '$expand: Retrieve has no expand (use a QueryExpression with a LinkEntity).'
        : `The query option ${extra.join(', ')} on a single row.`,
    );
  }
  let columns = selectColumns(i);
  if (path.segs.length) {
    // GET /<set>(<id>)/<column>[/$value] — one column, if metadata says it is one.
    const col = path.segs[0];
    if (path.segs.length > 2 || (path.segs[1] && path.segs[1] !== '$value')) {
      return no(`'/${path.segs.join('/')}' after the row.`);
    }
    if (!plainColumn(entity, col)) {
      return no(
        `'/${col}' after the row: related rows, a lookup's related row, or a column whose metadata isn't loaded.`,
      );
    }
    columns = [col];
  }
  if (!columns) notes.push(NO_SELECT);
  return { k: 'retrieve', entity, target, columns };
}

const NO_SELECT = 'No $select: every column comes back. Name the columns you need.';

// ── Query ───────────────────────────────────────────────────

const COMPARISON: Record<string, string> = {
  eq: 'Equal',
  ne: 'NotEqual',
  gt: 'GreaterThan',
  ge: 'GreaterEqual',
  lt: 'LessThan',
  le: 'LessEqual',
};

function queryModel(i: CodegenInputs, entity: string, opts: Opts, notes: string[]): QueryModel {
  const q: QueryModel = {
    entity,
    columns: null,
    filter: null,
    orders: [],
    count: false,
    countOnly: false,
  };
  for (const { key, value: rawValue } of i.built.queryParts) {
    const value = safeDecode(rawValue);
    switch (key) {
      case '$select':
        q.columns = selectColumns(i);
        break;
      case '$filter':
        q.filter = filterModel(entity, value);
        break;
      case '$orderby':
        for (const seg of value
          .split(',')
          .map((s) => s.trim())
          .filter(Boolean)) {
          const m = /^(\S+)(?:\s+(asc|desc))?$/i.exec(seg);
          if (!m) return no(`The $orderby item '${seg}'.`);
          if (m[1].includes('/'))
            return no(`Ordering by '${m[1]}' (a related column) needs a LinkEntity.`);
          q.orders.push({
            attr: checkName(bareAttr(m[1]), 'Column'),
            desc: /desc/i.test(m[2] ?? ''),
          });
        }
        break;
      case '$top':
        if (!INT_RE.test(value)) return no(`$top=${value}.`);
        q.top = Number(value);
        break;
      case '$count':
        q.count = value === 'true';
        break;
      case '$expand':
        return no('$expand: QueryExpression joins related tables with LinkEntity instead.');
      case '$apply':
        return no('$apply: aggregate with FetchXML (FetchExpression) instead.');
      case 'savedQuery':
      case 'userQuery':
        return no(`${key}: run the view's FetchXML with a FetchExpression.`);
      default:
        return no(`The query option ${key}.`);
    }
  }
  q.pageSize = opts.pageSize;
  if (!q.columns && !(q.count && q.top === 1)) notes.push(NO_SELECT);
  q.countOnly =
    q.count &&
    q.top === 1 &&
    Object.values(i.headers).some((v) => /totalrecordcountlimitexceeded/.test(v));
  return q;
}

function filterModel(entity: string, filter: string): Cond {
  const { group, quoted, warnings, errors } = parseFilterForCodegen(filter);
  const issue = errors[0] ?? warnings[0];
  if (issue) return no(`$filter: ${issue.message}`);
  return groupCond(entity, group, quoted);
}

function groupCond(entity: string, g: FilterGroup, quoted: Set<string>): Cond {
  if (g.negated) {
    return no('not (…): QueryExpression and FetchXML filters have no NOT over a group.');
  }
  return {
    type: 'group',
    or: g.combinator === 'or',
    children: g.rules.map((n) => nodeCond(entity, n, quoted)),
  };
}

function nodeCond(entity: string, n: FilterNode, quoted: Set<string>): Cond {
  if (n.type === 'group') return groupCond(entity, n, quoted);
  if (n.type === 'lambda') {
    return no(`${n.nav}/${n.lambda}(…): filter on related rows with a LinkEntity.`);
  }
  if (n.col.includes('/')) {
    return no(`'${n.col}' is a related column: filter on it through a LinkEntity.`);
  }
  const attr = checkName(bareAttr(n.col), 'Column');
  const type = colMeta(entity, attr)?.attributeType;
  if (n.type === 'rule') {
    if (n.op in COMPARISON) {
      if (n.valKind === 'column') {
        return {
          type: 'cond',
          attr,
          op: COMPARISON[n.op],
          values: [],
          column: checkName(bareAttr(n.val ?? ''), 'Column'),
        };
      }
      const value = literal(type, n.val ?? '', quoted.has(n.id));
      return { type: 'cond', attr, op: COMPARISON[n.op], values: [value] };
    }
    if (n.op === 'is-null') return { type: 'cond', attr, op: 'Null', values: [] };
    if (n.op === 'is-not-null') return { type: 'cond', attr, op: 'NotNull', values: [] };
    const v = n.val ?? '';
    const pattern =
      n.op === 'contains'
        ? `%${v}%`
        : n.op === 'startswith'
          ? `${v}%`
          : n.op === 'endswith'
            ? `%${v}`
            : null;
    if (pattern === null) return no(`The $filter operator ${n.op}.`);
    return {
      type: 'cond',
      attr,
      op: n.negated ? 'NotLike' : 'Like',
      values: [{ t: 'string', v: pattern }],
    };
  }
  // Dataverse query function → the ConditionOperator of the same name.
  const def = findOperator(n.op);
  if (!def || !def.kind.startsWith('dv-fn')) return no(`The query function ${n.op}.`);
  if (n.negated) return no(`not ${n.op}(…).`);
  const raws = n.values ?? n.vals ?? (n.val !== undefined && n.val !== '' ? [n.val] : []);
  const values = raws.map((r): SdkValue => {
    if (def.intValue || def.kind === 'dv-fn-1-int' || def.kind === 'dv-fn-2') {
      return INT_RE.test(r)
        ? { t: 'int', v: r }
        : no(`${n.op} expects a whole number, not '${r}'.`);
    }
    if (def.kind === 'dv-fn-guid') {
      return GUID_RE.test(r) ? { t: 'guid', v: r } : no(`${n.op} expects a GUID, not '${r}'.`);
    }
    return looseValue(type, r);
  });
  return { type: 'cond', attr, op: n.op, values };
}

// ── Associate / Disassociate ────────────────────────────────

function relationshipOf(entity: string, nav: string, notes: string[]) {
  checkName(nav, 'Navigation property');
  const meta = findTable(entity)?.navigationProperties.find((n) => n.name === nav);
  if (!meta) {
    notes.push(
      `Check: '${nav}' is used as the relationship name; the SDK wants the relationship's schema name (${entity} metadata isn't loaded).`,
    );
  }
  return {
    name: meta?.relationshipName ?? nav,
    target: meta?.targetEntity,
    cardinality: meta?.cardinality,
  };
}

/** One $ref request: POST <nav>/$ref, DELETE <nav>(<id>)/$ref or <nav>/$ref?$id=. */
function relateOne(
  i: CodegenInputs,
  entity: string,
  path: ParsedPath,
  k: 'associate' | 'disassociate',
  refUrl: unknown,
  notes: string[],
): Plan {
  const source = targetOf(entity, path.key);
  if (source.kind !== 'id') return no('Relating a row addressed by alternate key.');
  const m = /^([A-Za-z_][A-Za-z0-9_]*)(?:\(([^)]*)\))?$/.exec(path.segs[0]);
  if (!m) return no(`'${path.segs[0]}' isn't a navigation property.`);
  const rel = relationshipOf(entity, m[1], notes);
  let target: Extract<SdkValue, { t: 'ref' }>;
  if (typeof refUrl === 'string') {
    target = refFromUrl(refUrl, notes);
  } else if (m[2] && GUID_RE.test(m[2])) {
    if (!rel.target) {
      return no(`Disassociate needs the related table of '${m[1]}' (metadata isn't loaded).`);
    }
    target = { t: 'ref', entity: rel.target, id: m[2] };
  } else {
    return no('The $ref request has no related row.');
  }
  // A self-referencing relationship needs the source row's role. Through a
  // 1:N navigation property the source is the parent (Referenced); a wrong
  // role would make the related row the parent instead.
  let selfRole: 'Referenced' | 'Referencing' | undefined;
  if (target.entity === entity) {
    if (rel.cardinality === 'OneToMany') selfRole = 'Referenced';
    else if (rel.cardinality === 'ManyToMany') selfRole = 'Referencing';
    else {
      return no(
        `'${m[1]}' relates ${entity} rows to each other, and which side the row is on isn't known (metadata isn't loaded).`,
      );
    }
  }
  return {
    k,
    entity,
    id: source.id,
    relationship: rel.name,
    targets: [target],
    atomic: true,
    selfRole,
  };
}

/** Associate / Disassociate over a collection: one $ref request per target. */
function relatePlan(i: CodegenInputs, notes: string[]): Plan {
  const reqs = i.multiRequests ?? [];
  const plans = reqs.map((r) => {
    const path = parsePath(r.relativeUrl.replace(ENV.apiBase, ''));
    const entity = findTableBySet(path.set)?.logicalName ?? i.built.entityLogical;
    if (!entity) return no("The table of a multi-request step isn't known.");
    if (r.method === 'POST' && path.segs[path.segs.length - 1] === '$ref') {
      return relateOne(i, entity, path, 'associate', r.body?.['@odata.id'], notes);
    }
    if (r.method === 'DELETE' && path.segs[path.segs.length - 1] === '$ref') {
      return relateOne(i, entity, path, 'disassociate', undefined, notes);
    }
    return no(`A ${r.method} step in a multi-request.`);
  }) as Extract<Plan, { k: 'associate' | 'disassociate' }>[];
  const first = plans[0];
  const same = plans.every(
    (p) =>
      p.k === first.k &&
      p.entity === first.entity &&
      p.id === first.id &&
      p.relationship === first.relationship,
  );
  if (!same) return no("The requests don't all relate the same row through one relationship.");
  // Dedupe notes repeated per step.
  const unique = [...new Set(notes)];
  notes.length = 0;
  notes.push(...unique);
  return {
    ...first,
    targets: plans.flatMap((p) => p.targets),
    atomic: !!i.changeSet,
  };
}

// ── Actions / functions ─────────────────────────────────────

const WORKFLOW_NAMES = new Set(['ExecuteWorkflow']);

function executePlan(i: CodegenInputs, path: ParsedPath, notes: string[], isFn: boolean): Plan {
  // Bound: /<set>(<id>)/Microsoft.Dynamics.CRM.<name>(<args>). Unbound:
  // /<name>(<args>), where parsePath read the arguments as the key.
  const nsSeg = path.segs.find((s) => s.startsWith('Microsoft.Dynamics.CRM.'));
  if (!nsSeg && path.segs.length) return no(`'/${path.segs.join('/')}' after ${path.set}.`);
  const head = nsSeg
    ? nsSeg.slice('Microsoft.Dynamics.CRM.'.length)
    : `${path.set}${path.key !== undefined ? `(${path.key})` : ''}`;
  const boundKey = nsSeg ? path.key : undefined;
  const m = /^([A-Za-z_][A-Za-z0-9_]*)(?:\((.*)\))?$/s.exec(head);
  if (!m) return no(`'${head}' isn't a message name.`);
  const name = m[1];
  const csdl = findActionSync(name);
  const csdlParam = (p: string) => csdl?.parameters.find((x) => x.name === p);
  const params: [string, SdkValue][] = [];

  if (isFn) {
    const aliases = new Map(i.built.queryParts.map((p) => [p.key, safeDecode(p.value)]));
    const extra = [...aliases.keys()].filter((k) => !k.startsWith('@'));
    if (extra.length) return no(`The query option ${extra.join(', ')} on a function.`);
    for (const arg of m[2] ? topLevelSplit(m[2], ',') : []) {
      const eq = arg.indexOf('=');
      if (eq < 0) return no(`The function argument '${arg}'.`);
      const pname = checkName(arg.slice(0, eq).trim(), 'Parameter');
      let raw = arg.slice(eq + 1).trim();
      if (raw.startsWith('@')) raw = aliases.get(raw) ?? no(`No value for the alias ${raw}.`);
      params.push([pname, paramValue(csdlParam(pname), pname, fnLiteral(raw))]);
    }
  } else {
    if (i.built.queryParts.length) return no('A query string on an action.');
    for (const [pname, v] of Object.entries(i.body ?? {})) {
      if (v === undefined) continue;
      checkName(pname, 'Parameter');
      params.push([pname, paramValue(csdlParam(pname), pname, v)]);
    }
  }

  let target: Extract<SdkValue, { t: 'ref' }> | undefined;
  if (boundKey) {
    const entity = checkName(i.built.entityLogical || logicalOfSet(path.set, notes), 'Table');
    if (!GUID_RE.test(boundKey)) return no('A message bound to a row addressed by alternate key.');
    if (WORKFLOW_NAMES.has(name) && entity === 'workflow') {
      // ExecuteWorkflow takes the workflow as WorkflowId, and EntityId as a Guid.
      params.unshift(['WorkflowId', { t: 'guid', v: boundKey }]);
      for (const p of params) {
        if (p[0] === 'EntityId' && p[1].t === 'string' && GUID_RE.test(p[1].v)) {
          p[1] = { t: 'guid', v: p[1].v };
        }
      }
    } else {
      target = { t: 'ref', entity, id: boundKey };
      if (csdl?.source !== 'custom-api') {
        notes.push(
          `Check: the bound row is passed as Target. Custom APIs take it that way; a built-in message may name it differently.`,
        );
      }
    }
  } else if (nsSeg) {
    notes.push(
      `Check: ${name} is bound to a table, not a row; the SDK request has no Target for it.`,
    );
  }

  if (!csdl) {
    notes.push(
      `Parameter types follow the request's JSON (the ${name} definition isn't loaded): a GUID text stays a string, so pass a Guid where the message expects one.`,
    );
  } else if (csdl.source === 'oob') {
    notes.push(
      `${name} is a built-in message: Microsoft.Crm.Sdk.Messages has a typed ${name}Request, and an SDK parameter type can differ from the Web API's (e.g. a status as OptionSetValue).`,
    );
  }
  return { k: 'execute', name, params, target };
}

/**
 * A function-URL literal as a JSON-ish value. An unquoted GUID and a number
 * are marked, the number with its source text so nothing is rounded.
 */
function fnLiteral(raw: string): unknown {
  if (raw.length >= 2 && raw.startsWith("'") && raw.endsWith("'")) {
    return raw.slice(1, -1).replace(/''/g, "'");
  }
  if (GUID_RE.test(raw)) return { __guid: raw };
  if (raw === 'true' || raw === 'false') return raw === 'true';
  if (raw === 'null') return null;
  if (NUM_RE.test(raw)) return { __num: raw };
  if (raw.startsWith('[')) {
    try {
      return JSON.parse(raw);
    } catch {
      return no(`The function argument ${raw}.`);
    }
  }
  return no(`The function argument ${raw} (an enum or object literal).`);
}

const EDM: Partial<Record<ActionParam['type'], SdkValue['t']>> = {
  'Edm.String': 'string',
  'Edm.Int32': 'int',
  'Edm.Int64': 'long',
  'Edm.Boolean': 'bool',
  'Edm.Decimal': 'decimal',
  'Edm.Double': 'double',
  'Edm.Guid': 'guid',
  'Edm.DateTimeOffset': 'datetime',
  'Edm.Binary': 'bytes',
};

/** An action / function parameter, typed from its CSDL definition when loaded. */
function paramValue(def: ActionParam | undefined, name: string, v: unknown): SdkValue {
  const guid = v && typeof v === 'object' && '__guid' in v ? String(v.__guid) : null;
  const num =
    v && typeof v === 'object' && '__num' in v
      ? String(v.__num)
      : typeof v === 'number' && Number.isFinite(v)
        ? String(v)
        : null;
  if (v === null) return { t: 'null' };
  if (!def) {
    if (guid) return { t: 'guid', v: guid };
    if (num !== null) return numberFor(undefined, num);
    return jsonValue(v, `Parameter '${name}'`);
  }
  const t = EDM[def.type];
  const text = guid ?? num ?? (typeof v === 'string' ? v : null);
  if (t) {
    if (t === 'bool') return typeof v === 'boolean' ? { t, v } : no(`'${name}' isn't true/false.`);
    if (text === null) return no(`'${name}' isn't a ${def.type} value.`);
    if (t === 'int') {
      return INT_RE.test(text) && Math.abs(Number(text)) <= INT32_MAX
        ? { t, v: text }
        : no(`'${name}' isn't a 32-bit whole number.`);
    }
    if (t === 'long') {
      return INT_RE.test(text) && fitsLong(text)
        ? { t, v: text }
        : no(`'${name}' isn't a 64-bit whole number.`);
    }
    if (t === 'decimal' || t === 'double') {
      return NUM_RE.test(text) ? { t, v: text } : no(`'${name}' isn't a number.`);
    }
    if (t === 'guid' && !GUID_RE.test(text)) return no(`'${name}' isn't a GUID.`);
    return { t, v: text } as SdkValue;
  }
  if (def.type === 'OptionSetValue' && !def.enumType && num !== null && INT_RE.test(num)) {
    return { t: 'option', v: num };
  }
  const elem =
    def.type === 'Collection(Edm.String)'
      ? 'string'
      : def.type === 'Collection(Edm.Int32)'
        ? 'int'
        : def.type === 'Collection(Edm.Guid)'
          ? 'guid'
          : null;
  if (elem && Array.isArray(v)) {
    const items = v.map(String);
    const ok = items.every((x) =>
      elem === 'int'
        ? INT_RE.test(x) && Math.abs(Number(x)) <= INT32_MAX
        : elem === 'guid'
          ? GUID_RE.test(x)
          : true,
    );
    if (ok) return { t: 'array', elem, v: items };
  }
  return no(
    `Parameter '${name}' is ${def.enumType ? `an enum (${def.enumType})` : `a ${def.type}`}: its SDK type isn't known from the Web API.`,
  );
}

// ────────────────────────────────────────────────────────────
// C# rendering
// ────────────────────────────────────────────────────────────

interface CsCtx {
  usings: Set<string>;
}

const csStr = (s: string) => `"${csharpString(s)}"`;
const csGuid = (g: string) => `new Guid(${csStr(g)})`;

function csValue(v: SdkValue, c: CsCtx): string {
  switch (v.t) {
    case 'null':
      return 'null';
    case 'string':
      return csStr(v.v);
    case 'int':
      return v.v;
    case 'long':
      return `${v.v}L`;
    case 'decimal':
      return `${v.v}m`;
    case 'double':
      return `${v.v}d`;
    case 'bool':
      return String(v.v);
    case 'guid':
      return csGuid(v.v);
    case 'datetime':
      c.usings.add('System.Globalization');
      return `DateTime.Parse(${csStr(v.v)}, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind)`;
    case 'option':
      return `new OptionSetValue(${v.v})`;
    case 'options':
      return `new OptionSetValueCollection { ${v.v.map((n) => `new OptionSetValue(${n})`).join(', ')} }`;
    case 'money':
      return `new Money(${v.v}m)`;
    case 'ref':
      return `new EntityReference(${csStr(v.entity)}, ${csGuid(v.id)})`;
    case 'bytes':
      return `Convert.FromBase64String(${csStr(v.v)})`;
    case 'array': {
      const items = v.v.map((x) =>
        v.elem === 'string' ? csStr(x) : v.elem === 'guid' ? csGuid(x) : x,
      );
      const type = v.elem === 'string' ? 'string' : v.elem === 'guid' ? 'Guid' : 'int';
      return items.length ? `new ${type}[] { ${items.join(', ')} }` : `Array.Empty<${type}>()`;
    }
  }
}

const csKeys = (pairs: [string, SdkValue][], c: CsCtx) =>
  `new KeyAttributeCollection { ${pairs.map(([k, v]) => `[${csStr(k)}] = ${csValue(v, c)}`).join(', ')} }`;

function csRef(entity: string, t: Target, c: CsCtx): string {
  return t.kind === 'id'
    ? `new EntityReference(${csStr(entity)}, ${csGuid(t.id)})`
    : `new EntityReference(${csStr(entity)}, ${csKeys(t.pairs, c)})`;
}

function csEntity(entity: string, t: Target | undefined, attrs: Attr[], c: CsCtx, id?: string) {
  const ctor = t
    ? t.kind === 'id'
      ? `new Entity(${csStr(entity)}, ${csGuid(t.id)})`
      : `new Entity(${csStr(entity)}, ${csKeys(t.pairs, c)})`
    : id
      ? `new Entity(${csStr(entity)}, ${csGuid(id)})`
      : `new Entity(${csStr(entity)})`;
  if (!attrs.length) return `var entity = ${ctor};`;
  return `var entity = ${ctor}\n{\n${attrs
    .map((a) => `    [${csStr(a.name)}] = ${csValue(a.value, c)},`)
    .join('\n')}\n};`;
}

const csColumns = (cols: string[] | null) =>
  cols ? `new ColumnSet(${cols.map(csStr).join(', ')})` : 'new ColumnSet(true)';

/** `request.Parameters[...] = …` lines for the optional parameters. */
function csParams(o: Opts, c: CsCtx): string {
  return o.params.map(([k, v]) => `request.Parameters[${csStr(k)}] = ${csValue(v, c)};\n`).join('');
}

const CS_ROW_DUMP = (rowVar: string, indent: string) =>
  `${indent}Console.WriteLine($"{${rowVar}.LogicalName} {${rowVar}.Id}");
${indent}foreach (var attribute in ${rowVar}.Attributes)
${indent}{
${indent}    // Choices, lookups and money read better as their formatted value.
${indent}    object shown = ${rowVar}.FormattedValues.TryGetValue(attribute.Key, out var formatted) ? formatted : attribute.Value;
${indent}    Console.WriteLine($"  {attribute.Key}: {shown}");
${indent}}`;

function csFilterLines(cond: Cond, varName: string, lines: string[], n: { i: number }, c: CsCtx) {
  if (cond.type === 'cond') {
    lines.push(`${varName}.AddCondition(${csCondArgs(cond, c)});`);
    return;
  }
  for (const child of cond.children) {
    if (child.type === 'cond') {
      lines.push(`${varName}.AddCondition(${csCondArgs(child, c)});`);
    } else {
      const v = `filter${++n.i}`;
      lines.push(`var ${v} = new FilterExpression(LogicalOperator.${child.or ? 'Or' : 'And'});`);
      csFilterLines(child, v, lines, n, c);
      lines.push(`${varName}.AddFilter(${v});`);
    }
  }
}

function csCondArgs(cond: Extract<Cond, { type: 'cond' }>, c: CsCtx): string {
  if (cond.column) {
    return `new ConditionExpression(${csStr(cond.attr)}, ConditionOperator.${cond.op}, true, ${csStr(cond.column)})`;
  }
  return [
    csStr(cond.attr),
    `ConditionOperator.${cond.op}`,
    ...cond.values.map((v) => csValue(v, c)),
  ].join(', ');
}

function csQuery(q: QueryModel, o: Opts, c: CsCtx, notes: string[]): string {
  c.usings.add('Microsoft.Xrm.Sdk.Query');
  const init: string[] = [`    ColumnSet = ${csColumns(q.columns)},`];
  const paging = pagingOf(q, notes);
  if (paging.top !== undefined) init.push(`    TopCount = ${paging.top},`);
  if (paging.page) {
    init.push(
      `    PageInfo = new PagingInfo { PageNumber = 1, Count = ${paging.page.size}${paging.page.total ? ', ReturnTotalRecordCount = true' : ''} },`,
    );
  }
  const lines: string[] = [];
  if (q.filter && q.filter.type === 'group') {
    if (q.filter.or) lines.push('query.Criteria.FilterOperator = LogicalOperator.Or;');
    csFilterLines(q.filter, 'query.Criteria', lines, { i: 0 }, c);
  }
  for (const ord of q.orders) {
    lines.push(
      `query.AddOrder(${csStr(ord.attr)}, OrderType.${ord.desc ? 'Descending' : 'Ascending'});`,
    );
  }
  const head = `var query = new QueryExpression(${csStr(q.entity)})\n{\n${init.join('\n')}\n};\n${
    lines.length ? `${lines.join('\n')}\n` : ''
  }`;
  const exec = o.params.length
    ? (c.usings.add('Microsoft.Xrm.Sdk.Messages'),
      `\nvar request = new RetrieveMultipleRequest { Query = query };\n${csParams(o, c)}var response = (RetrieveMultipleResponse)service.Execute(request);\nEntityCollection results = response.EntityCollection;\n`)
    : `\nEntityCollection results = service.RetrieveMultiple(query);\n`;
  const total = `Console.WriteLine($"Total: {results.TotalRecordCount}{(results.TotalRecordCountLimitExceeded ? " (Dataverse stopped counting at its limit)" : "")}");`;
  if (q.countOnly) return `${head}${exec}${total}\n`;
  return `${head}${exec}
Console.WriteLine($"Returned {results.Entities.Count} record(s)");${paging.page?.total ? `\n${total}` : ''}
foreach (Entity row in results.Entities)
{
${CS_ROW_DUMP('row', '    ')}
}
if (results.MoreRecords)
    Console.WriteLine("More rows: set query.PageInfo.PagingCookie = results.PagingCookie, add 1 to PageNumber, and query again.");
`;
}

/**
 * TopCount and PageInfo can't be used together in the SDK, so $top wins over
 * a page size or a total count (with a note), as the snippet can't do both.
 */
function pagingOf(q: QueryModel, notes: string[]) {
  if (q.countOnly) return { page: { size: 1, total: true } };
  if (q.top !== undefined) {
    if (q.count)
      notes.push("The total count isn't requested: the SDK can't combine TopCount with PageInfo.");
    if (q.pageSize) notes.push('odata.maxpagesize is left out: TopCount already limits the rows.');
    return { top: q.top };
  }
  if (q.count || q.pageSize) return { page: { size: q.pageSize ?? 5000, total: q.count } };
  return {};
}

function csHeader(): string {
  return `// Dataverse SDK for .NET — ServiceClient (Microsoft.PowerPlatform.Dataverse.Client)
//   dotnet add package Microsoft.PowerPlatform.Dataverse.Client
// Change the connection string to the sign-in you use (client secret, certificate, …):
//   https://learn.microsoft.com/power-apps/developer/data-platform/xrm-tooling/use-connection-strings-xrm-tooling-connect`;
}

function csConnect(o: Opts): string {
  const host = ENV.host || '<your-org>.crm.dynamics.com';
  return `using var service = new ServiceClient(
    ${csStr(`AuthType=OAuth;Url=https://${host};`)} +
    "AppId=51f81489-12ee-4a9e-aaae-a2591f45987d;RedirectUri=http://localhost;LoginPrompt=Auto");
if (!service.IsReady) throw new InvalidOperationException(service.LastError);
${o.callerId ? `service.CallerId = ${csGuid(o.callerId)};  // MSCRMCallerID\n` : ''}${
    o.callerObjectId
      ? `service.CallerAADObjectId = ${csGuid(o.callerObjectId)};  // CallerObjectId\n`
      : ''
  }`;
}

function optNotes(plan: Plan, o: Opts): string[] {
  const out: string[] = [];
  const reads = plan.k === 'query' || plan.k === 'retrieve';
  if (o.annotations && reads) {
    out.push('Formatted values come back in FormattedValues without a Prefer header.');
  }
  if (
    o.returnRepresentation &&
    (plan.k === 'create' || plan.k === 'update' || plan.k === 'upsert')
  ) {
    out.push('The SDK returns no row from a write (return=representation): Retrieve it after.');
  }
  if (o.pageSize && plan.k !== 'query') out.push('odata.maxpagesize only applies to queries.');
  for (const [k, v] of o.dropped)
    out.push(`Not sent: the SDK has no equivalent of the "${k}: ${v}" header.`);
  return out;
}

const commentBlock = (lines: string[], prefix: string) =>
  lines.map((l) => `${prefix} ${lineComment(l)}`).join('\n');

export function genCsharpSdk(i: CodegenInputs): string {
  const { plan, opts, notes } = buildPlan(i);
  const c: CsCtx = {
    usings: new Set(['Microsoft.PowerPlatform.Dataverse.Client', 'Microsoft.Xrm.Sdk']),
  };
  if (plan.k === 'binary') return csBinary(plan.reason);
  if (plan.k === 'webapi') return csWebApi(i, plan.reason, c);
  const body = csBody(plan, opts, c, notes);
  const allNotes = [...new Set([...notes, ...optNotes(plan, opts)])];
  return `${csHeader()}${allNotes.length ? `\n//\n${commentBlock(allNotes, '//')}` : ''}

${[...c.usings]
  .sort()
  .map((u) => `using ${u};`)
  .join('\n')}

${csConnect(opts)}
${body}`;
}

function csBody(plan: Plan, o: Opts, c: CsCtx, notes: string[]): string {
  const req = (cls: string, init: string, after: string) => {
    c.usings.add('Microsoft.Xrm.Sdk.Messages');
    return `var request = new ${cls} { ${init} };\n${csParams(o, c)}${after}`;
  };
  switch (plan.k) {
    case 'query':
      return csQuery(plan.q, o, c, notes);
    case 'retrieve': {
      c.usings.add('Microsoft.Xrm.Sdk.Query');
      const cols = `var columns = ${csColumns(plan.columns)};\n`;
      const call =
        plan.target.kind === 'id' && !o.params.length
          ? `Entity row = service.Retrieve(${csStr(plan.entity)}, ${csGuid(plan.target.id)}, columns);\n`
          : req(
              'RetrieveRequest',
              `Target = ${csRef(plan.entity, plan.target, c)}, ColumnSet = columns`,
              'Entity row = ((RetrieveResponse)service.Execute(request)).Entity;\n',
            );
      return `${cols}${call}\n${CS_ROW_DUMP('row', '')}\n`;
    }
    case 'create': {
      const ent = `${csEntity(plan.entity, undefined, plan.attrs, c, plan.id)}\n`;
      const call = o.params.length
        ? req(
            'CreateRequest',
            'Target = entity',
            'var response = (CreateResponse)service.Execute(request);\nGuid id = response.id;\n',
          )
        : 'Guid id = service.Create(entity);\n';
      return `${ent}${plan.id ? '// Create only: fails if the row already exists (If-None-Match: *).\n' : ''}${call}Console.WriteLine($"Created ${csharpString(plan.entity)} {id}");\n`;
    }
    case 'update': {
      const ent = `${csEntity(plan.entity, plan.target, plan.attrs, c)}\n`;
      if (plan.rowVersion || o.params.length) {
        const rv = plan.rowVersion
          ? `entity.RowVersion = ${csStr(plan.rowVersion)};  // If-Match\n`
          : '';
        const init = plan.rowVersion
          ? 'Target = entity, ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches'
          : 'Target = entity';
        return `${ent}${rv}${req('UpdateRequest', init, 'service.Execute(request);\n')}Console.WriteLine("Updated");\n`;
      }
      return `${ent}service.Update(entity);\nConsole.WriteLine("Updated");\n`;
    }
    case 'upsert':
      return `${csEntity(plan.entity, plan.target, plan.attrs, c)}
// No If-Match: Dataverse creates the row if it doesn't exist (upsert).
${req('UpsertRequest', 'Target = entity', 'var response = (UpsertResponse)service.Execute(request);\n')}Console.WriteLine($"{(response.RecordCreated ? "Created" : "Updated")} {response.Target.Id}");
`;
    case 'delete': {
      if (plan.target.kind === 'id' && !plan.rowVersion && !o.params.length) {
        return `service.Delete(${csStr(plan.entity)}, ${csGuid(plan.target.id)});\nConsole.WriteLine("Deleted");\n`;
      }
      const ref = `var target = ${csRef(plan.entity, plan.target, c)};\n${
        plan.rowVersion ? `target.RowVersion = ${csStr(plan.rowVersion)};  // If-Match\n` : ''
      }`;
      const init = plan.rowVersion
        ? 'Target = target, ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches'
        : 'Target = target';
      return `${ref}${req('DeleteRequest', init, 'service.Execute(request);\n')}Console.WriteLine("Deleted");\n`;
    }
    case 'associate':
    case 'disassociate':
      return csRelate(plan, o, c);
    case 'execute': {
      const members = [
        ...plan.params.map(([k, v]) => `    [${csStr(k)}] = ${csValue(v, c)},`),
        ...(plan.target ? [`    ["Target"] = ${csValue(plan.target, c)},`] : []),
      ];
      return `var request = new OrganizationRequest(${csStr(plan.name)})${
        members.length ? `\n{\n${members.join('\n')}\n}` : ''
      };
${csParams(o, c)}OrganizationResponse response = service.Execute(request);
foreach (var result in response.Results)
    Console.WriteLine($"{result.Key}: {result.Value}");
`;
    }
  }
  return '';
}

function csRelate(
  plan: Extract<Plan, { k: 'associate' | 'disassociate' }>,
  o: Opts,
  c: CsCtx,
): string {
  const verb = plan.k === 'associate' ? 'Associate' : 'Disassociate';
  const rel = `var relationship = new Relationship(${csStr(plan.relationship)});\n${
    plan.selfRole
      ? `relationship.PrimaryEntityRole = EntityRole.${plan.selfRole};  // self-referencing\n`
      : ''
  }`;
  const refs = plan.targets.map((t) => `    ${csValue(t, c)},`).join('\n');
  const source = `${csStr(plan.entity)}, ${csGuid(plan.id)}`;
  if (plan.atomic || plan.targets.length === 1) {
    const related = `var related = new EntityReferenceCollection\n{\n${refs}\n};\n`;
    const call = o.params.length
      ? (c.usings.add('Microsoft.Xrm.Sdk.Messages'),
        `var request = new ${verb}Request\n{\n    Target = new EntityReference(${source}),\n    Relationship = relationship,\n    RelatedEntities = related,\n};\n${csParams(o, c)}service.Execute(request);\n`)
      : `service.${verb}(${source}, relationship, related);\n`;
    return `${rel}${related}${plan.targets.length > 1 ? '// One request: Dataverse applies all of them or none.\n' : ''}${call}Console.WriteLine("${verb}d ${plan.targets.length} row(s)");\n`;
  }
  const call = o.params.length
    ? (c.usings.add('Microsoft.Xrm.Sdk.Messages'),
      `        var request = new ${verb}Request\n        {\n            Target = new EntityReference(${source}),\n            Relationship = relationship,\n            RelatedEntities = new EntityReferenceCollection { target },\n        };\n${csParams(
        o,
        c,
      )
        .split('\n')
        .filter(Boolean)
        .map((l) => `        ${l}\n`)
        .join('')}        service.Execute(request);\n`)
    : `        service.${verb}(${source}, relationship, new EntityReferenceCollection { target });\n`;
  return `${rel}// One request per row, in order, as the Web API sends them.
var targets = new[]
{
${refs}
};
int ok = 0;
foreach (var target in targets)
{
    try
    {
${call}        ok++;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"{target.LogicalName} {target.Id}: {ex.Message}");
    }
}
Console.WriteLine($"{ok}/{targets.Length} requests succeeded");
`;
}

/** No SDK form: send the Web API request itself through ServiceClient. */
function csWebApi(i: CodegenInputs, reason: string, c: CsCtx): string {
  c.usings.add('System.Net.Http');
  const opts = readOpts({});
  const path =
    i.isNextLink && i.rawNextLink
      ? i.rawNextLink
          .trim()
          .replace(/^https?:\/\/[^/]+/, '')
          .replace(/^\/?api\/data\/v\d+(\.\d+)?\//, '')
      : encodeSpaces(i.built.relativeNoBase.replace(/^\//, ''));
  if (!path) {
    // Nothing to send yet (no table, action or record picked).
    return `${csHeader()}\n//\n// ${lineComment(reason)} The snippet appears once it is.\n`;
  }
  const headers = Object.entries(i.headers).filter(
    ([k, v]) => v && !isCredentialHeader(k) && k.toLowerCase() !== 'content-type',
  );
  const ctype = Object.entries(i.headers).find(([k]) => k.toLowerCase() === 'content-type')?.[1];
  const multi = i.multiRequests && i.multiRequests.length > 1;
  const hasBody = i.method !== 'GET' && i.method !== 'DELETE';
  const body = hasBody
    ? `@"${csharpVerbatimEscape(JSON.stringify(i.body ?? {}, null, 2))}"`
    : 'string.Empty';
  const headerInit = headers.length
    ? `new Dictionary<string, List<string>>\n{\n${headers
        .map(([k, v]) => `    [${csStr(k)}] = new List<string> { ${csStr(v)} },`)
        .join('\n')}\n}`
    : 'new Dictionary<string, List<string>>()';
  return `${csHeader()}
//
// No faithful SDK form for this request — ${lineComment(reason)}
// ServiceClient.ExecuteWebRequest sends the Web API request as-is, with the client's sign-in.${
    multi
      ? `\n// This is the first of ${i.multiRequests!.length} requests; the C# (HttpClient) tab sends them all.`
      : ''
  }

${[...c.usings]
  .sort()
  .map((u) => `using ${u};`)
  .join('\n')}

${csConnect(opts)}
var headers = ${headerInit};
using HttpResponseMessage response = service.ExecuteWebRequest(
    new HttpMethod(${csStr(i.method.toUpperCase())}),
    ${csStr(path)},
    ${body},
    headers${hasBody ? `,\n    ${csStr(ctype ?? 'application/json')}` : ''});
Console.WriteLine($"{(int)response.StatusCode} {response.ReasonPhrase}");
Console.WriteLine(await response.Content.ReadAsStringAsync());
`;
}

function csBinary(reason: string): string {
  return `${csHeader()}
//
// No faithful SDK form for this request — ${lineComment(reason)}
//   Upload:   InitializeFileBlocksUploadRequest → UploadBlockRequest (per block) → CommitFileBlocksUploadRequest
//   Download: InitializeFileBlocksDownloadRequest → DownloadBlockRequest
//   https://learn.microsoft.com/power-apps/developer/data-platform/file-column-data
// An image column can also be set as a byte[] column value with Update.
// The C# (HttpClient) tab sends this exact Web API request.
`;
}

// ────────────────────────────────────────────────────────────
// PowerShell rendering — Microsoft.Xrm.Data.PowerShell
// ────────────────────────────────────────────────────────────

const psGuid = (g: string) => `[guid]${psQuote(g)}`;

function psValue(v: SdkValue): string {
  switch (v.t) {
    case 'null':
      return '$null';
    case 'string':
      return psQuote(v.v);
    case 'int':
      return `[int]${v.v}`;
    case 'long':
      return `[long]${v.v}`;
    case 'decimal':
      return `[decimal]${psQuote(v.v)}`;
    case 'double':
      return `[double]${psQuote(v.v)}`;
    case 'bool':
      return v.v ? '$true' : '$false';
    case 'guid':
      return psGuid(v.v);
    case 'datetime':
      return `[datetime]::Parse(${psQuote(v.v)}, [cultureinfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)`;
    case 'option':
      return `(New-CrmOptionSetValue -Value ${v.v})`;
    case 'options':
      return `[Microsoft.Xrm.Sdk.OptionSetValueCollection]::new([Microsoft.Xrm.Sdk.OptionSetValue[]]@(${v.v
        .map((n) => `(New-CrmOptionSetValue -Value ${n})`)
        .join(', ')}))`;
    case 'money':
      // New-CrmMoney takes a [double]; a decimal keeps every digit.
      return `[Microsoft.Xrm.Sdk.Money]::new([decimal]${psQuote(v.v)})`;
    case 'ref':
      return `(New-CrmEntityReference -EntityLogicalName ${psQuote(v.entity)} -Id ${psQuote(v.id)})`;
    case 'bytes':
      return `[Convert]::FromBase64String(${psQuote(v.v)})`;
    case 'array': {
      const items = v.v.map((x) => (v.elem === 'int' ? x : psQuote(x)));
      return `[${v.elem}[]]@(${items.join(', ')})`;
    }
  }
}

function psFields(attrs: Attr[]): string {
  if (!attrs.length) return '$fields = @{}\n';
  return `$fields = @{\n${attrs.map((a) => `  ${psQuote(a.name)} = ${psValue(a.value)}`).join('\n')}\n}\n`;
}

const psParams = (o: Opts) =>
  o.params.map(([k, v]) => `$request.Parameters[${psQuote(k)}] = ${psValue(v)}\n`).join('');

function psHeader(): string {
  return `# Dataverse — Microsoft.Xrm.Data.PowerShell (Windows PowerShell 5.1)
#   Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser`;
}

function psConnect(o: Opts): string {
  const host = ENV.host || '<your-org>.crm.dynamics.com';
  return `Import-Module Microsoft.Xrm.Data.PowerShell
$conn = Connect-CrmOnline -ServerUrl ${psQuote(host)} -ForceOAuth
${o.callerId ? `$conn.CallerId = ${psGuid(o.callerId)}  # MSCRMCallerID\n` : ''}${
    o.callerObjectId
      ? `$conn.CallerAADObjectId = ${psGuid(o.callerObjectId)}  # CallerObjectId\n`
      : ''
  }`;
}

export function genPowershellXrm(i: CodegenInputs): string {
  const { plan, opts, notes } = buildPlan(i);
  if (plan.k === 'binary' || plan.k === 'webapi') return psFallback(plan.reason);
  let body: string;
  try {
    body = psBody(plan, opts, notes);
  } catch (e) {
    // FetchXML can't say everything a QueryExpression can (e.g. EqualRoleBusinessId).
    if (!(e instanceof Unsupported)) throw e;
    return psFallback(e.message);
  }
  const allNotes = [...new Set([...notes, ...optNotes(plan, opts)])];
  return `${psHeader()}${allNotes.length ? `\n#\n${commentBlock(allNotes, '#')}` : ''}

${psConnect(opts)}
${body}`;
}

function psFallback(reason: string): string {
  return `${psHeader()}
#
# No faithful SDK form for this request — ${lineComment(reason)}
# The PowerShell tab sends this exact Web API request with Invoke-RestMethod.
`;
}

/** Plans with headers the cmdlets can't pass go through SDK request objects. */
function psBody(plan: Plan, o: Opts, notes: string[]): string {
  const sdk = o.params.length > 0;
  const exec = (setup: string) => `${setup}${psParams(o)}$response = $conn.Execute($request)\n`;
  switch (plan.k) {
    case 'query':
      return psQuery(plan.q, o, notes);
    case 'retrieve': {
      if (plan.target.kind === 'id' && !sdk) {
        const fields = plan.columns ? plan.columns.map(psQuote).join(', ') : "'*'";
        return `$record = Get-CrmRecord -conn $conn -EntityLogicalName ${psQuote(plan.entity)} -Id ${psQuote(
          plan.target.id,
        )} -Fields ${fields}\n$record | Select-Object * -ExcludeProperty original, EntityReference | Format-List\n`;
      }
      const cols = plan.columns
        ? `[Microsoft.Xrm.Sdk.Query.ColumnSet]::new([string[]]@(${plan.columns.map(psQuote).join(', ')}))`
        : '[Microsoft.Xrm.Sdk.Query.ColumnSet]::new($true)';
      return (
        exec(
          `${psTarget(plan.entity, plan.target)}$request = New-Object Microsoft.Xrm.Sdk.Messages.RetrieveRequest
$request.Target = $target
$request.ColumnSet = ${cols}
`,
        ) + '$response.Entity.Attributes | Format-Table Key, Value\n'
      );
    }
    case 'create': {
      if (!sdk && !plan.id) {
        return `${psFields(plan.attrs)}$id = New-CrmRecord -conn $conn -EntityLogicalName ${psQuote(plan.entity)} -Fields $fields\n"Created ${plan.entity} $id"\n`;
      }
      return (
        exec(
          `${psEntity(plan.entity, plan.id ? { kind: 'id', id: plan.id } : undefined, plan.attrs)}${
            plan.id ? '# Create only: fails if the row already exists (If-None-Match: *).\n' : ''
          }$request = New-Object Microsoft.Xrm.Sdk.Messages.CreateRequest
$request.Target = $entity
`,
        ) + '"Created $($response.id)"\n'
      );
    }
    case 'update': {
      if (!sdk && !plan.rowVersion && plan.target.kind === 'id') {
        return `${psFields(plan.attrs)}Set-CrmRecord -conn $conn -EntityLogicalName ${psQuote(plan.entity)} -Id ${psQuote(
          plan.target.id,
        )} -Fields $fields\n"Updated"\n`;
      }
      return (
        exec(
          `${psEntity(plan.entity, plan.target, plan.attrs)}${
            plan.rowVersion ? `$entity.RowVersion = ${psQuote(plan.rowVersion)}  # If-Match\n` : ''
          }$request = New-Object Microsoft.Xrm.Sdk.Messages.UpdateRequest
$request.Target = $entity
${plan.rowVersion ? '$request.ConcurrencyBehavior = [Microsoft.Xrm.Sdk.ConcurrencyBehavior]::IfRowVersionMatches\n' : ''}`,
        ) + '"Updated"\n'
      );
    }
    case 'upsert':
      return (
        exec(
          `${psEntity(plan.entity, plan.target, plan.attrs)}# No If-Match: Dataverse creates the row if it doesn't exist (upsert).
$request = New-Object Microsoft.Xrm.Sdk.Messages.UpsertRequest
$request.Target = $entity
`,
        ) +
        'if ($response.RecordCreated) { "Created $($response.Target.Id)" } else { "Updated $($response.Target.Id)" }\n'
      );
    case 'delete': {
      if (!sdk && !plan.rowVersion && plan.target.kind === 'id') {
        return `Remove-CrmRecord -conn $conn -EntityLogicalName ${psQuote(plan.entity)} -Id ${psQuote(plan.target.id)}\n"Deleted"\n`;
      }
      return (
        exec(
          `${psTarget(plan.entity, plan.target)}${
            plan.rowVersion ? `$target.RowVersion = ${psQuote(plan.rowVersion)}  # If-Match\n` : ''
          }$request = New-Object Microsoft.Xrm.Sdk.Messages.DeleteRequest
$request.Target = $target
${plan.rowVersion ? '$request.ConcurrencyBehavior = [Microsoft.Xrm.Sdk.ConcurrencyBehavior]::IfRowVersionMatches\n' : ''}`,
        ) + '"Deleted"\n'
      );
    }
    case 'associate':
    case 'disassociate':
      return psRelate(plan, o);
    case 'execute': {
      const params = plan.params.length
        ? `@{\n${plan.params.map(([k, v]) => `  ${psQuote(k)} = ${psValue(v)}`).join('\n')}\n}`
        : '@{}';
      if (!sdk) {
        return `$parameters = ${params}
$response = Invoke-CrmAction -conn $conn -Name ${psQuote(plan.name)} -Parameters $parameters${
          plan.target ? ` -Target ${psValue(plan.target)}` : ''
        }
$response
`;
      }
      return `$request = New-Object Microsoft.Xrm.Sdk.OrganizationRequest -ArgumentList ${psQuote(plan.name)}
${plan.params.map(([k, v]) => `$request[${psQuote(k)}] = ${psValue(v)}\n`).join('')}${
        plan.target ? `$request['Target'] = ${psValue(plan.target)}\n` : ''
      }${psParams(o)}$response = $conn.Execute($request)
$response.Results
`;
    }
  }
  return '';
}

function psTarget(entity: string, t: Target): string {
  if (t.kind === 'id') {
    return `$target = New-CrmEntityReference -EntityLogicalName ${psQuote(entity)} -Id ${psQuote(t.id)}\n`;
  }
  return `$target = New-Object Microsoft.Xrm.Sdk.EntityReference
$target.LogicalName = ${psQuote(entity)}
${t.pairs.map(([k, v]) => `$target.KeyAttributes[${psQuote(k)}] = ${psValue(v)}\n`).join('')}`;
}

function psEntity(entity: string, t: Target | undefined, attrs: Attr[]): string {
  return `$entity = New-Object Microsoft.Xrm.Sdk.Entity -ArgumentList ${psQuote(entity)}
${t?.kind === 'id' ? `$entity.Id = ${psGuid(t.id)}\n` : ''}${
    t?.kind === 'key'
      ? t.pairs.map(([k, v]) => `$entity.KeyAttributes[${psQuote(k)}] = ${psValue(v)}\n`).join('')
      : ''
  }${attrs.map((a) => `$entity[${psQuote(a.name)}] = ${psValue(a.value)}\n`).join('')}`;
}

function psRelate(plan: Extract<Plan, { k: 'associate' | 'disassociate' }>, o: Opts): string {
  const verb = plan.k === 'associate' ? 'Associate' : 'Disassociate';
  const cmd = plan.k === 'associate' ? 'Add-CrmRecordAssociation' : 'Remove-CrmRecordAssociation';
  const cmdLine = (t: Extract<SdkValue, { t: 'ref' }>) =>
    `${cmd} -conn $conn -EntityLogicalName1 ${psQuote(plan.entity)} -Id1 ${psQuote(plan.id)} -EntityLogicalName2 ${psQuote(
      t.entity,
    )} -Id2 ${psQuote(t.id)} -RelationshipName ${psQuote(plan.relationship)}`;
  // The cmdlets can't set optional parameters or a self-referencing role.
  const cmdlets = !o.params.length && !plan.selfRole;
  if (cmdlets && (!plan.atomic || plan.targets.length === 1)) {
    const intro =
      plan.targets.length > 1
        ? '# One request per row, in order, as the Web API sends them.\n'
        : '';
    return `${intro}${plan.targets.map(cmdLine).join('\n')}\n`;
  }
  const request = (related: string, indent: string) =>
    [
      `$request = New-Object Microsoft.Xrm.Sdk.Messages.${verb}Request`,
      `$request.Target = New-CrmEntityReference -EntityLogicalName ${psQuote(plan.entity)} -Id ${psQuote(plan.id)}`,
      `$request.Relationship = New-Object Microsoft.Xrm.Sdk.Relationship -ArgumentList ${psQuote(plan.relationship)}`,
      ...(plan.selfRole
        ? [
            `$request.Relationship.PrimaryEntityRole = [Microsoft.Xrm.Sdk.EntityRole]::${plan.selfRole}  # self-referencing`,
          ]
        : []),
      `$request.RelatedEntities = ${related}`,
      ...o.params.map(([k, v]) => `$request.Parameters[${psQuote(k)}] = ${psValue(v)}`),
    ]
      .map((l) => `${indent}${l}\n`)
      .join('');
  if (!plan.atomic && plan.targets.length > 1) {
    return `# One request per row, in order, as the Web API sends them.
$targets = @(
${plan.targets.map((t) => `  ${psValue(t)}`).join(',\n')}
)
$ok = 0
foreach ($t in $targets) {
  $related = New-Object Microsoft.Xrm.Sdk.EntityReferenceCollection
  $related.Add($t)
${request('$related', '  ')}  try { $null = $conn.Execute($request); $ok++ }
  catch { "$($t.LogicalName) $($t.Id): $($_.Exception.Message)" }
}
"$ok/$($targets.Count) requests succeeded"
`;
  }
  const atomicNote =
    plan.targets.length > 1 ? '# One request: Dataverse applies all of them or none.\n' : '';
  return `$related = New-Object Microsoft.Xrm.Sdk.EntityReferenceCollection
${plan.targets.map((t) => `$related.Add(${psValue(t)})\n`).join('')}${request('$related', '')}${atomicNote}$response = $conn.Execute($request)
"${verb}d ${plan.targets.length} row(s)"
`;
}

// ── FetchXML ────────────────────────────────────────────────

/** ConditionOperator name → FetchXML operator. */
const FETCH_OPS: Record<string, string> = {
  Equal: 'eq',
  NotEqual: 'ne',
  GreaterThan: 'gt',
  GreaterEqual: 'ge',
  LessThan: 'lt',
  LessEqual: 'le',
  Like: 'like',
  NotLike: 'not-like',
  Null: 'null',
  NotNull: 'not-null',
  On: 'on',
  OnOrAfter: 'on-or-after',
  OnOrBefore: 'on-or-before',
  Today: 'today',
  Tomorrow: 'tomorrow',
  Yesterday: 'yesterday',
  Last7Days: 'last-seven-days',
  Next7Days: 'next-seven-days',
  LastWeek: 'last-week',
  ThisWeek: 'this-week',
  NextWeek: 'next-week',
  LastMonth: 'last-month',
  ThisMonth: 'this-month',
  NextMonth: 'next-month',
  LastYear: 'last-year',
  ThisYear: 'this-year',
  NextYear: 'next-year',
  LastXHours: 'last-x-hours',
  LastXDays: 'last-x-days',
  LastXWeeks: 'last-x-weeks',
  LastXMonths: 'last-x-months',
  LastXYears: 'last-x-years',
  NextXHours: 'next-x-hours',
  NextXDays: 'next-x-days',
  NextXWeeks: 'next-x-weeks',
  NextXMonths: 'next-x-months',
  NextXYears: 'next-x-years',
  OlderThanXMinutes: 'olderthan-x-minutes',
  OlderThanXHours: 'olderthan-x-hours',
  OlderThanXDays: 'olderthan-x-days',
  OlderThanXWeeks: 'olderthan-x-weeks',
  OlderThanXMonths: 'olderthan-x-months',
  OlderThanXYears: 'olderthan-x-years',
  ThisFiscalPeriod: 'this-fiscal-period',
  ThisFiscalYear: 'this-fiscal-year',
  LastFiscalPeriod: 'last-fiscal-period',
  LastFiscalYear: 'last-fiscal-year',
  NextFiscalPeriod: 'next-fiscal-period',
  NextFiscalYear: 'next-fiscal-year',
  LastXFiscalPeriods: 'last-x-fiscal-periods',
  LastXFiscalYears: 'last-x-fiscal-years',
  NextXFiscalPeriods: 'next-x-fiscal-periods',
  NextXFiscalYears: 'next-x-fiscal-years',
  InFiscalPeriod: 'in-fiscal-period',
  InFiscalYear: 'in-fiscal-year',
  InFiscalPeriodAndYear: 'in-fiscal-period-and-year',
  InOrAfterFiscalPeriodAndYear: 'in-or-after-fiscal-period-and-year',
  InOrBeforeFiscalPeriodAndYear: 'in-or-before-fiscal-period-and-year',
  Between: 'between',
  NotBetween: 'not-between',
  In: 'in',
  NotIn: 'not-in',
  ContainValues: 'contain-values',
  DoesNotContainValues: 'not-contain-values',
  Above: 'above',
  AboveOrEqual: 'eq-or-above',
  Under: 'under',
  UnderOrEqual: 'eq-or-under',
  NotUnder: 'not-under',
  EqualUserId: 'eq-userid',
  NotEqualUserId: 'ne-userid',
  EqualBusinessId: 'eq-businessid',
  NotEqualBusinessId: 'ne-businessid',
  EqualUserLanguage: 'eq-userlanguage',
  EqualUserTeams: 'eq-userteams',
  EqualUserOrUserTeams: 'eq-useroruserteams',
  EqualUserOrUserHierarchy: 'eq-useroruserhierarchy',
  EqualUserOrUserHierarchyAndTeams: 'eq-useroruserhierarchyandteams',
};

/** Operators whose values go in <value> children even when there is one. */
const LIST_OPS = new Set([
  'in',
  'not-in',
  'between',
  'not-between',
  'contain-values',
  'not-contain-values',
]);

/** XML attribute / text escaping; control characters as character references. */
export function xmlEscape(s: string): string {
  // XML 1.0 allows no other C0 control character, even as a reference.
  // eslint-disable-next-line no-control-regex
  if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(s)) {
    return no("A control character in a value: FetchXML (XML 1.0) can't carry it.");
  }
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
    .replace(
      // eslint-disable-next-line no-control-regex
      /[\u0000-\u001f\u0085\u2028\u2029]/g,
      (ch) => `&#${ch.charCodeAt(0)};`,
    );
}

function fetchValue(v: SdkValue): string {
  switch (v.t) {
    case 'bool':
      return v.v ? '1' : '0';
    case 'null':
      return '';
    case 'string':
    case 'int':
    case 'long':
    case 'decimal':
    case 'double':
    case 'guid':
    case 'datetime':
    case 'option':
    case 'money':
    case 'bytes':
      return v.v;
    default:
      return no('A list value in a FetchXML condition.');
  }
}

function fetchFilter(cond: Cond, indent: string, out: string[]): void {
  if (cond.type === 'group') {
    out.push(`${indent}<filter type="${cond.or ? 'or' : 'and'}">`);
    for (const c of cond.children) fetchFilter(c, `${indent}  `, out);
    out.push(`${indent}</filter>`);
    return;
  }
  const op = FETCH_OPS[cond.op] ?? no(`FetchXML has no operator for ${cond.op}.`);
  const attr = `attribute="${xmlEscape(cond.attr)}" operator="${op}"`;
  if (cond.column) {
    out.push(`${indent}<condition ${attr} valueof="${xmlEscape(cond.column)}" />`);
  } else if (cond.values.length === 0) {
    out.push(`${indent}<condition ${attr} />`);
  } else if (cond.values.length === 1 && !LIST_OPS.has(op)) {
    out.push(`${indent}<condition ${attr} value="${xmlEscape(fetchValue(cond.values[0]))}" />`);
  } else {
    out.push(`${indent}<condition ${attr}>`);
    for (const v of cond.values) out.push(`${indent}  <value>${xmlEscape(fetchValue(v))}</value>`);
    out.push(`${indent}</condition>`);
  }
}

function fetchXml(q: QueryModel, page?: { size: number; total: boolean }, top?: number): string {
  const fetchAttrs = page
    ? ` count="${page.size}" page="1"${page.total ? ' returntotalrecordcount="true"' : ''}`
    : top !== undefined
      ? ` top="${top}"`
      : '';
  const lines = [`<fetch${fetchAttrs}>`, `  <entity name="${xmlEscape(q.entity)}">`];
  if (q.columns) for (const c of q.columns) lines.push(`    <attribute name="${xmlEscape(c)}" />`);
  else lines.push('    <all-attributes />');
  if (q.filter && (q.filter.type === 'cond' || q.filter.children.length)) {
    fetchFilter(q.filter, '    ', lines);
  }
  for (const o of q.orders) {
    lines.push(
      `    <order attribute="${xmlEscape(o.attr)}"${o.desc ? ' descending="true"' : ''} />`,
    );
  }
  lines.push('  </entity>', '</fetch>');
  return lines.join('\n');
}

function psQuery(q: QueryModel, o: Opts, notes: string[]): string {
  const paging = pagingOf(q, notes);
  // Get-CrmRecordsByFetch pages the fetch itself (-TopCount is the page size)
  // and returns no total, so a count goes through RetrieveMultiple.
  if (paging.page?.total || o.params.length) {
    // RetrieveMultiple doesn't page the fetch, so $top goes in as `top`.
    const fetch = fetchXml(q, paging.page, paging.top);
    const exec = o.params.length
      ? `$request = New-Object Microsoft.Xrm.Sdk.Messages.RetrieveMultipleRequest
$request.Query = [Microsoft.Xrm.Sdk.Query.FetchExpression]::new($fetch)
${psParams(o)}$results = $conn.Execute($request).EntityCollection
`
      : `$results = $conn.RetrieveMultiple([Microsoft.Xrm.Sdk.Query.FetchExpression]::new($fetch))\n`;
    const total = `"Total: $($results.TotalRecordCount)$(if ($results.TotalRecordCountLimitExceeded) { ' (Dataverse stopped counting at its limit)' })"\n`;
    if (q.countOnly) return `$fetch = @'\n${fetch}\n'@\n${exec}${total}`;
    return `$fetch = @'\n${fetch}\n'@\n${exec}"Returned $($results.Entities.Count) record(s)"\n${
      paging.page?.total ? total : ''
    }foreach ($row in $results.Entities) {
  "$($row.LogicalName) $($row.Id)"
  $row.Attributes | Format-Table Key, Value
}
`;
  }
  const size = paging.top ?? paging.page?.size;
  return `$fetch = @'
${fetchXml(q)}
'@
$result = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch${size !== undefined ? ` -TopCount ${size}` : ''}
"Returned $($result.Count) record(s)"
$result.CrmRecords | Select-Object * -ExcludeProperty original, EntityReference | Format-List
${paging.top !== undefined ? '' : `if ($result.NextPage) { 'More rows: pass -AllRows, or -PageNumber 2 -PageCookie $result.PagingCookie.' }\n`}`;
}
