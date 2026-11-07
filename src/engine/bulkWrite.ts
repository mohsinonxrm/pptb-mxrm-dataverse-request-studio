// Bulk write — CreateMultiple / UpdateMultiple / UpsertMultiple / DeleteMultiple (#8).
//
// All four are actions bound to the table's entity collection:
//
//   POST /<entitySet>/Microsoft.Dynamics.CRM.<Operation>
//   { "Targets": [ { "@odata.type": "Microsoft.Dynamics.CRM.<logical>", … }, … ] }
//
// Every item in Targets MUST carry `@odata.type`. Row identity per operation:
//
//   CreateMultiple  — no key needed; a GUID key is sent as the primary key
//                     (create with a caller-chosen id).
//   UpdateMultiple  — GUID primary key required. Alternate keys aren't
//                     supported by UpdateMultiple in the Web API.
//   UpsertMultiple  — GUID primary key, or an alternate key segment
//                     (`col='v'` / `k1='a',k2=1`) sent as
//                     `"@odata.id": "<set>(<segment>)"`.
//   DeleteMultiple  — GUID primary key required (plus `partitionid` from the
//                     field set when the elastic table is partitioned).
//                     Elastic tables only; standard tables answer
//                     "DeleteMultiple has not yet been implemented" — use
//                     BulkDelete for those.
//
// Availability on standard tables varies per table: the
// `sdkmessagefilters` query below is the documented way to check.
// UpsertMultiple has no filter row of its own; a table that supports both
// CreateMultiple and UpdateMultiple supports it.
//
// https://learn.microsoft.com/en-us/power-apps/developer/data-platform/bulk-operations
// https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/createmultiple

import { findTable, findColumn, isLookupLike, type TableMeta } from '../metadata/metadata';
import type { ColumnMeta } from '../metadata/metadata';
import { ENV } from '../env/environment';
import type { HeaderItem } from '../editors/HeadersEditor';
import type {
  BypassOptions,
  CreateFieldValue,
  CreateState,
  LookupFieldValue,
} from '../state/writeState';
import { mustSupplyOnCreate } from './fieldBody';
import { buildCreateBody, type BuiltRequest } from './urlBuilder';

export type BulkOperation =
  | 'CreateMultiple'
  | 'UpdateMultiple'
  | 'UpsertMultiple'
  | 'DeleteMultiple';

export const BULK_OPERATIONS: BulkOperation[] = [
  'CreateMultiple',
  'UpdateMultiple',
  'UpsertMultiple',
  'DeleteMultiple',
];

/**
 * Rows above this count get a warning. Microsoft's guidance for standard
 * tables is 100–1,000 records per request (elastic: 100); there's no hard
 * cap, but large payloads hit message-size and time limits and the whole
 * request rolls back.
 */
export const BULK_WARN_THRESHOLD = 1000;
/** Elastic tables don't benefit from bigger requests — Microsoft suggests 100. */
export const BULK_ELASTIC_SUGGESTED = 100;

export interface BulkRow {
  /** Local id (React key / sidebar node), never sent. */
  id: string;
  /** GUID primary key, or (Upsert only) an alternate-key segment like `name='x'`. */
  key: string;
  fieldValues: Record<string, CreateFieldValue>;
  nullFields: string[];
}

export interface BulkWriteState {
  table: string;
  operation: BulkOperation;
  rows: BulkRow[];
  headers: HeaderItem[];
  bypass: BypassOptions;
  dirty: Set<string>;
}

const GUID_RE = /^\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}?$/i;
export const isGuid = (v: string): boolean => GUID_RE.test(v.trim());
const normalizeGuid = (v: string): string =>
  v
    .trim()
    .replace(/^\{|\}$/g, '')
    .toLowerCase();

let __rowSeq = 1;
export function newBulkRow(partial?: Partial<Omit<BulkRow, 'id'>>): BulkRow {
  const id =
    typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
      ? `row_${crypto.randomUUID()}`
      : `row_${Date.now().toString(36)}_${(__rowSeq++).toString(36)}`;
  return {
    id,
    key: partial?.key ?? '',
    fieldValues: partial?.fieldValues ?? {},
    nullFields: partial?.nullFields ?? [],
  };
}

export function duplicateBulkRow(row: BulkRow): BulkRow {
  // The key is cleared: two targets with the same key are an error for
  // UpsertMultiple and silently collapse to the first for UpdateMultiple.
  return newBulkRow({
    key: '',
    fieldValues: structuredCloneSafe(row.fieldValues),
    nullFields: [...row.nullFields],
  });
}

function structuredCloneSafe<T>(v: T): T {
  return JSON.parse(JSON.stringify(v)) as T;
}

// ── Per-operation rules ──────────────────────────────────────────────

export interface OperationInfo {
  label: string;
  /** Short description for the operation picker. */
  description: string;
  /** Whether each row must carry a key. */
  keyRequired: boolean;
  /** Whether an alternate-key segment is accepted in place of a GUID. */
  allowsAlternateKey: boolean;
  /** Whether rows carry column values (DeleteMultiple only needs the key + partitionid). */
  usesFields: boolean;
  /** FieldSetEditor eligibility filter. */
  fieldPurpose: 'create' | 'update' | 'upsert';
  /** HTTP status on success. */
  successStatus: number;
}

export const OPERATION_INFO: Record<BulkOperation, OperationInfo> = {
  CreateMultiple: {
    label: 'Create multiple',
    description: 'Creates every row in one request and returns the new ids.',
    keyRequired: false,
    allowsAlternateKey: false,
    usesFields: true,
    fieldPurpose: 'create',
    successStatus: 200,
  },
  UpdateMultiple: {
    label: 'Update multiple',
    description: 'Updates rows addressed by their primary key (GUID). Send only changed columns.',
    keyRequired: true,
    allowsAlternateKey: false,
    usesFields: true,
    fieldPurpose: 'update',
    successStatus: 204,
  },
  UpsertMultiple: {
    label: 'Upsert multiple',
    description:
      'Creates or updates each row, addressed by primary key or an alternate key. ' +
      'Needs a table that supports both CreateMultiple and UpdateMultiple.',
    keyRequired: true,
    allowsAlternateKey: true,
    usesFields: true,
    fieldPurpose: 'upsert',
    successStatus: 204,
  },
  DeleteMultiple: {
    label: 'Delete multiple',
    description:
      'Deletes rows by primary key. Elastic tables only — for standard tables use Bulk Delete.',
    keyRequired: true,
    allowsAlternateKey: false,
    usesFields: true,
    fieldPurpose: 'update',
    successStatus: 204,
  },
};

// ── Body builder ─────────────────────────────────────────────────────

export interface RowIssue {
  rowIndex: number;
  rowId: string;
  message: string;
}

/** Field values → wire properties, via the same encoder Create / Update use. */
function rowFields(table: string, row: BulkRow): Record<string, unknown> {
  return buildCreateBody({
    table,
    fieldValues: row.fieldValues,
    nullFields: row.nullFields,
  } as CreateState);
}

/**
 * Build one `Targets` item. Returns `null` (plus an issue) when the row
 * can't be addressed for the selected operation.
 */
export function buildBulkTarget(
  s: Pick<BulkWriteState, 'table' | 'operation'>,
  tbl: TableMeta,
  row: BulkRow,
): { target: Record<string, unknown>; issue?: string } {
  const info = OPERATION_INFO[s.operation];
  const target: Record<string, unknown> = {
    '@odata.type': `Microsoft.Dynamics.CRM.${tbl.logicalName}`,
  };
  const fields = info.usesFields ? rowFields(s.table, row) : {};
  Object.assign(target, fields);

  const key = row.key.trim();
  let issue: string | undefined;
  if (key) {
    if (isGuid(key)) {
      target[tbl.primaryKey] = normalizeGuid(key);
    } else if (info.allowsAlternateKey) {
      target['@odata.id'] = `${tbl.entitySetName}(${key})`;
    } else {
      issue =
        s.operation === 'UpdateMultiple'
          ? 'Key must be a GUID — UpdateMultiple in the Web API does not accept alternate keys.'
          : 'Key must be a GUID.';
    }
  } else if (info.keyRequired) {
    issue =
      s.operation === 'UpsertMultiple'
        ? "Missing key — a GUID or an alternate key such as name='value'."
        : 'Missing primary key (GUID).';
  }
  if (!issue && s.operation !== 'DeleteMultiple' && Object.keys(fields).length === 0) {
    issue = 'No column values set.';
  }
  if (!issue && s.operation === 'CreateMultiple') {
    // Same gate as Create Record: business-required columns only (system-
    // required ones are server-defaulted). One bad row rolls back the batch.
    const missing = missingRequired(tbl, row);
    if (missing.length > 0) {
      issue = `Required column${missing.length === 1 ? '' : 's'} unset: ${missing.slice(0, 3).join(', ')}${missing.length > 3 ? '…' : ''}`;
    }
  }
  return { target, issue };
}

// File / Image columns can't be set in a row, so they never count as missing (#305).
function missingRequired(tbl: TableMeta, row: BulkRow): string[] {
  return tbl.columns
    .filter((c) => mustSupplyOnCreate(c))
    .filter((c) => {
      const v = row.fieldValues[c.logicalName];
      if (v == null || v === '') return true;
      if (Array.isArray(v)) return v.length === 0;
      if (isLookupLike(c)) return !(v as LookupFieldValue)?.id;
      return false;
    })
    .map((c) => c.logicalName);
}

export interface BuiltBulkBody {
  body: { Targets: Record<string, unknown>[] };
  issues: RowIssue[];
}

export function buildBulkWriteBody(s: BulkWriteState): BuiltBulkBody {
  const tbl = findTable(s.table);
  if (!tbl) return { body: { Targets: [] }, issues: [] };
  const issues: RowIssue[] = [];
  const seenKeys = new Map<string, number>();
  const Targets = s.rows.map((row, i) => {
    const { target, issue } = buildBulkTarget(s, tbl, row);
    if (issue) issues.push({ rowIndex: i, rowId: row.id, message: issue });
    const k = row.key.trim().toLowerCase();
    if (k) {
      const first = seenKeys.get(k);
      if (first !== undefined) {
        issues.push({
          rowIndex: i,
          rowId: row.id,
          message:
            s.operation === 'UpdateMultiple'
              ? `Same key as row ${first + 1} — UpdateMultiple applies only the first of them.`
              : `Same key as row ${first + 1} — duplicate keys fail the request.`,
        });
      } else seenKeys.set(k, i);
    }
    return target;
  });
  return { body: { Targets }, issues };
}

/** The collection-bound action URL. */
export function buildBulkWrite(s: Pick<BulkWriteState, 'table' | 'operation'>): BuiltRequest {
  const tbl = findTable(s.table);
  if (!tbl) {
    return {
      relativeUrl: '',
      relativeNoBase: '',
      bytes: 0,
      queryParts: [],
      entitySet: '',
      entityLogical: '',
    };
  }
  const noBase = `/${tbl.entitySetName}/Microsoft.Dynamics.CRM.${s.operation}`;
  const relativeUrl = `${ENV.apiBase}${noBase}`;
  return {
    relativeUrl,
    relativeNoBase: noBase,
    bytes: relativeUrl.length,
    queryParts: [],
    entitySet: tbl.entitySetName,
    entityLogical: tbl.logicalName,
  };
}

// ── Availability check ───────────────────────────────────────────────

/**
 * Messages whose `sdkmessagefilters` rows decide availability. UpsertMultiple
 * has no filter row: it's available where Create- and UpdateMultiple both are.
 */
export function supportMessagesFor(op: BulkOperation): string[] {
  return op === 'UpsertMultiple' ? ['CreateMultiple', 'UpdateMultiple'] : [op];
}

/** `sdkmessagefilters` query path (relative to /api/data/v9.2/) for one message + table. */
export function supportQueryPath(message: string, table: string): string {
  // Message and logical names are plain identifiers; quotes are doubled
  // anyway so the literal can't break out of the filter. Spaces stay
  // literal, as in the other paths DRS hands to queryData.
  const lit = (v: string) => `'${v.replace(/'/g, "''")}'`;
  return (
    `sdkmessagefilters?$select=sdkmessagefilterid` +
    `&$filter=sdkmessageid/name eq ${lit(message)} and primaryobjecttypecode eq ${lit(table)}`
  );
}

// ── Disabled-reason / validation ─────────────────────────────────────

/** What blocks Execute, and the sidebar node that fixes it (#294). */
export interface BulkBlocker {
  id: string;
  reason: string;
  focusNode: string;
}

export function bulkBlocker(
  s: BulkWriteState,
  issues: RowIssue[],
  support: BulkSupport,
): BulkBlocker | null {
  const tbl = findTable(s.table);
  if (!s.table)
    return { id: 'bulk-table', reason: 'Pick a target table first.', focusNode: 'target' };
  if (!tbl) return { id: 'bulk-metadata', reason: 'Loading table metadata…', focusNode: 'target' };
  if (support.state === 'unsupported')
    return { id: 'bulk-unsupported', reason: support.message, focusNode: 'operation' };
  if (s.rows.length === 0)
    return { id: 'bulk-no-rows', reason: 'Add at least one row.', focusNode: 'rows' };
  if (issues.length > 0) {
    const first = issues[0];
    return {
      id: 'bulk-row',
      reason: `Row ${first.rowIndex + 1}: ${first.message}${issues.length > 1 ? ` (+${issues.length - 1} more)` : ''}`,
      focusNode: `row:${first.rowId}`,
    };
  }
  if (s.headers.some((h) => h.enabled && !h.name))
    return { id: 'bulk-header-name', reason: 'Fix empty header name.', focusNode: 'headers' };
  return null;
}

export function bulkDisabledReason(
  s: BulkWriteState,
  issues: RowIssue[],
  support: BulkSupport,
): string | null {
  return bulkBlocker(s, issues, support)?.reason ?? null;
}

export type BulkSupport =
  | { state: 'idle' }
  | { state: 'checking' }
  | { state: 'supported' }
  | { state: 'unsupported'; message: string }
  | { state: 'unknown'; message: string };

export function unsupportedMessage(op: BulkOperation, table: string, missing: string[]): string {
  if (op === 'DeleteMultiple') {
    return (
      `DeleteMultiple isn't available for '${table}'. It's supported on elastic tables only — ` +
      'use the Bulk Delete mode (an asynchronous job) for standard tables.'
    );
  }
  if (op === 'UpsertMultiple') {
    return (
      `UpsertMultiple needs both CreateMultiple and UpdateMultiple, and '${table}' ` +
      `doesn't support ${missing.join(' or ')}.`
    );
  }
  return `'${table}' doesn't support ${op} (no sdkmessagefilter for this table).`;
}

// ── Import (paste CSV / JSON) ────────────────────────────────────────

export interface ParsedRows {
  rows: BulkRow[];
  /** Source columns that didn't map to a writable column. */
  skippedColumns: string[];
  /** Per-cell problems (bad number, unknown choice label, …). */
  errors: string[];
}

/** Header names that address the row instead of a column. */
const KEY_HEADERS = new Set(['@key', 'key', 'id', '@odata.id']);

/**
 * Parse pasted rows. Accepts:
 *   • a JSON array of objects (or `{ "Targets": [ … ] }` / `{ "value": [ … ] }`)
 *   • CSV or tab-separated text with a header row (Excel copy-paste is TSV)
 *
 * Headers map to columns by logical name, then display name (case-
 * insensitive). The table's primary key column, `id`, `key`, `@key` or
 * `@odata.id` fill the row key. Lookups take `<guid>` (first target) or
 * `<table>:<guid>`; `<col>@odata.bind` values like `/accounts(<guid>)`
 * work too. Choices take the integer value or the option label; multi-
 * select choices take `1,2` / `1;2` or a JSON array.
 */
export function parseBulkRows(text: string, tbl: TableMeta): ParsedRows {
  const trimmed = text.trim();
  if (!trimmed) return { rows: [], skippedColumns: [], errors: [] };
  let records: Record<string, unknown>[];
  if (trimmed.startsWith('[') || trimmed.startsWith('{')) {
    let parsed: unknown;
    try {
      parsed = JSON.parse(trimmed);
    } catch (e) {
      return {
        rows: [],
        skippedColumns: [],
        errors: [`Invalid JSON: ${e instanceof Error ? e.message : String(e)}`],
      };
    }
    const arr = Array.isArray(parsed)
      ? parsed
      : parsed && typeof parsed === 'object'
        ? ((parsed as { Targets?: unknown; value?: unknown }).Targets ??
          (parsed as { value?: unknown }).value)
        : undefined;
    if (!Array.isArray(arr)) {
      return {
        rows: [],
        skippedColumns: [],
        errors: ['JSON must be an array of objects (or { "Targets": [ … ] }).'],
      };
    }
    records = arr.filter((r): r is Record<string, unknown> => !!r && typeof r === 'object');
  } else {
    const table = parseDelimited(trimmed);
    if (table.length < 2) {
      return {
        rows: [],
        skippedColumns: [],
        errors: ['CSV needs a header row and at least one data row.'],
      };
    }
    const [header, ...data] = table;
    records = data
      .filter((cells) => cells.some((c) => c.trim() !== ''))
      .map((cells) => {
        const rec: Record<string, unknown> = {};
        header.forEach((h, i) => {
          const v = cells[i];
          if (v !== undefined && v !== '') rec[h.trim()] = v;
        });
        return rec;
      });
  }

  const skipped = new Set<string>();
  const errors: string[] = [];
  const rows = records.map((rec, ri) => {
    const row = newBulkRow();
    for (const [rawName, rawVal] of Object.entries(rec)) {
      const name = rawName.trim();
      if (name === '@odata.type' || name === '@odata.etag') continue;
      if (rawVal === undefined) continue;
      if (KEY_HEADERS.has(name.toLowerCase()) || name.toLowerCase() === tbl.primaryKey) {
        row.key = stripODataId(String(rawVal ?? ''), tbl);
        continue;
      }
      const bindMatch = /^(.+)@odata\.bind$/i.exec(name);
      const col = resolveColumn(tbl, bindMatch ? bindMatch[1] : name);
      if (!col) {
        skipped.add(name);
        continue;
      }
      if (rawVal === null) {
        row.nullFields.push(col.logicalName);
        continue;
      }
      const res = coerceValue(col, rawVal);
      if ('error' in res) errors.push(`Row ${ri + 1}, ${name}: ${res.error}`);
      else row.fieldValues[col.logicalName] = res.value;
    }
    return row;
  });
  return { rows, skippedColumns: [...skipped], errors };
}

function stripODataId(v: string, tbl: TableMeta): string {
  // `accounts(<guid>)` / `accounts(name='x')` → inner segment.
  const m = /^\/?([A-Za-z0-9_]+)\((.*)\)$/.exec(v.trim());
  if (m && m[1] === tbl.entitySetName) return m[2];
  return v.trim();
}

function resolveColumn(tbl: TableMeta, name: string): ColumnMeta | undefined {
  const lower = name.toLowerCase();
  const byLogical = findColumn(tbl, lower);
  if (byLogical) return byLogical;
  // Bind properties can differ from the attribute name (polymorphic
  // `customerid_account`, schema-cased nav names); match on the prefix.
  const byBind = tbl.columns.find((c) => isLookupLike(c) && lower.startsWith(`${c.logicalName}_`));
  if (byBind) return byBind;
  return tbl.columns.find((c) => c.displayName.toLowerCase() === lower);
}

type Coerced = { value: CreateFieldValue } | { error: string };

function optionValue(
  options: { value: number; label: string }[],
  raw: unknown,
): number | undefined {
  if (typeof raw === 'number') return raw;
  const s = String(raw).trim();
  if (/^-?\d+$/.test(s)) return Number(s);
  return options.find((o) => o.label.toLowerCase() === s.toLowerCase())?.value;
}

export function coerceValue(col: ColumnMeta, raw: unknown): Coerced {
  const s = typeof raw === 'string' ? raw.trim() : raw;
  switch (col.attributeType) {
    case 'Integer':
    case 'BigInt':
    case 'Decimal':
    case 'Double':
    case 'Money': {
      const n = typeof s === 'number' ? s : Number(String(s).replace(/,/g, ''));
      if (!Number.isFinite(n)) return { error: `"${String(raw)}" is not a number` };
      if (
        (col.attributeType === 'Integer' || col.attributeType === 'BigInt') &&
        !Number.isInteger(n)
      )
        return { error: `"${String(raw)}" is not a whole number` };
      return { value: n };
    }
    case 'Boolean': {
      if (typeof s === 'boolean') return { value: s };
      const t = String(s).toLowerCase();
      if (['true', '1', 'yes', 'y'].includes(t)) return { value: true };
      if (['false', '0', 'no', 'n'].includes(t)) return { value: false };
      if (t === col.trueOption.label.toLowerCase()) return { value: true };
      if (t === col.falseOption.label.toLowerCase()) return { value: false };
      return { error: `"${String(raw)}" is not a yes/no value` };
    }
    case 'Picklist':
    case 'State':
    case 'Status':
    case 'EntityName': {
      const v = optionValue(col.options, s);
      return v === undefined ? { error: `unknown choice "${String(raw)}"` } : { value: v };
    }
    case 'MultiSelectPicklist': {
      const parts: unknown[] = Array.isArray(s)
        ? s
        : String(s)
            .split(/[,;]/)
            .map((p) => p.trim())
            .filter(Boolean);
      const out: number[] = [];
      for (const p of parts) {
        const v = optionValue(col.options, p);
        if (v === undefined) return { error: `unknown choice "${String(p)}"` };
        out.push(v);
      }
      return { value: out };
    }
    case 'Lookup':
    case 'Customer':
    case 'Owner': {
      const str = String(s);
      // `/accounts(<guid>)` (an @odata.bind value)
      const bind = /^\/?([A-Za-z0-9_]+)\(\{?([0-9a-fA-F-]{36})\}?\)$/.exec(str);
      if (bind) {
        const target = col.targets.find((t) => findTable(t)?.entitySetName === bind[1]);
        if (!target) return { error: `"${bind[1]}" is not a target of ${col.logicalName}` };
        return { value: { id: bind[2].toLowerCase(), targetEntity: target } as LookupFieldValue };
      }
      // `<table>:<guid>`
      const pair = /^([A-Za-z0-9_]+):\s*(.+)$/.exec(str);
      const target = pair ? pair[1].toLowerCase() : col.targets[0];
      const id = pair ? pair[2] : str;
      if (!isGuid(id)) return { error: `"${str}" is not a GUID` };
      if (!col.targets.includes(target))
        return { error: `"${target}" is not a target of ${col.logicalName}` };
      return { value: { id: normalizeGuid(id), targetEntity: target } as LookupFieldValue };
    }
    case 'File':
    case 'Image':
      return { error: 'file and image columns can’t be set in a bulk write' };
    default:
      return { value: typeof s === 'string' ? s : String(s) };
  }
}

// ── Save / load ──────────────────────────────────────────────────────

export type SerializedBulkWriteState = Omit<BulkWriteState, 'dirty'>;

export function serializeBulkWrite(s: BulkWriteState): SerializedBulkWriteState {
  const { dirty: _dirty, ...rest } = s;
  return rest;
}

export function deserializeBulkWrite(
  snap: SerializedBulkWriteState,
  defaults: Omit<BulkWriteState, 'dirty'>,
): BulkWriteState {
  return {
    ...defaults,
    ...snap,
    operation: BULK_OPERATIONS.includes(snap.operation) ? snap.operation : 'CreateMultiple',
    rows: (snap.rows ?? []).map((r) => ({
      id: r.id || newBulkRow().id,
      key: r.key ?? '',
      fieldValues: r.fieldValues ?? {},
      nullFields: r.nullFields ?? [],
    })),
    dirty: new Set(),
  };
}

/** Minimal RFC 4180 reader; tab-separated when the header row has tabs and no commas. */
export function parseDelimited(text: string): string[][] {
  const firstLine = text.split(/\r?\n/, 1)[0];
  const delim = firstLine.includes('\t') && !firstLine.includes(',') ? '\t' : ',';
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          cell += '"';
          i++;
        } else quoted = false;
      } else cell += ch;
      continue;
    }
    if (ch === '"' && cell === '') quoted = true;
    else if (ch === delim) {
      row.push(cell);
      cell = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
    } else cell += ch;
  }
  row.push(cell);
  rows.push(row);
  return rows;
}
