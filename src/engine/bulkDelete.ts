// Bulk Delete job (#10) — the BulkDelete action plus the $filter-tree →
// QueryExpression mapping it needs.
//
//   POST /BulkDelete
//   {
//     "QuerySet": [ { "EntityName": "account", "Criteria": { "FilterOperator": "And",
//                     "Conditions": [ { "AttributeName": "…", "Operator": "Equal",
//                                       "Values": [ { "Type": "System.String", "Value": "…" } ] } ],
//                     "Filters": [] } } ],
//     "JobName": "…", "SendEmailNotification": false, "ToRecipients": [], "CCRecipients": [],
//     "RecurrencePattern": "", "StartDateTime": "2026-10-17T10:00:00.000Z"
//   }
//   → 200 { "JobId": "<asyncoperationid>" }
//
// The job runs asynchronously; progress lives on asyncoperations(<JobId>)
// and the counts on the matching bulkdeleteoperations row.
//
// QueryExpression can't be run through the Web API, so DRS counts the
// matching rows with the equivalent OData $filter (the same tree encoded by
// the Retrieve Multiple encoder) before the user confirms.
//
// Mapping limits — these $filter-builder nodes have no QueryExpression
// equivalent here and block the request with an explanation:
//   • any/all lambdas over related records (would need LinkEntities)
//   • related-table paths (`nav/column`) — same reason
//   • column-to-column comparisons
//   • negated groups (`not (…)`) and negated query functions
// Negated contains / startswith / endswith map to NotLike.
//
// https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/bulkdelete
// https://learn.microsoft.com/en-us/power-apps/developer/data-platform/delete-data-bulk
// https://learn.microsoft.com/en-us/power-apps/developer/data-platform/auditing/delete-audit-data#use-bulkdelete-to-delete-audit-data

import { findTable, findColumn, type ColumnMeta, type TableMeta } from '../mock/metadata';
import { ENV } from '../mock/environment';
import type { HeaderItem } from '../editors/HeadersEditor';
import {
  collectInvalidLiterals,
  groupToOData,
  stripLeadingWildcards,
  type FilterGroup,
  type FilterNode,
} from '../editors/filter/filterTree';
import { findOperator } from '../editors/filter/operators';
import type { BuiltRequest } from './urlBuilder';

// ── QueryExpression JSON shapes ──────────────────────────────────────

export interface QeValue {
  Type: string;
  Value: string;
}
export interface QeCondition {
  AttributeName: string;
  Operator: string;
  Values: QeValue[];
}
export interface QeFilter {
  FilterOperator: 'And' | 'Or';
  Conditions: QeCondition[];
  Filters: QeFilter[];
}
export interface QueryExpressionJson {
  EntityName: string;
  Criteria: QeFilter;
}

/** OData comparison operator id → ConditionOperator. */
const COMPARISON: Record<string, string> = {
  eq: 'Equal',
  ne: 'NotEqual',
  gt: 'GreaterThan',
  ge: 'GreaterEqual',
  lt: 'LessThan',
  le: 'LessEqual',
};

/**
 * Dataverse query functions share their names with ConditionOperator
 * members (LastXDays, In, Under, EqualUserId, …). Only names on this list
 * are emitted; anything else is reported as unsupported.
 */
const CONDITION_OPERATORS = new Set([
  'On',
  'OnOrAfter',
  'OnOrBefore',
  'Today',
  'Tomorrow',
  'Yesterday',
  'Last7Days',
  'Next7Days',
  'LastWeek',
  'ThisWeek',
  'NextWeek',
  'LastMonth',
  'ThisMonth',
  'NextMonth',
  'LastYear',
  'ThisYear',
  'NextYear',
  'LastXHours',
  'LastXDays',
  'LastXWeeks',
  'LastXMonths',
  'LastXYears',
  'NextXHours',
  'NextXDays',
  'NextXWeeks',
  'NextXMonths',
  'NextXYears',
  'OlderThanXMinutes',
  'OlderThanXHours',
  'OlderThanXDays',
  'OlderThanXWeeks',
  'OlderThanXMonths',
  'OlderThanXYears',
  'ThisFiscalPeriod',
  'ThisFiscalYear',
  'LastFiscalPeriod',
  'LastFiscalYear',
  'NextFiscalPeriod',
  'NextFiscalYear',
  'LastXFiscalPeriods',
  'LastXFiscalYears',
  'NextXFiscalPeriods',
  'NextXFiscalYears',
  'InFiscalPeriod',
  'InFiscalYear',
  'InFiscalPeriodAndYear',
  'InOrAfterFiscalPeriodAndYear',
  'InOrBeforeFiscalPeriodAndYear',
  'Between',
  'NotBetween',
  'In',
  'NotIn',
  'ContainValues',
  'DoesNotContainValues',
  'Above',
  'AboveOrEqual',
  'Under',
  'UnderOrEqual',
  'NotUnder',
  'EqualUserId',
  'NotEqualUserId',
  'EqualBusinessId',
  'NotEqualBusinessId',
  'EqualRoleBusinessId',
  'EqualUserLanguage',
  'EqualUserTeams',
  'EqualUserOrUserTeams',
  'EqualUserOrUserHierarchy',
  'EqualUserOrUserHierarchyAndTeams',
]);

/** .NET type name for a condition value, by column type. */
export function valueTypeFor(col: ColumnMeta | undefined): string {
  switch (col?.attributeType) {
    case 'Integer':
    case 'Picklist':
    case 'State':
    case 'Status':
    case 'MultiSelectPicklist':
    case 'EntityName':
      return 'System.Int32';
    case 'BigInt':
      return 'System.Int64';
    case 'Decimal':
    case 'Money':
      return 'System.Decimal';
    case 'Double':
      return 'System.Double';
    case 'Boolean':
      return 'System.Boolean';
    case 'DateTime':
      return 'System.DateTime';
    case 'Uniqueidentifier':
    case 'Lookup':
    case 'Customer':
    case 'Owner':
      return 'System.Guid';
    default:
      return 'System.String';
  }
}

/** Escape LIKE wildcards so the user's text matches literally. */
const likeEscape = (v: string) => v.replace(/[[%_]/g, (m) => `[${m}]`);

export interface QueryMapping {
  query: QueryExpressionJson | null;
  /** Nodes that can't be expressed — any entry blocks the request. */
  problems: string[];
}

export function filterToQueryExpression(tbl: TableMeta, root: FilterGroup): QueryMapping {
  const problems: string[] = [];

  const condition = (node: FilterNode): QeCondition | null => {
    if (node.type === 'rule') {
      if (node.col.includes('/')) {
        problems.push(
          `"${node.col}" filters on a related table — not supported in a bulk delete query.`,
        );
        return null;
      }
      if (node.valKind === 'column') {
        problems.push(`"${node.col}" compares two columns — not supported in a bulk delete query.`);
        return null;
      }
      const op = findOperator(node.op);
      const col = findColumn(tbl, node.col);
      const v = node.val ?? '';
      if (op?.kind === 'null-check') {
        return {
          AttributeName: node.col,
          Operator: node.op === 'is-null' ? 'Null' : 'NotNull',
          Values: [],
        };
      }
      if (op?.kind === 'comparison' && COMPARISON[node.op]) {
        return {
          AttributeName: node.col,
          Operator: COMPARISON[node.op],
          Values: [{ Type: valueTypeFor(col), Value: v }],
        };
      }
      if (op?.kind === 'odata-fn') {
        const { value } = stripLeadingWildcards(v);
        const esc = likeEscape(value);
        const pattern =
          node.op === 'contains' ? `%${esc}%` : node.op === 'startswith' ? `${esc}%` : `%${esc}`;
        return {
          AttributeName: node.col,
          Operator: node.negated ? 'NotLike' : 'Like',
          Values: [{ Type: 'System.String', Value: pattern }],
        };
      }
      problems.push(`Operator "${node.op}" on "${node.col}" has no QueryExpression mapping.`);
      return null;
    }
    if (node.type === 'function') {
      const op = findOperator(node.op);
      if (!op || !CONDITION_OPERATORS.has(node.op)) {
        problems.push(`Query function "${node.op}" has no QueryExpression mapping.`);
        return null;
      }
      if (node.negated) {
        problems.push(`"not ${node.op}" can't be expressed in a bulk delete query.`);
        return null;
      }
      if (node.col.includes('/')) {
        problems.push(
          `"${node.col}" filters on a related table — not supported in a bulk delete query.`,
        );
        return null;
      }
      // Hierarchy conditions target the primary key, as in the OData encoder.
      const attr = op.category === 'hierarchy' ? tbl.primaryKey : node.col;
      const col = findColumn(tbl, node.col);
      const raw =
        op.kind === 'dv-fn-array'
          ? (node.values ?? [])
          : op.kind === 'dv-fn-2'
            ? (node.vals ?? ['', ''])
            : op.kind === 'dv-fn-0'
              ? []
              : [node.val ?? ''];
      const type =
        op.kind === 'dv-fn-1-int' || op.kind === 'dv-fn-2'
          ? 'System.Int32'
          : op.kind === 'dv-fn-guid'
            ? 'System.Guid'
            : valueTypeFor(col);
      return {
        AttributeName: attr,
        Operator: node.op,
        Values: raw.map((Value) => ({ Type: type, Value })),
      };
    }
    return null;
  };

  const group = (g: FilterGroup): QeFilter => {
    if (g.negated)
      problems.push('Negated groups ("not (…)") can’t be expressed in a bulk delete query.');
    const out: QeFilter = {
      FilterOperator: g.combinator === 'or' ? 'Or' : 'And',
      Conditions: [],
      Filters: [],
    };
    for (const n of g.rules) {
      if (n.type === 'group') out.Filters.push(group(n));
      else if (n.type === 'lambda') {
        problems.push(
          `"${n.nav}/${n.lambda}(…)" filters on related records — not supported in a bulk delete query.`,
        );
      } else {
        const c = condition(n);
        if (c) out.Conditions.push(c);
      }
    }
    return out;
  };

  const Criteria = group(root);
  // Values that don't match their column's type (#221): the job, and the
  // count query that previews it, would both fail on them.
  for (const l of collectInvalidLiterals(root, tbl)) {
    problems.push(`"${l.col}" needs ${l.expected}, not ${JSON.stringify(l.value)}.`);
  }
  return {
    query: problems.length ? null : { EntityName: tbl.logicalName, Criteria },
    problems,
  };
}

// ── Mode state + request ─────────────────────────────────────────────

export interface BulkDeleteState {
  table: string;
  filter: FilterGroup;
  jobName: string;
  sendEmail: boolean;
  /** Comma / whitespace separated systemuser ids. */
  toRecipients: string;
  ccRecipients: string;
  /** 'now' = StartDateTime is the moment you send. */
  startMode: 'now' | 'at';
  /** `datetime-local` value (local time) when startMode = 'at'. */
  startAt: string;
  /** null = run once; N = FREQ=DAILY;INTERVAL=N. */
  recurrenceDays: number | null;
  headers: HeaderItem[];
  dirty: Set<string>;
}

const GUID_RE = /^\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}?$/i;

export function parseRecipients(raw: string): { ids: string[]; invalid: string[] } {
  const parts = raw
    .split(/[\s,;]+/)
    .map((p) => p.trim())
    .filter(Boolean);
  const ids: string[] = [];
  const invalid: string[] = [];
  for (const p of parts) {
    if (GUID_RE.test(p)) ids.push(p.replace(/^\{|\}$/g, '').toLowerCase());
    else invalid.push(p);
  }
  return { ids: [...new Set(ids)], invalid };
}

/** Recipients are systemuser ids, sent as activityparty entries. */
const party = (id: string) => ({
  '@odata.type': 'Microsoft.Dynamics.CRM.activityparty',
  activitypartyid: id,
});

export function recurrencePattern(days: number | null): string {
  return days && days > 0 ? `FREQ=DAILY;INTERVAL=${Math.floor(days)}` : '';
}

export function startDateTime(
  s: Pick<BulkDeleteState, 'startMode' | 'startAt'>,
  now: Date,
): string {
  if (s.startMode === 'at' && s.startAt) {
    const d = new Date(s.startAt); // datetime-local → local time
    if (!Number.isNaN(d.getTime())) return d.toISOString();
  }
  return now.toISOString();
}

/** `yyyy-MM-dd HH:mm` in the user's local time — the job status shows local
 *  times, so a UTC stamp in the name read hours off (#200). */
function localStamp(d: Date): string {
  const p = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

export function defaultJobName(table: string, now: Date): string {
  return `DRS bulk delete · ${table || 'table'} · ${localStamp(now)}`;
}

export interface BuiltBulkDelete {
  body: Record<string, unknown> | null;
  /** Anything that blocks Send. */
  problems: string[];
  /** True when the query has no conditions — every row of the table. */
  deletesAll: boolean;
}

export function buildBulkDeleteBody(s: BulkDeleteState, now: Date): BuiltBulkDelete {
  const tbl = findTable(s.table);
  if (!tbl) return { body: null, problems: ['Pick a target table first.'], deletesAll: false };
  const { query, problems } = filterToQueryExpression(tbl, s.filter);
  const to = parseRecipients(s.toRecipients);
  const cc = parseRecipients(s.ccRecipients);
  const all = [...problems];
  if (!s.jobName.trim()) all.push('Give the job a name.');
  if (to.invalid.length || cc.invalid.length)
    all.push(`Not a systemuser id: ${[...to.invalid, ...cc.invalid].join(', ')}`);
  if (s.sendEmail && to.ids.length === 0)
    all.push('Email notification needs at least one To recipient (systemuser id).');
  if (s.startMode === 'at' && Number.isNaN(new Date(s.startAt).getTime()))
    all.push('Pick a valid start date and time.');
  if (s.recurrenceDays !== null && !(s.recurrenceDays >= 1))
    all.push('Recurrence must be every 1 day or more.');
  const deletesAll = !!query && isEmptyFilter(query.Criteria);
  return {
    body: query
      ? {
          QuerySet: [query],
          JobName: s.jobName.trim(),
          SendEmailNotification: s.sendEmail,
          ToRecipients: s.sendEmail ? to.ids.map(party) : [],
          CCRecipients: s.sendEmail ? cc.ids.map(party) : [],
          RecurrencePattern: recurrencePattern(s.recurrenceDays),
          StartDateTime: startDateTime(s, now),
        }
      : null,
    problems: all,
    deletesAll,
  };
}

function isEmptyFilter(f: QeFilter): boolean {
  return f.Conditions.length === 0 && f.Filters.every(isEmptyFilter);
}

/** Unbound action URL. */
export function buildBulkDelete(s: Pick<BulkDeleteState, 'table'>): BuiltRequest {
  const tbl = findTable(s.table);
  const noBase = '/BulkDelete';
  const relativeUrl = `${ENV.apiBase}${noBase}`;
  return {
    relativeUrl,
    relativeNoBase: noBase,
    bytes: relativeUrl.length,
    queryParts: [],
    entitySet: tbl?.entitySetName ?? '',
    entityLogical: tbl?.logicalName ?? '',
  };
}

/** Matching-row count query (relative to /api/data/v9.2/) — same filter, OData form. */
export function countPath(s: Pick<BulkDeleteState, 'table' | 'filter'>): string | null {
  const tbl = findTable(s.table);
  if (!tbl) return null;
  const filter = groupToOData(s.filter, tbl);
  // Same light encoding as the Retrieve Multiple URL builder: escape what
  // would change the URL's meaning (&, #, +, %…), keep the rest readable.
  const enc = (v: string) =>
    encodeURIComponent(v)
      .replace(/%2F/g, '/')
      .replace(/%2C/g, ',')
      .replace(/%3D/g, '=')
      .replace(/%24/g, '$')
      .replace(/%28/g, '(')
      .replace(/%29/g, ')')
      .replace(/%20/g, ' ')
      .replace(/%27/g, "'")
      .replace(/%3A/g, ':');
  return (
    `${tbl.entitySetName}?$select=${tbl.primaryKey}&$count=true&$top=1` +
    (filter ? `&$filter=${enc(filter)}` : '')
  );
}

/** `@odata.count` caps at 5,000. */
export const COUNT_CAP = 5000;

// ── Job status (asyncoperation) ──────────────────────────────────────

export const ASYNC_STATE: Record<number, string> = {
  0: 'Ready',
  1: 'Suspended',
  2: 'Locked',
  3: 'Completed',
};
export const ASYNC_STATUS: Record<number, string> = {
  0: 'Waiting for resources',
  10: 'Waiting',
  20: 'In progress',
  21: 'Pausing',
  22: 'Canceling',
  30: 'Succeeded',
  31: 'Failed',
  32: 'Canceled',
};

export function jobStatusPath(jobId: string): string {
  return (
    `asyncoperations(${jobId})?$select=name,statecode,statuscode,message,friendlymessage,` +
    `createdon,startedon,completedon,postponeuntil,recurrencepattern`
  );
}

export function bulkDeleteOperationPath(jobId: string): string {
  return (
    `bulkdeleteoperations?$select=successcount,failurecount,statecode,statuscode,isrecurring` +
    `&$filter=_asyncoperationid_value eq ${jobId}`
  );
}

// ── Save / load ──────────────────────────────────────────────────────

export type SerializedBulkDeleteState = Omit<BulkDeleteState, 'dirty'>;

export function serializeBulkDelete(s: BulkDeleteState): SerializedBulkDeleteState {
  const { dirty: _dirty, ...rest } = s;
  return rest;
}

export function deserializeBulkDelete(
  snap: SerializedBulkDeleteState,
  defaults: Omit<BulkDeleteState, 'dirty'>,
): BulkDeleteState {
  return { ...defaults, ...snap, dirty: new Set() };
}
