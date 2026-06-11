// Operator catalogue — every operator the FilterEditor exposes, with its
// per-AttributeTypeCode allowlist.
//
// Each operator declares:
//  - id            : stable identifier in the rule tree
//  - label         : UI display
//  - hint          : optional tooltip
//  - kind          : how to render + emit (comparison / odata-fn / null-check / dv-fn-*)
//  - category      : grouping in the operator menu
//  - odata         : canonical OData syntax fragment (or function name)
//  - arity         : 0 | 1 | 2 | 'n' — drives the value-input shape
//  - intValue      : true if value is integer (SpinButton)
//  - allowedTypes  : set of AttributeTypeCode values where this operator is valid
//  - requiresTable : optional per-entity gate (e.g. hierarchy)
//  - supportsColumnRhs : true when this operator can take a bare property name on RHS

import type { AttributeTypeCode } from '../../mock/metadata';

export type OpKind =
  | 'comparison' // eq/ne/gt/ge/lt/le — `<col> <op> <val>`
  | 'odata-fn' // contains/startswith/endswith (negation via FilterRule.negated)
  | 'null-check' // eq null / ne null
  | 'dv-fn-0' // Microsoft.Dynamics.CRM.Today(PropertyName='createdon')
  | 'dv-fn-1' // 1 PropertyValue scalar
  | 'dv-fn-1-int' // 1 PropertyValue int
  | 'dv-fn-2' // PropertyValue1 + PropertyValue2
  | 'dv-fn-array' // PropertyValues [...]
  | 'dv-fn-guid'; // PropertyValue GUID

export type OpCategory =
  | 'comparison'
  | 'string'
  | 'date-relative'
  | 'date-rolling'
  | 'fiscal'
  | 'range'
  | 'set'
  | 'choices'
  | 'hierarchy'
  | 'user-context';

export interface OperatorDef {
  id: string;
  label: string;
  hint?: string;
  kind: OpKind;
  category: OpCategory;
  odata: string;
  /** AttributeTypeCode values where this op is valid. Undefined = all */
  allowedTypes?: AttributeTypeCode[];
  /** Restrict to specific tables (e.g. hierarchy ops need parentaccountid) */
  requiresTable?: (table: string) => boolean;
  arity: 0 | 1 | 2 | 'n';
  intValue?: boolean;
  /** True if the RHS may be a property name (column-vs-column compare) */
  supportsColumnRhs?: boolean;
}

export const OP_CATEGORIES: { id: OpCategory; label: string }[] = [
  { id: 'comparison', label: 'Comparison' },
  { id: 'string', label: 'Text' },
  { id: 'date-relative', label: 'Date — relative' },
  { id: 'date-rolling', label: 'Date — rolling' },
  { id: 'fiscal', label: 'Fiscal calendar' },
  { id: 'range', label: 'Range' },
  { id: 'set', label: 'Set membership' },
  { id: 'choices', label: 'Multi-select choices' },
  { id: 'hierarchy', label: 'Hierarchy' },
  { id: 'user-context', label: 'Current user / business / team' },
];

// ── Sets of types per the spec's Table A ────────────────────
const ALL_FILTERABLE: AttributeTypeCode[] = [
  'BigInt',
  'Boolean',
  'Customer',
  'DateTime',
  'Decimal',
  'Double',
  'EntityName',
  'Integer',
  'Lookup',
  'Memo',
  'Money',
  'Owner',
  'Picklist',
  'State',
  'Status',
  'String',
  'Uniqueidentifier',
  'MultiSelectPicklist',
];
const COMPARE_EQ_NE: AttributeTypeCode[] = ALL_FILTERABLE.filter(
  (t) => t !== 'MultiSelectPicklist',
);
const COMPARE_ORDERED: AttributeTypeCode[] = [
  'BigInt',
  'DateTime',
  'Decimal',
  'Double',
  'Integer',
  'Memo',
  'Money',
  'Picklist',
  'String',
];
const STRING_FN: AttributeTypeCode[] = ['String', 'Memo'];
const DATE: AttributeTypeCode[] = ['DateTime'];
const RANGE_OK: AttributeTypeCode[] = [
  'BigInt',
  'DateTime',
  'Decimal',
  'Double',
  'Integer',
  'Money',
];
const SET_OK: AttributeTypeCode[] = [
  'BigInt',
  'DateTime',
  'Decimal',
  'Double',
  'Integer',
  'Lookup',
  'Customer',
  'Owner',
  'Memo',
  'Money',
  'Picklist',
  'State',
  'Status',
  'String',
  'Uniqueidentifier',
  'EntityName',
];
const MULTI_CHOICE: AttributeTypeCode[] = ['MultiSelectPicklist'];
const LOOKUP_FAMILY: AttributeTypeCode[] = ['Lookup', 'Customer', 'Owner'];

// ============================================================
// Operator catalogue
// ============================================================
export const OPERATORS: OperatorDef[] = [
  // ── Comparison ────────────────────────────────────────────
  {
    id: 'eq',
    label: 'equals',
    kind: 'comparison',
    category: 'comparison',
    odata: 'eq',
    arity: 1,
    allowedTypes: COMPARE_EQ_NE,
    supportsColumnRhs: true,
    hint: 'Exact match. Same-type column on the RHS is also allowed (column-vs-column compare).',
  },
  {
    id: 'ne',
    label: 'not equals',
    kind: 'comparison',
    category: 'comparison',
    odata: 'ne',
    arity: 1,
    allowedTypes: COMPARE_EQ_NE,
    supportsColumnRhs: true,
  },
  {
    id: 'gt',
    label: '>  greater than',
    kind: 'comparison',
    category: 'comparison',
    odata: 'gt',
    arity: 1,
    allowedTypes: COMPARE_ORDERED,
    supportsColumnRhs: true,
  },
  {
    id: 'ge',
    label: '≥  greater or equal',
    kind: 'comparison',
    category: 'comparison',
    odata: 'ge',
    arity: 1,
    allowedTypes: COMPARE_ORDERED,
    supportsColumnRhs: true,
  },
  {
    id: 'lt',
    label: '<  less than',
    kind: 'comparison',
    category: 'comparison',
    odata: 'lt',
    arity: 1,
    allowedTypes: COMPARE_ORDERED,
    supportsColumnRhs: true,
  },
  {
    id: 'le',
    label: '≤  less or equal',
    kind: 'comparison',
    category: 'comparison',
    odata: 'le',
    arity: 1,
    allowedTypes: COMPARE_ORDERED,
    supportsColumnRhs: true,
  },
  {
    id: 'is-null',
    label: 'is null',
    kind: 'null-check',
    category: 'comparison',
    odata: 'eq null',
    arity: 0,
    allowedTypes: ALL_FILTERABLE,
  },
  {
    id: 'is-not-null',
    label: 'is not null',
    kind: 'null-check',
    category: 'comparison',
    odata: 'ne null',
    arity: 0,
    allowedTypes: ALL_FILTERABLE,
  },

  // ── String functions ──────────────────────────────────────
  // (Negation handled via FilterRule.negated — a separate "not" checkbox per rule.
  //  This is the only documented way Dataverse supports per-condition negation.)
  {
    id: 'contains',
    label: 'contains',
    kind: 'odata-fn',
    category: 'string',
    odata: 'contains',
    arity: 1,
    allowedTypes: STRING_FN,
    hint: 'Substring match. Leading wildcards (e.g. %foo) are not supported and will be stripped.',
  },
  {
    id: 'startswith',
    label: 'starts with',
    kind: 'odata-fn',
    category: 'string',
    odata: 'startswith',
    arity: 1,
    allowedTypes: STRING_FN,
    hint: 'Sargable — uses an index seek. Preferable to contains/endswith on large tables.',
  },
  {
    id: 'endswith',
    label: 'ends with',
    kind: 'odata-fn',
    category: 'string',
    odata: 'endswith',
    arity: 1,
    allowedTypes: STRING_FN,
    hint: 'endswith is always non-sargable; expect a table scan on large data.',
  },

  // ── Date — fixed relative ────────────────────────────────
  {
    id: 'On',
    label: 'on',
    kind: 'dv-fn-1',
    category: 'date-relative',
    odata: 'Microsoft.Dynamics.CRM.On',
    arity: 1,
    allowedTypes: DATE,
  },
  {
    id: 'OnOrAfter',
    label: 'on or after',
    kind: 'dv-fn-1',
    category: 'date-relative',
    odata: 'Microsoft.Dynamics.CRM.OnOrAfter',
    arity: 1,
    allowedTypes: DATE,
  },
  {
    id: 'OnOrBefore',
    label: 'on or before',
    kind: 'dv-fn-1',
    category: 'date-relative',
    odata: 'Microsoft.Dynamics.CRM.OnOrBefore',
    arity: 1,
    allowedTypes: DATE,
  },
  {
    id: 'Today',
    label: 'today',
    kind: 'dv-fn-0',
    category: 'date-relative',
    odata: 'Microsoft.Dynamics.CRM.Today',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'Tomorrow',
    label: 'tomorrow',
    kind: 'dv-fn-0',
    category: 'date-relative',
    odata: 'Microsoft.Dynamics.CRM.Tomorrow',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'Yesterday',
    label: 'yesterday',
    kind: 'dv-fn-0',
    category: 'date-relative',
    odata: 'Microsoft.Dynamics.CRM.Yesterday',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'Last7Days',
    label: 'last 7 days',
    kind: 'dv-fn-0',
    category: 'date-relative',
    odata: 'Microsoft.Dynamics.CRM.Last7Days',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'Next7Days',
    label: 'next 7 days',
    kind: 'dv-fn-0',
    category: 'date-relative',
    odata: 'Microsoft.Dynamics.CRM.Next7Days',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'LastWeek',
    label: 'last week',
    kind: 'dv-fn-0',
    category: 'date-relative',
    odata: 'Microsoft.Dynamics.CRM.LastWeek',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'ThisWeek',
    label: 'this week',
    kind: 'dv-fn-0',
    category: 'date-relative',
    odata: 'Microsoft.Dynamics.CRM.ThisWeek',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'NextWeek',
    label: 'next week',
    kind: 'dv-fn-0',
    category: 'date-relative',
    odata: 'Microsoft.Dynamics.CRM.NextWeek',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'LastMonth',
    label: 'last month',
    kind: 'dv-fn-0',
    category: 'date-relative',
    odata: 'Microsoft.Dynamics.CRM.LastMonth',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'ThisMonth',
    label: 'this month',
    kind: 'dv-fn-0',
    category: 'date-relative',
    odata: 'Microsoft.Dynamics.CRM.ThisMonth',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'NextMonth',
    label: 'next month',
    kind: 'dv-fn-0',
    category: 'date-relative',
    odata: 'Microsoft.Dynamics.CRM.NextMonth',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'LastYear',
    label: 'last year',
    kind: 'dv-fn-0',
    category: 'date-relative',
    odata: 'Microsoft.Dynamics.CRM.LastYear',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'ThisYear',
    label: 'this year',
    kind: 'dv-fn-0',
    category: 'date-relative',
    odata: 'Microsoft.Dynamics.CRM.ThisYear',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'NextYear',
    label: 'next year',
    kind: 'dv-fn-0',
    category: 'date-relative',
    odata: 'Microsoft.Dynamics.CRM.NextYear',
    arity: 0,
    allowedTypes: DATE,
  },

  // ── Date — rolling X-units ───────────────────────────────
  {
    id: 'LastXHours',
    label: 'in last X hours',
    kind: 'dv-fn-1-int',
    category: 'date-rolling',
    odata: 'Microsoft.Dynamics.CRM.LastXHours',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'LastXDays',
    label: 'in last X days',
    kind: 'dv-fn-1-int',
    category: 'date-rolling',
    odata: 'Microsoft.Dynamics.CRM.LastXDays',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'LastXWeeks',
    label: 'in last X weeks',
    kind: 'dv-fn-1-int',
    category: 'date-rolling',
    odata: 'Microsoft.Dynamics.CRM.LastXWeeks',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'LastXMonths',
    label: 'in last X months',
    kind: 'dv-fn-1-int',
    category: 'date-rolling',
    odata: 'Microsoft.Dynamics.CRM.LastXMonths',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'LastXYears',
    label: 'in last X years',
    kind: 'dv-fn-1-int',
    category: 'date-rolling',
    odata: 'Microsoft.Dynamics.CRM.LastXYears',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'NextXHours',
    label: 'in next X hours',
    kind: 'dv-fn-1-int',
    category: 'date-rolling',
    odata: 'Microsoft.Dynamics.CRM.NextXHours',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'NextXDays',
    label: 'in next X days',
    kind: 'dv-fn-1-int',
    category: 'date-rolling',
    odata: 'Microsoft.Dynamics.CRM.NextXDays',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'NextXWeeks',
    label: 'in next X weeks',
    kind: 'dv-fn-1-int',
    category: 'date-rolling',
    odata: 'Microsoft.Dynamics.CRM.NextXWeeks',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'NextXMonths',
    label: 'in next X months',
    kind: 'dv-fn-1-int',
    category: 'date-rolling',
    odata: 'Microsoft.Dynamics.CRM.NextXMonths',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'NextXYears',
    label: 'in next X years',
    kind: 'dv-fn-1-int',
    category: 'date-rolling',
    odata: 'Microsoft.Dynamics.CRM.NextXYears',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'OlderThanXMinutes',
    label: 'older than X minutes',
    kind: 'dv-fn-1-int',
    category: 'date-rolling',
    odata: 'Microsoft.Dynamics.CRM.OlderThanXMinutes',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'OlderThanXHours',
    label: 'older than X hours',
    kind: 'dv-fn-1-int',
    category: 'date-rolling',
    odata: 'Microsoft.Dynamics.CRM.OlderThanXHours',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'OlderThanXDays',
    label: 'older than X days',
    kind: 'dv-fn-1-int',
    category: 'date-rolling',
    odata: 'Microsoft.Dynamics.CRM.OlderThanXDays',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'OlderThanXWeeks',
    label: 'older than X weeks',
    kind: 'dv-fn-1-int',
    category: 'date-rolling',
    odata: 'Microsoft.Dynamics.CRM.OlderThanXWeeks',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'OlderThanXMonths',
    label: 'older than X months',
    kind: 'dv-fn-1-int',
    category: 'date-rolling',
    odata: 'Microsoft.Dynamics.CRM.OlderThanXMonths',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'OlderThanXYears',
    label: 'older than X years',
    kind: 'dv-fn-1-int',
    category: 'date-rolling',
    odata: 'Microsoft.Dynamics.CRM.OlderThanXYears',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },

  // ── Fiscal calendar ──────────────────────────────────────
  {
    id: 'ThisFiscalPeriod',
    label: 'this fiscal period',
    kind: 'dv-fn-0',
    category: 'fiscal',
    odata: 'Microsoft.Dynamics.CRM.ThisFiscalPeriod',
    arity: 0,
    allowedTypes: DATE,
    hint: 'Requires fiscal calendar configured on the org.',
  },
  {
    id: 'ThisFiscalYear',
    label: 'this fiscal year',
    kind: 'dv-fn-0',
    category: 'fiscal',
    odata: 'Microsoft.Dynamics.CRM.ThisFiscalYear',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'LastFiscalPeriod',
    label: 'last fiscal period',
    kind: 'dv-fn-0',
    category: 'fiscal',
    odata: 'Microsoft.Dynamics.CRM.LastFiscalPeriod',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'LastFiscalYear',
    label: 'last fiscal year',
    kind: 'dv-fn-0',
    category: 'fiscal',
    odata: 'Microsoft.Dynamics.CRM.LastFiscalYear',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'NextFiscalPeriod',
    label: 'next fiscal period',
    kind: 'dv-fn-0',
    category: 'fiscal',
    odata: 'Microsoft.Dynamics.CRM.NextFiscalPeriod',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'NextFiscalYear',
    label: 'next fiscal year',
    kind: 'dv-fn-0',
    category: 'fiscal',
    odata: 'Microsoft.Dynamics.CRM.NextFiscalYear',
    arity: 0,
    allowedTypes: DATE,
  },
  {
    id: 'LastXFiscalPeriods',
    label: 'in last X fiscal periods',
    kind: 'dv-fn-1-int',
    category: 'fiscal',
    odata: 'Microsoft.Dynamics.CRM.LastXFiscalPeriods',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'LastXFiscalYears',
    label: 'in last X fiscal years',
    kind: 'dv-fn-1-int',
    category: 'fiscal',
    odata: 'Microsoft.Dynamics.CRM.LastXFiscalYears',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'NextXFiscalPeriods',
    label: 'in next X fiscal periods',
    kind: 'dv-fn-1-int',
    category: 'fiscal',
    odata: 'Microsoft.Dynamics.CRM.NextXFiscalPeriods',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'NextXFiscalYears',
    label: 'in next X fiscal years',
    kind: 'dv-fn-1-int',
    category: 'fiscal',
    odata: 'Microsoft.Dynamics.CRM.NextXFiscalYears',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'InFiscalPeriod',
    label: 'in fiscal period N',
    kind: 'dv-fn-1-int',
    category: 'fiscal',
    odata: 'Microsoft.Dynamics.CRM.InFiscalPeriod',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'InFiscalYear',
    label: 'in fiscal year YYYY',
    kind: 'dv-fn-1-int',
    category: 'fiscal',
    odata: 'Microsoft.Dynamics.CRM.InFiscalYear',
    arity: 1,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'InFiscalPeriodAndYear',
    label: 'in fiscal period & year',
    kind: 'dv-fn-2',
    category: 'fiscal',
    odata: 'Microsoft.Dynamics.CRM.InFiscalPeriodAndYear',
    arity: 2,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'InOrAfterFiscalPeriodAndYear',
    label: 'in or after fiscal period & year',
    kind: 'dv-fn-2',
    category: 'fiscal',
    odata: 'Microsoft.Dynamics.CRM.InOrAfterFiscalPeriodAndYear',
    arity: 2,
    intValue: true,
    allowedTypes: DATE,
  },
  {
    id: 'InOrBeforeFiscalPeriodAndYear',
    label: 'in or before fiscal period & year',
    kind: 'dv-fn-2',
    category: 'fiscal',
    odata: 'Microsoft.Dynamics.CRM.InOrBeforeFiscalPeriodAndYear',
    arity: 2,
    intValue: true,
    allowedTypes: DATE,
  },

  // ── Range ────────────────────────────────────────────────
  {
    id: 'Between',
    label: 'between (inclusive)',
    kind: 'dv-fn-array',
    category: 'range',
    odata: 'Microsoft.Dynamics.CRM.Between',
    arity: 2,
    allowedTypes: RANGE_OK,
  },
  {
    id: 'NotBetween',
    label: 'not between',
    kind: 'dv-fn-array',
    category: 'range',
    odata: 'Microsoft.Dynamics.CRM.NotBetween',
    arity: 2,
    allowedTypes: RANGE_OK,
  },

  // ── Set membership ───────────────────────────────────────
  {
    id: 'In',
    label: 'is one of',
    kind: 'dv-fn-array',
    category: 'set',
    odata: 'Microsoft.Dynamics.CRM.In',
    arity: 'n',
    allowedTypes: SET_OK,
    hint: 'Compress dozens of OR-equality conditions into a single In(...) — helps with the 500-condition limit.',
  },
  {
    id: 'NotIn',
    label: 'not one of',
    kind: 'dv-fn-array',
    category: 'set',
    odata: 'Microsoft.Dynamics.CRM.NotIn',
    arity: 'n',
    allowedTypes: SET_OK,
  },

  // ── Multi-select choices ─────────────────────────────────
  {
    id: 'ContainValues',
    label: 'contains any of',
    kind: 'dv-fn-array',
    category: 'choices',
    odata: 'Microsoft.Dynamics.CRM.ContainValues',
    arity: 'n',
    intValue: true,
    allowedTypes: MULTI_CHOICE,
    hint: 'Choices columns only. Function name is ContainValues (not ContainsValues).',
  },
  {
    id: 'DoesNotContainValues',
    label: 'contains none of',
    kind: 'dv-fn-array',
    category: 'choices',
    odata: 'Microsoft.Dynamics.CRM.DoesNotContainValues',
    arity: 'n',
    intValue: true,
    allowedTypes: MULTI_CHOICE,
  },

  // ── Hierarchy ────────────────────────────────────────────
  // Hierarchy functions require the entity to have a hierarchical
  // relationship configured (e.g. account.parentaccountid). The
  // encoder rewrites the user's picked column to the entity's PRIMARY
  // KEY at emit time — that's what Dataverse demands for `PropertyName`
  // (test 5.13b succeeds with PK; test 5.13 fails with the lookup).
  // We don't pre-gate the entity here: most orgs have hierarchies
  // configured on more than just `account` (e.g. user, businessunit,
  // territory). Dataverse rejects with 0x80047020 at runtime if the
  // entity isn't hierarchical (test 13.11). `validateRequest` emits a
  // soft hint advisory.
  {
    id: 'Above',
    label: 'above (ancestor of)',
    kind: 'dv-fn-guid',
    category: 'hierarchy',
    odata: 'Microsoft.Dynamics.CRM.Above',
    arity: 1,
    allowedTypes: LOOKUP_FAMILY,
  },
  {
    id: 'AboveOrEqual',
    label: 'above or equal',
    kind: 'dv-fn-guid',
    category: 'hierarchy',
    odata: 'Microsoft.Dynamics.CRM.AboveOrEqual',
    arity: 1,
    allowedTypes: LOOKUP_FAMILY,
  },
  {
    id: 'Under',
    label: 'under (descendant of)',
    kind: 'dv-fn-guid',
    category: 'hierarchy',
    odata: 'Microsoft.Dynamics.CRM.Under',
    arity: 1,
    allowedTypes: LOOKUP_FAMILY,
  },
  {
    id: 'UnderOrEqual',
    label: 'under or equal',
    kind: 'dv-fn-guid',
    category: 'hierarchy',
    odata: 'Microsoft.Dynamics.CRM.UnderOrEqual',
    arity: 1,
    allowedTypes: LOOKUP_FAMILY,
  },
  {
    id: 'NotUnder',
    label: 'not under',
    kind: 'dv-fn-guid',
    category: 'hierarchy',
    odata: 'Microsoft.Dynamics.CRM.NotUnder',
    arity: 1,
    allowedTypes: LOOKUP_FAMILY,
  },

  // ── Current user / business / team ───────────────────────
  {
    id: 'EqualUserId',
    label: 'equals current user',
    kind: 'dv-fn-0',
    category: 'user-context',
    odata: 'Microsoft.Dynamics.CRM.EqualUserId',
    arity: 0,
    allowedTypes: LOOKUP_FAMILY,
  },
  {
    id: 'NotEqualUserId',
    label: 'does not equal current user',
    kind: 'dv-fn-0',
    category: 'user-context',
    odata: 'Microsoft.Dynamics.CRM.NotEqualUserId',
    arity: 0,
    allowedTypes: LOOKUP_FAMILY,
  },
  {
    id: 'EqualBusinessId',
    label: "equals user's business unit",
    kind: 'dv-fn-0',
    category: 'user-context',
    odata: 'Microsoft.Dynamics.CRM.EqualBusinessId',
    arity: 0,
    allowedTypes: LOOKUP_FAMILY,
  },
  {
    id: 'NotEqualBusinessId',
    label: "not user's business unit",
    kind: 'dv-fn-0',
    category: 'user-context',
    odata: 'Microsoft.Dynamics.CRM.NotEqualBusinessId',
    arity: 0,
    allowedTypes: LOOKUP_FAMILY,
  },
  {
    id: 'EqualRoleBusinessId',
    label: "equals role's business unit",
    kind: 'dv-fn-0',
    category: 'user-context',
    odata: 'Microsoft.Dynamics.CRM.EqualRoleBusinessId',
    arity: 0,
    allowedTypes: LOOKUP_FAMILY,
  },
  {
    id: 'EqualUserLanguage',
    label: "equals user's language",
    kind: 'dv-fn-0',
    category: 'user-context',
    odata: 'Microsoft.Dynamics.CRM.EqualUserLanguage',
    arity: 0,
    allowedTypes: ['Integer'],
  },
  {
    id: 'EqualUserTeams',
    label: "equals one of user's teams",
    kind: 'dv-fn-0',
    category: 'user-context',
    odata: 'Microsoft.Dynamics.CRM.EqualUserTeams',
    arity: 0,
    allowedTypes: LOOKUP_FAMILY,
  },
  {
    id: 'EqualUserOrUserTeams',
    label: 'equals user or user teams',
    kind: 'dv-fn-0',
    category: 'user-context',
    odata: 'Microsoft.Dynamics.CRM.EqualUserOrUserTeams',
    arity: 0,
    allowedTypes: LOOKUP_FAMILY,
  },
  {
    id: 'EqualUserOrUserHierarchy',
    label: 'equals user or hierarchy',
    kind: 'dv-fn-0',
    category: 'user-context',
    odata: 'Microsoft.Dynamics.CRM.EqualUserOrUserHierarchy',
    arity: 0,
    allowedTypes: LOOKUP_FAMILY,
  },
  {
    id: 'EqualUserOrUserHierarchyAndTeams',
    label: 'equals user / hierarchy / teams',
    kind: 'dv-fn-0',
    category: 'user-context',
    odata: 'Microsoft.Dynamics.CRM.EqualUserOrUserHierarchyAndTeams',
    arity: 0,
    allowedTypes: LOOKUP_FAMILY,
  },
];

export const findOperator = (id: string): OperatorDef | undefined =>
  OPERATORS.find((o) => o.id === id);

/** Return operators valid for the given AttributeTypeCode and table. */
export function operatorsFor(atc: AttributeTypeCode, table: string): OperatorDef[] {
  return OPERATORS.filter((op) => {
    if (op.allowedTypes && !op.allowedTypes.includes(atc)) return false;
    if (op.requiresTable && !op.requiresTable(table)) return false;
    return true;
  });
}

/** Operators that allow the RHS to be a bare property name (column-vs-column). */
export const isColumnComparable = (op: OperatorDef | undefined): boolean => !!op?.supportsColumnRhs;
