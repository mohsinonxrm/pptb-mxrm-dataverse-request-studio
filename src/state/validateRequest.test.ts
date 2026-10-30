// validateRequest: the execute-time checks on the filter / expand / $apply
// tree that block Execute (errors) or only advise (warnings, infos) (#235).
// Literal-value and name checks (#221) are covered in filterTree.test.ts.
import { beforeAll, describe, expect, it } from 'vitest';
import { validateRequest, type ValidateRequestInput } from './validateRequest';
import { __registerLiveTables, type TableMeta } from '../mock/metadata';
import type { FilterGroup, FilterNode } from '../editors/filter/filterTree';
import type { ExpandSpec } from '../editors/ExpandEditor';
import type { ApplyAgg, ApplySpec } from '../editors/ApplyEditor';
import { disabledReasonFromAdvisories } from '../primitives/advisories';

const account: TableMeta = {
  logicalName: 'account',
  entitySetName: 'accounts',
  displayName: 'Account',
  primaryKey: 'accountid',
  primaryName: 'name',
  columns: [
    { logicalName: 'accountid', displayName: 'Account', attributeType: 'Uniqueidentifier' },
    { logicalName: 'name', displayName: 'Name', attributeType: 'String', format: 'Text' },
    { logicalName: 'revenue', displayName: 'Revenue', attributeType: 'Money' },
    {
      logicalName: 'createdon',
      displayName: 'Created On',
      attributeType: 'DateTime',
      format: 'DateAndTime',
    },
    { logicalName: 'primarycontactid', displayName: 'Primary Contact', attributeType: 'Lookup' },
    { logicalName: 'customerid', displayName: 'Customer', attributeType: 'Customer' },
    { logicalName: 'ownerid', displayName: 'Owner', attributeType: 'Owner' },
  ] as TableMeta['columns'],
  navigationProperties: [
    { name: 'primarycontactid', cardinality: 'ManyToOne', targetEntity: 'contact' },
    { name: 'contact_customer_accounts', cardinality: 'OneToMany', targetEntity: 'contact' },
    { name: 'accountleads_association', cardinality: 'ManyToMany', targetEntity: 'lead' },
  ] as TableMeta['navigationProperties'],
};

const contact: TableMeta = {
  logicalName: 'contact',
  entitySetName: 'contacts',
  displayName: 'Contact',
  primaryKey: 'contactid',
  primaryName: 'fullname',
  columns: [
    { logicalName: 'contactid', displayName: 'Contact', attributeType: 'Uniqueidentifier' },
    { logicalName: 'fullname', displayName: 'Full Name', attributeType: 'String', format: 'Text' },
    { logicalName: 'ownerid', displayName: 'Owner', attributeType: 'Owner' },
  ] as TableMeta['columns'],
  navigationProperties: [
    { name: 'parentcustomerid_account', cardinality: 'ManyToOne', targetEntity: 'account' },
  ] as TableMeta['navigationProperties'],
};

// `lead` is deliberately not registered: its metadata "hasn't loaded".
beforeAll(() => __registerLiveTables([account, contact]));

let n = 0;
const group = (rules: FilterNode[], extra: Partial<FilterGroup> = {}): FilterGroup => ({
  id: `g${++n}`,
  type: 'group',
  combinator: 'and',
  rules,
  ...extra,
});
const rule = (col: string, op = 'eq', val = 'x'): FilterNode => ({
  id: `r${++n}`,
  type: 'rule',
  col,
  op,
  val,
});
const fn = (op: string, col: string, extra: object = {}): FilterNode =>
  ({ id: `f${++n}`, type: 'function', op, col, ...extra }) as FilterNode;
const lambda = (nav: string, alias: string, inner: FilterNode[]): FilterNode => ({
  id: `l${++n}`,
  type: 'lambda',
  nav,
  lambda: 'any',
  alias,
  inner: group(inner),
});
const expand = (nav: string, nestedExpand?: ExpandSpec[]): ExpandSpec => ({
  id: `e${++n}`,
  nav,
  select: [],
  orderby: [],
  nestedExpand,
});
const apply = (extra: Partial<ApplySpec>): ApplySpec => ({
  enabled: true,
  prefilter: group([]),
  groupby: [],
  aggregates: [],
  ...extra,
});
const agg = (fnName: ApplyAgg['fn'], col: string, alias: string): ApplyAgg => ({
  fn: fnName,
  col,
  alias,
});

const run = (input: Omit<ValidateRequestInput, 'table'>) =>
  validateRequest({ table: 'account', ...input });
const ids = (input: Omit<ValidateRequestInput, 'table'>) => run(input).map((a) => a.id);
const errors = (input: Omit<ValidateRequestInput, 'table'>) =>
  run(input).filter((a) => a.severity === 'error');

describe('when there is nothing to check', () => {
  it('returns nothing without a table, or before its metadata has loaded', () => {
    const filter = group([fn('In', 'name', { negated: true })]);
    expect(validateRequest({ filter })).toEqual([]);
    expect(validateRequest({ table: 'lead', filter })).toEqual([]);
  });

  it('returns nothing for a clean request', () => {
    expect(
      run({
        filter: group([rule('name'), fn('In', 'name', { values: ['a'] })]),
        expand: [expand('primarycontactid')],
        orderby: [{ col: 'name', dir: 'asc' }] as ValidateRequestInput['orderby'],
      }),
    ).toEqual([]);
  });
});

describe('negation', () => {
  it('blocks `not` on a Dataverse function', () => {
    const f = fn('In', 'name', { negated: true });
    expect(errors({ filter: group([f]) }).map((a) => a.id)).toEqual([`fn-negated-In-${f.id}`]);
  });

  it('allows `not` on a plain comparison', () => {
    expect(run({ filter: group([{ ...rule('name'), negated: true } as FilterNode]) })).toEqual([]);
  });

  it('allows a negated group of plain comparisons', () => {
    expect(run({ filter: group([group([rule('name')], { negated: true })]) })).toEqual([]);
  });

  it('blocks a negated group with a Dataverse function anywhere inside it', () => {
    const direct = group([fn('In', 'name')], { negated: true });
    expect(ids({ filter: group([direct]) })).toEqual([`group-not-with-dvfn-${direct.id}`]);

    const nested = group([group([rule('name'), fn('In', 'name')])], { negated: true });
    expect(ids({ filter: group([nested]) })).toEqual([`group-not-with-dvfn-${nested.id}`]);

    const inLambda = group([lambda('contact_customer_accounts', 'c', [fn('In', 'c/fullname')])], {
      negated: true,
    });
    expect(ids({ filter: group([inLambda]) })).toEqual([`group-not-with-dvfn-${inLambda.id}`]);
  });

  it('checks a negated root group too', () => {
    const root = group([fn('In', 'name')], { negated: true });
    expect(ids({ filter: root })).toEqual([`group-not-with-dvfn-${root.id}`]);
  });
});

describe('Dataverse function columns', () => {
  it('blocks a navigation path in PropertyName', () => {
    const f = fn('EqualUserId', 'primarycontactid/ownerid');
    const [a] = errors({ filter: group([f]) });
    expect(a.id).toBe(`fn-nav-path-EqualUserId-${f.id}`);
    expect(a.title).toBe('Nav-path `primarycontactid/ownerid` in `PropertyName` is rejected.');
  });

  it("allows a lambda alias in front of the inner table's column", () => {
    const filter = group([
      lambda('contact_customer_accounts', 'c', [fn('EqualUserId', 'c/ownerid')]),
    ]);
    expect(run({ filter })).toEqual([]);
  });

  it('marks a hierarchy function as information only, so Execute is not blocked', () => {
    const f = fn('Above', 'accountid', { val: '00000000-0000-0000-0000-000000000001' });
    const out = run({ filter: group([f]) });
    expect(out.map((a) => [a.id, a.severity])).toEqual([[`fn-hierarchy-Above-${f.id}`, 'info']]);
    expect(out[0].title).toContain('`account`');
    expect(disabledReasonFromAdvisories(out)).toBeNull();
  });
});

describe('lambdas', () => {
  it('blocks a navigation property the table does not have', () => {
    const l = lambda('gone_nav', 'c', [rule('c/fullname')]);
    expect(ids({ filter: group([l]) })).toEqual([`lambda-unknown-nav-${l.id}`]);
  });

  it('blocks a single-valued (N:1) navigation property', () => {
    const l = lambda('primarycontactid', 'c', [rule('c/fullname')]);
    expect(ids({ filter: group([l]) })).toEqual([`lambda-on-n1-${l.id}`]);
  });

  it('blocks a missing alias', () => {
    const l = lambda('contact_customer_accounts', '  ', [rule('fullname')]);
    expect(errors({ filter: group([l]) }).map((a) => a.id)).toContain(`lambda-no-alias-${l.id}`);
  });

  it('checks the inner filter against the target table', () => {
    const inner = fn('In', 'c/fullname', { negated: true });
    const l = lambda('contact_customer_accounts', 'c', [inner]);
    expect(ids({ filter: group([l]) })).toEqual([`fn-negated-In-${inner.id}`]);
    // The hierarchy hint names the inner table.
    const h = fn('Above', 'c/contactid');
    const [hint] = run({ filter: group([lambda('contact_customer_accounts', 'c', [h])]) });
    expect(hint.title).toContain('`contact`');
  });

  it("skips the inner checks while the target table's metadata is loading", () => {
    const l = lambda('accountleads_association', 'l', [fn('In', 'l/subject', { negated: true })]);
    expect(run({ filter: group([l]) })).toEqual([]);
  });
});

describe('the 500-condition cap', () => {
  const rules = (count: number) => Array.from({ length: count }, () => rule('name'));

  it('allows 500 conditions', () => {
    expect(run({ filter: group(rules(500)) })).toEqual([]);
  });

  it('blocks 501, counting conditions in nested groups', () => {
    const out = errors({ filter: group([...rules(400), group(rules(101))]) });
    expect(out.map((a) => [a.id, a.title])).toEqual([
      ['filter-cond-cap', 'Filter has 501 conditions, exceeds the 500 hard cap.'],
    ]);
  });
});

describe('$expand', () => {
  const many = (count: number) => Array.from({ length: count }, () => expand('primarycontactid'));

  it('warns, without blocking, past the documented 15 expands', () => {
    expect(run({ expand: many(15) })).toEqual([]);
    const out = run({
      expand: [...many(14), expand('primarycontactid', [expand('x'), expand('y')])],
    });
    expect(out.map((a) => [a.id, a.severity])).toEqual([['expand-cap', 'warning']]);
    expect(out[0].title).toBe('17 expand options exceeds the documented 15 cap.');
  });

  it('notes the 5,000-row cap on a collection expand with nothing nested', () => {
    for (const nav of ['contact_customer_accounts', 'accountleads_association']) {
      const out = run({ expand: [expand(nav)] });
      expect(out.map((a) => [a.id, a.severity])).toEqual([['expand-flat-collection', 'info']]);
    }
  });

  it('does not note it for a single-valued expand, a nested one, or an unknown nav', () => {
    expect(run({ expand: [expand('primarycontactid')] })).toEqual([]);
    expect(
      run({ expand: [expand('contact_customer_accounts', [expand('parentcustomerid_account')])] }),
    ).toEqual([]);
    expect(run({ expand: [expand('no_such_nav')] })).toEqual([]);
  });
});

describe('$apply', () => {
  it('checks the prefilter, not the $filter, when $apply is on', () => {
    const blocked = group([fn('In', 'name', { negated: true })]);
    expect(run({ filter: blocked, apply: apply({}) })).toEqual([]);
    expect(errors({ filter: group([]), apply: apply({ prefilter: blocked }) })).toHaveLength(1);
  });

  it('ignores the $apply rules while $apply is off', () => {
    expect(
      run({
        orderby: [{ col: 'name', dir: 'asc' }] as ValidateRequestInput['orderby'],
        apply: apply({ enabled: false, groupby: ['createdon'], aggregates: [agg('sum', '', '')] }),
      }),
    ).toEqual([]);
  });

  it('blocks $orderby on aggregate output', () => {
    expect(
      ids({
        orderby: [{ col: 'name', dir: 'asc' }] as ValidateRequestInput['orderby'],
        apply: apply({}),
      }),
    ).toEqual(['apply-orderby-conflict']);
    expect(ids({ orderby: [], apply: apply({}) })).toEqual([]);
  });

  it('blocks groupby on a DateTime column', () => {
    expect(ids({ apply: apply({ groupby: ['createdon'] }) })).toEqual([
      'apply-groupby-datetime-createdon',
    ]);
  });

  it('blocks groupby on a bare lookup, customer or owner column', () => {
    expect(
      ids({ apply: apply({ groupby: ['primarycontactid', 'customerid', 'ownerid'] }) }),
    ).toEqual([
      'apply-groupby-lookup-primarycontactid',
      'apply-groupby-lookup-customerid',
      'apply-groupby-lookup-ownerid',
    ]);
  });

  it('blocks groupby on the _<attr>_value form', () => {
    expect(ids({ apply: apply({ groupby: ['_primarycontactid_value'] }) })).toEqual([
      'apply-groupby-value-_primarycontactid_value',
    ]);
  });

  it('allows groupby on a scalar column or through a lookup to one', () => {
    expect(run({ apply: apply({ groupby: ['name', 'primarycontactid/fullname'] }) })).toEqual([]);
  });

  it('blocks an aggregate with no alias', () => {
    const out = errors({ apply: apply({ aggregates: [agg('$count', '', ' ')] }) });
    expect(out.map((a) => a.title)).toEqual(['Aggregate is missing an alias.']);
  });

  it('blocks two aggregates with the same alias', () => {
    expect(
      ids({
        apply: apply({ aggregates: [agg('sum', 'revenue', 'x'), agg('max', 'revenue', 'x')] }),
      }),
    ).toEqual(['apply-agg-dup-alias-x']);
  });

  it('blocks a scalar aggregate with no column, but not $count', () => {
    expect(ids({ apply: apply({ aggregates: [agg('average', ' ', 'avg')] }) })).toEqual([
      'apply-agg-no-col-avg',
    ]);
    expect(run({ apply: apply({ aggregates: [agg('$count', '', 'total')] }) })).toEqual([]);
  });
});

describe('gating Execute', () => {
  it('every blocker is an error, so the Execute button names it', () => {
    const out = run({
      orderby: [{ col: 'name', dir: 'asc' }] as ValidateRequestInput['orderby'],
      apply: apply({ groupby: ['createdon'] }),
    });
    expect(out.every((a) => a.severity === 'error' && a.source === 'validation')).toBe(true);
    expect(disabledReasonFromAdvisories(out)).toBe('2 blockers — open Advisories to review.');
  });

  it('warnings and infos alone do not block', () => {
    const out = run({
      expand: Array.from({ length: 16 }, () => expand('contact_customer_accounts')),
    });
    expect(out.map((a) => a.severity).sort()).toEqual(['info', 'warning']);
    expect(disabledReasonFromAdvisories(out)).toBeNull();
  });
});
