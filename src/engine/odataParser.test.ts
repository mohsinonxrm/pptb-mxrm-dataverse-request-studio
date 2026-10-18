import { beforeEach, describe, expect, it } from 'vitest';
import { parseODataUrl, topLevelSplit, type ParsedRequest } from './odataParser';
import { buildRetrieveMultiple } from './urlBuilder';
import { emptyApply } from '../editors/ApplyEditor';
import type {
  FilterFunctionNode,
  FilterGroup,
  FilterLambdaNode,
  FilterRule,
} from '../editors/filter/filterTree';
import { findTable } from '../mock/metadata';
import {
  account,
  contact,
  group,
  registerFixtureTables,
  retrieveMultipleState,
  rule,
} from '../test/fixtures';

const resolve = (set: string) =>
  [account, contact].find((t) => t.entitySetName === set)?.logicalName;
const loadTable = async (logical: string) => findTable(logical);

async function parseOk(url: string): Promise<ParsedRequest> {
  const r = await parseODataUrl(url, resolve);
  expect(r.errors).toEqual([]);
  expect(r.ok).toBe(true);
  return r.parsed!;
}

/** Strip generated ids so trees compare structurally. */
function shape(node: unknown): unknown {
  if (Array.isArray(node)) return node.map(shape);
  if (node && typeof node === 'object') {
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(node)) if (k !== 'id') out[k] = shape(v);
    return out;
  }
  return node;
}

beforeEach(() => {
  registerFixtureTables();
});

describe('parseODataUrl — URL handling', () => {
  it('accepts absolute, api-relative and bare entity-set URLs', async () => {
    for (const url of [
      'https://contoso.crm.dynamics.com/api/data/v9.2/accounts?$select=name',
      '/api/data/v9.2/accounts?$select=name',
      'accounts?$select=name',
    ]) {
      const p = await parseOk(url);
      expect(p.table).toBe('account');
      expect(p.entitySet).toBe('accounts');
      expect(p.select).toEqual(['name']);
    }
  });

  it('decodes a percent-encoded paste', async () => {
    const p = await parseOk('accounts?$filter=name%20eq%20%27A%26B%27');
    expect(shape(p.filter.rules)).toEqual([
      { type: 'rule', col: 'name', op: 'eq', val: 'A&B', valKind: 'literal' },
    ]);
  });

  it('rejects an entity set the environment does not know', async () => {
    const r = await parseODataUrl('widgets?$top=1', resolve);
    expect(r.ok).toBe(false);
    expect(r.errors[0].message).toMatch(/widgets/);
  });

  it('rejects unsupported query options', async () => {
    const r = await parseODataUrl('accounts?$search=contoso&$top=1', resolve);
    expect(r.ok).toBe(false);
    expect(r.errors.map((e) => e.message).join()).toMatch(/\$search/);
  });

  it('drops $skiptoken from a pasted nextLink with a warning instead of failing (#83)', async () => {
    const r = await parseODataUrl(
      'https://contoso.crm.dynamics.com/api/data/v9.2/accounts?$select=name,revenue&$orderby=name' +
        '&$skiptoken=%253Ccookie%2520pagenumber=%25222%2522%2520/%253E',
      resolve,
    );
    expect(r.ok).toBe(true);
    expect(r.errors).toEqual([]);
    expect(r.warnings).toHaveLength(1);
    expect(r.warnings[0].message).toMatch(/\$skiptoken/);
    expect(r.warnings[0].message).toMatch(/Retrieve NextLink/);
    expect(r.parsed?.select).toEqual(['name', 'revenue']);
    expect(r.parsed?.orderby.map((o) => [o.col, o.dir])).toEqual([['name', 'asc']]);
  });

  it('does not warn when there is no $skiptoken', async () => {
    const r = await parseODataUrl('accounts?$select=name', resolve);
    expect(r.warnings).toEqual([]);
  });
});

describe('parseODataUrl — $select / $orderby / scalars', () => {
  it('maps _<col>_value back to a lookup column but preserves unknown wire forms', async () => {
    const p = await parseOk('accounts?$select=name,_primarycontactid_value,_nosuch_value');
    expect(p.select).toEqual(['name', 'primarycontactid', '_nosuch_value']);
  });

  it('parses $orderby directions, $top and $count', async () => {
    const p = await parseOk('accounts?$orderby=name desc,createdon&$top=25&$count=true');
    expect(p.orderby.map((o) => [o.col, o.dir])).toEqual([
      ['name', 'desc'],
      ['createdon', 'asc'],
    ]);
    expect(p.top).toBe(25);
    expect(p.countOn).toBe(true);
  });
});

describe('parseODataUrl — $filter', () => {
  it('builds and/or groups and keeps escaped apostrophes', async () => {
    const p = await parseOk(
      "accounts?$filter=name eq 'O''Brien' and (numberofemployees gt 5 or donotemail eq true)",
    );
    expect(shape(p.filter)).toEqual({
      type: 'group',
      combinator: 'and',
      rules: [
        { type: 'rule', col: 'name', op: 'eq', val: "O'Brien", valKind: 'literal' },
        {
          type: 'group',
          combinator: 'or',
          rules: [
            { type: 'rule', col: 'numberofemployees', op: 'gt', val: '5', valKind: 'literal' },
            { type: 'rule', col: 'donotemail', op: 'eq', val: 'true', valKind: 'literal' },
          ],
        },
      ],
    });
  });

  it('classifies null checks, datetimes, GUIDs and column comparisons', async () => {
    const guid = '00000000-0000-0000-0000-0000000000c1';
    const p = await parseOk(
      `accounts?$filter=revenue eq null and createdon ge 2024-01-01T00:00:00Z and _primarycontactid_value eq ${guid} and name ne accountnumber`,
    );
    const rules = p.filter.rules as FilterRule[];
    expect(rules.map((r) => [r.col, r.op, r.val, r.valKind])).toEqual([
      ['revenue', 'is-null', '', undefined],
      ['createdon', 'ge', '2024-01-01T00:00:00Z', 'literal'],
      ['_primarycontactid_value', 'eq', guid, 'literal'],
      ['name', 'ne', 'accountnumber', 'column'],
    ]);
  });

  it('maps not on string functions and groups', async () => {
    const p = await parseOk(
      "accounts?$filter=not contains(name,'test') and not (revenue lt 1 or revenue gt 9)",
    );
    const [fn, grp] = p.filter.rules;
    expect(fn).toMatchObject({ type: 'rule', op: 'contains', val: 'test', negated: true });
    expect(grp).toMatchObject({ type: 'group', combinator: 'or', negated: true });
  });

  it('wraps a lone negated comparison in a negated group of one', async () => {
    const p = await parseOk('accounts?$filter=not numberofemployees eq 0');
    expect(p.filter).toMatchObject({
      type: 'group',
      negated: true,
      rules: [{ type: 'rule', op: 'eq', val: '0' }],
    });
  });

  // #117: parseAndExpr / parseOrExpr used to flatten any child group with the
  // same combinator — including a *negated* one — so the `not` was lost and
  // the pasted query was silently inverted.
  it('keeps a negated comparison negated inside an and-chain', async () => {
    const p = await parseOk("accounts?$filter=name eq 'x' and not numberofemployees eq 0");
    expect(shape(p.filter)).toEqual(
      shape(
        group('and', [
          rule('name', 'eq', 'x', { valKind: 'literal' }),
          group('and', [rule('numberofemployees', 'eq', '0', { valKind: 'literal' })], {
            negated: true,
          }),
        ]),
      ),
    );
  });

  it('keeps a negated same-combinator group intact (and / or)', async () => {
    const and = await parseOk(
      "accounts?$filter=name eq 'x' and not (revenue gt 1 and numberofemployees eq 0)",
    );
    expect(and.filter.rules).toHaveLength(2);
    expect(and.filter.rules[1]).toMatchObject({
      type: 'group',
      combinator: 'and',
      negated: true,
      rules: [{ col: 'revenue' }, { col: 'numberofemployees' }],
    });

    const or = await parseOk(
      "accounts?$filter=name eq 'x' or not (revenue gt 1 or numberofemployees eq 0)",
    );
    expect(or.filter).toMatchObject({ combinator: 'or' });
    expect(or.filter.rules).toHaveLength(2);
    expect(or.filter.rules[1]).toMatchObject({
      type: 'group',
      combinator: 'or',
      negated: true,
      rules: [{ col: 'revenue' }, { col: 'numberofemployees' }],
    });
  });

  it('keeps a not inside an or-branch', async () => {
    const p = await parseOk("accounts?$filter=name eq 'x' or not numberofemployees eq 0");
    expect(p.filter).toMatchObject({ combinator: 'or' });
    expect(p.filter.negated).toBeFalsy();
    expect(p.filter.rules[1]).toMatchObject({
      type: 'group',
      negated: true,
      rules: [{ col: 'numberofemployees', op: 'eq', val: '0' }],
    });
  });

  it('keeps nested nots at every level', async () => {
    const p = await parseOk(
      "accounts?$filter=name eq 'x' and not (revenue gt 1 and not (numberofemployees eq 0 and donotemail eq true))",
    );
    const outer = p.filter.rules[1] as FilterGroup;
    expect(outer).toMatchObject({ type: 'group', combinator: 'and', negated: true });
    expect(outer.rules).toHaveLength(2);
    expect(outer.rules[1]).toMatchObject({
      type: 'group',
      combinator: 'and',
      negated: true,
      rules: [{ col: 'numberofemployees' }, { col: 'donotemail' }],
    });
  });

  it('cancels a double not, and only then flattens the group', async () => {
    const p = await parseOk(
      "accounts?$filter=name eq 'x' and not not (revenue gt 1 and numberofemployees eq 0)",
    );
    expect(p.filter.negated).toBeFalsy();
    expect((p.filter.rules as FilterRule[]).map((r) => [r.type, r.col])).toEqual([
      ['rule', 'name'],
      ['rule', 'revenue'],
      ['rule', 'numberofemployees'],
    ]);
  });

  it('round-trips negated groups through the URL builder', async () => {
    const rebuild = async (url: string) => {
      const p = await parseOk(url);
      return decodeURIComponent(
        buildRetrieveMultiple(retrieveMultipleState({ ...p, apply: p.apply ?? emptyApply() }))
          .relativeUrl,
      );
    };
    const nots = (s: string) => s.match(/\bnot \(/g)?.length ?? 0;
    for (const filter of [
      "name eq 'x' and not (numberofemployees eq 0)",
      "name eq 'x' or not (revenue gt 1 or numberofemployees eq 0)",
      "name eq 'x' and not (revenue gt 1 and not (numberofemployees eq 0))",
    ]) {
      const once = await rebuild(`accounts?$filter=${filter}`);
      expect(nots(once)).toBe(nots(filter));
      // Parsing the rebuilt URL must be a fixed point.
      expect(await rebuild(once)).toBe(once);
    }
  });

  it('parses Dataverse query functions of each arity', async () => {
    const p = await parseOk(
      "accounts?$filter=Microsoft.Dynamics.CRM.LastXDays(PropertyName='createdon',PropertyValue=30)" +
        " and Microsoft.Dynamics.CRM.In(PropertyName='industrycode',PropertyValues=['1','2'])" +
        " and Microsoft.Dynamics.CRM.Today(PropertyName='createdon')",
    );
    const fns = p.filter.rules as FilterFunctionNode[];
    expect(fns.map((f) => [f.type, f.op, f.col, f.val, f.values])).toEqual([
      ['function', 'LastXDays', 'createdon', '30', undefined],
      ['function', 'In', 'industrycode', undefined, ['1', '2']],
      ['function', 'Today', 'createdon', undefined, undefined],
    ]);
  });

  it('drops `not` before a Dataverse function with a warning', async () => {
    const r = await parseODataUrl(
      "accounts?$filter=not Microsoft.Dynamics.CRM.Today(PropertyName='createdon')",
      resolve,
    );
    expect(r.ok).toBe(true);
    expect(r.warnings[0].message).toMatch(/Not\*/);
    expect(r.parsed?.filter.rules[0]).toMatchObject({ type: 'function', op: 'Today' });
  });

  it('rejects unknown and un-prefixed functions', async () => {
    const unknown = await parseODataUrl(
      "accounts?$filter=Microsoft.Dynamics.CRM.NoSuchFn(PropertyName='name')",
      resolve,
    );
    expect(unknown.ok).toBe(false);
    expect(unknown.errors[0].message).toMatch(/NoSuchFn/);

    const bare = await parseODataUrl("accounts?$filter=Today(PropertyName='createdon')", resolve);
    expect(bare.ok).toBe(false);
    expect(bare.errors[0].message).toMatch(/Microsoft\.Dynamics\.CRM\.Today/);
  });

  it('parses lambdas, including nested ones with an outer alias', async () => {
    const p = await parseOk(
      "accounts?$filter=contact_customer_accounts/any(c:c/jobtitle eq 'CEO' and c/Contact_Tasks/all(t:t/subject ne null))",
    );
    const lambda = p.filter.rules[0] as FilterLambdaNode;
    expect(lambda).toMatchObject({
      type: 'lambda',
      nav: 'contact_customer_accounts',
      lambda: 'any',
      alias: 'c',
    });
    const inner = lambda.inner.rules;
    expect(inner[0]).toMatchObject({ type: 'rule', col: 'c/jobtitle', val: 'CEO' });
    expect(inner[1]).toMatchObject({
      type: 'lambda',
      nav: 'Contact_Tasks',
      lambda: 'all',
      alias: 't',
    });
  });

  it('reports a syntax error instead of throwing', async () => {
    const r = await parseODataUrl('accounts?$filter=name eq', resolve);
    expect(r.ok).toBe(false);
    expect(r.errors[0].message).toMatch(/missing its RHS/);
  });
});

describe('parseODataUrl — $expand / $apply', () => {
  it('parses nested $expand with inner options', async () => {
    const p = await parseOk(
      "accounts?$expand=contact_customer_accounts($select=fullname;$filter=jobtitle eq 'CEO';$orderby=fullname desc;$top=3;$expand=parentcustomerid_account($select=name)),primarycontactid",
    );
    expect(p.expand.map((e) => e.nav)).toEqual(['contact_customer_accounts', 'primarycontactid']);
    const [first] = p.expand;
    expect(first.select).toEqual(['fullname']);
    expect(first.top).toBe(3);
    expect(first.orderby.map((o) => [o.col, o.dir])).toEqual([['fullname', 'desc']]);
    expect((first.filter as FilterGroup).rules[0]).toMatchObject({ col: 'jobtitle', val: 'CEO' });
    expect(first.nestedExpand?.map((e) => [e.nav, e.select])).toEqual([
      ['parentcustomerid_account', ['name']],
    ]);
  });

  it('rejects unsupported inner $expand options', async () => {
    const r = await parseODataUrl('accounts?$expand=primarycontactid($count=true)', resolve);
    expect(r.ok).toBe(false);
    expect(r.errors[0].message).toMatch(/\$count/);
  });

  it('parses a filter / groupby / aggregate pipeline', async () => {
    const p = await parseOk(
      'accounts?$apply=filter(numberofemployees gt 0)/groupby((industrycode),aggregate(revenue with sum as total,$count as n))',
    );
    expect(p.apply).toMatchObject({
      enabled: true,
      groupby: ['industrycode'],
      aggregates: [
        { col: 'revenue', fn: 'sum', alias: 'total' },
        { col: '$count', fn: '$count', alias: 'n' },
      ],
    });
    expect(p.apply?.prefilter.rules[0]).toMatchObject({ col: 'numberofemployees', op: 'gt' });
  });

  it('rejects unmodelled $apply stages and aggregate functions', async () => {
    const stage = await parseODataUrl('accounts?$apply=topcount(5,revenue)', resolve);
    expect(stage.ok).toBe(false);
    const fn = await parseODataUrl(
      'accounts?$apply=aggregate(revenue with countdistinct as d)',
      resolve,
    );
    expect(fn.ok).toBe(false);
    expect(fn.errors[0].message).toMatch(/countdistinct/);
  });
});

describe('parseODataUrl — metadata validation', () => {
  it('flags columns and navs that do not exist on the table', async () => {
    const r = await parseODataUrl(
      "accounts?$select=name,nosuchcol&$filter=bogus eq 'x'&$expand=nosuchnav",
      resolve,
      loadTable,
    );
    expect(r.ok).toBe(false);
    const msgs = r.errors.map((e) => e.message).join('\n');
    expect(msgs).toMatch(/nosuchcol/);
    expect(msgs).toMatch(/bogus/);
    expect(msgs).toMatch(/nosuchnav/);
  });

  it('passes a query whose references all resolve', async () => {
    const r = await parseODataUrl(
      "accounts?$select=name,_primarycontactid_value&$filter=contact_customer_accounts/any(c:c/jobtitle eq 'CEO')",
      resolve,
      loadTable,
    );
    expect(r.errors).toEqual([]);
    expect(r.ok).toBe(true);
  });
});

describe('builder ↔ parser round trip', () => {
  it('re-builds the same URL from a parsed builder URL', async () => {
    const original = buildRetrieveMultiple(
      retrieveMultipleState({
        select: ['name', 'primarycontactid', 'revenue'],
        filter: group('and', [
          rule('name', 'startswith', "O'Br"),
          group('or', [rule('numberofemployees', 'ge', '10'), rule('donotemail', 'eq', 'false')]),
        ]),
        orderby: [{ col: 'revenue', dir: 'desc' }],
        top: 20,
        countOn: true,
      }),
    );
    const p = await parseOk(original.relativeUrl);
    const rebuilt = buildRetrieveMultiple(
      retrieveMultipleState({ ...p, apply: p.apply ?? emptyApply() }),
    );
    expect(rebuilt.relativeUrl).toBe(original.relativeUrl);
  });
});

describe('topLevelSplit', () => {
  it('splits only outside parens, brackets and quoted strings', () => {
    expect(topLevelSplit("a($select=x,y),b,c('1,2'),d=[1,2]", ',')).toEqual([
      'a($select=x,y)',
      'b',
      "c('1,2')",
      'd=[1,2]',
    ]);
    expect(topLevelSplit("$filter=name eq 'it''s & co'&$top=1", '&')).toEqual([
      "$filter=name eq 'it''s & co'",
      '$top=1',
    ]);
  });
});
