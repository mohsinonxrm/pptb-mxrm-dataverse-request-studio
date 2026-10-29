// #221: typed filter literals go unquoted only when they have the column's
// shape; anything else is quoted and escaped, and flagged so Execute blocks.
import { beforeAll, describe, expect, it } from 'vitest';
import { __registerLiveTables, findTable, type TableMeta } from '../../mock/metadata';
import {
  _setTableLookup,
  collectInvalidLiterals,
  collectInvalidNames,
  groupToOData,
  type FilterGroup,
  type FilterNode,
} from './filterTree';
import { validateRequest } from '../../state/validateRequest';
import { filterToQueryExpression } from '../../engine/bulkDelete';

const ID = '00000000-0000-0000-0000-0000000000a1';

const account: TableMeta = {
  logicalName: 'account',
  entitySetName: 'accounts',
  displayName: 'Account',
  primaryKey: 'accountid',
  primaryName: 'name',
  columns: [
    { logicalName: 'accountid', displayName: 'Account', attributeType: 'Uniqueidentifier' },
    { logicalName: 'name', displayName: 'Name', attributeType: 'String', format: 'Text' },
    { logicalName: 'numberofemployees', displayName: 'Employees', attributeType: 'Integer' },
    { logicalName: 'revenue', displayName: 'Revenue', attributeType: 'Money' },
    {
      logicalName: 'createdon',
      displayName: 'Created On',
      attributeType: 'DateTime',
      format: 'DateAndTime',
      dateTimeBehavior: 'UserLocal',
    },
    { logicalName: 'objecttypecode', displayName: 'Type', attributeType: 'EntityName' },
    {
      logicalName: 'parentaccountid',
      displayName: 'Parent',
      attributeType: 'Lookup',
      targets: ['account'],
    },
  ] as TableMeta['columns'],
  navigationProperties: [
    {
      name: 'contact_customer_accounts',
      cardinality: 'OneToMany',
      targetEntity: 'account',
    },
  ] as TableMeta['navigationProperties'],
};

beforeAll(() => {
  __registerLiveTables([account]);
  _setTableLookup(findTable);
});

const tree = (...rules: FilterNode[]): FilterGroup => ({
  id: 'root',
  type: 'group',
  combinator: 'and',
  rules,
});
const rule = (col: string, val: string, op = 'eq'): FilterNode => ({
  id: `r-${col}`,
  type: 'rule',
  col,
  op,
  val,
});
const fn = (op: string, col: string, v: Partial<FilterNode>): FilterNode =>
  ({ id: `f-${op}`, type: 'function', op, col, ...v }) as FilterNode;

describe('typed literals', () => {
  it('stay unquoted when they have the column shape', () => {
    expect(
      groupToOData(
        tree(
          rule('numberofemployees', '42'),
          rule('revenue', '1.5', 'gt'),
          rule('createdon', '2024-01-31T10:00:00.000Z', 'gt'),
          rule('accountid', ID),
        ),
        account,
      ),
    ).toBe(
      `numberofemployees eq 42 and revenue gt 1.5 and createdon gt 2024-01-31T10:00:00.000Z and accountid eq ${ID}`,
    );
  });

  it('are quoted and escaped when they do not, so they cannot extend the expression', () => {
    expect(groupToOData(tree(rule('numberofemployees', "1 or name ne 'x'")), account)).toBe(
      "numberofemployees eq '1 or name ne ''x'''",
    );
    expect(groupToOData(tree(rule('accountid', `${ID}) or (1 eq 1`)), account)).toBe(
      `accountid eq '${ID}) or (1 eq 1'`,
    );
    // The DateTime fallback used to quote without escaping.
    expect(groupToOData(tree(rule('createdon', "today' or name ne '")), account)).toBe(
      "createdon eq 'today'' or name ne '''",
    );
  });

  it('Dataverse functions: whole numbers and GUIDs are checked and escaped', () => {
    expect(
      groupToOData(tree(fn('LastXDays', 'createdon', { val: "7) or (name ne 'x'" })), account),
    ).toBe(
      "Microsoft.Dynamics.CRM.LastXDays(PropertyName='createdon',PropertyValue='7) or (name ne ''x''')",
    );
    expect(
      groupToOData(tree(fn('Under', 'parentaccountid', { val: "x') or (name ne 'y" })), account),
    ).toBe(
      "Microsoft.Dynamics.CRM.Under(PropertyName='accountid',PropertyValue='x'') or (name ne ''y')",
    );
    expect(
      groupToOData(
        tree(fn('InFiscalPeriodAndYear', 'createdon', { vals: ['3', '2024'] })),
        account,
      ),
    ).toBe(
      "Microsoft.Dynamics.CRM.InFiscalPeriodAndYear(PropertyName='createdon',PropertyValue1=3,PropertyValue2=2024)",
    );
  });
});

describe('collectInvalidLiterals', () => {
  it('lists each value without its column shape, and nothing for valid ones', () => {
    const bad = tree(
      rule('numberofemployees', 'many'),
      rule('name', 'anything goes'),
      rule('createdon', 'yesterday'),
      fn('LastXDays', 'createdon', { val: 'seven' }),
      fn('Under', 'parentaccountid', { val: 'not-a-guid' }),
      fn('On', 'createdon', { val: '31/01/2024' }),
    );
    expect(collectInvalidLiterals(bad, account).map((l) => [l.col, l.expected])).toEqual([
      ['numberofemployees', 'a whole number'],
      ['createdon', 'an ISO 8601 date or date-time'],
      ['createdon', 'a whole number'],
      ['parentaccountid', 'a GUID'],
      ['createdon', 'an ISO 8601 date or date-time'],
    ]);
    const good = tree(
      rule('numberofemployees', '5'),
      fn('LastXDays', 'createdon', { val: '7' }),
      fn('Under', 'parentaccountid', { val: ID }),
      fn('On', 'createdon', { val: '2024-01-31' }),
    );
    expect(collectInvalidLiterals(good, account)).toEqual([]);
  });

  it('blocks Execute in Retrieve Multiple, naming the value', () => {
    const errors = validateRequest({ table: 'account', filter: tree(rule('revenue', 'lots')) })
      .filter((a) => a.severity === 'error')
      .map((a) => a.title);
    expect(errors).toEqual(['`revenue` needs a number, not "lots".']);
  });

  it('is a Bulk Delete problem too', () => {
    expect(filterToQueryExpression(account, tree(rule('numberofemployees', 'x'))).problems).toEqual(
      ['"numberofemployees" needs a whole number, not "x".'],
    );
  });
});

describe('review follow-ups', () => {
  it('EntityName is a string: quoted, and not flagged', () => {
    const t = tree(rule('objecttypecode', 'account'));
    expect(groupToOData(t, account)).toBe("objecttypecode eq 'account'");
    expect(collectInvalidLiterals(t, account)).toEqual([]);
  });

  it('names: a column, a column-to-column right-hand side and a lambda nav must be names', () => {
    const t = tree(
      rule("name eq 'a' or name", 'x'),
      {
        id: 'c',
        type: 'rule',
        col: 'name',
        op: 'eq',
        valKind: 'column',
        val: 'name or createdon ne null',
      },
      {
        id: 'l',
        type: 'lambda',
        nav: 'contact_customer_accounts/any(x:true) or x',
        lambda: 'any',
        alias: 'c',
        inner: tree(),
      },
      rule('parentaccountid/name', 'fine'),
    );
    expect(collectInvalidNames(t)).toEqual([
      `Filter column "name eq 'a' or name" isn't a valid name.`,
      `Filter column "name or createdon ne null" isn't a valid name.`,
      `Filter navigation property "contact_customer_accounts/any(x:true) or…" isn't a valid name (letters, digits and underscores).`,
    ]);
    const blockers = validateRequest({ table: 'account', filter: t }).filter(
      (a) => a.severity === 'error' && a.id.startsWith('query-name-'),
    );
    expect(blockers).toHaveLength(3);
  });

  it("checks $expand's navigation property and the literals and names of its inner filter", () => {
    const advisories = validateRequest({
      table: 'account',
      expand: [
        {
          id: 'e',
          nav: 'contact_customer_accounts',
          select: [],
          orderby: [],
          filter: tree(rule('numberofemployees', 'lots')),
          nestedExpand: [{ id: 'n', nav: 'bad nav', select: [], orderby: [] }],
        },
      ],
    });
    const titles = advisories.filter((a) => a.severity === 'error').map((a) => a.title);
    expect(titles).toContain('`numberofemployees` needs a whole number, not "lots".');
    expect(titles.some((t) => /Expanded navigation property "bad nav"/.test(String(t)))).toBe(true);
  });
});
