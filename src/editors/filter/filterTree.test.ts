// #221: typed filter literals go unquoted only when they have the column's
// shape; anything else is quoted and escaped, and flagged so Execute blocks.
import { beforeAll, describe, expect, it } from 'vitest';
import { __registerLiveTables, findTable, type TableMeta } from '../../mock/metadata';
import {
  _setTableLookup,
  collectInvalidLiterals,
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
    {
      logicalName: 'parentaccountid',
      displayName: 'Parent',
      attributeType: 'Lookup',
      targets: ['account'],
    },
  ] as TableMeta['columns'],
  navigationProperties: [],
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
