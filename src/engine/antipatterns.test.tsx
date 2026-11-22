// Query anti-pattern advisories for Retrieve Multiple (#235): each detector's
// trigger and non-trigger, how filter columns are resolved through
// nav-paths and lambdas, and that none of them blocks Execute (they are
// warnings or infos, never errors).
import { renderToStaticMarkup } from 'react-dom/server';
import { beforeAll, describe, expect, it } from 'vitest';
import {
  buildWildcardAdvisory,
  detectColumnAntipatterns,
  detectComputedColumnFilters,
  detectLargeSelect,
  detectLargeTextFilters,
  detectLogicalFilters,
  detectLogicalSelect,
  detectRetrieveMultipleAntipatterns,
} from './antipatterns';
import { __registerLiveTables, type ColumnMeta, type TableMeta } from '../metadata/metadata';
import type { FilterGroup, FilterNode } from '../editors/filter/filterTree';
import { disabledReasonFromAdvisories, type Advisory } from '../primitives/advisories';

const col = (c: Partial<ColumnMeta> & { logicalName: string }): ColumnMeta =>
  ({ displayName: c.logicalName, attributeType: 'String', maxLength: 100, ...c }) as ColumnMeta;

const account: TableMeta = {
  logicalName: 'account',
  entitySetName: 'accounts',
  displayName: 'Account',
  primaryKey: 'accountid',
  primaryName: 'name',
  columns: [
    col({ logicalName: 'name', displayName: 'Account Name', maxLength: 160 }),
    col({ logicalName: 'description', displayName: 'Description', attributeType: 'Memo' }),
    col({ logicalName: 'notes_long', displayName: 'Long Notes', maxLength: 2000 }),
    col({ logicalName: 'notes_edge', displayName: 'Edge Notes', maxLength: 850 }),
    col({ logicalName: 'address1_city', displayName: 'City', isLogical: true }),
    col({ logicalName: 'address1_composite', displayName: 'Address', isLogical: true }),
    col({ logicalName: 'new_logical', displayName: 'Logical', isLogical: true }),
    col({ logicalName: 'new_calc', displayName: 'Calc', attributeType: 'Money', sourceType: 1 }),
    col({
      logicalName: 'new_rollup',
      displayName: 'Rollup',
      attributeType: 'Money',
      sourceType: 2,
    }),
    col({
      logicalName: 'new_simple',
      displayName: 'Simple',
      attributeType: 'Money',
      sourceType: 0,
    }),
  ],
  navigationProperties: [
    { name: 'primarycontactid', cardinality: 'ManyToOne', targetEntity: 'contact' },
    { name: 'contact_customer_accounts', cardinality: 'OneToMany', targetEntity: 'contact' },
    { name: 'lead_customer_accounts', cardinality: 'OneToMany', targetEntity: 'lead' },
  ] as TableMeta['navigationProperties'],
};

const contact: TableMeta = {
  logicalName: 'contact',
  entitySetName: 'contacts',
  displayName: 'Contact',
  primaryKey: 'contactid',
  primaryName: 'fullname',
  columns: [
    col({ logicalName: 'fullname' }),
    col({ logicalName: 'description', displayName: 'Description', attributeType: 'Memo' }),
    // Only on contact, so finding it proves the column was resolved there.
    col({ logicalName: 'new_bio', displayName: 'Bio', attributeType: 'Memo' }),
    col({ logicalName: 'new_formula', displayName: 'Formula', sourceType: 3 }),
  ],
  navigationProperties: [
    { name: 'parentcustomerid_account', cardinality: 'ManyToOne', targetEntity: 'account' },
  ] as TableMeta['navigationProperties'],
};

// `lead` is deliberately not registered: its metadata "hasn't loaded".
beforeAll(() => __registerLiveTables([account, contact]));

let n = 0;
const group = (...rules: FilterNode[]): FilterGroup => ({
  id: `g${++n}`,
  type: 'group',
  combinator: 'and',
  rules,
});
const rule = (c: string, op = 'eq'): FilterNode => ({
  id: `r${++n}`,
  type: 'rule',
  col: c,
  op,
  val: 'x',
});
const lambda = (nav: string, alias: string, ...inner: FilterNode[]): FilterNode => ({
  id: `l${++n}`,
  type: 'lambda',
  nav,
  lambda: 'any',
  alias,
  inner: group(...inner),
});

const summary = (out: Advisory[]) => out.map((a) => [a.id, a.severity, a.title]);
/** The advisory body as the user reads it. */
function bodyText(a: Advisory): string {
  const div = document.createElement('div');
  div.innerHTML = renderToStaticMarkup(<>{a.body}</>);
  return div.textContent ?? '';
}
const names = (count: number) => Array.from({ length: count }, (_, i) => `c${i}`);

describe('detectLargeSelect', () => {
  it('is quiet up to 25 columns', () => {
    expect(detectLargeSelect(names(25))).toEqual([]);
    expect(detectLargeSelect([])).toEqual([]);
  });

  it('warns past 25', () => {
    expect(summary(detectLargeSelect(names(26)))).toEqual([
      ['select-too-many', 'warning', '26 columns in $select — performance risk'],
    ]);
  });
});

describe('detectLogicalSelect', () => {
  it('is quiet with no logical columns, and ignores names it cannot find', () => {
    expect(detectLogicalSelect(['name', 'description', 'nope'], account)).toEqual([]);
  });

  it('warns on a logical column', () => {
    expect(summary(detectLogicalSelect(['name', 'new_logical'], account))).toEqual([
      ['select-logical', 'warning', '1 logical column in $select'],
    ]);
  });

  it('counts address columns in $select too, and lists each one', () => {
    const [a] = detectLogicalSelect(
      ['address1_city', 'address1_composite', 'new_logical'],
      account,
    );
    expect(a.title).toBe('3 logical columns in $select');
    expect(bodyText(a)).toContain('address1_city, address1_composite, new_logical');
  });
});

describe('detectLargeTextFilters', () => {
  it('warns on contains / startswith / endswith against a Memo column', () => {
    for (const op of ['contains', 'startswith', 'endswith']) {
      expect(summary(detectLargeTextFilters(group(rule('description', op)), account))).toEqual([
        ['filter-large-text', 'warning', 'contains/startswith/endswith on large-text column'],
      ]);
    }
  });

  it('is quiet for an exact match on a Memo column', () => {
    expect(detectLargeTextFilters(group(rule('description', 'eq')), account)).toEqual([]);
  });

  it('counts a String as large only above 850 characters', () => {
    expect(detectLargeTextFilters(group(rule('notes_long', 'contains')), account)).toHaveLength(1);
    expect(detectLargeTextFilters(group(rule('notes_edge', 'contains')), account)).toEqual([]);
    expect(detectLargeTextFilters(group(rule('name', 'contains')), account)).toEqual([]);
  });

  it.each([
    ['a nav-path', rule('primarycontactid/new_bio', 'contains'), 'new_bio'],
    [
      'a lambda alias',
      lambda('contact_customer_accounts', 'c', group(rule('c/new_bio', 'endswith'))),
      'new_bio',
    ],
    ['a nested group', group(rule('description', 'startswith')), 'description'],
  ])('resolves a column through %s', (_how, node, column) => {
    // One at a time, on a column only the target table has, so one path
    // can't stand in for another.
    const [a] = detectLargeTextFilters(group(node), account);
    expect(a?.title).toMatch(/on large-text column$/);
    expect(bodyText(a)).toMatch(new RegExp(`^${column} can't be`));
  });

  it("doesn't find a contact-only column on the account itself", () => {
    expect(detectLargeTextFilters(group(rule('new_bio', 'contains')), account)).toEqual([]);
  });

  it('lists each distinct column once, with a plural title', () => {
    const [a] = detectLargeTextFilters(
      group(
        rule('description', 'contains'),
        rule('description', 'startswith'),
        rule('notes_long', 'contains'),
      ),
      account,
    );
    expect(a.title).toBe('contains/startswith/endswith on large-text columns');
    expect(bodyText(a)).toMatch(/^description, notes_long can't be/);
  });

  it("skips columns it can't resolve", () => {
    expect(
      detectLargeTextFilters(
        group(
          rule('nowhere/description', 'contains'),
          rule('missing', 'contains'),
          lambda('lead_customer_accounts', 'l', rule('l/description', 'contains')),
          { id: 'blank', type: 'rule', col: '', op: 'contains' },
        ),
        account,
      ),
    ).toEqual([]);
  });
});

describe('detectLogicalFilters', () => {
  it('warns on a logical column', () => {
    expect(summary(detectLogicalFilters(group(rule('new_logical')), account))).toEqual([
      ['filter-logical', 'warning', 'Filter on logical column'],
    ]);
  });

  it('exempts the address columns', () => {
    expect(
      detectLogicalFilters(group(rule('address1_city'), rule('address1_composite')), account),
    ).toEqual([]);
  });

  it('is quiet for physical columns', () => {
    expect(detectLogicalFilters(group(rule('name')), account)).toEqual([]);
  });
});

describe('detectComputedColumnFilters', () => {
  it('names the kind of a single computed column', () => {
    expect(summary(detectComputedColumnFilters(group(rule('new_calc')), account))).toEqual([
      ['filter-computed', 'warning', 'Filter on calculated column'],
    ]);
    expect(detectComputedColumnFilters(group(rule('new_rollup')), account)[0].title).toBe(
      'Filter on rollup column',
    );
  });

  it('says "computed columns" for more than one, including through a nav-path', () => {
    const [a] = detectComputedColumnFilters(
      group(rule('new_calc'), rule('primarycontactid/new_formula'), rule('new_calc', 'ne')),
      account,
    );
    expect(a.title).toBe('Filter on computed columns');
    expect(bodyText(a)).toMatch(/^new_calc \(calculated\), new_formula \(formula\) — /);
  });

  it('is quiet for sourceType 0 and for plain columns', () => {
    expect(detectComputedColumnFilters(group(rule('new_simple'), rule('name')), account)).toEqual(
      [],
    );
  });
});

describe('buildWildcardAdvisory', () => {
  const entry = (c: string) => ({ col: c, raw: '%abc', cleaned: 'abc', kinds: ['%'] });

  it('is quiet when nothing was stripped', () => {
    expect(buildWildcardAdvisory([])).toEqual([]);
  });

  it('is an info listing each rewrite', () => {
    expect(summary(buildWildcardAdvisory([entry('name')]))).toEqual([
      ['wildcard-stripped', 'info', 'Stripped leading wildcard from 1 filter'],
    ]);
    const [a] = buildWildcardAdvisory([entry('name'), entry('fullname')]);
    expect(a.title).toBe('Stripped leading wildcards from 2 filters');
    expect(bodyText(a)).toContain('name: %abc → abc');
    expect(bodyText(a)).toContain('fullname: %abc → abc');
  });
});

describe('detectColumnAntipatterns', () => {
  const find = (c: string, t = account) => t.columns.find((x) => x.logicalName === c)!;

  it('flags a computed column', () => {
    expect(detectColumnAntipatterns(find('new_calc'))).toMatchObject([
      { kind: 'calculated', title: 'Filtering on a calculated column' },
    ]);
  });

  it('flags a logical column, except an address column', () => {
    expect(detectColumnAntipatterns(find('new_logical'))).toMatchObject([{ kind: 'logical' }]);
    expect(detectColumnAntipatterns(find('address1_city'))).toEqual([]);
  });

  it('flags large text only for a scan operator', () => {
    const memo = find('description');
    expect(detectColumnAntipatterns(memo)).toEqual([]);
    expect(detectColumnAntipatterns(memo, 'eq')).toEqual([]);
    const [a] = detectColumnAntipatterns(memo, 'contains');
    expect(a).toMatchObject({ kind: 'largeText', title: 'contains() on a large-text column' });
    expect(a.body).toContain('a Memo (multi-line text)');
    expect(detectColumnAntipatterns(find('notes_long'), 'startswith')[0].body).toContain(
      'a long String (MaxLength > 850)',
    );
  });

  it('reports every kind that applies', () => {
    const c = col({
      logicalName: 'new_everything',
      attributeType: 'Memo',
      isLogical: true,
      sourceType: 2,
    });
    expect(detectColumnAntipatterns(c, 'endswith').map((a) => a.kind)).toEqual([
      'calculated',
      'logical',
      'largeText',
    ]);
  });

  it('is quiet for a plain column', () => {
    expect(detectColumnAntipatterns(find('name'), 'contains')).toEqual([]);
  });
});

describe('detectRetrieveMultipleAntipatterns', () => {
  it('returns nothing before the table metadata has loaded', () => {
    expect(
      detectRetrieveMultipleAntipatterns({
        table: 'lead',
        select: names(30),
        filter: group(rule('description', 'contains')),
      }),
    ).toEqual([]);
  });

  it('returns nothing for a clean query', () => {
    expect(
      detectRetrieveMultipleAntipatterns({
        table: 'account',
        select: ['name'],
        filter: group(rule('name')),
      }),
    ).toEqual([]);
  });

  it('combines every detector, and none of them blocks Execute', () => {
    const out = detectRetrieveMultipleAntipatterns({
      table: 'account',
      select: [...names(25), 'new_logical'],
      filter: group(rule('description', 'contains'), rule('new_logical'), rule('new_calc')),
      strippedWildcards: [{ col: 'description', raw: '%a', cleaned: 'a', kinds: ['%'] }],
    });
    expect(out.map((a) => [a.id, a.severity])).toEqual([
      ['select-too-many', 'warning'],
      ['select-logical', 'warning'],
      ['filter-large-text', 'warning'],
      ['filter-logical', 'warning'],
      ['filter-computed', 'warning'],
      ['wildcard-stripped', 'info'],
    ]);
    expect(disabledReasonFromAdvisories(out)).toBeNull();
  });
});
