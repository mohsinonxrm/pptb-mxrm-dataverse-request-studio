// urlCaretContext / urlSuggestions: completion in hand-typed URLs (#268).
import { beforeAll, describe, expect, it } from 'vitest';
import {
  applySuggestion,
  urlCaretContext,
  urlSuggestions,
  type CaretContext,
  type SuggestionDeps,
  type UrlMode,
} from './urlCompletion';
import { registerFixtureTables } from '../test/fixtures';
import { findTable } from '../metadata/metadata';
import type { CsdlAction } from '../metadata/actionsCsdl';

beforeAll(registerFixtureTables);

/** `marked` without its one `|` caret marker, and where the marker was. */
function unmark(marked: string): { text: string; caret: number } {
  const caret = marked.indexOf('|');
  return { text: marked.slice(0, caret) + marked.slice(caret + 1), caret };
}

/** The context at `|` in `marked`. */
function at(marked: string, mode: UrlMode = 'url', set?: string): CaretContext | null {
  const { text, caret } = unmark(marked);
  return urlCaretContext(text, caret, mode, set);
}

const op = (
  name: string,
  kind: 'Action' | 'Function',
  binding: CsdlAction['binding'],
  params: string[] = [],
): CsdlAction =>
  ({
    name,
    kind,
    namespace: 'Microsoft.Dynamics.CRM',
    binding,
    parameters: params.map((p) => ({ name: p, type: 'Edm.String', required: true })),
    returnType: { kind: 'void', typeName: '' },
    source: 'oob',
  }) as CsdlAction;

const deps: SuggestionDeps = {
  entities: [
    { logicalName: 'account', entitySetName: 'accounts', displayName: 'Account' },
    { logicalName: 'contact', entitySetName: 'contacts', displayName: 'Contact' },
  ],
  tableOf: findTable,
  operations: [
    op('WhoAmI', 'Function', { kind: 'unbound' }),
    op('RetrieveVersion', 'Function', { kind: 'unbound' }),
    op('WinOpportunity', 'Action', { kind: 'unbound' }, ['Status']),
    op('CalculateRollupField', 'Function', { kind: 'unbound' }, ['Target', 'FieldName']),
    op('Merge', 'Action', { kind: 'entity', entityType: 'account' }, ['entity']),
    op('RetrieveCount', 'Function', { kind: 'collection', entityType: 'account' }),
    {
      ...op('ContainValues', 'Function', { kind: 'entity', entityType: 'account' }),
      isQueryFunction: true,
    },
  ],
};
const labels = (marked: string, mode: UrlMode = 'url', set?: string) => {
  const ctx = at(marked, mode, set);
  return ctx ? urlSuggestions(ctx, deps).items.map((s) => s.label) : null;
};

describe('urlCaretContext', () => {
  it('reads the first path segment as a table, after any host and API base', () => {
    expect(at('acc|')).toMatchObject({ want: 'tables', from: 0, to: 3, prefix: 'acc' });
    expect(at('https://org.crm.dynamics.com/api/data/v9.2/acc|')).toMatchObject({
      want: 'tables',
      prefix: 'acc',
      from: 'https://org.crm.dynamics.com/api/data/v9.2/'.length,
    });
    expect(at('/api/data/v9.2/|')).toMatchObject({ want: 'tables', prefix: '' });
  });

  it('replaces the rest of the word after the caret', () => {
    expect(at('acc|ounts?$top=1')).toMatchObject({ from: 0, to: 8 });
  });

  it('offers nothing in a key, a Content-ID or a string', () => {
    expect(at('accounts(00|')).toBeNull();
    expect(at('$1/|')).toBeNull();
    expect(at("accounts?$filter=name eq 'ab|")).toBeNull();
  });

  it('reads later segments against the table the path reaches', () => {
    expect(at('accounts(1)/prim|')).toMatchObject({
      want: 'segment',
      prefix: 'prim',
      keyed: true,
      table: { entitySet: 'accounts', navs: [] },
    });
    expect(at('accounts(1)/primarycontactid/|')).toMatchObject({
      want: 'segment',
      keyed: false,
      table: { entitySet: 'accounts', navs: ['primarycontactid'] },
    });
  });

  it('reads query option names and values', () => {
    expect(at('accounts?$se|')).toMatchObject({ want: 'options', prefix: '$se' });
    expect(at('accounts?$select=name,acc|')).toMatchObject({
      want: 'columns',
      prefix: 'acc',
      from: 'accounts?$select=name,'.length,
    });
    expect(at('accounts?$top=5&$orderby=name |')).toMatchObject({ want: 'orderDir', prefix: '' });
    expect(at('accounts?$count=t|')).toMatchObject({ want: 'booleans', prefix: 't' });
    expect(at('accounts?$top=|')).toBeNull();
  });

  it('follows $expand into the nav’s own options, and past an earlier &', () => {
    expect(at('accounts?$top=1&$expand=primarycontactid($select=full|')).toMatchObject({
      want: 'columns',
      prefix: 'full',
      table: { entitySet: 'accounts', navs: ['primarycontactid'] },
    });
    expect(at('accounts?$expand=primarycontactid($select=fullname;$to|')).toMatchObject({
      want: 'nestedOptions',
      prefix: '$to',
    });
    expect(
      at('accounts?$expand=primarycontactid($expand=parentcustomerid_account($select=na|'),
    ).toMatchObject({
      want: 'columns',
      table: { entitySet: 'accounts', navs: ['primarycontactid', 'parentcustomerid_account'] },
    });
    expect(at('accounts?$expand=contact_customer_accounts,prim|')).toMatchObject({
      want: 'navs',
      prefix: 'prim',
    });
  });

  it('tells operands, operators and logical operators apart in $filter', () => {
    expect(at('accounts?$filter=na|')).toMatchObject({ want: 'operand', prefix: 'na' });
    expect(at('accounts?$filter=(name eq null) and (rev|')).toMatchObject({
      want: 'operand',
      prefix: 'rev',
    });
    expect(at('accounts?$filter=contains(na|')).toMatchObject({ want: 'operand' });
    expect(at('accounts?$filter=name |')).toMatchObject({ want: 'operators' });
    expect(at("accounts?$filter=name eq 'x' |")).toMatchObject({ want: 'logical' });
    expect(at('accounts?$filter=revenue gt 5 o|')).toMatchObject({ want: 'logical', prefix: 'o' });
    expect(at('accounts?$filter=name eq |')).toBeNull();
    expect(at('accounts?$filter=primarycontactid/full|')).toMatchObject({
      want: 'operand',
      prefix: 'full',
      from: 'accounts?$filter=primarycontactid/'.length,
      table: { entitySet: 'accounts', navs: ['primarycontactid'] },
    });
  });

  it('reads query mode against the given table', () => {
    expect(at('$select=na|', 'query', 'accounts')).toMatchObject({
      want: 'columns',
      table: { entitySet: 'accounts', navs: [] },
    });
    expect(at('$top=5&$|', 'query', 'accounts')).toMatchObject({ want: 'options', prefix: '$' });
    expect(at('$select=na|', 'query', '')).toBeNull();
  });
});

describe('urlSuggestions', () => {
  it('lists tables and unbound operations, not query functions', () => {
    expect(labels('|')).toEqual([
      'accounts',
      'contacts',
      'WhoAmI',
      'RetrieveVersion',
      'WinOpportunity',
      'CalculateRollupField',
    ]);
    const who = urlSuggestions(at('Who|')!, deps).items[0];
    expect(who).toMatchObject({ label: 'WhoAmI', insert: 'WhoAmI()', kind: 'operation' });
    const calc = urlSuggestions(at('Calc|')!, deps).items[0];
    expect(calc.insert).toBe('CalculateRollupField(');
  });

  it('matches by display name too, starts-with first', () => {
    expect(labels('accounts?$select=acc|')).toEqual(['accountid', 'accountnumber']);
    expect(labels('accounts?$select=emp|')).toEqual(['numberofemployees']);
  });

  it('offers a row’s navs and entity-bound operations, a collection’s count and collection operations', () => {
    expect(labels('accounts(1)/|')).toEqual([
      'contact_customer_accounts',
      'primarycontactid',
      'Microsoft.Dynamics.CRM.Merge',
    ]);
    expect(labels('accounts/|')).toEqual(['$count', 'Microsoft.Dynamics.CRM.RetrieveCount']);
    expect(labels('accounts(1)/primarycontactid/|')).toEqual([
      'parentcustomerid_account',
      'parentcustomerid_contact',
    ]);
  });

  it('inserts lookups as their _value property in $select', () => {
    const items = urlSuggestions(at('contacts?$select=parent|')!, deps).items;
    expect(items.map((s) => s.insert)).toEqual(['_parentcustomerid_value']);
  });

  it('sorts and matches a lookup’s _value property by its column name', () => {
    expect(labels('accounts?$select=|')!.slice(0, 2)).toEqual(['accountid', 'accountnumber']);
    expect(labels('accounts?$select=|')!.indexOf('_primarycontactid_value')).toBeGreaterThan(
      labels('accounts?$select=|')!.indexOf('numberofemployees'),
    );
    expect(labels('accounts?$select=pri|')).toEqual(['_primarycontactid_value']);
  });

  it('offers $filter operands: columns, single-valued navs as paths, functions', () => {
    const items = urlSuggestions(at('accounts?$filter=|')!, deps).items;
    expect(items.find((s) => s.label === 'primarycontactid')?.insert).toBe('primarycontactid/');
    expect(items.some((s) => s.label === 'contact_customer_accounts')).toBe(false);
    expect(items.find((s) => s.label === 'contains')?.insert).toBe('contains(');
    expect(labels('accounts?$filter=primarycontactid/full|')).toEqual(['fullname']);
    expect(labels('accounts?$filter=name |')).toEqual(['eq', 'ne', 'gt', 'ge', 'lt', 'le']);
  });

  it('names the table still to load, and lists nothing for it meanwhile', () => {
    const r = urlSuggestions(at('accounts?$select=|')!, { ...deps, tableOf: () => undefined });
    expect(r).toMatchObject({ items: [], needTable: 'account' });
    const nav = urlSuggestions(at('accounts?$expand=primarycontactid($select=|')!, {
      ...deps,
      tableOf: (t) => (t === 'account' ? findTable(t) : undefined),
    });
    expect(nav.needTable).toBe('contact');
  });

  it('honours the scope filters', () => {
    const r = urlSuggestions(at('accounts?$select=|')!, {
      ...deps,
      columnInScope: (c) => c.logicalName !== 'revenue',
    });
    expect(r.items.map((s) => s.label)).not.toContain('revenue');
  });

  it('drops a list that only repeats what’s typed', () => {
    expect(labels('accounts?$orderby=name desc|')).toEqual([]);
  });

  it('caps the list and counts the rest', () => {
    const many = Array.from({ length: 60 }, (_, i) => ({
      logicalName: `t${i}`,
      entitySetName: `t${i}s`,
      displayName: `T${i}`,
    }));
    const r = urlSuggestions(at('t|')!, { ...deps, entities: many, operations: [] });
    expect(r.items).toHaveLength(50);
    expect(r.hiddenCount).toBe(10);
  });
});

describe('review fixes', () => {
  it('offers nothing while the host or the API base is typed', () => {
    expect(at('https://org.cr|')).toBeNull();
    expect(at('https://org.crm.dynamics.com/api/da|')).toBeNull();
    expect(at('https://org.crm.dynamics.com/acc|')).toBeNull();
    expect(at('http|')).toBeNull();
    expect(at('api/data/v9.2|')).toBeNull();
    expect(at('/api/data/v9.2/acc|')).toMatchObject({ want: 'tables', prefix: 'acc' });
    expect(at('https://org.crm.dynamics.com:443/api/data/v9.2/acc|')).toMatchObject({
      want: 'tables',
    });
    // A table can still start with "ap"
    expect(at('ap|')).toMatchObject({ want: 'tables', prefix: 'ap' });
  });

  it('inserts mid-word without repeating what follows the word', () => {
    const apply = (marked: string, insert: string) => {
      const ctx = at(marked)!;
      return applySuggestion(unmark(marked).text, ctx, {
        label: insert,
        insert,
        kind: 'keyword',
      });
    };
    expect(apply('accounts?$sele|ct=name', '$select=')).toEqual({
      value: 'accounts?$select=name',
      caret: 'accounts?$select='.length,
    });
    expect(
      apply('accounts?$filter=primary|contactid/fullname eq null', 'primarycontactid/'),
    ).toEqual({
      value: 'accounts?$filter=primarycontactid/fullname eq null',
      caret: 'accounts?$filter=primarycontactid/'.length,
    });
    expect(apply('Who|AmI()', 'WhoAmI()')).toEqual({ value: 'WhoAmI()', caret: 8 });
    expect(apply('acc|', 'accounts')).toEqual({ value: 'accounts', caret: 8 });
  });

  it('reads an unquoted GUID as a value', () => {
    expect(
      at('accounts?$filter=accountid eq 1a2b3c4d-1a2b-1a2b-1a2b-a1b2c3d4e5f6 |'),
    ).toMatchObject({ want: 'logical' });
  });

  it('offers no columns through a collection nav, which needs any / all', () => {
    expect(labels('accounts?$filter=contact_customer_accounts/|')).toEqual([]);
    // In a collection nav's own $expand options, its columns are fine
    expect(labels('accounts?$expand=contact_customer_accounts($filter=full|')).toEqual([
      'fullname',
    ]);
  });

  it('reads a function’s second argument as a value, and a finished sort item as done', () => {
    expect(at('accounts?$filter=contains(name, |')).toBeNull();
    expect(at('accounts?$orderby=name desc |')).toBeNull();
    expect(at('accounts?$orderby=name d|')).toMatchObject({ want: 'orderDir', prefix: 'd' });
  });

  it('offers operations bound to any table (crmbaseentity)', () => {
    const r = urlSuggestions(at('contacts(1)/|')!, {
      ...deps,
      operations: [
        ...deps.operations,
        op('ExportPdfDocument', 'Action', { kind: 'entity', entityType: 'crmbaseentity' }),
      ],
    });
    expect(r.items.map((s) => s.label)).toContain('Microsoft.Dynamics.CRM.ExportPdfDocument');
  });
});

describe('ranking', () => {
  it('puts the shortest starts-with match first once something is typed', () => {
    const many = ['contactinvoicescollection', 'contactleadscollection', 'contacts'].map((s) => ({
      logicalName: s,
      entitySetName: s,
      displayName: s,
    }));
    const r = urlSuggestions(at('cont|')!, { ...deps, entities: many, operations: [] });
    expect(r.items[0].label).toBe('contacts');
    const none = urlSuggestions(at('|')!, { ...deps, entities: many, operations: [] });
    expect(none.items[0].label).toBe('contactinvoicescollection');
  });
});

describe('query functions', () => {
  it('leaves out $filter query functions that the CSDL declares unbound', () => {
    const last7 = {
      ...op('Last7Days', 'Function', { kind: 'unbound' }, ['PropertyName']),
      returnType: { kind: 'primitive', typeName: 'Edm.Boolean' },
    } as CsdlAction;
    const r = urlSuggestions(at('La|')!, { ...deps, operations: [...deps.operations, last7] });
    expect(r.items.map((s) => s.label)).not.toContain('Last7Days');
  });
});
