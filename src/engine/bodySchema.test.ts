// buildBodySchema: the JSON view's completions, hover and type checks (#268).
import { beforeAll, describe, expect, it } from 'vitest';
import { buildBodySchema, type BodyJsonSchema } from './bodySchema';
import { decodeFieldBody } from './fieldBody';
import { account, contact, registerFixtureTables } from '../test/fixtures';
import type { ColumnMeta, TableMeta } from '../metadata/metadata';

beforeAll(registerFixtureTables);

const sets: Record<string, string> = { account: 'accounts', contact: 'contacts' };
const entitySetOf = (t: string) => sets[t];
const writable = (t: TableMeta) =>
  t.columns.filter((c) => c.attributeType !== 'Uniqueidentifier' && c.logicalName !== 'createdon');
const props = (s: BodyJsonSchema) => s.properties ?? {};

describe('buildBodySchema', () => {
  it('offers each writable column, and lookups only as @odata.bind keys', () => {
    const s = buildBodySchema(account, writable(account), { nullSupported: true, entitySetOf });
    expect(Object.keys(props(s)).sort()).toEqual(
      [
        'name',
        'accountnumber',
        'numberofemployees',
        'revenue',
        'donotemail',
        'industrycode',
        'sample_tags',
        'primarycontactid@odata.bind',
      ].sort(),
    );
    expect(props(s).name.description).toMatch(/^Name \(String\)\./);
    expect(props(s).name.type).toEqual(['string', 'null']);
    expect(props(s).revenue.type).toEqual(['number', 'null']);
  });

  it('completes a lookup to its target entity set', () => {
    const s = buildBodySchema(account, writable(account), { nullSupported: true, entitySetOf });
    const bind = props(s)['primarycontactid@odata.bind'];
    expect(bind.type).toEqual(['string', 'null']);
    expect(bind.defaultSnippets).toEqual([
      {
        label: '/contacts(<guid>)',
        description: 'A contact row',
        body: '/contacts(${1:00000000-0000-0000-0000-000000000000})',
      },
    ]);
    expect(bind.description).toContain('"/contacts(<guid>)", or null to clear it');
  });

  it('leaves the entity set as a placeholder when it isn’t known', () => {
    const s = buildBodySchema(account, writable(account), { nullSupported: false });
    const bind = props(s)['primarycontactid@odata.bind'];
    expect(bind.type).toBe('string');
    expect(bind.defaultSnippets?.[0].body).toBe(
      '/${1:contacts}(${2:00000000-0000-0000-0000-000000000000})',
    );
    expect(bind.description).toContain('"/<contact entity set>(<guid>)".');
  });

  it('gives a polymorphic lookup one key per target, named after its nav property', () => {
    const s = buildBodySchema(contact, writable(contact), { nullSupported: true, entitySetOf });
    expect(props(s)['parentcustomerid_account@odata.bind'].defaultSnippets?.[0].label).toBe(
      '/accounts(<guid>)',
    );
    expect(props(s)['parentcustomerid_contact@odata.bind'].defaultSnippets?.[0].label).toBe(
      '/contacts(<guid>)',
    );
    expect(props(s)['parentcustomerid@odata.bind']).toBeUndefined();
  });

  it('offers every target of an Owner column, whose nav targets "owner"', () => {
    const owner = {
      logicalName: 'ownerid',
      displayName: 'Owner',
      attributeType: 'Owner',
      targets: ['systemuser', 'team'],
    } as ColumnMeta;
    const tbl: TableMeta = {
      ...account,
      columns: [owner],
      navigationProperties: [
        {
          name: 'ownerid',
          relationshipName: 'owner_accounts',
          cardinality: 'ManyToOne',
          targetEntity: 'owner',
          referencingAttribute: 'ownerid',
        },
      ],
    };
    const s = buildBodySchema(tbl, [owner], {
      nullSupported: true,
      entitySetOf: (t) => ({ systemuser: 'systemusers', team: 'teams' })[t],
    });
    expect(props(s)['ownerid@odata.bind'].defaultSnippets?.map((d) => d.label)).toEqual([
      '/systemusers(<guid>)',
      '/teams(<guid>)',
    ]);
  });

  it('falls back to <column>@odata.bind when no nav property references the column', () => {
    const tbl: TableMeta = { ...account, navigationProperties: [] };
    const s = buildBodySchema(tbl, writable(tbl), { nullSupported: true, entitySetOf });
    expect(props(s)['primarycontactid@odata.bind']).toBeDefined();
  });

  it('completes choices to their values, labelled in the list', () => {
    const s = buildBodySchema(account, writable(account), { nullSupported: true, entitySetOf });
    expect(props(s).industrycode).toMatchObject({
      type: ['integer', 'null'],
      defaultSnippets: [
        { label: '1 · Accounting', description: 'Accounting', body: 1 },
        { label: '2 · Agriculture', description: 'Agriculture', body: 2 },
      ],
    });
    expect(props(s).industrycode.description).toContain('Values: 1 = Accounting, 2 = Agriculture.');
    expect(props(s).sample_tags.defaultSnippets).toBeUndefined();
    expect(props(s).sample_tags.description).toContain('Comma-separated values: 100 = A, 200 = B.');
  });

  it('names a boolean’s labels, and leaves true / false to its type', () => {
    const s = buildBodySchema(account, writable(account), { nullSupported: true, entitySetOf });
    expect(props(s).donotemail.type).toEqual(['boolean', 'null']);
    expect(props(s).donotemail.defaultSnippets).toBeUndefined();
    expect(props(s).donotemail.description).toContain('true = Do Not Allow, false = Allow.');
  });

  it('caps a long option list in the description', () => {
    const many = {
      logicalName: 'big',
      displayName: 'Big',
      attributeType: 'Picklist',
      options: Array.from({ length: 30 }, (_, i) => ({ value: i, label: `O${i}` })),
    } as ColumnMeta;
    const s = buildBodySchema(account, [many], { nullSupported: false });
    expect(props(s).big.description).toMatch(/24 = O24, and 5 more\.$/);
    expect(props(s).big.defaultSnippets).toHaveLength(30);
  });

  it('leaves null out where the field set has no explicit-null support', () => {
    const s = buildBodySchema(account, writable(account), { nullSupported: false, entitySetOf });
    expect(props(s).name.type).toBe('string');
    expect(props(s).industrycode.type).toBe('integer');
  });

  it('keeps out-of-scope columns for hover and checks, but doesn’t suggest them', () => {
    const s = buildBodySchema(account, writable(account), {
      nullSupported: true,
      entitySetOf,
      suggest: (c) => c.logicalName !== 'revenue' && c.logicalName !== 'primarycontactid',
    });
    expect(props(s).revenue.doNotSuggest).toBe(true);
    expect(props(s)['primarycontactid@odata.bind'].doNotSuggest).toBe(true);
    expect(props(s).name.doNotSuggest).toBeUndefined();
  });
});

// The schema's types, checked the way Monaco's JSON service does
function accepts(s: BodyJsonSchema, v: unknown): boolean {
  if (!s.type) return true;
  const types = Array.isArray(s.type) ? s.type : [s.type];
  return types.some((t) =>
    t === 'null'
      ? v === null
      : t === 'integer'
        ? Number.isInteger(v)
        : t === 'array'
          ? Array.isArray(v)
          : typeof v === t,
  );
}

describe('buildBodySchema agrees with decodeFieldBody', () => {
  const samples: unknown[] = ['text', '', 3, 2.5, 0, true, false, null, [100], { a: 1 }];
  const cols = writable(account).filter((c) => c.attributeType !== 'Lookup');
  const schema = buildBodySchema(account, cols, { nullSupported: true, entitySetOf });

  for (const c of cols) {
    it(`${c.logicalName} (${c.attributeType})`, () => {
      for (const v of samples) {
        const decoded = decodeFieldBody(account, { [c.logicalName]: v });
        const maps = decoded.unmapped.length === 0;
        // A multi-select array decodes, but Dataverse wants "100,200" (see bodySchema)
        if (c.attributeType === 'MultiSelectPicklist' && Array.isArray(v)) {
          expect(accepts(props(schema)[c.logicalName], v)).toBe(false);
          continue;
        }
        // A multi-select string must be numbers; the schema only checks it's a string
        if (c.attributeType === 'MultiSelectPicklist' && typeof v === 'string') continue;
        expect([c.logicalName, v, accepts(props(schema)[c.logicalName], v)]).toEqual([
          c.logicalName,
          v,
          maps,
        ]);
      }
    });
  }
});
