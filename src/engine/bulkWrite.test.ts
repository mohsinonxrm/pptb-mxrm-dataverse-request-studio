import { beforeAll, describe, expect, it } from 'vitest';
import { __registerLiveTables, type TableMeta } from '../mock/metadata';
import { defaultBypassOptions } from '../state/writeState';
import {
  buildBulkWrite,
  buildBulkWriteBody,
  bulkDisabledReason,
  coerceValue,
  deserializeBulkWrite,
  newBulkRow,
  parseBulkRows,
  parseDelimited,
  serializeBulkWrite,
  supportMessagesFor,
  supportQueryPath,
  type BulkOperation,
  type BulkRow,
  type BulkWriteState,
} from './bulkWrite';

const A1 = '00000000-0000-0000-0000-0000000000a1';
const A2 = '00000000-0000-0000-0000-0000000000a2';
const C1 = '00000000-0000-0000-0000-0000000000c1';

const account: TableMeta = {
  logicalName: 'account',
  entitySetName: 'accounts',
  displayName: 'Account',
  primaryKey: 'accountid',
  primaryName: 'name',
  columns: [
    { logicalName: 'accountid', displayName: 'Account', attributeType: 'Uniqueidentifier' },
    {
      logicalName: 'name',
      displayName: 'Account Name',
      attributeType: 'String',
      format: 'Text',
      maxLength: 160,
    },
    {
      logicalName: 'numberofemployees',
      displayName: 'Employees',
      attributeType: 'Integer',
      format: 'None',
    },
    {
      logicalName: 'donotemail',
      displayName: 'Do not allow Emails',
      attributeType: 'Boolean',
      trueOption: { value: 1, label: 'Do Not Allow' },
      falseOption: { value: 0, label: 'Allow' },
    },
    {
      logicalName: 'industrycode',
      displayName: 'Industry',
      attributeType: 'Picklist',
      options: [
        { value: 1, label: 'Accounting' },
        { value: 2, label: 'Agriculture' },
      ],
    },
    {
      logicalName: 'cr_tags',
      displayName: 'Tags',
      attributeType: 'MultiSelectPicklist',
      options: [
        { value: 10, label: 'Red' },
        { value: 20, label: 'Blue' },
      ],
    },
    {
      logicalName: 'primarycontactid',
      displayName: 'Primary Contact',
      attributeType: 'Lookup',
      targets: ['contact'],
    },
  ],
  navigationProperties: [
    {
      name: 'primarycontactid',
      cardinality: 'ManyToOne',
      targetEntity: 'contact',
      referencingAttribute: 'primarycontactid',
    } as never,
  ],
};

const contact: TableMeta = {
  logicalName: 'contact',
  entitySetName: 'contacts',
  displayName: 'Contact',
  primaryKey: 'contactid',
  primaryName: 'fullname',
  columns: [
    { logicalName: 'contactid', displayName: 'Contact', attributeType: 'Uniqueidentifier' },
  ],
  navigationProperties: [],
};

beforeAll(() => {
  __registerLiveTables([account, contact]);
});

const state = (operation: BulkOperation, rows: BulkRow[]): BulkWriteState => ({
  table: 'account',
  operation,
  rows,
  headers: [],
  bypass: defaultBypassOptions(),
  dirty: new Set(),
});

describe('buildBulkWrite', () => {
  it('targets the collection-bound action on the entity set', () => {
    const b = buildBulkWrite({ table: 'account', operation: 'CreateMultiple' });
    expect(b.relativeNoBase).toBe('/accounts/Microsoft.Dynamics.CRM.CreateMultiple');
    expect(b.relativeUrl.endsWith('/accounts/Microsoft.Dynamics.CRM.CreateMultiple')).toBe(true);
    expect(b.entityLogical).toBe('account');
  });
  it('is empty until the table is known', () => {
    expect(buildBulkWrite({ table: 'nope', operation: 'UpdateMultiple' }).relativeUrl).toBe('');
  });
});

describe('buildBulkWriteBody', () => {
  it('CreateMultiple: @odata.type on every target, lookups as @odata.bind', () => {
    const rows = [
      newBulkRow({ fieldValues: { name: 'Contoso', numberofemployees: 10 } }),
      newBulkRow({
        fieldValues: { name: 'Fabrikam', primarycontactid: { id: C1, targetEntity: 'contact' } },
      }),
    ];
    const { body, issues } = buildBulkWriteBody(state('CreateMultiple', rows));
    expect(issues).toEqual([]);
    expect(body.Targets).toEqual([
      { '@odata.type': 'Microsoft.Dynamics.CRM.account', name: 'Contoso', numberofemployees: 10 },
      {
        '@odata.type': 'Microsoft.Dynamics.CRM.account',
        name: 'Fabrikam',
        'primarycontactid@odata.bind': `/contacts(${C1})`,
      },
    ]);
  });

  it('CreateMultiple: a GUID key becomes the primary key; no key is fine', () => {
    const { body, issues } = buildBulkWriteBody(
      state('CreateMultiple', [
        newBulkRow({ key: `{${A1.toUpperCase()}}`, fieldValues: { name: 'x' } }),
      ]),
    );
    expect(issues).toEqual([]);
    expect(body.Targets[0].accountid).toBe(A1);
  });

  it('UpdateMultiple: requires a GUID key and rejects alternate keys', () => {
    const rows = [
      newBulkRow({ key: A1, fieldValues: { name: 'a' } }),
      newBulkRow({ key: '', fieldValues: { name: 'b' } }),
      newBulkRow({ key: "name='c'", fieldValues: { name: 'c' } }),
    ];
    const { body, issues } = buildBulkWriteBody(state('UpdateMultiple', rows));
    expect(body.Targets[0]).toEqual({
      '@odata.type': 'Microsoft.Dynamics.CRM.account',
      name: 'a',
      accountid: A1,
    });
    expect(issues.map((i) => i.rowIndex)).toEqual([1, 2]);
    expect(issues[1].message).toMatch(/alternate keys/);
  });

  it('UpsertMultiple: alternate key goes to @odata.id', () => {
    const { body, issues } = buildBulkWriteBody(
      state('UpsertMultiple', [
        newBulkRow({ key: "accountnumber='001'", fieldValues: { name: 'a' } }),
      ]),
    );
    expect(issues).toEqual([]);
    expect(body.Targets[0]['@odata.id']).toBe("accounts(accountnumber='001')");
    expect(body.Targets[0].accountid).toBeUndefined();
  });

  it('flags duplicate keys', () => {
    const rows = [
      newBulkRow({ key: A1, fieldValues: { name: 'a' } }),
      newBulkRow({ key: A1.toUpperCase(), fieldValues: { name: 'b' } }),
    ];
    const { issues } = buildBulkWriteBody(state('UpsertMultiple', rows));
    expect(issues).toHaveLength(1);
    expect(issues[0].message).toMatch(/Same key as row 1/);
  });

  it('DeleteMultiple: key only, no column values needed', () => {
    const { body, issues } = buildBulkWriteBody(
      state('DeleteMultiple', [newBulkRow({ key: A1 }), newBulkRow({ key: A2 })]),
    );
    expect(issues).toEqual([]);
    expect(body.Targets).toEqual([
      { '@odata.type': 'Microsoft.Dynamics.CRM.account', accountid: A1 },
      { '@odata.type': 'Microsoft.Dynamics.CRM.account', accountid: A2 },
    ]);
  });

  it('rows without values are an issue for create/update', () => {
    const { issues } = buildBulkWriteBody(state('CreateMultiple', [newBulkRow()]));
    expect(issues[0].message).toMatch(/No column values/);
  });

  it('explicit nulls are sent', () => {
    const { body } = buildBulkWriteBody(
      state('UpdateMultiple', [newBulkRow({ key: A1, nullFields: ['primarycontactid', 'name'] })]),
    );
    expect(body.Targets[0]).toMatchObject({ 'primarycontactid@odata.bind': null, name: null });
  });
});

describe('bulkDisabledReason', () => {
  it('blocks an unsupported table with the support message', () => {
    const s = state('CreateMultiple', [newBulkRow({ fieldValues: { name: 'a' } })]);
    expect(bulkDisabledReason(s, [], { state: 'unsupported', message: 'nope' })).toBe('nope');
    expect(bulkDisabledReason(s, [], { state: 'supported' })).toBeNull();
    expect(bulkDisabledReason(s, [], { state: 'unknown', message: 'x' })).toBeNull();
  });
  it('asks for rows', () => {
    expect(bulkDisabledReason(state('CreateMultiple', []), [], { state: 'supported' })).toMatch(
      /at least one row/,
    );
  });
});

describe('availability query', () => {
  it('UpsertMultiple is checked through Create- and UpdateMultiple', () => {
    expect(supportMessagesFor('UpsertMultiple')).toEqual(['CreateMultiple', 'UpdateMultiple']);
    expect(supportMessagesFor('DeleteMultiple')).toEqual(['DeleteMultiple']);
  });
  it('builds the documented sdkmessagefilters filter', () => {
    expect(supportQueryPath('CreateMultiple', 'account')).toBe(
      "sdkmessagefilters?$select=sdkmessagefilterid&$filter=sdkmessageid/name eq 'CreateMultiple' and primaryobjecttypecode eq 'account'",
    );
    expect(supportQueryPath("x'y", 'account')).toContain("'x''y'");
  });
});

describe('parseDelimited', () => {
  it('handles quotes, escaped quotes and embedded newlines', () => {
    expect(parseDelimited('a,b\n"x, y","he said ""hi""\nthere"\r\n1,2')).toEqual([
      ['a', 'b'],
      ['x, y', 'he said "hi"\nthere'],
      ['1', '2'],
    ]);
  });
  it('detects tab-separated (Excel) paste', () => {
    expect(parseDelimited('a\tb\n1\t2')).toEqual([
      ['a', 'b'],
      ['1', '2'],
    ]);
  });
});

describe('parseBulkRows', () => {
  it('maps CSV headers by logical and display name and coerces values', () => {
    const res = parseBulkRows(
      `accountid,Account Name,numberofemployees,donotemail,industrycode,cr_tags,primarycontactid,bogus\n` +
        `${A1},Contoso,"1,200",yes,Agriculture,"Red;20",${C1},zzz\n` +
        `,Fabrikam,,,,,,\n`,
      account,
    );
    expect(res.errors).toEqual([]);
    expect(res.skippedColumns).toEqual(['bogus']);
    expect(res.rows).toHaveLength(2);
    expect(res.rows[0].key).toBe(A1);
    expect(res.rows[0].fieldValues).toEqual({
      name: 'Contoso',
      numberofemployees: 1200,
      donotemail: true,
      industrycode: 2,
      cr_tags: [10, 20],
      primarycontactid: { id: C1, targetEntity: 'contact' },
    });
    expect(res.rows[1].fieldValues).toEqual({ name: 'Fabrikam' });
  });

  it('reads a JSON array (or Targets envelope) including @odata.bind and nulls', () => {
    const res = parseBulkRows(
      JSON.stringify({
        Targets: [
          {
            '@odata.type': 'Microsoft.Dynamics.CRM.account',
            '@odata.id': "accounts(accountnumber='7')",
            name: 'A',
            'primarycontactid@odata.bind': `/contacts(${C1})`,
            numberofemployees: null,
          },
        ],
      }),
      account,
    );
    expect(res.errors).toEqual([]);
    expect(res.rows[0].key).toBe("accountnumber='7'");
    expect(res.rows[0].fieldValues.primarycontactid).toEqual({ id: C1, targetEntity: 'contact' });
    expect(res.rows[0].nullFields).toEqual(['numberofemployees']);
  });

  it('reports bad cells and bad JSON', () => {
    const res = parseBulkRows('name,numberofemployees,industrycode\nA,abc,Mining', account);
    expect(res.errors).toEqual([
      'Row 1, numberofemployees: "abc" is not a number',
      'Row 1, industrycode: unknown choice "Mining"',
    ]);
    expect(parseBulkRows('[{', account).errors[0]).toMatch(/Invalid JSON/);
    expect(parseBulkRows('{"a":1}', account).errors[0]).toMatch(/array of objects/);
  });

  it('rejects a lookup to a table that is not a target', () => {
    const col = account.columns.find((c) => c.logicalName === 'primarycontactid')!;
    expect(coerceValue(col, `account:${A1}`)).toEqual({
      error: '"account" is not a target of primarycontactid',
    });
  });
});

describe('save round-trip', () => {
  it('drops dirty and restores rows', () => {
    const s = state('UpsertMultiple', [newBulkRow({ key: A1, fieldValues: { name: 'a' } })]);
    const snap = JSON.parse(JSON.stringify(serializeBulkWrite(s)));
    expect(snap.dirty).toBeUndefined();
    const back = deserializeBulkWrite(snap, {
      table: '',
      operation: 'CreateMultiple',
      rows: [],
      headers: [],
      bypass: defaultBypassOptions(),
    });
    expect(back.operation).toBe('UpsertMultiple');
    expect(back.rows[0].fieldValues).toEqual({ name: 'a' });
    expect(back.dirty.size).toBe(0);
  });
});
