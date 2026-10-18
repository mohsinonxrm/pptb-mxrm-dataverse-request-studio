import { beforeAll, describe, expect, it } from 'vitest';
import { __registerLiveTables, findTable, type TableMeta } from '../mock/metadata';
import { _setTableLookup, type FilterGroup } from '../editors/filter/filterTree';
import {
  buildBulkDelete,
  buildBulkDeleteBody,
  countPath,
  deserializeBulkDelete,
  filterToQueryExpression,
  parseRecipients,
  recurrencePattern,
  serializeBulkDelete,
  startDateTime,
  type BulkDeleteState,
} from './bulkDelete';

const U1 = '00000000-0000-0000-0000-0000000000b1';

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
      displayName: 'Name',
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
      logicalName: 'statecode',
      displayName: 'Status',
      attributeType: 'State',
      options: [
        { value: 0, label: 'Active' },
        { value: 1, label: 'Inactive' },
      ],
    },
    {
      logicalName: 'createdon',
      displayName: 'Created On',
      attributeType: 'DateTime',
      format: 'DateAndTime',
      dateTimeBehavior: 'UserLocal',
    } as never,
    {
      logicalName: 'parentaccountid',
      displayName: 'Parent',
      attributeType: 'Lookup',
      targets: ['account'],
    },
  ],
  navigationProperties: [],
};

beforeAll(() => {
  __registerLiveTables([account]);
  _setTableLookup(findTable);
});

const tree = (rules: FilterGroup['rules'], combinator: 'and' | 'or' = 'and'): FilterGroup => ({
  id: 'root',
  type: 'group',
  combinator,
  rules,
});

const state = (filter: FilterGroup, patch: Partial<BulkDeleteState> = {}): BulkDeleteState => ({
  table: 'account',
  filter,
  jobName: 'Cleanup',
  sendEmail: false,
  toRecipients: '',
  ccRecipients: '',
  startMode: 'now',
  startAt: '',
  recurrenceDays: null,
  headers: [],
  dirty: new Set(),
  ...patch,
});

describe('filterToQueryExpression', () => {
  it('maps comparisons, null checks, Like and nested groups with typed values', () => {
    const { query, problems } = filterToQueryExpression(
      account,
      tree([
        { id: 'r1', type: 'rule', col: 'statecode', op: 'eq', val: '1' },
        { id: 'r2', type: 'rule', col: 'name', op: 'contains', val: '50%_off' },
        { id: 'r3', type: 'rule', col: 'parentaccountid', op: 'is-null' },
        {
          id: 'g1',
          type: 'group',
          combinator: 'or',
          rules: [
            { id: 'r4', type: 'rule', col: 'numberofemployees', op: 'lt', val: '5' },
            { id: 'r5', type: 'rule', col: 'name', op: 'startswith', val: 'test', negated: true },
          ],
        },
      ]),
    );
    expect(problems).toEqual([]);
    expect(query).toEqual({
      EntityName: 'account',
      Criteria: {
        FilterOperator: 'And',
        Conditions: [
          {
            AttributeName: 'statecode',
            Operator: 'Equal',
            Values: [{ Type: 'System.Int32', Value: '1' }],
          },
          {
            AttributeName: 'name',
            Operator: 'Like',
            Values: [{ Type: 'System.String', Value: '%50[%][_]off%' }],
          },
          { AttributeName: 'parentaccountid', Operator: 'Null', Values: [] },
        ],
        Filters: [
          {
            FilterOperator: 'Or',
            Conditions: [
              {
                AttributeName: 'numberofemployees',
                Operator: 'LessThan',
                Values: [{ Type: 'System.Int32', Value: '5' }],
              },
              {
                AttributeName: 'name',
                Operator: 'NotLike',
                Values: [{ Type: 'System.String', Value: 'test%' }],
              },
            ],
            Filters: [],
          },
        ],
      },
    });
  });

  it('maps Dataverse query functions to same-named ConditionOperators', () => {
    const { query, problems } = filterToQueryExpression(
      account,
      tree([
        { id: 'f1', type: 'function', op: 'OlderThanXDays', col: 'createdon', val: '30' },
        { id: 'f2', type: 'function', op: 'In', col: 'statecode', values: ['0', '1'] },
        { id: 'f3', type: 'function', op: 'EqualUserId', col: 'parentaccountid' },
        { id: 'f4', type: 'function', op: 'Under', col: 'parentaccountid', val: U1 },
      ]),
    );
    expect(problems).toEqual([]);
    expect(query!.Criteria.Conditions).toEqual([
      {
        AttributeName: 'createdon',
        Operator: 'OlderThanXDays',
        Values: [{ Type: 'System.Int32', Value: '30' }],
      },
      {
        AttributeName: 'statecode',
        Operator: 'In',
        Values: [
          { Type: 'System.Int32', Value: '0' },
          { Type: 'System.Int32', Value: '1' },
        ],
      },
      { AttributeName: 'parentaccountid', Operator: 'EqualUserId', Values: [] },
      // Hierarchy conditions target the primary key.
      {
        AttributeName: 'accountid',
        Operator: 'Under',
        Values: [{ Type: 'System.Guid', Value: U1 }],
      },
    ]);
  });

  it('reports what QueryExpression cannot express', () => {
    const { query, problems } = filterToQueryExpression(account, {
      ...tree([
        { id: 'r1', type: 'rule', col: 'parentaccountid/name', op: 'eq', val: 'x' },
        { id: 'r2', type: 'rule', col: 'name', op: 'eq', val: 'accountnumber', valKind: 'column' },
        { id: 'f1', type: 'function', op: 'Today', col: 'createdon', negated: true },
        {
          id: 'l1',
          type: 'lambda',
          nav: 'contact_customer_accounts',
          lambda: 'any',
          alias: 'c',
          inner: tree([]),
        },
      ]),
      negated: true,
    });
    expect(query).toBeNull();
    expect(problems).toHaveLength(5);
  });
});

describe('buildBulkDeleteBody', () => {
  const now = new Date('2026-10-17T12:00:00.000Z');

  it('builds the documented BulkDelete body', () => {
    const { body, problems, deletesAll } = buildBulkDeleteBody(
      state(tree([{ id: 'r1', type: 'rule', col: 'statecode', op: 'eq', val: '1' }])),
      now,
    );
    expect(problems).toEqual([]);
    expect(deletesAll).toBe(false);
    expect(body).toMatchObject({
      JobName: 'Cleanup',
      SendEmailNotification: false,
      ToRecipients: [],
      CCRecipients: [],
      RecurrencePattern: '',
      StartDateTime: '2026-10-17T12:00:00.000Z',
    });
    expect((body!.QuerySet as unknown[]).length).toBe(1);
  });

  it('flags an empty query as deleting every row', () => {
    expect(buildBulkDeleteBody(state(tree([])), now).deletesAll).toBe(true);
  });

  it('recipients become activityparty entries and are required for email', () => {
    const r = buildBulkDeleteBody(
      state(tree([]), { sendEmail: true, toRecipients: `{${U1.toUpperCase()}}, nope` }),
      now,
    );
    expect(r.problems).toContain('Not a systemuser id: nope');
    const ok = buildBulkDeleteBody(state(tree([]), { sendEmail: true, toRecipients: U1 }), now);
    expect(ok.body!.ToRecipients).toEqual([
      { '@odata.type': 'Microsoft.Dynamics.CRM.activityparty', activitypartyid: U1 },
    ]);
    expect(buildBulkDeleteBody(state(tree([]), { sendEmail: true }), now).problems[0]).toMatch(
      /at least one To recipient/,
    );
  });

  it('requires a job name and a valid schedule', () => {
    const r = buildBulkDeleteBody(
      state(tree([]), { jobName: ' ', startMode: 'at', startAt: 'bad', recurrenceDays: 0 }),
      now,
    );
    expect(r.problems).toEqual([
      'Give the job a name.',
      'Pick a valid start date and time.',
      'Recurrence must be every 1 day or more.',
    ]);
  });
});

describe('helpers', () => {
  it('recurrence pattern', () => {
    expect(recurrencePattern(null)).toBe('');
    expect(recurrencePattern(7)).toBe('FREQ=DAILY;INTERVAL=7');
  });
  it('start time: now or the picked local time', () => {
    const now = new Date('2026-10-17T12:00:00.000Z');
    expect(startDateTime({ startMode: 'now', startAt: '' }, now)).toBe(now.toISOString());
    expect(startDateTime({ startMode: 'at', startAt: '2026-11-01T08:30' }, now)).toBe(
      new Date('2026-11-01T08:30').toISOString(),
    );
  });
  it('parseRecipients dedupes and splits invalid entries', () => {
    expect(parseRecipients(`${U1}\n${U1.toUpperCase()}; x`)).toEqual({ ids: [U1], invalid: ['x'] });
  });
  it('count query reuses the OData filter', () => {
    expect(
      countPath(state(tree([{ id: 'r1', type: 'rule', col: 'name', op: 'eq', val: 'A&B' }]))),
    ).toBe("accounts?$select=accountid&$count=true&$top=1&$filter=name eq 'A%26B'");
    expect(countPath(state(tree([])))).toBe('accounts?$select=accountid&$count=true&$top=1');
  });
  it('unbound action URL', () => {
    expect(buildBulkDelete({ table: 'account' }).relativeNoBase).toBe('/BulkDelete');
  });
  it('save round-trip', () => {
    const s = state(tree([]), { recurrenceDays: 30 });
    const back = deserializeBulkDelete(JSON.parse(JSON.stringify(serializeBulkDelete(s))), {
      ...s,
      recurrenceDays: null,
    });
    expect(back.recurrenceDays).toBe(30);
    expect(back.dirty.size).toBe(0);
  });
});
