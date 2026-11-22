import { beforeEach, describe, expect, it, vi } from 'vitest';
import {
  buildAssociate,
  buildAssociateBody,
  buildAssociateRequests,
  buildCreate,
  buildCreateBody,
  buildDelete,
  buildDisassociate,
  buildDisassociateRequests,
  buildMergeBody,
  buildPredefinedQuery,
  buildRetrieveMultiple,
  buildRetrieveNextLink,
  buildRetrieveSingle,
  buildUpdate,
  buildUpdateBody,
  buildUpsert,
} from './urlBuilder';
import { makeExpandSpec } from '../editors/ExpandEditor';
import { emptyPrefer } from '../editors/PreferEditor';
import { defaultBypassOptions, type MergeState } from '../state/writeState';
import {
  ACCOUNT_ID,
  CONTACT_ID,
  createState,
  group,
  registerFixtureTables,
  retrieveMultipleState,
  rule,
} from '../test/fixtures';

vi.mock('../env/environment', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../env/environment')>()),
  ENV: {
    name: 'Test',
    host: 'contoso.crm.dynamics.com',
    apiBase: '/api/data/v9.2',
    user: '',
  },
}));

const API = '/api/data/v9.2';

beforeEach(() => {
  registerFixtureTables();
});

describe('buildRetrieveMultiple', () => {
  it('returns an empty request for an unknown table', () => {
    const built = buildRetrieveMultiple(retrieveMultipleState({ table: 'nope' }));
    expect(built.relativeUrl).toBe('');
    expect(built.queryParts).toEqual([]);
  });

  it('emits the bare entity-set path when no options are set', () => {
    const built = buildRetrieveMultiple(retrieveMultipleState());
    expect(built.relativeUrl).toBe(`${API}/accounts`);
    expect(built.relativeNoBase).toBe('/accounts');
    expect(built.entitySet).toBe('accounts');
    expect(built.entityLogical).toBe('account');
    expect(built.bytes).toBe(built.relativeUrl.length);
  });

  it('rewrites lookup columns in $select to the _<col>_value form', () => {
    const built = buildRetrieveMultiple(
      retrieveMultipleState({ select: ['name', 'primarycontactid'] }),
    );
    expect(built.relativeUrl).toBe(`${API}/accounts?$select=name,_primarycontactid_value`);
  });

  it('composes the query options in a stable order', () => {
    const built = buildRetrieveMultiple(
      retrieveMultipleState({
        select: ['name'],
        filter: group('and', [rule('numberofemployees', 'gt', '10')]),
        expand: [{ ...makeExpandSpec('primarycontactid'), select: ['fullname'] }],
        orderby: [
          { col: 'name', dir: 'asc' },
          { col: 'createdon', dir: 'desc' },
        ],
        top: 5,
        countOn: true,
      }),
    );
    expect(built.queryParts.map((p) => p.key)).toEqual([
      '$select',
      '$filter',
      '$expand',
      '$orderby',
      '$top',
      '$count',
    ]);
    expect(built.relativeNoBase).toBe(
      '/accounts?$select=name&$filter=numberofemployees gt 10' +
        '&$expand=primarycontactid($select=fullname)&$orderby=name,createdon desc&$top=5&$count=true',
    );
  });

  it('quotes string literals per column type and doubles embedded apostrophes', () => {
    const built = buildRetrieveMultiple(
      retrieveMultipleState({
        filter: group('and', [
          rule('name', 'eq', "O'Brien"),
          rule('revenue', 'ge', '1000'),
          rule('donotemail', 'eq', 'true'),
          rule('primarycontactid', 'eq', CONTACT_ID),
        ]),
      }),
    );
    const filter = built.queryParts.find((p) => p.key === '$filter')?.value;
    expect(filter).toBe(
      `name eq 'O''Brien' and revenue ge 1000 and donotemail eq true and _primarycontactid_value eq ${CONTACT_ID}`,
    );
  });

  it('percent-encodes characters that would break the query string', () => {
    // `&` would start a new option, `#` a fragment, `%` an escape and `+`
    // would read as a space — they must not appear raw in a value.
    const built = buildRetrieveMultiple(
      retrieveMultipleState({
        filter: group('and', [rule('name', 'contains', 'A&B #1 50% C+D')]),
      }),
    );
    const filter = built.queryParts.find((p) => p.key === '$filter')?.value;
    expect(filter).toBe("contains(name,'A%26B %231 50%25 C%2BD')");
    expect(decodeURIComponent(filter!)).toBe("contains(name,'A&B #1 50% C+D')");
  });

  it('parenthesises nested groups and honours a group-level not', () => {
    const built = buildRetrieveMultiple(
      retrieveMultipleState({
        filter: group('and', [
          rule('name', 'startswith', 'Contoso'),
          group('or', [rule('numberofemployees', 'lt', '5'), rule('industrycode', 'eq', '1')], {
            negated: true,
          }),
        ]),
      }),
    );
    const filter = built.queryParts.find((p) => p.key === '$filter')?.value;
    expect(filter).toBe(
      "startswith(name,'Contoso') and (not (numberofemployees lt 5 or industrycode eq 1))",
    );
  });

  it('emits lambda filters against the related table', () => {
    const built = buildRetrieveMultiple(
      retrieveMultipleState({
        filter: group('and', [
          {
            id: 'l1',
            type: 'lambda',
            nav: 'contact_customer_accounts',
            lambda: 'any',
            alias: 'c',
            inner: group('and', [rule('c/jobtitle', 'eq', 'CEO')]),
          },
        ]),
      }),
    );
    const filter = built.queryParts.find((p) => p.key === '$filter')?.value;
    // The lambda's `:` is percent-encoded on the wire.
    expect(filter).toBe("contact_customer_accounts/any(c%3Ac/jobtitle eq 'CEO')");
    expect(decodeURIComponent(filter!)).toBe(
      "contact_customer_accounts/any(c:c/jobtitle eq 'CEO')",
    );
  });

  it('drops $top when odata.maxpagesize is set (they are mutually exclusive)', () => {
    const built = buildRetrieveMultiple(
      retrieveMultipleState({ top: 10, prefer: { ...emptyPrefer(), maxpagesize: 50 } }),
    );
    expect(built.queryParts.find((p) => p.key === '$top')).toBeUndefined();
  });

  it('keeps $top next to a page size when asked (Execute on a host without Prefer, #219)', () => {
    const built = buildRetrieveMultiple(
      retrieveMultipleState({ top: 10, prefer: { ...emptyPrefer(), maxpagesize: 50 } }),
      { keepTop: true },
    );
    expect(built.relativeNoBase).toBe('/accounts?$top=10');
  });

  it('lets $apply replace $select / $filter / $expand / $orderby', () => {
    const built = buildRetrieveMultiple(
      retrieveMultipleState({
        select: ['name'],
        orderby: [{ col: 'name', dir: 'asc' }],
        apply: {
          enabled: true,
          prefilter: group('and', [rule('numberofemployees', 'gt', '0')]),
          groupby: ['industrycode'],
          aggregates: [{ col: 'revenue', fn: 'sum', alias: 'total' }],
        },
      }),
    );
    expect(built.queryParts.map((p) => p.key)).toEqual(['$apply']);
    expect(built.queryParts[0].value).toBe(
      'filter(numberofemployees gt 0)/groupby((industrycode),aggregate(revenue with sum as total))',
    );
  });

  it('encodes nested $expand with the inner options separated by semicolons', () => {
    const built = buildRetrieveMultiple(
      retrieveMultipleState({
        expand: [
          {
            ...makeExpandSpec('contact_customer_accounts'),
            select: ['fullname', 'parentcustomerid'],
            orderby: [{ col: 'fullname', dir: 'desc' }],
            top: 3,
            nestedExpand: [{ ...makeExpandSpec('parentcustomerid_account'), select: ['name'] }],
          },
        ],
      }),
    );
    expect(built.queryParts[0]).toEqual({
      key: '$expand',
      value:
        'contact_customer_accounts($select=fullname,_parentcustomerid_value;$orderby=fullname desc;' +
        '$top=3;$expand=parentcustomerid_account($select=name))',
    });
  });
});

describe('buildRetrieveSingle', () => {
  it('uses an <id> placeholder until a record is picked', () => {
    const built = buildRetrieveSingle({
      table: 'account',
      recordId: null,
      select: ['name'],
      expand: [],
      prefer: emptyPrefer(),
      headers: [],
      dirty: new Set(),
    });
    expect(built.relativeUrl).toBe(`${API}/accounts(<id>)?$select=name`);
    expect(built.recordId).toBeUndefined();
  });

  it('addresses the picked record', () => {
    const built = buildRetrieveSingle({
      table: 'account',
      recordId: ACCOUNT_ID,
      select: [],
      expand: [],
      prefer: emptyPrefer(),
      headers: [],
      dirty: new Set(),
    });
    expect(built.relativeNoBase).toBe(`/accounts(${ACCOUNT_ID})`);
    expect(built.recordId).toBe(ACCOUNT_ID);
  });
});

describe('buildRetrieveNextLink', () => {
  it('strips the host and derives the entity set from an opaque nextLink', () => {
    const url =
      'https://contoso.crm.dynamics.com/api/data/v9.2/opportunities?$select=name&$skiptoken=%253Ccookie%253E';
    const built = buildRetrieveNextLink({
      url,
      prefer: emptyPrefer(),
      headers: [],
      dirty: new Set(),
    });
    expect(built.relativeUrl).toBe(
      '/api/data/v9.2/opportunities?$select=name&$skiptoken=%253Ccookie%253E',
    );
    expect(built.relativeNoBase).toBe('/opportunities?$select=name&$skiptoken=%253Ccookie%253E');
    expect(built.entitySet).toBe('opportunities');
    // Not loaded: no guess from the plural (#224).
    expect(built.entityLogical).toBe('');
  });

  const nextLink = (set: string) =>
    buildRetrieveNextLink({
      url: `https://contoso.crm.dynamics.com/api/data/v9.2/${set}?$skiptoken=x`,
      prefer: emptyPrefer(),
      headers: [],
      dirty: new Set(),
    });

  it('takes the table from the loaded metadata by entity set name (#224)', () => {
    expect(nextLink('accounts').entityLogical).toBe('account');
    expect(nextLink('contacts').entityLogical).toBe('contact');
  });

  it('uses the entity list for sets whose table is not loaded, irregular plurals included', () => {
    const list: Record<string, string> = {
      activitypointers: 'activitypointer',
      addresses: 'customeraddress',
    };
    const built = (set: string) =>
      buildRetrieveNextLink(
        {
          url: `https://contoso.crm.dynamics.com/api/data/v9.2/${set}?$skiptoken=x`,
          prefer: emptyPrefer(),
          headers: [],
          dirty: new Set(),
        },
        (s) => list[s],
      ).entityLogical;
    expect(built('activitypointers')).toBe('activitypointer');
    expect(built('addresses')).toBe('customeraddress');
    // Not in the list: the loaded metadata still answers.
    expect(built('accounts')).toBe('account');
  });

  it("doesn't strip a plural it can't look up", () => {
    // The old guess turned these into "addresse" and "activitypointer".
    expect(nextLink('addresses').entityLogical).toBe('');
    expect(nextLink('activitypointers').entityLogical).toBe('');
  });
});

describe('write builders', () => {
  it('Create only carries $select when return=representation is on', () => {
    const off = buildCreate(createState({ returnSelect: ['name'] }));
    expect(off.relativeUrl).toBe(`${API}/accounts`);
    const on = buildCreate(
      createState({
        returnSelect: ['name', 'primarycontactid'],
        prefer: { ...emptyPrefer(), returnRepresentation: true },
      }),
    );
    expect(on.relativeUrl).toBe(`${API}/accounts?$select=name,_primarycontactid_value`);
  });

  it('Create body binds lookups, joins multi-select values and drops empty strings', () => {
    const body = buildCreateBody(
      createState({
        fieldValues: {
          name: 'Contoso',
          accountnumber: '',
          numberofemployees: 0,
          donotemail: false,
          sample_tags: [100, 200],
          primarycontactid: { id: CONTACT_ID, targetEntity: 'contact' },
        },
      }),
    );
    expect(body).toEqual({
      name: 'Contoso',
      numberofemployees: 0,
      donotemail: false,
      sample_tags: '100,200',
      'primarycontactid@odata.bind': `/contacts(${CONTACT_ID})`,
    });
  });

  it('Create body uses the target-disambiguated nav for polymorphic lookups', () => {
    const body = buildCreateBody(
      createState({
        table: 'contact',
        fieldValues: { parentcustomerid: { id: ACCOUNT_ID, targetEntity: 'account' } },
      }),
    );
    expect(body).toEqual({ 'parentcustomerid_account@odata.bind': `/accounts(${ACCOUNT_ID})` });
  });

  it('Create body emits explicit nulls, and null wins over a typed value', () => {
    const body = buildCreateBody(
      createState({
        table: 'contact',
        fieldValues: { jobtitle: 'CEO' },
        nullFields: ['jobtitle', 'parentcustomerid'],
      }),
    );
    expect(body).toEqual({ jobtitle: null, 'parentcustomerid_account@odata.bind': null });
  });

  it('Update PUT addresses a single column and wraps the value', () => {
    const state = {
      table: 'account',
      recordId: ACCOUNT_ID,
      method: 'PUT' as const,
      putColumn: 'name',
      fieldValues: { name: 'Fabrikam' },
      nullFields: [],
      prefer: { ...emptyPrefer(), returnRepresentation: true },
      headers: [],
      returnSelect: ['name'],
      concurrency: { kind: 'update-only' as const },
      bypass: defaultBypassOptions(),
      dirty: new Set<string>(),
    };
    // PUT on a property URL takes no query options, even with return=representation.
    expect(buildUpdate(state).relativeUrl).toBe(`${API}/accounts(${ACCOUNT_ID})/name`);
    expect(buildUpdateBody(state)).toEqual({ value: 'Fabrikam' });
    expect(buildUpdate({ ...state, method: 'PATCH' }).relativeUrl).toBe(
      `${API}/accounts(${ACCOUNT_ID})?$select=name`,
    );
  });

  it('Upsert addresses rows by alternate key with OData literals', () => {
    const base = {
      table: 'account',
      fieldValues: {},
      nullFields: [],
      prefer: emptyPrefer(),
      headers: [],
      returnSelect: [],
      concurrency: { kind: 'none' as const },
      bypass: defaultBypassOptions(),
      dirty: new Set<string>(),
    };
    const single = buildUpsert({
      ...base,
      key: {
        kind: 'alternate',
        keyName: 'akey_accountnumber',
        keyValues: { accountnumber: "A'1" },
      },
    });
    expect(single.relativeNoBase).toBe("/accounts(accountnumber='A''1')");
    expect(single.recordId).toBeUndefined();

    const composite = buildUpsert({
      ...base,
      key: {
        kind: 'alternate',
        keyName: 'akey_composite',
        keyValues: { numberofemployees: '42', name: 'Contoso' },
      },
    });
    expect(composite.relativeNoBase).toBe("/accounts(name='Contoso',numberofemployees=42)");

    const guid = buildUpsert({ ...base, key: { kind: 'guid', recordId: ACCOUNT_ID } });
    expect(guid.relativeNoBase).toBe(`/accounts(${ACCOUNT_ID})`);
    expect(guid.recordId).toBe(ACCOUNT_ID);
  });

  it('Delete of a single property targets the column URL', () => {
    const built = buildDelete({
      table: 'account',
      recordId: ACCOUNT_ID,
      scope: { kind: 'single-property', column: 'revenue' },
      concurrency: { kind: 'none' },
      headers: [],
      bypassCustomPlugins: false,
      bypass: defaultBypassOptions(),
      confirmText: '',
      acknowledged: false,
      dirty: new Set(),
    });
    expect(built.relativeUrl).toBe(`${API}/accounts(${ACCOUNT_ID})/revenue`);
  });

  it('Merge body copies subordinate / custom values and skips target choices', () => {
    const state: MergeState = {
      table: 'account',
      targetId: ACCOUNT_ID,
      subordinateId: '00000000-0000-0000-0000-0000000000a2',
      fieldChoices: {
        name: 'target',
        revenue: 'subordinate',
        sample_tags: 'subordinate',
        accountnumber: 'custom',
      },
      customValues: { accountnumber: 'ACC-9' },
      subordinateSnapshot: { name: 'Sub', revenue: 5000, sample_tags: '100,200' },
      performParentingChecks: false,
      suppressDuplicateDetection: false,
      bypass: defaultBypassOptions(),
      headers: [],
      dirty: new Set(),
    };
    const body = buildMergeBody(state);
    expect(body).toEqual({
      Target: { '@odata.type': 'Microsoft.Dynamics.CRM.account', accountid: ACCOUNT_ID },
      Subordinate: {
        '@odata.type': 'Microsoft.Dynamics.CRM.account',
        accountid: '00000000-0000-0000-0000-0000000000a2',
      },
      UpdateContent: {
        '@odata.type': 'Microsoft.Dynamics.CRM.account',
        revenue: 5000,
        sample_tags: '100,200',
        accountnumber: 'ACC-9',
      },
      PerformParentingChecks: false,
    });
  });
});

describe('relate builders', () => {
  const relate = {
    table: 'account',
    sourceId: ACCOUNT_ID,
    targetNames: {},
    headers: [],
    dirty: new Set<string>(),
  };
  const C2 = '00000000-0000-0000-0000-0000000000c2';

  it('Associate on a collection nav POSTs one $ref per target with an absolute @odata.id', () => {
    const state = {
      ...relate,
      navProperty: 'contact_customer_accounts',
      targets: [CONTACT_ID, '', C2],
    };
    const reqs = buildAssociateRequests(state);
    expect(reqs).toEqual([
      {
        method: 'POST',
        relativeUrl: `${API}/accounts(${ACCOUNT_ID})/contact_customer_accounts/$ref`,
        relativeNoBase: `/accounts(${ACCOUNT_ID})/contact_customer_accounts/$ref`,
        body: { '@odata.id': `https://contoso.crm.dynamics.com${API}/contacts(${CONTACT_ID})` },
        targetId: CONTACT_ID,
      },
      expect.objectContaining({ targetId: C2 }),
    ]);
    expect(buildAssociate(state).relativeNoBase).toBe(
      `/accounts(${ACCOUNT_ID})/contact_customer_accounts/$ref`,
    );
  });

  it('Associate on a single-valued nav PATCHes the source with a relative @odata.bind', () => {
    const state = { ...relate, navProperty: 'primarycontactid', targets: [CONTACT_ID, C2] };
    expect(buildAssociateRequests(state)).toEqual([
      {
        method: 'PATCH',
        relativeUrl: `${API}/accounts(${ACCOUNT_ID})`,
        relativeNoBase: `/accounts(${ACCOUNT_ID})`,
        body: { 'primarycontactid@odata.bind': `contacts(${CONTACT_ID})` },
        targetId: CONTACT_ID,
      },
    ]);
    expect(buildAssociate(state).relativeNoBase).toBe(`/accounts(${ACCOUNT_ID})`);
    expect(buildAssociateBody({ ...state, targets: [] })).toEqual({
      'primarycontactid@odata.bind': 'contacts(<target-id>)',
    });
  });

  it('Disassociate on a collection nav DELETEs each link by target id', () => {
    const state = {
      ...relate,
      navProperty: 'contact_customer_accounts',
      targetIds: [CONTACT_ID, C2],
    };
    expect(buildDisassociateRequests(state).map((r) => [r.method, r.relativeNoBase])).toEqual([
      ['DELETE', `/accounts(${ACCOUNT_ID})/contact_customer_accounts(${CONTACT_ID})/$ref`],
      ['DELETE', `/accounts(${ACCOUNT_ID})/contact_customer_accounts(${C2})/$ref`],
    ]);
    expect(buildDisassociate(state).relativeNoBase).toBe(
      `/accounts(${ACCOUNT_ID})/contact_customer_accounts(${CONTACT_ID})/$ref`,
    );
  });

  it('Disassociate on a single-valued nav PATCHes the bind to null', () => {
    const state = { ...relate, navProperty: 'primarycontactid', targetIds: [] };
    expect(buildDisassociateRequests(state)).toEqual([
      {
        method: 'PATCH',
        relativeUrl: `${API}/accounts(${ACCOUNT_ID})`,
        relativeNoBase: `/accounts(${ACCOUNT_ID})`,
        targetId: null,
        body: { 'primarycontactid@odata.bind': null },
      },
    ]);
  });
});

describe('buildPredefinedQuery', () => {
  it('emits savedQuery / userQuery with a placeholder until a query is picked', () => {
    const base = {
      table: 'account',
      queryType: 'savedQuery' as const,
      top: null,
      prefer: emptyPrefer(),
      headers: [],
      dirty: new Set<string>(),
    };
    expect(buildPredefinedQuery({ ...base, queryId: null }).relativeNoBase).toBe(
      '/accounts?savedQuery=<query-guid>',
    );
    expect(
      buildPredefinedQuery({ ...base, queryType: 'userQuery', queryId: 'abc', top: 3 })
        .relativeNoBase,
    ).toBe('/accounts?userQuery=abc&$top=3');
  });
});
