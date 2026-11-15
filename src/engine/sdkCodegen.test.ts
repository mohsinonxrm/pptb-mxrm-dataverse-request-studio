// SDK code generation (#232): C# ServiceClient and Microsoft.Xrm.Data.PowerShell.
// Golden fragments per mode group, the no-faithful-form fallbacks, and
// literals that C# / PowerShell / FetchXML read back as the exact value.
import { beforeAll, describe, expect, it, vi } from 'vitest';
import { generateCode, type CodegenInputs } from './codeGenerators';
import type { BuiltRequest } from './urlBuilder';
import { __registerLiveTables, type TableMeta } from '../metadata/metadata';
import type { CsdlAction } from '../metadata/actionsCsdl';
import { parseFilterForCodegen } from './odataParser';
import { psQuote } from './codegenStrings';
import { xmlEscape } from './sdkCodegen';

vi.mock('../env/environment', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../env/environment')>()),
  ENV: {
    name: 'Test',
    host: 'contoso.crm.dynamics.com',
    apiBase: '/api/data/v9.2',
    user: '',
  },
}));

const ACTIONS = new Map<string, CsdlAction>();
vi.mock('../host/csdlProvider', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/csdlProvider')>()),
  findActionSync: (name: string) => ACTIONS.get(name),
}));

const API = '/api/data/v9.2';
const ID = '00000000-0000-0000-0000-0000000000a1';
const T1 = '00000000-0000-0000-0000-0000000000b1';
const T2 = '00000000-0000-0000-0000-0000000000b2';
const LS = String.fromCharCode(0x2028);

const account = {
  logicalName: 'account',
  entitySetName: 'accounts',
  displayName: 'Account',
  primaryKey: 'accountid',
  primaryName: 'name',
  columns: [
    { logicalName: 'name', attributeType: 'String' },
    { logicalName: 'description', attributeType: 'Memo' },
    { logicalName: 'revenue', attributeType: 'Money' },
    { logicalName: 'industrycode', attributeType: 'Picklist' },
    { logicalName: 'statecode', attributeType: 'State' },
    { logicalName: 'createdon', attributeType: 'DateTime' },
    { logicalName: 'primarycontactid', attributeType: 'Lookup', targets: ['contact'] },
    { logicalName: 'numberofemployees', attributeType: 'Integer' },
    { logicalName: 'exchangerate', attributeType: 'Decimal' },
    { logicalName: 'creditonhold', attributeType: 'Boolean' },
    { logicalName: 'new_tags', attributeType: 'MultiSelectPicklist' },
    { logicalName: 'new_document', attributeType: 'File' },
    { logicalName: 'entityimage', attributeType: 'Image' },
  ],
  navigationProperties: [
    {
      name: 'primarycontactid',
      relationshipName: 'account_primary_contact',
      cardinality: 'ManyToOne',
      targetEntity: 'contact',
      referencingAttribute: 'primarycontactid',
    },
    {
      name: 'contact_customer_accounts',
      relationshipName: 'contact_customer_accounts',
      cardinality: 'OneToMany',
      targetEntity: 'contact',
    },
    {
      name: 'Referencedaccount_parent_account',
      relationshipName: 'account_parent_account',
      cardinality: 'OneToMany',
      targetEntity: 'account',
    },
    {
      name: 'new_account_account',
      relationshipName: 'new_account_account',
      cardinality: 'ManyToMany',
      targetEntity: 'account',
    },
  ],
} as unknown as TableMeta;
const contact = {
  logicalName: 'contact',
  entitySetName: 'contacts',
  displayName: 'Contact',
  primaryKey: 'contactid',
  primaryName: 'fullname',
  columns: [{ logicalName: 'fullname', attributeType: 'String' }],
  navigationProperties: [],
} as unknown as TableMeta;

beforeAll(() => __registerLiveTables([account, contact]));

function built(
  path: string,
  parts: [string, string][] = [],
  extra: Partial<BuiltRequest> = {},
): BuiltRequest {
  const q = parts.map(([k, v]) => `${k}=${v}`).join('&');
  const noBase = q ? `${path}?${q}` : path;
  return {
    relativeUrl: `${API}${noBase}`,
    relativeNoBase: noBase,
    bytes: 0,
    queryParts: parts.map(([key, value]) => ({ key, value })),
    entitySet: 'accounts',
    entityLogical: 'account',
    ...extra,
  };
}
const unbound = { entitySet: '', entityLogical: '' };

const cs = (method: string, b: BuiltRequest, extra: Partial<CodegenInputs> = {}) =>
  generateCode('csharp-sdk', { method, built: b, headers: {}, ...extra });
const ps = (method: string, b: BuiltRequest, extra: Partial<CodegenInputs> = {}) =>
  generateCode('powershell-xrm', { method, built: b, headers: {}, ...extra });

const query = (parts: [string, string][], extra: Partial<CodegenInputs> = {}) =>
  cs('GET', built('/accounts', parts), extra);

/** Read a regular C# string literal ("…") the way the compiler does. */
function csUnquote(lit: string): string {
  expect(lit.startsWith('"') && lit.endsWith('"')).toBe(true);
  const SIMPLE: Record<string, string> = { '\\': '\\', '"': '"', n: '\n', r: '\r', t: '\t' };
  let out = '';
  for (let p = 1; p < lit.length - 1; p++) {
    const c = lit[p];
    if (c === '"' || /[\n\r\u0085]/.test(c) || c === LS || c === String.fromCharCode(0x2029)) {
      throw new Error(`bare ${JSON.stringify(c)} in a C# literal`);
    }
    if (c !== '\\') {
      out += c;
      continue;
    }
    const e = lit[++p];
    if (e === 'u') {
      out += String.fromCharCode(parseInt(lit.slice(p + 1, p + 5), 16));
      p += 4;
    } else if (e in SIMPLE) out += SIMPLE[e];
    else throw new Error(`unknown escape \\${e}`);
  }
  return out;
}

/** Read a PowerShell single-quoted literal: any single-quote character doubled is one. */
function psUnquote(lit: string): string {
  const QUOTES = new Set([
    "'",
    ...[0x2018, 0x2019, 0x201a, 0x201b].map((c) => String.fromCharCode(c)),
  ]);
  expect(lit[0]).toBe("'");
  let out = '';
  for (let p = 1; p < lit.length; p++) {
    const c = lit[p];
    if (QUOTES.has(c)) {
      if (QUOTES.has(lit[p + 1] ?? '')) {
        out += c;
        p++;
        continue;
      }
      expect(p).toBe(lit.length - 1); // the literal ends here, nowhere earlier
      return out;
    }
    out += c;
  }
  throw new Error('unterminated PowerShell literal');
}

describe('C# SDK — queries (QueryExpression)', () => {
  it('maps $select, $filter, $orderby and the page size', () => {
    const out = query(
      [
        ['$select', 'name,_primarycontactid_value'],
        [
          '$filter',
          "statecode eq 0 and (contains(name,'Contoso') or revenue gt 1000) and createdon ge 2024-01-01T00:00:00Z",
        ],
        ['$orderby', 'name desc,createdon'],
      ],
      { headers: { Prefer: 'odata.maxpagesize=50' } },
    );
    expect(out).toContain('var query = new QueryExpression("account")');
    expect(out).toContain('ColumnSet = new ColumnSet("name", "primarycontactid"),');
    expect(out).toContain('PageInfo = new PagingInfo { PageNumber = 1, Count = 50 },');
    expect(out).toContain('query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);');
    expect(out).toContain('var filter1 = new FilterExpression(LogicalOperator.Or);');
    expect(out).toContain('filter1.AddCondition("name", ConditionOperator.Like, "%Contoso%");');
    // revenue is a Money column: the condition value is a decimal.
    expect(out).toContain('filter1.AddCondition("revenue", ConditionOperator.GreaterThan, 1000m);');
    expect(out).toContain('query.Criteria.AddFilter(filter1);');
    expect(out).toContain(
      'ConditionOperator.GreaterEqual, DateTime.Parse("2024-01-01T00:00:00Z", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind)',
    );
    expect(out).toContain('using System.Globalization;');
    expect(out).toContain('query.AddOrder("name", OrderType.Descending);');
    expect(out).toContain('query.AddOrder("createdon", OrderType.Ascending);');
    expect(out).toContain('EntityCollection results = service.RetrieveMultiple(query);');
    expect(out).toContain('using var service = new ServiceClient(');
    expect(out).toContain('Url=https://contoso.crm.dynamics.com;');
  });

  it('sets the root filter to Or and keeps quoted numbers as strings', () => {
    const out = query([['$filter', "name eq '42' or numberofemployees eq 42"]]);
    expect(out).toContain('query.Criteria.FilterOperator = LogicalOperator.Or;');
    expect(out).toContain('AddCondition("name", ConditionOperator.Equal, "42");');
    expect(out).toContain('AddCondition("numberofemployees", ConditionOperator.Equal, 42);');
  });

  it('types literals by shape when the column has no metadata', () => {
    const out = cs(
      'GET',
      built(
        '/widgets',
        [
          [
            '$filter',
            `a eq 5 and b eq 2.5 and c eq true and d eq ${ID} and e eq 2024-05-01 and f eq 9999999999`,
          ],
        ],
        { entitySet: 'widgets', entityLogical: 'new_widget' },
      ),
    );
    expect(out).toContain('AddCondition("a", ConditionOperator.Equal, 5);');
    expect(out).toContain('AddCondition("b", ConditionOperator.Equal, 2.5m);');
    expect(out).toContain('AddCondition("c", ConditionOperator.Equal, true);');
    expect(out).toContain(`AddCondition("d", ConditionOperator.Equal, new Guid("${ID}"));`);
    expect(out).toContain('AddCondition("e", ConditionOperator.Equal, DateTime.Parse("2024-05-01"');
    expect(out).toContain('AddCondition("f", ConditionOperator.Equal, 9999999999L);');
  });

  it('maps null checks, negated string functions and column comparisons', () => {
    const out = query([
      [
        '$filter',
        "name eq null and description ne null and not startswith(name,'A') and revenue gt exchangerate",
      ],
    ]);
    expect(out).toContain('AddCondition("name", ConditionOperator.Null);');
    expect(out).toContain('AddCondition("description", ConditionOperator.NotNull);');
    expect(out).toContain('AddCondition("name", ConditionOperator.NotLike, "A%");');
    expect(out).toContain(
      'AddCondition(new ConditionExpression("revenue", ConditionOperator.GreaterThan, true, "exchangerate"));',
    );
  });

  it('maps Dataverse query functions to the ConditionOperator of the same name', () => {
    const out = query([
      [
        '$filter',
        `Microsoft.Dynamics.CRM.LastXDays(PropertyName='createdon',PropertyValue=7) and Microsoft.Dynamics.CRM.In(PropertyName='industrycode',PropertyValues=['1','2']) and Microsoft.Dynamics.CRM.EqualUserId(PropertyName='ownerid') and Microsoft.Dynamics.CRM.Under(PropertyName='accountid',PropertyValue='${ID}') and Microsoft.Dynamics.CRM.Between(PropertyName='revenue',PropertyValues=['10','20'])`,
      ],
    ]);
    expect(out).toContain('AddCondition("createdon", ConditionOperator.LastXDays, 7);');
    // industrycode is a choice: In takes ints.
    expect(out).toContain('AddCondition("industrycode", ConditionOperator.In, 1, 2);');
    expect(out).toContain('AddCondition("ownerid", ConditionOperator.EqualUserId);');
    expect(out).toContain(`AddCondition("accountid", ConditionOperator.Under, new Guid("${ID}"));`);
    expect(out).toContain('AddCondition("revenue", ConditionOperator.Between, 10m, 20m);');
  });

  it('requests the total count with PageInfo when there is no $top', () => {
    const out = query([
      ['$select', 'name'],
      ['$count', 'true'],
    ]);
    expect(out).toContain('Count = 5000, ReturnTotalRecordCount = true');
    expect(out).toContain('results.TotalRecordCount');
  });

  it('keeps TopCount over a count or page size, and says so', () => {
    const out = query(
      [
        ['$select', 'name'],
        ['$top', '5'],
        ['$count', 'true'],
      ],
      { headers: { Prefer: 'odata.maxpagesize=10' } },
    );
    expect(out).toContain('TopCount = 5,');
    expect(out).not.toContain('PageInfo = new');
    expect(out).toContain("// The total count isn't requested");
    expect(out).toContain('// odata.maxpagesize is left out');
  });

  it('sends Count only as a one-row page with the total', () => {
    const out = query(
      [
        ['$top', '1'],
        ['$count', 'true'],
      ],
      {
        headers: {
          Prefer:
            'odata.include-annotations="Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded"',
        },
      },
    );
    expect(out).toContain(
      'PageInfo = new PagingInfo { PageNumber = 1, Count = 1, ReturnTotalRecordCount = true },',
    );
    expect(out).not.toContain('TopCount');
    expect(out).toContain('Console.WriteLine($"Total: {results.TotalRecordCount}');
    expect(out).not.toContain('foreach (Entity row');
  });

  it('asks for all columns, with a note, when there is no $select', () => {
    const out = query([]);
    expect(out).toContain('ColumnSet = new ColumnSet(true),');
    expect(out).toContain('// No $select: every column comes back.');
  });

  it.each([
    ['$expand', 'primarycontactid($select=fullname)', 'LinkEntity'],
    ['$apply', 'groupby((statecode))', 'FetchXML'],
    ['savedQuery', ID, 'FetchExpression'],
    ['$filter', "primarycontactid/fullname eq 'x'", 'LinkEntity'],
    ['$filter', "contact_customer_accounts/any(c:c/fullname eq 'x')", 'LinkEntity'],
    ['$filter', 'not (statecode eq 0 or statecode eq 1)', 'NOT over a group'],
    ['$orderby', 'primarycontactid/fullname', 'LinkEntity'],
  ])('falls back to ExecuteWebRequest for %s=%s', (key, value, why) => {
    const out = query([[key, value]]);
    expect(out).toContain('// No faithful SDK form for this request');
    expect(out).toContain(why);
    expect(out).toContain('service.ExecuteWebRequest(');
    expect(out).toContain('new HttpMethod("GET")');
    expect(out).not.toContain('QueryExpression("account")');
  });

  it('sends the fallback URL and headers as-is, without the credential', () => {
    const out = query([['$expand', 'primarycontactid($select=fullname)']], {
      headers: { Authorization: 'Bearer SECRET', Prefer: 'odata.include-annotations="*"' },
    });
    expect(out).toContain('"accounts?$expand=primarycontactid($select=fullname)"');
    expect(out).toContain('["Prefer"] = new List<string> { "odata.include-annotations=\\"*\\"" },');
    expect(out).not.toContain('SECRET');
  });

  it('notes that formatted values need no Prefer header', () => {
    const out = query([['$select', 'name']], {
      headers: { Prefer: 'odata.include-annotations="*"' },
    });
    expect(out).toContain('// Formatted values come back in FormattedValues');
    expect(out).toContain('row.FormattedValues.TryGetValue');
  });
});

describe('C# SDK — single row', () => {
  it('retrieves by id with the selected columns', () => {
    const out = cs(
      'GET',
      built(`/accounts(${ID})`, [['$select', 'name,revenue']], { recordId: ID }),
    );
    expect(out).toContain('var columns = new ColumnSet("name", "revenue");');
    expect(out).toContain(`Entity row = service.Retrieve("account", new Guid("${ID}"), columns);`);
  });

  it('retrieves by alternate key with a RetrieveRequest', () => {
    const out = cs('GET', built(`/accounts(accountnumber='A-1')`, [['$select', 'name']]));
    expect(out).toContain(
      'Target = new EntityReference("account", new KeyAttributeCollection { ["accountnumber"] = "A-1" }), ColumnSet = columns',
    );
    expect(out).toContain('((RetrieveResponse)service.Execute(request)).Entity');
  });

  it('retrieves one column for /<column>, and falls back for related rows', () => {
    expect(cs('GET', built(`/accounts(${ID})/name`, [], { recordId: ID }))).toContain(
      'new ColumnSet("name")',
    );
    expect(
      cs('GET', built(`/accounts(${ID})/contact_customer_accounts`, [], { recordId: ID })),
    ).toContain('service.ExecuteWebRequest(');
  });

  it('falls back for $expand on a single row', () => {
    const out = cs(
      'GET',
      built(`/accounts(${ID})`, [['$expand', 'primarycontactid']], { recordId: ID }),
    );
    expect(out).toContain('Retrieve has no expand');
    expect(out).toContain('service.ExecuteWebRequest(');
  });
});

describe('C# SDK — writes', () => {
  it('creates an Entity with typed SDK values from the metadata', () => {
    const out = cs('POST', built('/accounts'), {
      body: {
        name: 'Contoso',
        revenue: 12.5,
        industrycode: 3,
        statecode: 0,
        numberofemployees: 10,
        exchangerate: 1.25,
        creditonhold: false,
        createdon: '2024-01-02T03:04:05Z',
        new_tags: '1,2',
        entityimage: 'AAEC',
        'primarycontactid@odata.bind': `/contacts(${T1})`,
        description: null,
      },
    });
    expect(out).toContain('var entity = new Entity("account")');
    expect(out).toContain('["name"] = "Contoso",');
    expect(out).toContain('["revenue"] = new Money(12.5m),');
    expect(out).toContain('["industrycode"] = new OptionSetValue(3),');
    expect(out).toContain('["statecode"] = new OptionSetValue(0),');
    expect(out).toContain('["numberofemployees"] = 10,');
    expect(out).toContain('["exchangerate"] = 1.25m,');
    expect(out).toContain('["creditonhold"] = false,');
    expect(out).toContain('["createdon"] = DateTime.Parse("2024-01-02T03:04:05Z"');
    expect(out).toContain(
      '["new_tags"] = new OptionSetValueCollection { new OptionSetValue(1), new OptionSetValue(2) },',
    );
    expect(out).toContain('["entityimage"] = Convert.FromBase64String("AAEC"),');
    expect(out).toContain(
      `["primarycontactid"] = new EntityReference("contact", new Guid("${T1}")),`,
    );
    expect(out).toContain('["description"] = null,');
    expect(out).toContain('Guid id = service.Create(entity);');
    expect(out).not.toContain('values follow the JSON');
  });

  it('follows the JSON types, with a note, when the columns have no metadata', () => {
    const out = cs(
      'POST',
      built('/widgets', [], { entitySet: 'widgets', entityLogical: 'new_widget' }),
      {
        body: { new_name: 'x', new_count: 3, new_rate: 2.5 },
      },
    );
    expect(out).toContain('["new_count"] = 3,');
    expect(out).toContain('["new_rate"] = 2.5m,');
    expect(out).toContain(
      'values follow the JSON: wrap a choice in OptionSetValue and money in Money',
    );
  });

  it('flags a guessed table for a bind to an entity set with no metadata', () => {
    const out = cs('POST', built('/accounts'), {
      body: { 'new_categoryid@odata.bind': `/new_categories(${T1})` },
    });
    expect(out).toContain(`new EntityReference("new_category", new Guid("${T1}"))`);
    expect(out).toContain(
      "// Check: 'new_category' is a guess at the table behind 'new_categories'",
    );
  });

  it('maps the bypass headers to optional parameters on a CreateRequest', () => {
    const out = cs('POST', built('/accounts'), {
      headers: {
        'MSCRM.BypassBusinessLogicExecution': 'CustomSync,CustomAsync',
        'MSCRM.BypassCustomPluginExecution': 'true',
        'MSCRM.SuppressDuplicateDetection': 'false',
        'MSCRM.BypassBusinessLogicExecutionStepIds': T1,
        MSCRMCallerID: T2,
        'X-Custom': 'v',
      },
      body: { name: 'x' },
    });
    expect(out).toContain('var request = new CreateRequest { Target = entity };');
    expect(out).toContain(
      'request.Parameters["BypassBusinessLogicExecution"] = "CustomSync,CustomAsync";',
    );
    expect(out).toContain('request.Parameters["BypassCustomPluginExecution"] = true;');
    expect(out).toContain('request.Parameters["SuppressDuplicateDetection"] = false;');
    expect(out).toContain(`request.Parameters["BypassBusinessLogicExecutionStepIds"] = "${T1}";`);
    expect(out).toContain(`service.CallerId = new Guid("${T2}");`);
    expect(out).toContain('// Not sent: the SDK has no equivalent of the "X-Custom: v" header.');
    expect(out).toContain('using Microsoft.Xrm.Sdk.Messages;');
  });

  it('notes that return=representation has no SDK equivalent', () => {
    const out = cs('POST', built('/accounts', [['$select', 'name']]), {
      headers: { Prefer: 'return=representation' },
      body: { name: 'x' },
    });
    expect(out).toContain('// The SDK returns no row from a write');
  });

  it('updates with If-Match: * and uses RowVersion for an ETag', () => {
    const b = built(`/accounts(${ID})`, [], { recordId: ID });
    const plain = cs('PATCH', b, { headers: { 'If-Match': '*' }, body: { name: 'x' } });
    expect(plain).toContain(`var entity = new Entity("account", new Guid("${ID}"))`);
    expect(plain).toContain('service.Update(entity);');
    const etag = cs('PATCH', b, { headers: { 'If-Match': 'W/"12345"' }, body: { name: 'x' } });
    expect(etag).toContain('entity.RowVersion = "12345";');
    expect(etag).toContain(
      'new UpdateRequest { Target = entity, ConcurrencyBehavior = ConcurrencyBehavior.IfRowVersionMatches }',
    );
  });

  it('upserts a PATCH without If-Match, including by alternate key', () => {
    const out = cs('PATCH', built(`/accounts(accountnumber='A''1',numberofemployees=5)`), {
      body: { name: 'x' },
    });
    expect(out).toContain(
      'new Entity("account", new KeyAttributeCollection { ["accountnumber"] = "A\'1", ["numberofemployees"] = 5 })',
    );
    expect(out).toContain('var request = new UpsertRequest { Target = entity };');
    expect(out).toContain('response.RecordCreated');
  });

  it('creates with the id for If-None-Match: *, and falls back by alternate key', () => {
    const out = cs('PATCH', built(`/accounts(${ID})`, [], { recordId: ID }), {
      headers: { 'If-None-Match': '*' },
      body: { name: 'x' },
    });
    expect(out).toContain(`new Entity("account", new Guid("${ID}"))`);
    expect(out).toContain('Guid id = service.Create(entity);');
    expect(
      cs('PATCH', built(`/accounts(accountnumber='A1')`), {
        headers: { 'If-None-Match': '*' },
        body: { name: 'x' },
      }),
    ).toContain('Create-only by alternate key');
  });

  it('deletes by id, by alternate key and with an ETag', () => {
    expect(cs('DELETE', built(`/accounts(${ID})`, [], { recordId: ID }))).toContain(
      `service.Delete("account", new Guid("${ID}"));`,
    );
    const key = cs('DELETE', built(`/accounts(accountnumber='A1')`));
    expect(key).toContain(
      'var target = new EntityReference("account", new KeyAttributeCollection { ["accountnumber"] = "A1" });',
    );
    expect(key).toContain('new DeleteRequest { Target = target }');
    const etag = cs('DELETE', built(`/accounts(${ID})`, [], { recordId: ID }), {
      headers: { 'If-Match': '"7"' },
    });
    expect(etag).toContain('target.RowVersion = "7";');
    expect(etag).toContain('ConcurrencyBehavior.IfRowVersionMatches');
  });

  it('sets one column for PUT /<column>, and clears one for DELETE /<column>', () => {
    const b = (col: string) => built(`/accounts(${ID})/${col}`, [], { recordId: ID });
    expect(cs('PUT', b('revenue'), { body: { value: 99 } })).toContain(
      '["revenue"] = new Money(99m),',
    );
    expect(cs('DELETE', b('description'))).toContain('["description"] = null,');
    // File / image columns have their own delete message.
    expect(cs('DELETE', b('new_document'))).toContain('No faithful SDK form');
  });

  it.each([
    [{ primarycontact: { fullname: 'x' } }, 'deep insert'],
    [{ email_activity_parties: [{ partyid: 'x' }] }, 'list of rows'],
    [{ 'name@OData.Community.Display.V1.FormattedValue': 'x' }, 'annotation'],
    [{ primarycontactid: T1 }, '@odata.bind'],
  ])('falls back for the body %j', (body, why) => {
    const out = cs('POST', built('/accounts'), { body });
    expect(out).toContain('No faithful SDK form');
    expect(out).toContain(why);
    expect(out).toContain('new HttpMethod("POST")');
  });
});

describe('C# SDK — associate / disassociate', () => {
  const ref = (t: string) => `https://contoso.crm.dynamics.com${API}/contacts(${t})`;
  const multi = (
    method: string,
    changeSet: boolean,
    url: (t: string) => string,
    body?: (t: string) => Record<string, unknown>,
  ) => ({
    multiRequests: [T1, T2].map((t) => ({
      method,
      relativeUrl: `${API}${url(t)}`,
      body: body?.(t),
    })),
    changeSet,
  });

  it('associates one row with the relationship schema name', () => {
    const out = cs('POST', built(`/accounts(${ID})/contact_customer_accounts/$ref`), {
      body: { '@odata.id': ref(T1) },
    });
    expect(out).toContain('var relationship = new Relationship("contact_customer_accounts");');
    expect(out).toContain(`new EntityReference("contact", new Guid("${T1}")),`);
    expect(out).toContain(
      `service.Associate("account", new Guid("${ID}"), relationship, related);`,
    );
  });

  it('sends one request per row, or one for All or nothing', () => {
    const url = () => `/accounts(${ID})/contact_customer_accounts/$ref`;
    const body = (t: string) => ({ '@odata.id': ref(t) });
    const each = cs('POST', built(url()), multi('POST', false, url, body));
    expect(each).toContain('foreach (var target in targets)');
    expect(each).toContain('new EntityReferenceCollection { target }');
    expect(each).toContain('requests succeeded');
    const atomic = cs('POST', built(url()), multi('POST', true, url, body));
    expect(atomic).toContain('// One request: Dataverse applies all of them or none.');
    expect(atomic).toContain(`new EntityReference("contact", new Guid("${T2}")),`);
    expect(atomic).not.toContain('foreach');
  });

  it('disassociates using the related table from the metadata', () => {
    const url = (t: string) => `/accounts(${ID})/contact_customer_accounts(${t})/$ref`;
    const out = cs('DELETE', built(url(T1)), multi('DELETE', false, url));
    expect(out).toContain('service.Disassociate("account"');
    expect(out).toContain(`new EntityReference("contact", new Guid("${T2}")),`);
    const one = cs('DELETE', built(url(T1)));
    expect(one).toContain(
      `service.Disassociate("account", new Guid("${ID}"), relationship, related);`,
    );
  });

  it('marks a self-referencing relationship', () => {
    const out = cs('POST', built(`/accounts(${ID})/new_account_account/$ref`), {
      body: { '@odata.id': `https://contoso.crm.dynamics.com${API}/accounts(${T1})` },
    });
    expect(out).toContain('relationship.PrimaryEntityRole = EntityRole.Referencing;');
  });

  it('sets, re-points and clears a single-valued lookup', () => {
    const b = built(`/accounts(${ID})`, [], { recordId: ID });
    // DRS's single-valued Associate: PATCH with a relative bind (no leading slash).
    const patch = cs('PATCH', b, {
      headers: { 'If-Match': '*' },
      body: { 'primarycontactid@odata.bind': `contacts(${T1})` },
    });
    expect(patch).toContain(
      `["primarycontactid"] = new EntityReference("contact", new Guid("${T1}")),`,
    );
    expect(patch).toContain('service.Update(entity);');
    const clear = cs('PATCH', b, {
      headers: { 'If-Match': '*' },
      body: { 'primarycontactid@odata.bind': null },
    });
    expect(clear).toContain('["primarycontactid"] = null,');
    const put = cs('PUT', built(`/accounts(${ID})/primarycontactid/$ref`, [], { recordId: ID }), {
      body: { '@odata.id': ref(T2) },
    });
    expect(put).toContain(
      `["primarycontactid"] = new EntityReference("contact", new Guid("${T2}")),`,
    );
    const del = cs('DELETE', built(`/accounts(${ID})/primarycontactid/$ref`, [], { recordId: ID }));
    expect(del).toContain('["primarycontactid"] = null,');
  });

  it('notes a guessed relationship name when the metadata is missing', () => {
    const out = cs(
      'POST',
      built(`/widgets(${ID})/new_widget_parts/$ref`, [], {
        entitySet: 'widgets',
        entityLogical: 'new_widget',
      }),
      { body: { '@odata.id': ref(T1) } },
    );
    expect(out).toContain('new Relationship("new_widget_parts")');
    expect(out).toContain("the SDK wants the relationship's schema name");
  });
});

describe('C# SDK — actions and functions', () => {
  it('sends an unbound action as an OrganizationRequest with the JSON types', () => {
    const out = cs('POST', built('/new_DoThing', [], unbound), {
      body: { Name: 'x', Count: 3, Flag: true },
    });
    expect(out).toContain('var request = new OrganizationRequest("new_DoThing")');
    expect(out).toContain('["Name"] = "x",');
    expect(out).toContain('["Count"] = 3,');
    expect(out).toContain('["Flag"] = true,');
    expect(out).toContain('OrganizationResponse response = service.Execute(request);');
    expect(out).toContain("Parameter types follow the request's JSON");
  });

  it('types parameters from the CSDL and passes a bound row as Target', () => {
    ACTIONS.set('new_Approve', {
      name: 'new_Approve',
      kind: 'Action',
      namespace: 'Microsoft.Dynamics.CRM',
      binding: { kind: 'entity', entityType: 'account' },
      parameters: [
        { name: 'ApproverId', type: 'Edm.Guid', required: true },
        { name: 'Amount', type: 'Edm.Decimal', required: false },
        { name: 'When', type: 'Edm.DateTimeOffset', required: false },
        { name: 'Codes', type: 'Collection(Edm.Int32)', required: false },
      ],
      returnType: { kind: 'void', typeName: '' },
      source: 'custom-api',
    });
    const out = cs(
      'POST',
      built(`/accounts(${ID})/Microsoft.Dynamics.CRM.new_Approve`, [], { recordId: ID }),
      { body: { ApproverId: T1, Amount: 10, When: '2024-01-01T00:00:00Z', Codes: [1, 2] } },
    );
    expect(out).toContain(`["ApproverId"] = new Guid("${T1}"),`);
    expect(out).toContain('["Amount"] = 10m,');
    expect(out).toContain('["When"] = DateTime.Parse("2024-01-01T00:00:00Z"');
    expect(out).toContain('["Codes"] = new int[] { 1, 2 },');
    expect(out).toContain(`["Target"] = new EntityReference("account", new Guid("${ID}")),`);
    expect(out).not.toContain('Check: the bound row is passed as Target');
  });

  it('flags a built-in message whose SDK types can differ', () => {
    ACTIONS.set('WinOpportunity', {
      name: 'WinOpportunity',
      kind: 'Action',
      namespace: 'Microsoft.Dynamics.CRM',
      binding: { kind: 'unbound' },
      parameters: [{ name: 'Status', type: 'Edm.Int32', required: true }],
      returnType: { kind: 'void', typeName: '' },
      source: 'oob',
    });
    const out = cs('POST', built('/WinOpportunity', [], unbound), { body: { Status: 3 } });
    expect(out).toContain('["Status"] = 3,');
    expect(out).toContain('typed WinOpportunityRequest');
  });

  it('falls back when a parameter is an entity or object', () => {
    const out = cs('POST', built('/Merge', [], unbound), {
      body: { Target: { '@odata.type': 'Microsoft.Dynamics.CRM.account', accountid: ID } },
    });
    expect(out).toContain('No faithful SDK form');
    expect(out).toContain('service.ExecuteWebRequest(');
    expect(out).toContain('"Merge"');
  });

  it('reads function arguments from the URL and its aliases', () => {
    const fn = cs(
      'GET',
      built(
        '/RetrieveTotalRecordCount(EntityNames=@p1)',
        [['@p1', '["account","contact"]']],
        unbound,
      ),
    );
    expect(fn).toContain('var request = new OrganizationRequest("RetrieveTotalRecordCount")');
    expect(fn).toContain('["EntityNames"] = new string[] { "account", "contact" },');
    const inline = cs('GET', built(`/new_Lookup(Code='O''B',Id=${ID})`, [], unbound));
    expect(inline).toContain('["Code"] = "O\'B",');
    expect(inline).toContain(`["Id"] = new Guid("${ID}"),`);
    const whoAmI = cs('GET', built('/WhoAmI()', [], unbound));
    expect(whoAmI).toContain('var request = new OrganizationRequest("WhoAmI");');
  });

  it('falls back for an enum function argument', () => {
    const out = cs(
      'GET',
      built(
        "/RetrieveAllEntities(EntityFilters=Microsoft.Dynamics.CRM.EntityFilters'Entity',RetrieveAsIfPublished=false)",
        [],
        unbound,
      ),
    );
    expect(out).toContain('an enum or object literal');
  });

  it('passes ExecuteWorkflow the workflow as WorkflowId and EntityId as a Guid', () => {
    const out = cs(
      'POST',
      built(`/workflows(${ID})/Microsoft.Dynamics.CRM.ExecuteWorkflow`, [], {
        entitySet: 'workflows',
        entityLogical: 'workflow',
        recordId: ID,
      }),
      { body: { EntityId: T1 } },
    );
    expect(out).toContain(`["WorkflowId"] = new Guid("${ID}"),`);
    expect(out).toContain(`["EntityId"] = new Guid("${T1}"),`);
    expect(out).not.toContain('["Target"]');
  });

  it('falls back for a collection-valued action body (CreateMultiple)', () => {
    const out = cs('POST', built('/accounts/Microsoft.Dynamics.CRM.CreateMultiple'), {
      body: { Targets: [{ name: 'a' }] },
    });
    expect(out).toContain('No faithful SDK form');
  });
});

describe('C# SDK — no faithful form', () => {
  it('describes the SDK file messages for binary content, without code', () => {
    const out = cs('PATCH', built(`/accounts(${ID})/new_document`, [], { recordId: ID }), {
      rawBody: '<file bytes>',
    });
    expect(out).toContain('InitializeFileBlocksUploadRequest');
    expect(out).not.toContain('ServiceClient(');
  });

  it('follows a nextLink with ExecuteWebRequest on the API-relative path', () => {
    const out = cs('GET', built('/accounts'), {
      isNextLink: true,
      rawNextLink: `https://contoso.crm.dynamics.com${API}/accounts?$skiptoken=abc\n`,
    });
    expect(out).toContain('PagingInfo.PagingCookie');
    expect(out).toContain('"accounts?$skiptoken=abc"');
  });
});

describe('C# SDK — escaping', () => {
  const nasty = `O'Brien "q" \\ C:\\x\n\r\t ${LS} {braces} $x`;

  it('round-trips string values through C# literals', () => {
    const out = cs('POST', built('/accounts'), { body: { name: nasty } });
    const lit = out.match(/\["name"\] = ("(?:[^"\\]|\\.)*"),/)?.[1];
    expect(lit).toBeDefined();
    expect(csUnquote(lit!)).toBe(nasty);
  });

  it('round-trips filter values and keeps comments on one line', () => {
    const out = query([['$filter', `name eq '${nasty.replace(/'/g, "''")}'`]], {
      headers: { 'X-Note': `a\nusing Evil; ${LS}b` },
    });
    const lit = out.match(/ConditionOperator\.Equal, ("(?:[^"\\]|\\.)*")\);/)?.[1];
    expect(csUnquote(lit!)).toBe(nasty);
    const noteLine = out.split('\n').find((l) => l.includes('X-Note'));
    expect(noteLine?.startsWith('// ')).toBe(true);
    expect(out).not.toMatch(/^using Evil/m);
    expect(out).not.toContain(LS);
  });

  it('refuses names that would leave the string', () => {
    const out = query([['$select', 'name") ; Evil(("']]);
    expect(out).toContain('No faithful SDK form');
    expect(out).not.toContain('ColumnSet(');
  });

  it('never carries the typed credential', () => {
    const out = cs('POST', built('/accounts'), {
      headers: { Authorization: 'Bearer SECRET', Cookie: 'a=b' },
      body: { name: 'x' },
    });
    expect(out).not.toContain('SECRET');
    expect(out).not.toContain('a=b');
  });
});

describe('PowerShell (Xrm) — Microsoft.Xrm.Data.PowerShell', () => {
  it('queries with FetchXML through Get-CrmRecordsByFetch', () => {
    const out = ps(
      'GET',
      built('/accounts', [
        ['$select', 'name,_primarycontactid_value'],
        [
          '$filter',
          "statecode eq 0 and (contains(name,'A<&\">') or revenue gt exchangerate) and Microsoft.Dynamics.CRM.In(PropertyName='industrycode',PropertyValues=['1','2']) and Microsoft.Dynamics.CRM.LastXDays(PropertyName='createdon',PropertyValue=7) and creditonhold eq true",
        ],
        ['$orderby', 'name desc'],
        ['$top', '10'],
      ]),
    );
    expect(out).toContain('Import-Module Microsoft.Xrm.Data.PowerShell');
    expect(out).toContain(
      "$conn = Connect-CrmOnline -ServerUrl 'contoso.crm.dynamics.com' -ForceOAuth",
    );
    expect(out).toContain('$fetch = @\'\n<fetch>\n  <entity name="account">');
    expect(out).toContain('<attribute name="primarycontactid" />');
    expect(out).toContain('<condition attribute="statecode" operator="eq" value="0" />');
    expect(out).toContain('<filter type="or">');
    expect(out).toContain(
      '<condition attribute="name" operator="like" value="%A&lt;&amp;&quot;&gt;%" />',
    );
    expect(out).toContain('<condition attribute="revenue" operator="gt" valueof="exchangerate" />');
    expect(out).toContain(
      '<condition attribute="industrycode" operator="in">\n        <value>1</value>',
    );
    expect(out).toContain('<condition attribute="createdon" operator="last-x-days" value="7" />');
    expect(out).toContain('<condition attribute="creditonhold" operator="eq" value="1" />');
    expect(out).toContain('<order attribute="name" descending="true" />');
    // $top is the page size: the cmdlet pages the fetch itself.
    expect(out).toContain('$result = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch -TopCount 10');
    expect(out).not.toContain('top="');
  });

  it("reads a total count with RetrieveMultiple, which the cmdlet doesn't return", () => {
    const out = ps(
      'GET',
      built('/accounts', [
        ['$select', 'name'],
        ['$count', 'true'],
      ]),
    );
    expect(out).toContain('<fetch count="5000" page="1" returntotalrecordcount="true">');
    expect(out).toContain(
      '$conn.RetrieveMultiple([Microsoft.Xrm.Sdk.Query.FetchExpression]::new($fetch))',
    );
    expect(out).toContain('$results.TotalRecordCount');
  });

  it('keeps a value with a line break from ending the here-string', () => {
    const out = ps('GET', built('/accounts', [['$filter', "name eq 'x\n''@\nRemove-Item'"]]));
    const body = out.slice(out.indexOf("@'\n") + 3, out.indexOf("\n'@"));
    expect(body).toContain('value="x&#10;&apos;@&#10;Remove-Item');
    expect(out.match(/^'@$/gm)).toHaveLength(1);
  });

  it('falls back when FetchXML has no operator for the condition', () => {
    const out = ps(
      'GET',
      built('/accounts', [
        [
          '$filter',
          "Microsoft.Dynamics.CRM.EqualRoleBusinessId(PropertyName='owningbusinessunit')",
        ],
      ]),
    );
    expect(out).toContain('No faithful SDK form');
    expect(out).toContain('The PowerShell tab sends this exact Web API request');
    // C# has the ConditionOperator.
    expect(
      cs(
        'GET',
        built('/accounts', [
          [
            '$filter',
            "Microsoft.Dynamics.CRM.EqualRoleBusinessId(PropertyName='owningbusinessunit')",
          ],
        ]),
      ),
    ).toContain('ConditionOperator.EqualRoleBusinessId');
  });

  it('retrieves a row with Get-CrmRecord', () => {
    const out = ps(
      'GET',
      built(`/accounts(${ID})`, [['$select', 'name,revenue']], { recordId: ID }),
    );
    expect(out).toContain(
      `Get-CrmRecord -conn $conn -EntityLogicalName 'account' -Id '${ID}' -Fields 'name', 'revenue'`,
    );
    expect(ps('GET', built(`/accounts(${ID})`, [], { recordId: ID }))).toContain("-Fields '*'");
  });

  it('creates with New-CrmRecord and typed field values', () => {
    const out = ps('POST', built('/accounts'), {
      body: {
        name: "it's",
        revenue: 12.5,
        industrycode: 3,
        exchangerate: 1.25,
        creditonhold: true,
        new_tags: '1,2',
        'primarycontactid@odata.bind': `/contacts(${T1})`,
      },
    });
    expect(out).toContain("'name' = 'it''s'");
    // New-CrmMoney takes a [double]; Money is built from a [decimal].
    expect(out).toContain("'revenue' = [Microsoft.Xrm.Sdk.Money]::new([decimal]'12.5')");
    expect(out).toContain("'industrycode' = (New-CrmOptionSetValue -Value 3)");
    expect(out).toContain("'exchangerate' = [decimal]'1.25'");
    expect(out).toContain("'creditonhold' = $true");
    expect(out).toContain(
      "'new_tags' = [Microsoft.Xrm.Sdk.OptionSetValueCollection]::new([Microsoft.Xrm.Sdk.OptionSetValue[]]@((New-CrmOptionSetValue -Value 1), (New-CrmOptionSetValue -Value 2)))",
    );
    expect(out).toContain(
      `'primarycontactid' = (New-CrmEntityReference -EntityLogicalName 'contact' -Id '${T1}')`,
    );
    expect(out).toContain(
      "$id = New-CrmRecord -conn $conn -EntityLogicalName 'account' -Fields $fields",
    );
  });

  it('updates with Set-CrmRecord, or an UpdateRequest for an ETag', () => {
    const b = built(`/accounts(${ID})`, [], { recordId: ID });
    expect(ps('PATCH', b, { headers: { 'If-Match': '*' }, body: { name: 'x' } })).toContain(
      `Set-CrmRecord -conn $conn -EntityLogicalName 'account' -Id '${ID}' -Fields $fields`,
    );
    const etag = ps('PATCH', b, { headers: { 'If-Match': 'W/"9"' }, body: { name: 'x' } });
    expect(etag).toContain("$entity.RowVersion = '9'");
    expect(etag).toContain('[Microsoft.Xrm.Sdk.ConcurrencyBehavior]::IfRowVersionMatches');
    expect(etag).toContain('$response = $conn.Execute($request)');
  });

  it('upserts and passes bypass parameters through SDK request objects', () => {
    const out = ps('PATCH', built(`/accounts(accountnumber='A1')`), {
      headers: { 'MSCRM.BypassBusinessLogicExecution': 'CustomAsync', CallerObjectId: T1 },
      body: { name: 'x' },
    });
    expect(out).toContain("$entity.KeyAttributes['accountnumber'] = 'A1'");
    expect(out).toContain('New-Object Microsoft.Xrm.Sdk.Messages.UpsertRequest');
    expect(out).toContain("$request.Parameters['BypassBusinessLogicExecution'] = 'CustomAsync'");
    expect(out).toContain(`$conn.CallerAADObjectId = [guid]'${T1}'`);
  });

  it('deletes with Remove-CrmRecord, or a DeleteRequest by alternate key', () => {
    expect(ps('DELETE', built(`/accounts(${ID})`, [], { recordId: ID }))).toContain(
      `Remove-CrmRecord -conn $conn -EntityLogicalName 'account' -Id '${ID}'`,
    );
    const key = ps('DELETE', built(`/accounts(accountnumber='A1')`));
    expect(key).toContain("$target.KeyAttributes['accountnumber'] = 'A1'");
    expect(key).toContain('New-Object Microsoft.Xrm.Sdk.Messages.DeleteRequest');
  });

  it('associates with Add-CrmRecordAssociation per row, or one request for All or nothing', () => {
    const url = `/accounts(${ID})/contact_customer_accounts/$ref`;
    const multiRequests = [T1, T2].map((t) => ({
      method: 'POST',
      relativeUrl: `${API}${url}`,
      body: { '@odata.id': `https://contoso.crm.dynamics.com${API}/contacts(${t})` },
    }));
    const each = ps('POST', built(url), { multiRequests });
    expect(each).toContain(
      `Add-CrmRecordAssociation -conn $conn -EntityLogicalName1 'account' -Id1 '${ID}' -EntityLogicalName2 'contact' -Id2 '${T2}' -RelationshipName 'contact_customer_accounts'`,
    );
    const atomic = ps('POST', built(url), { multiRequests, changeSet: true });
    expect(atomic).toContain('New-Object Microsoft.Xrm.Sdk.Messages.AssociateRequest');
    expect(atomic).toContain('# One request: Dataverse applies all of them or none.');
    const del = ps('DELETE', built(`/accounts(${ID})/contact_customer_accounts(${T1})/$ref`));
    expect(del).toContain('Remove-CrmRecordAssociation -conn $conn');
  });

  it('runs a message with Invoke-CrmAction', () => {
    const out = ps(
      'POST',
      built(`/accounts(${ID})/Microsoft.Dynamics.CRM.new_Ping`, [], { recordId: ID }),
      { body: { Note: 'hi', At: '2024-01-01' } },
    );
    expect(out).toContain("'Note' = 'hi'");
    expect(out).toContain(
      `$response = Invoke-CrmAction -conn $conn -Name 'new_Ping' -Parameters $parameters -Target (New-CrmEntityReference -EntityLogicalName 'account' -Id '${ID}')`,
    );
  });

  it('points at the PowerShell tab when there is no faithful form', () => {
    const out = ps('GET', built('/accounts', [['$expand', 'primarycontactid']]));
    expect(out).toContain('# No faithful SDK form for this request');
    expect(out).not.toContain('Connect-CrmOnline');
  });

  it('round-trips values through PowerShell literals, curly quotes included', () => {
    const value = `a'b${String.fromCharCode(0x2019)}c${String.fromCharCode(0x2018)} $env:X \`n "d"`;
    const out = ps('POST', built('/accounts'), { body: { name: value } });
    const line = out.split('\n').find((l) => l.startsWith("  'name' = "))!;
    expect(psUnquote(line.slice("  'name' = ".length))).toBe(value);
    expect(psUnquote(psQuote(value))).toBe(value);
  });

  it('never carries the typed credential', () => {
    const out = ps('GET', built('/accounts', [['$select', 'name']]), {
      headers: { Authorization: 'Bearer SECRET' },
    });
    expect(out).not.toContain('SECRET');
  });
});

describe('review fixes', () => {
  const ref = (set: string, t: string) => `https://contoso.crm.dynamics.com${API}/${set}(${t})`;

  it('treats /BulkDelete and /Merge as messages, never a Create on the table', () => {
    // Bulk Delete and Merge carry the table in entitySet but post to the message.
    const empty = cs('POST', built('/BulkDelete'), { body: {} });
    expect(empty).toContain('new OrganizationRequest("BulkDelete")');
    expect(empty).not.toContain('service.Create');
    expect(ps('POST', built('/BulkDelete'), { body: {} })).not.toContain('New-CrmRecord');
    const job = cs('POST', built('/BulkDelete'), {
      body: { QuerySet: [{ EntityName: 'account' }], JobName: 'x' },
    });
    expect(job).toContain('No faithful SDK form');
    expect(job).toContain('"BulkDelete"');
    const merge = cs('POST', built('/Merge'), {
      body: { Target: { '@odata.type': 'Microsoft.Dynamics.CRM.account', accountid: ID } },
    });
    expect(merge).toContain("Parameter 'Target'");
  });

  it('makes the source row the parent of a self-referencing 1:N relationship', () => {
    const b = built(`/accounts(${ID})/Referencedaccount_parent_account/$ref`);
    const body = { '@odata.id': ref('accounts', T1) };
    expect(cs('POST', b, { body })).toContain(
      'relationship.PrimaryEntityRole = EntityRole.Referenced;',
    );
    expect(ps('POST', b, { body })).toContain('[Microsoft.Xrm.Sdk.EntityRole]::Referenced ');
    // Without the relationship's metadata the side isn't known: fall back.
    const unknown = cs(
      'POST',
      built(`/widgets(${ID})/new_widget_widget/$ref`, [], {
        entitySet: 'widgets',
        entityLogical: 'new_widget',
      }),
      { body: { '@odata.id': ref('new_widgets', T1) } },
    );
    expect(unknown).toContain('which side the row is on');
  });

  it('loops per row through SDK requests for a self-referencing relationship in PowerShell', () => {
    const url = `/accounts(${ID})/new_account_account/$ref`;
    const out = ps('POST', built(url), {
      multiRequests: [T1, T2].map((t) => ({
        method: 'POST',
        relativeUrl: `${API}${url}`,
        body: { '@odata.id': ref('accounts', t) },
      })),
    });
    expect(out).toContain('foreach ($t in $targets)');
    expect(out).toContain('[Microsoft.Xrm.Sdk.EntityRole]::Referencing');
    expect(out).not.toContain('Add-CrmRecordAssociation');
  });

  it('keeps $top in the PowerShell fetch sent with optional parameters', () => {
    const out = ps(
      'GET',
      built('/accounts', [
        ['$select', 'name'],
        ['$top', '5'],
      ]),
      {
        headers: { 'MSCRM.BypassBusinessLogicExecution': 'CustomSync' },
      },
    );
    expect(out).toContain('<fetch top="5">');
    expect(out).toContain('New-Object Microsoft.Xrm.Sdk.Messages.RetrieveMultipleRequest');
    expect(out).toContain("$request.Parameters['BypassBusinessLogicExecution'] = 'CustomSync'");
  });

  it('runs a PowerShell action with a bypass header through $conn.Execute', () => {
    const out = ps('POST', built('/new_DoThing', [], unbound), {
      headers: { 'MSCRM.BypassCustomPluginExecution': 'true' },
      body: { Name: 'x' },
    });
    expect(out).toContain(
      "$request = New-Object Microsoft.Xrm.Sdk.OrganizationRequest -ArgumentList 'new_DoThing'",
    );
    expect(out).toContain("$request.Parameters['BypassCustomPluginExecution'] = $true");
    expect(out).toContain('$response = $conn.Execute($request)');
  });

  it('reads Count only in PowerShell from a one-row page', () => {
    const out = ps(
      'GET',
      built('/accounts', [
        ['$top', '1'],
        ['$count', 'true'],
      ]),
      {
        headers: {
          Prefer:
            'odata.include-annotations="Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded"',
        },
      },
    );
    expect(out).toContain('<fetch count="1" page="1" returntotalrecordcount="true">');
    expect(out).toContain('$results.TotalRecordCount');
    expect(out).not.toContain('foreach ($row');
  });

  it("doesn't read a lookup's navigation segment as its column", () => {
    const b = built(`/accounts(${ID})/primarycontactid`, [], { recordId: ID });
    expect(cs('GET', b)).toContain('No faithful SDK form');
    expect(cs('DELETE', b)).toContain('No faithful SDK form');
  });

  it('reads /<column>/$value of a plain column with Retrieve', () => {
    const out = cs('GET', built(`/accounts(${ID})/name/$value`, [], { recordId: ID }));
    expect(out).toContain('new ColumnSet("name")');
    expect(out).not.toContain('InitializeFileBlocks');
  });

  it('keeps function number literals exactly', () => {
    const out = cs(
      'GET',
      built(
        '/CalcIt(Amount=12345678901234567.891,Big=12345678901234567890,Two=2.0,N=7)',
        [],
        unbound,
      ),
    );
    expect(out).toContain('["Amount"] = 12345678901234567.891m,');
    // Past a long's range, the whole number stays exact as a decimal.
    expect(out).toContain('["Big"] = 12345678901234567890m,');
    expect(out).toContain('["Two"] = 2.0m,');
    expect(out).toContain('["N"] = 7,');
  });

  it('refuses a fraction for an Int32 / Int64 parameter', () => {
    ACTIONS.set('new_Count', {
      name: 'new_Count',
      kind: 'Action',
      namespace: 'Microsoft.Dynamics.CRM',
      binding: { kind: 'unbound' },
      parameters: [
        { name: 'Big', type: 'Edm.Int64', required: true },
        { name: 'N', type: 'Edm.Int32', required: true },
      ],
      returnType: { kind: 'void', typeName: '' },
      source: 'custom-api',
    });
    const b = built('/new_Count', [], unbound);
    expect(cs('POST', b, { body: { Big: 1.5, N: 2 } })).toContain('64-bit whole number');
    expect(cs('POST', b, { body: { Big: 3, N: 2.5 } })).toContain('32-bit whole number');
    const ok = cs('POST', b, { body: { Big: 3, N: 2 } });
    expect(ok).toContain('["Big"] = 3L,');
    expect(ok).toContain('["N"] = 2,');
  });

  it('falls back for an alternate key that uses a lookup', () => {
    const out = cs('PATCH', built(`/accounts(_primarycontactid_value=${T1},name='a')`), {
      body: { name: 'x' },
    });
    expect(out).toContain("uses the lookup '_primarycontactid_value'");
  });

  it('keeps a leading-zero query-function value as text without metadata', () => {
    const out = cs(
      'GET',
      built(
        '/widgets',
        [
          [
            '$filter',
            "Microsoft.Dynamics.CRM.In(PropertyName='code',PropertyValues=['00123','7'])",
          ],
        ],
        { entitySet: 'widgets', entityLogical: 'new_widget' },
      ),
    );
    expect(out).toContain('AddCondition("code", ConditionOperator.In, "00123", 7);');
  });

  it('falls back in PowerShell for a control character FetchXML cannot carry', () => {
    const bell = String.fromCharCode(7);
    const b = built('/accounts', [['$filter', `name eq 'a${bell}b'`]]);
    expect(ps('GET', b)).toContain('FetchXML (XML 1.0)');
    expect(cs('GET', b)).toContain('ConditionOperator.Equal, "a\\u0007b"');
  });

  it('disassociates through ?$id=', () => {
    const out = cs(
      'DELETE',
      built(`/accounts(${ID})/contact_customer_accounts/$ref`, [['$id', ref('contacts', T1)]]),
    );
    expect(out).toContain(
      `service.Disassociate("account", new Guid("${ID}"), relationship, related);`,
    );
    expect(out).toContain(`new EntityReference("contact", new Guid("${T1}")),`);
  });

  it('shows a comment, not a program, until the request is complete', () => {
    const out = cs('POST', built('', [], { ...unbound, relativeUrl: '', relativeNoBase: '' }));
    expect(out).toContain("The request isn't complete yet.");
    expect(out).not.toContain('ServiceClient(');
  });
});

describe('helpers', () => {
  it('parseFilterForCodegen marks quoted literals', () => {
    const r = parseFilterForCodegen("a eq '5' and b eq 5");
    const [qa, qb] = r.group.rules as { id: string }[];
    expect(r.quoted.has(qa.id)).toBe(true);
    expect(r.quoted.has(qb.id)).toBe(false);
  });

  it('xmlEscape escapes markup, quotes and control characters', () => {
    expect(xmlEscape(`<a b="c" d='e'>&\n`)).toBe(
      '&lt;a b=&quot;c&quot; d=&apos;e&apos;&gt;&amp;&#10;',
    );
  });
});
