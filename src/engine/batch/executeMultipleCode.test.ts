import { describe, expect, it } from 'vitest';
import { compileBatch, newGroup, newItem } from './batchModel';
import { generateExecuteMultipleCSharp, type CodeMetadata } from './executeMultipleCode';

const md: CodeMetadata = {
  logicalOf: (set) => ({ accounts: 'account', contacts: 'contact' })[set] ?? null,
  attrType: (logical, attr) =>
    (
      ({
        account: {
          name: 'String',
          industrycode: 'Picklist',
          revenue: 'Money',
          primarycontactid: 'Lookup',
          customerid: 'Customer',
        },
      }) as Record<string, Record<string, string>>
    )[logical]?.[attr] ?? null,
};

const ID = '11111111-2222-3333-4444-555555555555';

function code(
  items: ReturnType<typeof newItem>[],
  opts = { continueOnError: true, returnResponses: false },
) {
  return generateExecuteMultipleCSharp(compileBatch([newGroup(false, items)]).items, opts, md);
}

describe('generateExecuteMultipleCSharp', () => {
  it('maps create with typed columns and a lookup bind', () => {
    const out = code([
      newItem('create', {
        entitySet: 'accounts',
        body: JSON.stringify({
          name: 'A',
          industrycode: 3,
          revenue: 10.5,
          'primarycontactid@odata.bind': `contacts(${ID})`,
          '@odata.type': 'Microsoft.Dynamics.CRM.account',
        }),
      }),
    ]);
    expect(out).toContain('requests.Add(new CreateRequest');
    expect(out).toContain('Target = new Entity("account")');
    expect(out).toContain('["name"] = "A",');
    expect(out).toContain('["industrycode"] = new OptionSetValue(3),');
    expect(out).toContain('["revenue"] = new Money(10.5m),');
    expect(out).toContain(
      `["primarycontactid"] = new EntityReference("contact", new Guid("${ID}")),`,
    );
    expect(out).not.toContain('@odata.type');
    expect(out).toContain('ContinueOnError = true,');
    expect(out).toContain('ReturnResponses = false,');
  });

  it('maps update, delete, retrieve and associate', () => {
    const out = code([
      newItem('update', { entitySet: 'accounts', recordId: ID, body: '{"name":"B"}' }),
      newItem('delete', { entitySet: 'contacts', recordId: `{${ID}}` }),
      newItem('retrieve', {
        entitySet: 'accounts',
        recordId: ID,
        query: '$select=name,_primarycontactid_value',
      }),
      newItem('retrieve', {
        entitySet: 'contacts',
        query: '$select=fullname&$top=5&$filter=statecode eq 0',
      }),
      newItem('associate', {
        entitySet: 'accounts',
        recordId: ID,
        navProperty: 'contact_customer_accounts',
        targetSet: 'contacts',
        targetId: ID,
      }),
    ]);
    expect(out).toContain(`Target = new Entity("account", new Guid("${ID}"))`);
    expect(out).toContain(
      `new DeleteRequest { Target = new EntityReference("contact", new Guid("${ID}")) }`,
    );
    expect(out).toContain('ColumnSet = new ColumnSet("name", "primarycontactid"),');
    expect(out).toContain('Query = new QueryExpression("contact")');
    expect(out).toContain('TopCount = 5,');
    expect(out).toContain('// TODO translate $filter');
    expect(out).toContain('Relationship = new Relationship("contact_customer_accounts"),');
  });

  it('turns an unbound action into an OrganizationRequest and flags what it can’t map', () => {
    const out = code([
      newItem('custom', { method: 'POST', url: 'WinOpportunity', body: '{"Status":3}' }),
      newItem('custom', { method: 'GET', url: 'WhoAmI' }),
      newItem('create', {
        entitySet: 'accounts',
        body: '{"customerid_account@odata.bind":"$1","x":1}',
      }),
    ]);
    expect(out).toContain('requests.Add(new OrganizationRequest("WinOpportunity")');
    expect(out).toContain('["Status"] = 3,');
    expect(out).toContain('// No SDK request generated for this Web API call');
    expect(out).toContain('// TODO customerid: $1 referenced an earlier request');
    expect(out).toContain('// TODO x: confirm the type');
  });
});
