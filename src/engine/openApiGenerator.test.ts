import { beforeEach, describe, expect, it, vi } from 'vitest';
import { columnSchema, generateOpenApi } from './openApiGenerator';
import type { CodegenInputs } from './codeGenerators';
import type { BuiltRequest } from './urlBuilder';
import type { ColumnMeta } from '../mock/metadata';
import { ACCOUNT_ID, CONTACT_ID, account, registerFixtureTables } from '../test/fixtures';

vi.mock('../mock/environment', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../mock/environment')>()),
  ENV: {
    name: 'Test',
    host: 'contoso.crm.dynamics.com',
    apiBase: '/api/data/v9.2',
    user: '',
  },
}));

const API = '/api/data/v9.2';

/** A BuiltRequest for `<API><path>` with the given encoded query parts. */
function built(
  path: string,
  queryParts: { key: string; value: string }[] = [],
  extra: Partial<BuiltRequest> = {},
): BuiltRequest {
  const qs = queryParts.map((p) => `${p.key}=${p.value}`).join('&');
  const noBase = qs ? `${path}?${qs}` : path;
  return {
    relativeUrl: `${API}${noBase}`,
    relativeNoBase: noBase,
    bytes: 0,
    queryParts,
    entitySet: 'accounts',
    entityLogical: 'account',
    ...extra,
  };
}

function inputs(
  method: string,
  b: BuiltRequest,
  extra: Partial<CodegenInputs> = {},
): CodegenInputs {
  return { method, built: b, headers: {}, ...extra };
}

type Doc = any;
const doc = (i: CodegenInputs): Doc => JSON.parse(generateOpenApi(i));
const col = (name: string) => account.columns.find((c) => c.logicalName === name)!;

beforeEach(() => registerFixtureTables());

describe('columnSchema', () => {
  it('maps Edm types to JSON Schema', () => {
    expect(columnSchema(col('name'))).toMatchObject({
      title: 'Name',
      type: 'string',
      maxLength: 160,
    });
    expect(columnSchema(col('numberofemployees'))).toMatchObject({
      type: 'integer',
      format: 'int32',
    });
    expect(columnSchema(col('revenue'))).toMatchObject({ type: 'number', format: 'decimal' });
    expect(columnSchema(col('donotemail'))).toMatchObject({ type: 'boolean' });
    expect(columnSchema(col('createdon'))).toMatchObject({ type: 'string', format: 'date-time' });
    expect(columnSchema(col('primarycontactid'))).toMatchObject({
      type: 'string',
      format: 'uuid',
    });
  });

  it('maps a DateOnly column to format: date', () => {
    const birthday: ColumnMeta = {
      logicalName: 'birthdate',
      displayName: 'Birthday',
      attributeType: 'DateTime',
      format: 'DateOnly',
      dateTimeBehavior: 'DateOnly',
    };
    expect(columnSchema(birthday)).toMatchObject({ type: 'string', format: 'date' });
  });

  it('lists choice values as an enum with their labels', () => {
    expect(columnSchema(col('industrycode'))).toMatchObject({
      type: 'integer',
      enum: [1, 2],
      'x-enum-labels': { '1': 'Accounting', '2': 'Agriculture' },
    });
    // Multi-select choices travel as a comma-separated string.
    const tags = columnSchema(col('sample_tags'));
    expect(tags.type).toBe('string');
    expect(String(tags.description)).toContain('100 = A');
  });
});

describe('generateOpenApi', () => {
  it('describes an entity-set GET: server, path, query params, row schema', () => {
    const d = doc(
      inputs(
        'GET',
        built('/accounts', [
          { key: '$select', value: 'name,revenue' },
          { key: '$filter', value: 'statecode%20eq%200' },
          { key: '$top', value: '5' },
        ]),
      ),
    );
    expect(d.openapi).toBe('3.0.3');
    expect(d.info.title).toBe('Dataverse Web API: Account');
    expect(d.servers).toEqual([{ url: 'https://contoso.crm.dynamics.com/api/data/v9.2' }]);
    const op = d.paths['/accounts'].get;
    expect(op.tags).toEqual(['account']);
    // Query options become query parameters with the decoded built value as example.
    expect(op.parameters).toEqual([
      expect.objectContaining({ name: '$select', in: 'query', example: 'name,revenue' }),
      expect.objectContaining({ name: '$filter', in: 'query', example: 'statecode eq 0' }),
      expect.objectContaining({ name: '$top', in: 'query', example: '5', required: false }),
    ]);
    const ok = op.responses['200'].content['application/json'].schema;
    expect(ok.properties.value.items).toEqual({ $ref: '#/components/schemas/account' });
    expect(op.responses.default.content['application/json'].schema).toEqual({
      $ref: '#/components/schemas/ODataError',
    });
    // $select narrows the row schema; the primary key and etag always come back.
    expect(Object.keys(d.components.schemas.account.properties).sort()).toEqual(
      ['@odata.etag', 'accountid', 'name', 'revenue'].sort(),
    );
  });

  it('lists every readable column when there is no $select, lookups as _x_value', () => {
    const props = doc(inputs('GET', built('/accounts'))).components.schemas.account.properties;
    expect(props).toHaveProperty('_primarycontactid_value');
    expect(props).not.toHaveProperty('primarycontactid');
    expect(props).toHaveProperty('industrycode');
  });

  it('templates a record GUID as a path parameter and returns the row', () => {
    const d = doc(inputs('GET', built(`/accounts(${ACCOUNT_ID})`)));
    const op = d.paths['/accounts({accountid})'].get;
    expect(op.parameters[0]).toMatchObject({
      name: 'accountid',
      in: 'path',
      required: true,
      schema: { type: 'string', format: 'uuid' },
      example: ACCOUNT_ID,
    });
    expect(op.responses['200'].content['application/json'].schema).toEqual({
      $ref: '#/components/schemas/account',
    });
  });

  it('types a Create body from metadata and returns 201 with return=representation', () => {
    const body = {
      name: 'Contoso',
      numberofemployees: 10,
      'primarycontactid@odata.bind': `/contacts(${CONTACT_ID})`,
      sample_custom: true,
    };
    const d = doc(
      inputs('POST', built('/accounts'), {
        body,
        headers: { Prefer: 'return=representation', Authorization: 'Bearer x' },
      }),
    );
    const op = d.paths['/accounts'].post;
    const rb = op.requestBody.content['application/json'];
    expect(rb.example).toEqual(body);
    expect(rb.schema.properties.name).toMatchObject({ type: 'string', maxLength: 160 });
    expect(rb.schema.properties.numberofemployees).toMatchObject({ type: 'integer' });
    expect(rb.schema.properties['primarycontactid@odata.bind']).toMatchObject({
      type: 'string',
      example: `/contacts(${CONTACT_ID})`,
    });
    // Unknown keys are inferred from the value.
    expect(rb.schema.properties.sample_custom).toEqual({ type: 'boolean' });
    expect(op.responses['201'].content['application/json'].schema).toEqual({
      $ref: '#/components/schemas/account',
    });
    // Prefer is documented as a header; Authorization is covered by the security schemes.
    const headers = op.parameters.filter((p: Doc) => p.in === 'header').map((p: Doc) => p.name);
    expect(headers).toEqual(['Prefer']);
  });

  it('returns 204 with OData-EntityId for a Create without return=representation', () => {
    const op = doc(inputs('POST', built('/accounts'), { body: { name: 'x' } })).paths['/accounts']
      .post;
    expect(op.responses['204'].headers['OData-EntityId']).toBeDefined();
    expect(op.responses['201']).toBeUndefined();
  });

  it('documents an If-Match header and a 204 for a PATCH', () => {
    const op = doc(
      inputs('PATCH', built(`/accounts(${ACCOUNT_ID})`), {
        body: { name: 'y' },
        headers: { 'If-Match': '*' },
      }),
    ).paths['/accounts({accountid})'].patch;
    expect(op.parameters).toContainEqual(
      expect.objectContaining({ name: 'If-Match', in: 'header', example: '*' }),
    );
    expect(Object.keys(op.responses).sort()).toEqual(['204', 'default']);
  });

  it('marks function parameter aliases as required', () => {
    const d = doc(
      inputs(
        'GET',
        built('/WhoAmI()', [{ key: '@p1', value: "'x'" }], { entitySet: '', entityLogical: '' }),
      ),
    );
    const op = d.paths['/WhoAmI()'].get;
    expect(op.parameters).toEqual([
      expect.objectContaining({ name: '@p1', in: 'query', required: true, example: "'x'" }),
    ]);
    expect(d.info.title).toBe('Dataverse Web API: request');
    expect(d.components.schemas).not.toHaveProperty('account');
  });

  it('merges Associate targets on the same URL into one operation; numbers distinct ones', () => {
    const same = doc(
      inputs('POST', built(`/accounts(${ACCOUNT_ID})/contact_customer_accounts/$ref`), {
        multiRequests: [
          {
            method: 'POST',
            relativeUrl: `${API}/accounts(${ACCOUNT_ID})/contact_customer_accounts/$ref`,
            body: { '@odata.id': 'a' },
          },
          {
            method: 'POST',
            relativeUrl: `${API}/accounts(${ACCOUNT_ID})/contact_customer_accounts/$ref`,
            body: { '@odata.id': 'b' },
          },
        ],
      }),
    );
    expect(Object.keys(same.paths)).toEqual([
      '/accounts({accountid})/contact_customer_accounts/$ref',
    ]);
    expect(
      same.paths['/accounts({accountid})/contact_customer_accounts/$ref'].post.responses,
    ).toHaveProperty('204');

    const two = doc(
      inputs('DELETE', built(`/accounts(${ACCOUNT_ID})`), {
        multiRequests: [
          { method: 'DELETE', relativeUrl: `${API}/accounts(${ACCOUNT_ID})/a/$ref` },
          { method: 'DELETE', relativeUrl: `${API}/accounts(${ACCOUNT_ID})/b/$ref` },
        ],
      }),
    );
    const ids = Object.values(two.paths).map((p: Doc) => p.delete.operationId);
    expect(ids).toEqual(['delete_1_accounts', 'delete_2_accounts']);
  });

  it('declares bearer and Entra ID OAuth2 security for the org', () => {
    const d = doc(inputs('GET', built('/accounts')));
    expect(d.security).toEqual([
      { bearerAuth: [] },
      { entraId: ['https://contoso.crm.dynamics.com/user_impersonation'] },
    ]);
    expect(d.components.securitySchemes.bearerAuth).toMatchObject({
      type: 'http',
      scheme: 'bearer',
    });
    expect(d.components.securitySchemes.entraId.flows.authorizationCode.scopes).toHaveProperty(
      'https://contoso.crm.dynamics.com/user_impersonation',
    );
  });

  it('uses the raw nextLink for Retrieve NextLink', () => {
    const d = doc(
      inputs('GET', built('/accounts'), {
        isNextLink: true,
        rawNextLink: `https://contoso.crm.dynamics.com${API}/accounts?$skiptoken=abc`,
      }),
    );
    expect(d.paths['/accounts'].get.parameters).toEqual([
      expect.objectContaining({ name: '$skiptoken', example: 'abc' }),
    ]);
  });
});
