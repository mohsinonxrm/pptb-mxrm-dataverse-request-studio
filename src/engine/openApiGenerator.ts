// OpenAPI 3.0 document for the request being built (#28).
//
// Describes exactly the request(s) on screen — path + method, the query
// options as query parameters (with the built values as examples), custom
// headers, the JSON body, and the response — against the org's Web API
// base URL. Request and response schemas come from the table's column
// metadata: each column maps to a JSON Schema type (Edm → OpenAPI), lookups
// to `_<name>_value` GUIDs in responses and `<nav>@odata.bind` strings in
// bodies. Auth is Microsoft Entra ID: a bearer token, or the OAuth2
// authorization-code flow with the org's `user_impersonation` scope.
//
// 3.0.3 rather than 3.1: it is what most generators and API tools import.

import { API_VERSION, ENV } from '../mock/environment';
import { findTable, type ColumnMeta, type TableMeta } from '../mock/metadata';
import type { CodegenInputs } from './codeGenerators';
import { impersonationVar, isCredentialHeader } from './sensitiveHeaders';

type Schema = Record<string, unknown>;

interface OpRequest {
  method: string;
  /** Path + query, no host, may include the API base. */
  relativeUrl: string;
  body?: Record<string, unknown>;
  description?: string;
}

const GUID_RE =
  /\(([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}|<[^>]+>)\)/g;

/** JSON Schema for one column's value, as the Web API returns / accepts it. */
export function columnSchema(c: ColumnMeta): Schema {
  const base: Schema =
    c.displayName && c.displayName !== c.logicalName ? { title: c.displayName } : {};
  if (c.description) base.description = c.description;
  switch (c.attributeType) {
    case 'String':
    case 'Memo':
      return { ...base, type: 'string', ...(c.maxLength ? { maxLength: c.maxLength } : {}) };
    case 'Integer':
      return {
        ...base,
        type: 'integer',
        format: 'int32',
        ...(c.minValue != null ? { minimum: c.minValue } : {}),
        ...(c.maxValue != null ? { maximum: c.maxValue } : {}),
      };
    case 'BigInt':
      return { ...base, type: 'integer', format: 'int64' };
    case 'Decimal':
    case 'Money':
      return { ...base, type: 'number', format: 'decimal' };
    case 'Double':
      return { ...base, type: 'number', format: 'double' };
    case 'Boolean':
      return { ...base, type: 'boolean' };
    case 'DateTime':
      return { ...base, type: 'string', format: c.format === 'DateOnly' ? 'date' : 'date-time' };
    case 'Picklist':
    case 'State':
    case 'Status':
      return {
        ...base,
        type: 'integer',
        enum: c.options.map((o) => o.value),
        'x-enum-labels': Object.fromEntries(c.options.map((o) => [String(o.value), o.label])),
      };
    case 'MultiSelectPicklist':
      return {
        ...base,
        type: 'string',
        description: `Comma-separated option values: ${c.options.map((o) => `${o.value} = ${o.label}`).join(', ')}`,
      };
    case 'Uniqueidentifier':
    case 'Lookup':
    case 'Customer':
    case 'Owner':
      return { ...base, type: 'string', format: 'uuid' };
    case 'Image':
      return { ...base, type: 'string', format: 'byte' };
    case 'File':
      return { ...base, type: 'string', format: 'uuid', description: 'File id' };
    default:
      return { ...base, type: 'string' };
  }
}

function isLookup(c: ColumnMeta): boolean {
  return (
    c.attributeType === 'Lookup' || c.attributeType === 'Customer' || c.attributeType === 'Owner'
  );
}

/** Response property name: `_<name>_value` for lookups, else the logical name. */
function oDataName(c: ColumnMeta): string {
  return c.oDataName ?? (isLookup(c) ? `_${c.logicalName}_value` : c.logicalName);
}

/** Entity schema for responses — every readable column, or just `$select`ed ones. */
function entitySchema(t: TableMeta, select: string[] | null): Schema {
  const props: Record<string, Schema> = {
    '@odata.etag': { type: 'string', readOnly: true },
  };
  for (const c of t.columns) {
    if (c.isValidForRead === false || c.isValidOData === false || c.isLogical) continue;
    const name = oDataName(c);
    if (select && !select.includes(name) && !select.includes(c.logicalName)) continue;
    props[name] = columnSchema(c);
  }
  // The primary key always comes back, even when not selected.
  const pk = t.columns.find((c) => c.logicalName === t.primaryKey);
  if (pk && !props[pk.logicalName]) props[pk.logicalName] = columnSchema(pk);
  return {
    type: 'object',
    title: t.displayName,
    properties: props,
    additionalProperties: true,
  };
}

/** Request-body schema: known columns typed from metadata, binds as strings. */
function bodySchema(body: Record<string, unknown>, t: TableMeta | undefined): Schema {
  const props: Record<string, Schema> = {};
  for (const [k, v] of Object.entries(body)) {
    if (k.endsWith('@odata.bind')) {
      props[k] = {
        type: 'string',
        description: 'Related row: `/<entityset>(<id>)`',
        example: v,
      };
      continue;
    }
    const col = t?.columns.find((c) => c.logicalName === k);
    props[k] = col ? columnSchema(col) : inferSchema(v);
  }
  return { type: 'object', properties: props };
}

function inferSchema(v: unknown): Schema {
  if (v === null) return { nullable: true };
  if (Array.isArray(v)) return { type: 'array', items: v.length ? inferSchema(v[0]) : {} };
  switch (typeof v) {
    case 'string':
      return { type: 'string' };
    case 'number':
      return Number.isInteger(v) ? { type: 'integer' } : { type: 'number' };
    case 'boolean':
      return { type: 'boolean' };
    case 'object': {
      const props: Record<string, Schema> = {};
      for (const [k, x] of Object.entries(v as Record<string, unknown>)) props[k] = inferSchema(x);
      return { type: 'object', properties: props };
    }
    default:
      return {};
  }
}

function safeDecode(v: string): string {
  try {
    return decodeURIComponent(v);
  } catch {
    return v;
  }
}

/** Split `/api/data/v9.2/accounts(<guid>)?$select=x` into a templated path + params. */
function splitRequest(
  relativeUrl: string,
  t: TableMeta | undefined,
): { path: string; pathParams: Schema[]; queryParams: Schema[] } {
  const url = relativeUrl.replace(/^https?:\/\/[^/]+/i, '');
  const q = url.indexOf('?');
  let path = q >= 0 ? url.slice(0, q) : url;
  const qs = q >= 0 ? url.slice(q + 1) : '';
  if (path.startsWith(ENV.apiBase)) path = path.slice(ENV.apiBase.length);
  if (!path.startsWith('/')) path = `/${path}`;
  const pathParams: Schema[] = [];
  path = path.replace(GUID_RE, (_m, id: string) => {
    const name = pathParams.length === 0 ? (t?.primaryKey ?? 'id') : `id${pathParams.length + 1}`;
    pathParams.push({
      name,
      in: 'path',
      required: true,
      schema: { type: 'string', format: 'uuid' },
      ...(id.startsWith('<') ? {} : { example: id }),
    });
    return `({${name}})`;
  });
  const queryParams: Schema[] = qs
    ? qs
        .split('&')
        .filter(Boolean)
        .map((pair) => {
          const eq = pair.indexOf('=');
          const name = safeDecode(eq >= 0 ? pair.slice(0, eq) : pair);
          const value = safeDecode(eq >= 0 ? pair.slice(eq + 1) : '');
          return {
            name,
            in: 'query',
            required: name.startsWith('@'),
            schema: { type: 'string' },
            example: value,
          };
        })
    : [];
  return { path, pathParams, queryParams };
}

// Credential headers (#217) are covered by the security scheme.
const SKIP_HEADERS = new Set(['accept', 'content-type']);

function headerParams(headers: Record<string, string>): Schema[] {
  return Object.entries(headers)
    .filter(([k, v]) => v && !SKIP_HEADERS.has(k.toLowerCase()) && !isCredentialHeader(k))
    .map(([name, value]) =>
      // An impersonation header's value is a user id: not written into a
      // file meant to be shared (#217).
      impersonationVar(name)
        ? {
            name,
            in: 'header',
            required: false,
            description: 'Id of the user to impersonate.',
            schema: { type: 'string', format: 'uuid' },
          }
        : { name, in: 'header', required: false, schema: { type: 'string' }, example: value },
    );
}

const ERROR_RESPONSE = {
  description: 'Error',
  content: { 'application/json': { schema: { $ref: '#/components/schemas/ODataError' } } },
};

function successResponse(
  op: OpRequest,
  path: string,
  i: CodegenInputs,
  entityRef: string | null,
): Record<string, unknown> {
  const method = op.method.toUpperCase();
  const preferRep = /return=representation/i.test(i.headers['Prefer'] ?? '');
  const set = i.built.entitySet;
  const onSet = !!set && path === `/${set}`;
  const onRow = !!set && new RegExp(`^/${set}\\(\\{[^}]+\\}\\)$`).test(path);
  const json = (schema: Schema, description: string) => ({
    description,
    content: { 'application/json': { schema } },
  });
  if (method === 'GET') {
    if (entityRef && onSet) {
      return {
        '200': json(
          {
            type: 'object',
            properties: {
              '@odata.context': { type: 'string' },
              '@odata.count': { type: 'integer' },
              '@odata.nextLink': { type: 'string' },
              value: { type: 'array', items: { $ref: entityRef } },
            },
          },
          'Rows',
        ),
      };
    }
    if (entityRef && onRow) return { '200': json({ $ref: entityRef }, 'Row') };
    return { '200': json({ type: 'object', additionalProperties: true }, 'OK') };
  }
  if ((method === 'POST' || method === 'PATCH') && preferRep && entityRef && (onSet || onRow)) {
    return { [method === 'POST' ? '201' : '200']: json({ $ref: entityRef }, 'Row as written') };
  }
  if (method === 'POST' && onSet) {
    return {
      '204': {
        description: 'Created',
        headers: {
          'OData-EntityId': {
            description: 'URL of the created row',
            schema: { type: 'string' },
          },
        },
      },
    };
  }
  if (method === 'POST' && !path.endsWith('/$ref')) {
    return {
      '200': json({ type: 'object', additionalProperties: true }, 'OK'),
      '204': { description: 'No content' },
    };
  }
  return { '204': { description: 'No content' } };
}

/** OpenAPI 3.0.3 document (JSON) for the current request. */
export function generateOpenApi(i: CodegenInputs): string {
  const table = i.built.entityLogical ? findTable(i.built.entityLogical) : undefined;
  const ops: OpRequest[] = i.isNextLink
    ? [{ method: 'GET', relativeUrl: (i.rawNextLink ?? '').trim() }]
    : i.multiRequests && i.multiRequests.length > 1
      ? i.multiRequests
      : [{ method: i.method, relativeUrl: i.built.relativeUrl, body: i.body }];

  const schemas: Record<string, Schema> = {
    ODataError: {
      type: 'object',
      properties: {
        error: {
          type: 'object',
          properties: { code: { type: 'string' }, message: { type: 'string' } },
        },
      },
    },
  };
  let entityRef: string | null = null;
  if (table) {
    const select = i.built.queryParts.find((p) => p.key === '$select');
    schemas[table.logicalName] = entitySchema(
      table,
      select ? safeDecode(select.value).split(',').filter(Boolean) : null,
    );
    entityRef = `#/components/schemas/${table.logicalName}`;
  }

  const paths: Record<string, Record<string, unknown>> = {};
  ops.forEach((op, n) => {
    if (!op.relativeUrl || op.relativeUrl.startsWith('<')) return;
    const { path, pathParams, queryParams } = splitRequest(op.relativeUrl, table);
    const method = op.method.toLowerCase();
    const item = (paths[path] ??= {});
    // Several targets on the same URL (Associate) share one operation.
    if (item[method]) return;
    const body = op.body;
    item[method] = {
      operationId: `${method}${ops.length > 1 ? `_${n + 1}` : ''}_${(i.built.entitySet || 'request').replace(/[^A-Za-z0-9]/g, '_')}`,
      summary: op.description ?? `${op.method.toUpperCase()} ${path}`,
      ...(table ? { tags: [table.logicalName] } : {}),
      parameters: [...pathParams, ...queryParams, ...headerParams(i.headers)],
      ...(body
        ? {
            requestBody: {
              required: true,
              content: { 'application/json': { schema: bodySchema(body, table), example: body } },
            },
          }
        : {}),
      responses: { ...successResponse(op, path, i, entityRef), default: ERROR_RESPONSE },
    };
  });

  const host = ENV.host || '<org>.crm.dynamics.com';
  const doc = {
    openapi: '3.0.3',
    info: {
      title: `Dataverse Web API: ${table?.displayName ?? (i.built.entitySet || 'request')}`,
      version: API_VERSION,
      description: 'Generated by Dataverse Request Studio for the request in the builder.',
    },
    servers: [{ url: `https://${host}${ENV.apiBase}` }],
    security: [{ bearerAuth: [] }, { entraId: [`https://${host}/user_impersonation`] }],
    paths,
    components: {
      schemas,
      securitySchemes: {
        bearerAuth: {
          type: 'http',
          scheme: 'bearer',
          bearerFormat: 'JWT',
          description: 'Microsoft Entra ID access token for the Dataverse org.',
        },
        entraId: {
          type: 'oauth2',
          flows: {
            authorizationCode: {
              authorizationUrl: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
              tokenUrl: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
              scopes: {
                [`https://${host}/user_impersonation`]: 'Access Dataverse as the signed-in user',
              },
            },
          },
        },
      },
    },
  };
  return JSON.stringify(doc, null, 2);
}
