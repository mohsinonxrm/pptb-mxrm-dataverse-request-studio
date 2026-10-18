// Request-collection file formats (#27): Postman v2.1, Insomnia v4 and
// Bruno's collection JSON. Pure functions over a neutral request list, so
// the same input yields all three files and the output can be checked
// without a host.
//
// Every URL is written against a `baseUrl` variable (https://<org host>)
// and every request inherits bearer auth from an `accessToken` variable,
// so a collection carries no secrets and can be pointed at another org by
// changing one variable.

/** One HTTP request in an exported collection. */
export interface CollectionRequest {
  name: string;
  method: string;
  /** Path + query, no host — e.g. `/api/data/v9.2/accounts?$select=name`. */
  path: string;
  headers: Record<string, string>;
  /** JSON object for JSON requests, or a text placeholder (binary steps). */
  body?: Record<string, unknown> | string;
  description?: string;
}

/** A top-level entry: a single request, or an ordered folder of steps. */
export type CollectionItem =
  | ({ kind: 'request' } & CollectionRequest)
  | { kind: 'folder'; name: string; description?: string; steps: CollectionRequest[] };

export interface CollectionInput {
  name: string;
  /** `https://<org host>` — becomes the `baseUrl` variable. */
  baseUrl: string;
  items: CollectionItem[];
}

/** Path placeholder for a step whose URL comes from a previous response (chunked upload). */
export const UPLOAD_URL_VAR = '{{uploadUrl}}';

export type CollectionFormat = 'postman' | 'insomnia' | 'bruno';

export const COLLECTION_FORMAT_LABELS: Record<CollectionFormat, string> = {
  postman: 'Postman collection (v2.1)',
  insomnia: 'Insomnia export (v4)',
  bruno: 'Bruno collection (JSON)',
};

export function buildCollection(fmt: CollectionFormat, input: CollectionInput): string {
  switch (fmt) {
    case 'postman':
      return JSON.stringify(toPostman(input), null, 2);
    case 'insomnia':
      return JSON.stringify(toInsomnia(input), null, 2);
    case 'bruno':
      return JSON.stringify(toBruno(input), null, 2);
  }
}

/** File name for a collection export, e.g. `drs-contoso.crm.dynamics.com.postman_collection.json`. */
export function collectionFileName(fmt: CollectionFormat, stem: string): string {
  const safe = stem.replace(/[^a-z0-9._-]+/gi, '_');
  switch (fmt) {
    case 'postman':
      return `${safe}.postman_collection.json`;
    case 'insomnia':
      return `${safe}.insomnia.json`;
    case 'bruno':
      return `${safe}.bruno.json`;
  }
}

// ── helpers ──────────────────────────────────────────────────────────

function splitPath(path: string): { segments: string[]; query: { key: string; value: string }[] } {
  const q = path.indexOf('?');
  const p = q >= 0 ? path.slice(0, q) : path;
  const qs = q >= 0 ? path.slice(q + 1) : '';
  const segments = p.split('/').filter(Boolean);
  const query = qs
    ? qs.split('&').map((pair) => {
        const eq = pair.indexOf('=');
        return eq >= 0
          ? { key: pair.slice(0, eq), value: pair.slice(eq + 1) }
          : { key: pair, value: '' };
      })
    : [];
  return { segments, query };
}

const usesUploadUrl = (items: CollectionItem[]) =>
  flatSteps(items).some((r) => r.path === UPLOAD_URL_VAR);

function bodyText(body: CollectionRequest['body']): string | undefined {
  if (body == null) return undefined;
  return typeof body === 'string' ? body : JSON.stringify(body, null, 2);
}

function isJsonBody(body: CollectionRequest['body']): boolean {
  return body != null && typeof body === 'object';
}

function flatSteps(items: CollectionItem[]): CollectionRequest[] {
  return items.flatMap((it) => (it.kind === 'folder' ? it.steps : [it]));
}

let idCounter = 0;
function rid(prefix: string): string {
  idCounter += 1;
  const rand =
    typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
      ? crypto.randomUUID().replace(/-/g, '')
      : `${Date.now().toString(16)}${idCounter.toString(16)}`;
  return `${prefix}_${rand.slice(0, 24)}`;
}

function uuid(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  const h = (n: number) =>
    Array.from({ length: n }, () => Math.floor(Math.random() * 16).toString(16)).join('');
  return `${h(8)}-${h(4)}-4${h(3)}-a${h(3)}-${h(12)}`;
}

// ── Postman v2.1 ─────────────────────────────────────────────────────
// https://schema.postman.com/collection/json/v2.1.0/draft-07/collection.json

function postmanRequest(r: CollectionRequest) {
  const { segments, query } = splitPath(r.path);
  const text = bodyText(r.body);
  return {
    name: r.name,
    request: {
      method: r.method,
      header: Object.entries(r.headers).map(([key, value]) => ({ key, value, type: 'text' })),
      url:
        r.path === UPLOAD_URL_VAR
          ? { raw: UPLOAD_URL_VAR, host: [UPLOAD_URL_VAR] }
          : {
              raw: `{{baseUrl}}${r.path}`,
              host: ['{{baseUrl}}'],
              path: segments,
              ...(query.length ? { query } : {}),
            },
      ...(text != null
        ? {
            body: {
              mode: 'raw',
              raw: text,
              ...(isJsonBody(r.body) ? { options: { raw: { language: 'json' } } } : {}),
            },
          }
        : {}),
      ...(r.description ? { description: r.description } : {}),
    },
    response: [],
  };
}

function toPostman(input: CollectionInput) {
  return {
    info: {
      _postman_id: uuid(),
      name: input.name,
      description: 'Exported from Dataverse Request Studio. Set accessToken to a bearer token.',
      schema: 'https://schema.getpostman.com/json/collection/v2.1.0/collection.json',
    },
    auth: {
      type: 'bearer',
      bearer: [{ key: 'token', value: '{{accessToken}}', type: 'string' }],
    },
    variable: [
      { key: 'baseUrl', value: input.baseUrl, type: 'string' },
      { key: 'accessToken', value: '', type: 'string' },
      ...(usesUploadUrl(input.items) ? [{ key: 'uploadUrl', value: '', type: 'string' }] : []),
    ],
    item: input.items.map((it) =>
      it.kind === 'folder'
        ? {
            name: it.name,
            ...(it.description ? { description: it.description } : {}),
            item: it.steps.map(postmanRequest),
          }
        : postmanRequest(it),
    ),
  };
}

// ── Insomnia v4 ──────────────────────────────────────────────────────

function toInsomnia(input: CollectionInput) {
  const wrk = rid('wrk');
  const resources: Record<string, unknown>[] = [
    { _id: wrk, _type: 'workspace', parentId: null, name: input.name, scope: 'collection' },
    {
      _id: rid('env'),
      _type: 'environment',
      parentId: wrk,
      name: 'Base Environment',
      data: {
        baseUrl: input.baseUrl,
        accessToken: '',
        ...(usesUploadUrl(input.items) ? { uploadUrl: '' } : {}),
      },
    },
  ];
  let sort = 0;
  const request = (r: CollectionRequest, parentId: string) => {
    const text = bodyText(r.body);
    resources.push({
      _id: rid('req'),
      _type: 'request',
      parentId,
      name: r.name,
      description: r.description ?? '',
      method: r.method,
      url: r.path === UPLOAD_URL_VAR ? '{{ _.uploadUrl }}' : `{{ _.baseUrl }}${r.path}`,
      headers: Object.entries(r.headers).map(([name, value]) => ({ name, value })),
      body:
        text != null
          ? { mimeType: isJsonBody(r.body) ? 'application/json' : 'text/plain', text }
          : {},
      authentication: { type: 'bearer', token: '{{ _.accessToken }}' },
      metaSortKey: sort++,
    });
  };
  for (const it of input.items) {
    if (it.kind === 'folder') {
      const fld = rid('fld');
      resources.push({
        _id: fld,
        _type: 'request_group',
        parentId: wrk,
        name: it.name,
        description: it.description ?? '',
        metaSortKey: sort++,
      });
      it.steps.forEach((s) => request(s, fld));
    } else {
      request(it, wrk);
    }
  }
  return {
    _type: 'export',
    __export_format: 4,
    __export_date: new Date().toISOString(),
    __export_source: 'dataverse-request-studio',
    resources,
  };
}

// ── Bruno collection JSON ────────────────────────────────────────────
// The shape Bruno writes from "Export collection" and reads back through
// "Import collection → Bruno collection".

function brunoRequest(r: CollectionRequest, seq: number) {
  const text = bodyText(r.body);
  const { query } = splitPath(r.path);
  return {
    type: 'http-request',
    name: r.name,
    seq,
    request: {
      url: r.path === UPLOAD_URL_VAR ? UPLOAD_URL_VAR : `{{baseUrl}}${r.path}`,
      method: r.method,
      headers: Object.entries(r.headers).map(([name, value]) => ({ name, value, enabled: true })),
      params: query.map((q) => ({ name: q.key, value: q.value, type: 'query', enabled: true })),
      body:
        text == null
          ? { mode: 'none' }
          : isJsonBody(r.body)
            ? { mode: 'json', json: text }
            : { mode: 'text', text },
      auth: { mode: 'inherit' },
      script: {},
      vars: {},
      assertions: [],
      tests: '',
      docs: r.description ?? '',
    },
  };
}

function toBruno(input: CollectionInput) {
  return {
    name: input.name,
    version: '1',
    items: input.items.map((it, i) =>
      it.kind === 'folder'
        ? {
            type: 'folder',
            name: it.name,
            seq: i + 1,
            items: it.steps.map((s, j) => brunoRequest(s, j + 1)),
          }
        : brunoRequest(it, i + 1),
    ),
    environments: [
      {
        name: 'Dataverse',
        variables: [
          { name: 'baseUrl', value: input.baseUrl, enabled: true, secret: false, type: 'text' },
          { name: 'accessToken', value: '', enabled: true, secret: true, type: 'text' },
          ...(usesUploadUrl(input.items)
            ? [{ name: 'uploadUrl', value: '', enabled: true, secret: false, type: 'text' }]
            : []),
        ],
      },
    ],
    root: {
      request: { auth: { mode: 'bearer', bearer: { token: '{{accessToken}}' } } },
    },
    brunoConfig: {
      version: '1',
      name: input.name,
      type: 'collection',
      ignore: ['node_modules', '.git'],
    },
  };
}

/** Number of HTTP requests in a collection (folders count their steps). */
export function countRequests(items: CollectionItem[]): number {
  return flatSteps(items).length;
}
