// "Execute sends exactly what the Code tab shows" (#235). For each mode:
// build its state, generate the Code tab's request (the JSON format:
// method, URL, headers, body) from the inputs the mode hands CodeView, run
// the mode's Execute call against a recording `window.dataverseAPI`, turn
// each host call back into the HTTP request the host sends, and compare.
//
// Nothing is mocked between the state and the host: the real builders, code
// generator, executor, header policy and bridge run, on a ToolBox 1.2.7+
// host (header pass-through). The modes' own glue is copied, not rendered:
// each test hands both sides the header rows as the mode composes them
// (the Bypass and Concurrency panes' rows typed in), so those panes'
// composers aren't covered here. The only differences allowed are the
// documented ones, each named in `sentHeaders` below:
//   • the host builds Accept / Content-Type / OData-* / Prefer itself;
//   • a read sends only its allow-listed rows, and a Prefer that keeps the
//     host's annotations and carries the page size (#109, #219);
//   • Execute Function sends only impersonation (#109);
//   • Associate / Disassociate send impersonation and bypass only, plus
//     If-Match: * on the single-valued PATCH (#109, #254).
// Covered: bound and collection-bound routing (#55), function parameters
// (#61), the page size and $top (#219). Credential headers (#217) are
// covered in `requestHeaderPolicy.test.ts`.
import { beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';

type Headers = Record<string, string>;
interface HostCall {
  op: string;
  args: unknown[];
}

// A ToolBox 1.2.7+ host, in place before any module reads it: the bridge
// detects embedding and header pass-through once, at load.
const host = vi.hoisted(() => {
  const calls: HostCall[] = [];
  const record =
    (op: string, result: unknown = undefined) =>
    async (...args: unknown[]) => {
      calls.push({ op, args });
      return typeof result === 'function' ? (result as () => unknown)() : result;
    };
  const dataverseAPI = {
    create: record('create', { id: '00000000-0000-0000-0000-00000000ffff' }),
    update: record('update'),
    delete: record('delete'),
    retrieve: record('retrieve', {}),
    // Associate first asks which targets are already linked: none are.
    queryData: record('queryData', () => ({ value: [] })),
    execute: record('execute', {}),
    associate: record('associate'),
    disassociate: record('disassociate'),
    executeBatch: record('executeBatch', []),
    executeTransaction: record('executeTransaction', []),
  };
  const w = window as unknown as Record<string, unknown>;
  w.dataverseAPI = dataverseAPI;
  w.toolboxAPI = {
    utils: { getCurrentTheme: async () => 'light' },
    connections: { getActiveConnection: async () => null },
    events: { on: () => {} },
  };
  return { calls };
});

vi.mock('../mock/environment', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../mock/environment')>()),
  ENV: { name: 'Test', host: 'contoso.crm.dynamics.com', apiBase: '/api/data/v9.2', user: '' },
}));

// The CSDL operations the Execute modes pick from.
const catalog = vi.hoisted(() => {
  const op = (
    name: string,
    kind: 'Action' | 'Function',
    binding: Record<string, string>,
    parameters: Record<string, unknown>[],
  ) => ({
    name,
    kind,
    namespace: 'Microsoft.Dynamics.CRM',
    binding,
    parameters,
    returnType: { kind: 'void', typeName: 'void' },
    source: 'custom-api',
  });
  const list = [
    op('new_Notify', 'Action', { kind: 'unbound' }, [
      { name: 'Message', type: 'Edm.String', required: true },
      { name: 'Count', type: 'Edm.Int32', required: false },
      { name: 'Customer', type: 'EntityReference', required: false, entityType: 'account' },
    ]),
    op('new_Score', 'Action', { kind: 'entity', entityType: 'account' }, [
      { name: 'Weight', type: 'Edm.Decimal', required: false },
    ]),
    op('new_Recalculate', 'Action', { kind: 'collection', entityType: 'account' }, []),
    op('new_Find', 'Function', { kind: 'unbound' }, [
      { name: 'Name', type: 'Edm.String', required: true },
      { name: 'Since', type: 'Edm.DateTimeOffset', required: false },
      { name: 'Owner', type: 'Edm.Guid', required: false },
    ]),
    op('new_Rank', 'Function', { kind: 'entity', entityType: 'account' }, []),
  ];
  return { find: (name: string) => list.find((a) => a.name === name) };
});
vi.mock('../host/csdlProvider', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/csdlProvider')>()),
  findActionSync: catalog.find,
  actions: { find: async (name: string) => catalog.find(name) },
}));

// Table metadata comes from the fixtures, as the builders read it. Filled in
// `beforeAll`: importing the fixtures from the mock factory would cycle.
const tables = vi.hoisted(() => new Map<string, unknown>());
vi.mock('../host/metadataProvider', () => ({
  metadata: {
    peekTable: (n: string) => tables.get(n),
    getTable: async (n: string) => tables.get(n),
  },
}));

import { runtime } from './runtime';
import { generateCode, type CodegenInputs } from './codeGenerators';
import {
  buildAssociate,
  buildAssociateBody,
  buildAssociateRequests,
  buildCreate,
  buildCreateBody,
  buildDelete,
  buildDisassociate,
  buildDisassociateRequests,
  buildRetrieveMultiple,
  buildRetrieveSingle,
  buildUpdate,
  buildUpdateBody,
  buildUpsert,
  buildUpsertBody,
  relateRequestHeaders,
} from './urlBuilder';
import {
  buildExecuteAction,
  buildExecuteActionBody,
  buildExecuteFunction,
  buildExecuteWorkflow,
  buildExecuteWorkflowBody,
} from './executeBuilders';
import { headerItemsToObject, type HeaderItem } from '../editors/HeadersEditor';
import { emptyPrefer, preferToHeaderString } from '../editors/PreferEditor';
import { makeExpandSpec } from '../editors/ExpandEditor';
import {
  ACCOUNT_ID,
  CONTACT_ID,
  account,
  contact,
  createState,
  group,
  registerFixtureTables,
  retrieveMultipleState,
  rule,
} from '../test/fixtures';
import { defaultBypassOptions } from '../state/writeState';
import type { DeleteState, UpdateState, UpsertState } from '../state/writeState';
import type { RetrieveSingleState } from '../state/readState';
import type { AssociateState, DisassociateState } from '../state/relateState';
import type {
  ExecuteActionState,
  ExecuteFunctionState,
  ExecuteWorkflowState,
} from '../state/executeState';

const API = '/api/data/v9.2/';
const CALLER = '00000000-0000-0000-0000-0000000000e1';
const WORKFLOW_ID = '00000000-0000-0000-0000-0000000000f1';
const CONTACT_2 = '00000000-0000-0000-0000-0000000000c2';

/** A request as it goes over the wire, relative to the Web API base. */
interface Wire {
  method: string;
  path: string;
  headers: Headers;
  body?: unknown;
}

// ── The Code tab side ────────────────────────────────────────────────────

/** The Code tab's request(s), from its JSON format. */
function codeTab(inputs: CodegenInputs): Wire[] {
  const out = JSON.parse(generateCode('json', inputs));
  const rel = (u: string) => (u.startsWith(API) ? u.slice(API.length) : u);
  if (out.mode === 'multi-request' || out.mode === 'change-set') {
    return out.requests.map((r: { method: string; relativeUrl: string; body?: unknown }) => ({
      method: r.method,
      path: rel(r.relativeUrl),
      headers: out.headers,
      body: r.body,
    }));
  }
  return [{ method: out.method, path: rel(out.relativeUrl), headers: out.headers, body: out.body }];
}

// ── The Execute side ─────────────────────────────────────────────────────

const setOf = (logical: string) =>
  [account, contact].find((t) => t.logicalName === logical)!.entitySetName;

/**
 * The HTTP request the host sends for a recorded `dataverseAPI` call, per
 * PPTB's documented mapping (#55 for the three action bindings). Headers
 * are the call's `additionalHeaders` (the last argument).
 */
function toWire({ op, args }: HostCall): Wire {
  const last = args[args.length - 1];
  const headers = (last && typeof last === 'object' ? last : {}) as Headers;
  switch (op) {
    case 'queryData':
      return { method: 'GET', path: args[0] as string, headers: args.length > 1 ? headers : {} };
    case 'create':
      return {
        method: 'POST',
        path: setOf(args[0] as string),
        body: args[1],
        headers: args.length > 2 ? headers : {},
      };
    case 'update':
      return {
        method: 'PATCH',
        path: `${setOf(args[0] as string)}(${args[1]})`,
        body: args[2],
        headers: args.length > 3 ? headers : {},
      };
    case 'delete':
      return {
        method: 'DELETE',
        path: `${setOf(args[0] as string)}(${args[1]})`,
        headers: args.length > 2 ? headers : {},
      };
    case 'execute': {
      const r = args[0] as {
        operationName: string;
        operationType: string;
        entityName?: string;
        entityId?: string;
        parameters?: Record<string, unknown>;
      };
      // A function sent through execute() would go out as a GET the host
      // composes itself; the modes send functions through queryData.
      if (r.operationType !== 'action') throw new Error(`execute() got a ${r.operationType}`);
      // The host gets only the logical name and guesses the entity set by
      // pluralizing it, so model that rather than the real set name. The
      // fixture tables pluralize regularly; an irregular one (opportunity)
      // would need the executor to pass the set name.
      const set = r.entityName && `${r.entityName}s`;
      const op = r.entityName ? `Microsoft.Dynamics.CRM.${r.operationName}` : r.operationName;
      const path = !set ? op : r.entityId ? `${set}(${r.entityId})/${op}` : `${set}/${op}`;
      return { method: 'POST', path, body: r.parameters, headers: args.length > 1 ? headers : {} };
    }
    case 'associate': {
      const [primary, id, rel, target, targetId] = args as string[];
      return {
        method: 'POST',
        path: `${setOf(primary)}(${id})/${rel}/$ref`,
        body: { '@odata.id': `${setOf(target)}(${targetId})` },
        headers: args.length > 5 ? headers : {},
      };
    }
    case 'disassociate': {
      const [primary, id, rel, targetId] = args as string[];
      return {
        method: 'DELETE',
        path: `${setOf(primary)}(${id})/${rel}(${targetId})/$ref`,
        headers: args.length > 4 ? headers : {},
      };
    }
  }
  throw new Error(`unmapped host call ${op}`);
}

// ── Comparing ────────────────────────────────────────────────────────────

/** A URL path + query, decoded, so the Code tab's readable form and the
 *  wire's percent-encoded one compare equal when they mean the same. */
function decoded(path: string): { path: string; query: [string, string][] } {
  const [p, q = ''] = path.split(/\?(.*)/s);
  return {
    path: decodeURIComponent(p).replace(/^\//, ''),
    query: [...new URLSearchParams(q.replace(/ /g, '%20'))],
  };
}

const HOST_OWNED = ['accept', 'content-type', 'odata-maxversion', 'odata-version', 'prefer'];
const lower = (h: Headers) =>
  Object.fromEntries(Object.entries(h ?? {}).map(([k, v]) => [k.toLowerCase(), v]));
const pick = (h: Headers, names: string[]) =>
  Object.fromEntries(Object.entries(lower(h)).filter(([k]) => names.includes(k)));
const without = (h: Headers, names: string[]) =>
  Object.fromEntries(Object.entries(lower(h)).filter(([k]) => !names.includes(k)));

const IMPERSONATION = ['mscrmcallerid', 'callerobjectid'];
const BYPASS = [
  'mscrm.bypasscustompluginexecution',
  'mscrm.bypassbusinesslogicexecution',
  'mscrm.suppresscallbackregistrationexpanderjob',
];

/** What Execute must send, given the Code tab's headers: the documented differences. */
const sentHeaders = {
  /** Writes and actions: every row but the ones the host builds itself. */
  write: (code: Headers) => without(code, HOST_OWNED),
  /** Reads: the allow-list, plus a Prefer for the page size (#109, #219). */
  read: (code: Headers, maxpagesize: number | null) => ({
    ...pick(code, [...IMPERSONATION, 'if-none-match', 'consistency']),
    // The host's own annotations first, then the page size: Dataverse
    // rejects the other order (0x80060888).
    ...(maxpagesize
      ? { prefer: `odata.include-annotations="*",odata.maxpagesize=${maxpagesize}` }
      : {}),
  }),
  impersonation: (code: Headers) => pick(code, IMPERSONATION),
  relate: (code: Headers) => pick(code, [...IMPERSONATION, ...BYPASS]),
  /** Single-valued Associate / Disassociate: a PATCH that must not create the source (#254). */
  relatePatch: (code: Headers) => pick(code, [...IMPERSONATION, ...BYPASS, 'if-match']),
};

function expectSame(code: Wire, sent: Wire, expectedHeaders: Headers) {
  expect(sent.method).toBe(code.method);
  expect(decoded(sent.path)).toEqual(decoded(code.path));
  expect(sent.body ?? null).toEqual(code.body ?? null);
  expect(lower(sent.headers)).toEqual(expectedHeaders);
}

// ── Fixtures ─────────────────────────────────────────────────────────────

let hid = 0;
const rows = (h: Headers): HeaderItem[] =>
  Object.entries(h).map(([name, value]) => ({ id: `h${++hid}`, name, value, enabled: true }));

/** What every mode's Headers pane starts with, plus the given rows. */
const pane = (extra: Headers = {}): HeaderItem[] =>
  rows({
    Accept: 'application/json',
    'OData-MaxVersion': '4.0',
    'OData-Version': '4.0',
    ...extra,
  });
const writePane = (extra: Headers = {}) =>
  pane({ 'Content-Type': 'application/json; charset=utf-8', ...extra });

beforeAll(() => {
  registerFixtureTables();
  for (const t of [account, contact]) tables.set(t.logicalName, t);
});
beforeEach(() => {
  host.calls.length = 0;
});

async function run(exec: () => Promise<{ ok: boolean; status: number; body?: unknown }>) {
  const res = await exec();
  expect(res.ok, JSON.stringify(res.body)).toBe(true);
}

// ── Reads ────────────────────────────────────────────────────────────────

describe('reads', () => {
  it('Retrieve Multiple: select, filter, orderby, expand, top, impersonation', async () => {
    const s = retrieveMultipleState({
      select: ['name', 'revenue', 'createdon'],
      filter: group('and', [
        rule('name', 'eq', "O'Brien & Sons"),
        group('or', [rule('numberofemployees', 'gt', '10'), rule('industrycode', 'eq', '2')]),
      ]),
      orderby: [{ col: 'name', dir: 'desc' }],
      top: 5,
      expand: [{ ...makeExpandSpec('primarycontactid'), select: ['fullname'] }],
      headers: pane({ MSCRMCallerID: CALLER, 'X-Custom': 'kept in the Code tab only' }),
    });
    const [code] = codeTab({
      method: 'GET',
      built: buildRetrieveMultiple(s),
      headers: headerItemsToObject(s.headers, preferToHeaderString(s.prefer)),
    });
    await run(() => runtime.retrieveMultiple(s, headerItemsToObject(s.headers, null)));
    expect(host.calls).toHaveLength(1);
    expectSame(code, toWire(host.calls[0]), sentHeaders.read(code.headers, null));
    expect(code.path).toContain('$top=5');
  });

  it('Retrieve Multiple with a page size: both leave $top out, Execute sends the page size', async () => {
    const s = retrieveMultipleState({
      select: ['name'],
      top: 5,
      prefer: { ...emptyPrefer(), maxpagesize: 50 },
      headers: pane(),
    });
    const [code] = codeTab({
      method: 'GET',
      built: buildRetrieveMultiple(s),
      headers: headerItemsToObject(s.headers, preferToHeaderString(s.prefer)),
    });
    await run(() => runtime.retrieveMultiple(s, headerItemsToObject(s.headers, null)));
    expectSame(code, toWire(host.calls[0]), sentHeaders.read(code.headers, 50));
    expect(code.path).not.toContain('$top');
    expect(code.headers.Prefer).toContain('odata.maxpagesize=50');
  });

  it('Retrieve Single: select, expand, If-None-Match', async () => {
    const s = {
      table: 'account',
      recordId: ACCOUNT_ID,
      select: ['name', 'revenue'],
      expand: [{ ...makeExpandSpec('primarycontactid'), select: ['fullname', 'jobtitle'] }],
      prefer: emptyPrefer(),
      headers: pane({ 'If-None-Match': 'W/"12345"' }),
      dirty: new Set(),
    } as unknown as RetrieveSingleState;
    const [code] = codeTab({
      method: 'GET',
      built: buildRetrieveSingle(s),
      headers: headerItemsToObject(s.headers, preferToHeaderString(s.prefer)),
    });
    await run(() => runtime.retrieveSingle(s, headerItemsToObject(s.headers, null)));
    expectSame(code, toWire(host.calls[0]), sentHeaders.read(code.headers, null));
  });
});

// ── Writes ───────────────────────────────────────────────────────────────

describe('writes', () => {
  const fields = {
    name: "O'Brien & Sons",
    numberofemployees: 12,
    revenue: 1500.5,
    donotemail: true,
    industrycode: 2,
    sample_tags: [100, 200],
    createdon: '2024-03-01T09:30:00Z',
    primarycontactid: { id: CONTACT_ID, targetEntity: 'contact' },
  };
  const extra = {
    MSCRMCallerID: CALLER,
    'MSCRM.BypassBusinessLogicExecution': 'CustomSync',
    'X-Custom': 'sent as typed',
  };

  it('Create: every column shape, lookups as @odata.bind, explicit nulls', async () => {
    const s = createState({
      fieldValues: { ...fields, accountnumber: 'kept' },
      nullFields: ['accountnumber'],
      prefer: { ...emptyPrefer(), returnRepresentation: true },
      headers: writePane(extra),
    });
    const [code] = codeTab({
      method: 'POST',
      built: buildCreate(s),
      headers: headerItemsToObject(s.headers, preferToHeaderString(s.prefer)),
      body: buildCreateBody(s),
      entityLogical: 'account',
    });
    await run(() => runtime.create(s, headerItemsToObject(s.headers, null)));
    expectSame(code, toWire(host.calls[0]), sentHeaders.write(code.headers));
    expect(code.body).toMatchObject({
      'primarycontactid@odata.bind': `/contacts(${CONTACT_ID})`,
      sample_tags: '100,200',
      accountnumber: null,
    });
  });

  it('Update (PATCH) with If-Match: *', async () => {
    const s = {
      table: 'account',
      recordId: ACCOUNT_ID,
      method: 'PATCH',
      putColumn: null,
      fieldValues: fields,
      nullFields: [],
      rawBody: null,
      prefer: emptyPrefer(),
      headers: writePane({ ...extra, 'If-Match': '*' }),
      returnSelect: [],
      concurrency: { kind: 'none' },
      bypass: defaultBypassOptions(),
      dirty: new Set(),
    } as unknown as UpdateState;
    const [code] = codeTab({
      method: 'PATCH',
      built: buildUpdate(s),
      headers: headerItemsToObject(s.headers, preferToHeaderString(s.prefer)),
      body: buildUpdateBody(s),
      entityLogical: 'account',
    });
    await run(() => runtime.update(s, headerItemsToObject(s.headers, null)));
    expect(host.calls).toHaveLength(1);
    expectSame(code, toWire(host.calls[0]), sentHeaders.write(code.headers));
  });

  it('Upsert by GUID', async () => {
    const s = {
      table: 'account',
      key: { kind: 'guid', recordId: ACCOUNT_ID },
      fieldValues: fields,
      nullFields: [],
      rawBody: null,
      prefer: emptyPrefer(),
      headers: writePane(extra),
      returnSelect: [],
      concurrency: { kind: 'none' },
      bypass: defaultBypassOptions(),
      dirty: new Set(),
    } as unknown as UpsertState;
    const [code] = codeTab({
      method: 'PATCH',
      built: buildUpsert(s),
      headers: headerItemsToObject(s.headers, preferToHeaderString(s.prefer)),
      body: buildUpsertBody(s),
      entityLogical: 'account',
    });
    await run(() => runtime.upsert(s, headerItemsToObject(s.headers, null)));
    expectSame(code, toWire(host.calls[0]), sentHeaders.write(code.headers));
  });

  it('Delete (whole row) with If-Match', async () => {
    const s = {
      table: 'account',
      recordId: ACCOUNT_ID,
      scope: { kind: 'whole-row' },
      concurrency: { kind: 'none' },
      headers: pane({ ...extra, 'If-Match': 'W/"777"' }),
      bypassCustomPlugins: false,
      bypass: defaultBypassOptions(),
      confirmText: '',
      acknowledged: true,
      dirty: new Set(),
    } as unknown as DeleteState;
    const [code] = codeTab({
      method: 'DELETE',
      built: buildDelete(s),
      headers: headerItemsToObject(s.headers, null),
      entityLogical: 'account',
    });
    await run(() => runtime.delete(s, headerItemsToObject(s.headers, null)));
    expectSame(code, toWire(host.calls[0]), sentHeaders.write(code.headers));
  });
});

// ── Execute ──────────────────────────────────────────────────────────────

describe('Execute Action / Function / Workflow', () => {
  const action = (over: Partial<ExecuteActionState>) =>
    ({
      actionName: null,
      boundRecordId: null,
      paramValues: {},
      headers: writePane({ MSCRMCallerID: CALLER }),
      category: 'custom-api',
      dirty: new Set(),
      ...over,
    }) as ExecuteActionState;

  async function actionParity(s: ExecuteActionState) {
    const [code] = codeTab({
      method: 'POST',
      built: buildExecuteAction(s),
      headers: headerItemsToObject(s.headers, null),
      body: buildExecuteActionBody(s),
    });
    await run(() => runtime.action(s, headerItemsToObject(s.headers, null)));
    expect(host.calls).toHaveLength(1);
    expectSame(code, toWire(host.calls[0]), sentHeaders.write(code.headers));
    return code;
  }

  it('an unbound action with string, number and record-reference parameters', async () => {
    const code = await actionParity(
      action({
        actionName: 'new_Notify',
        paramValues: {
          Message: "It's done & dusted",
          Count: 3,
          Customer: { id: ACCOUNT_ID, entityType: 'account' },
        } as never,
      }),
    );
    expect(code.path).toBe('new_Notify');
  });

  it('an entity-bound action goes to the record (#55)', async () => {
    const code = await actionParity(
      action({ actionName: 'new_Score', boundRecordId: ACCOUNT_ID, paramValues: { Weight: 1.5 } }),
    );
    expect(code.path).toBe(`accounts(${ACCOUNT_ID})/Microsoft.Dynamics.CRM.new_Score`);
  });

  it('a collection-bound action goes to the collection, not the unbound form (#55)', async () => {
    const code = await actionParity(action({ actionName: 'new_Recalculate' }));
    expect(code.path).toBe('accounts/Microsoft.Dynamics.CRM.new_Recalculate');
  });

  const fn = (over: Partial<ExecuteFunctionState>) =>
    ({
      functionName: null,
      boundRecordId: null,
      paramValues: {},
      useParamAliases: false,
      headers: pane({ MSCRMCallerID: CALLER, 'X-Custom': 'Code tab only' }),
      category: 'custom-api',
      dirty: new Set(),
      ...over,
    }) as ExecuteFunctionState;

  async function functionParity(s: ExecuteFunctionState) {
    const [code] = codeTab({
      method: 'GET',
      built: buildExecuteFunction(s),
      headers: headerItemsToObject(s.headers, null),
    });
    await run(() => runtime.function(s, headerItemsToObject(s.headers, null)));
    expect(host.calls).toHaveLength(1);
    expectSame(code, toWire(host.calls[0]), sentHeaders.impersonation(code.headers));
    return code;
  }

  const awkward = {
    Name: "O'Brien & Co + #1 50% ?",
    Since: '2024-03-01T09:30:00Z',
    Owner: CALLER,
  };

  it('a function with parameters inline (#61)', async () => {
    await functionParity(fn({ functionName: 'new_Find', paramValues: awkward }));
  });

  it('a function with parameter aliases (#61)', async () => {
    const code = await functionParity(
      fn({ functionName: 'new_Find', paramValues: awkward, useParamAliases: true }),
    );
    expect(code.path).toMatch(/^new_Find\(Name=@p1,Since=@p2,Owner=@p3\)\?@p1=/);
  });

  it('an entity-bound function uses the real entity set (#55)', async () => {
    const code = await functionParity(fn({ functionName: 'new_Rank', boundRecordId: ACCOUNT_ID }));
    expect(code.path).toBe(`accounts(${ACCOUNT_ID})/Microsoft.Dynamics.CRM.new_Rank()`);
  });

  it('Execute Workflow', async () => {
    const s = {
      workflowId: WORKFLOW_ID,
      entityId: ACCOUNT_ID,
      headers: writePane({ MSCRMCallerID: CALLER }),
      dirty: new Set(),
    } as ExecuteWorkflowState;
    const [code] = codeTab({
      method: 'POST',
      built: buildExecuteWorkflow(s),
      headers: headerItemsToObject(s.headers, null),
      body: buildExecuteWorkflowBody(s),
    });
    await run(() => runtime.workflow(s, headerItemsToObject(s.headers, null)));
    expectSame(code, toWire(host.calls[0]), sentHeaders.write(code.headers));
  });
});

// ── Relate ───────────────────────────────────────────────────────────────

describe('Associate / Disassociate', () => {
  const relatePane = () =>
    writePane({
      MSCRMCallerID: CALLER,
      'MSCRM.BypassBusinessLogicExecution': 'CustomSync',
      'X-Custom': 'Code tab only',
    });

  it('Associate over a collection: one POST $ref per target', async () => {
    const s = {
      table: 'account',
      sourceId: ACCOUNT_ID,
      navProperty: 'contact_customer_accounts',
      targets: [CONTACT_ID, CONTACT_2],
      targetNames: {},
      headers: relatePane(),
      dirty: new Set(),
    } as unknown as AssociateState;
    const requests = buildAssociateRequests(s);
    const code = codeTab({
      method: 'POST',
      built: buildAssociate(s),
      headers: headerItemsToObject(s.headers, null),
      body: buildAssociateBody(s) as Record<string, unknown>,
      multiRequests: requests.map((r) => ({
        method: r.method,
        relativeUrl: r.relativeUrl,
        body: r.body as unknown as Record<string, unknown>,
      })),
    });
    await run(() => runtime.associate(s, headerItemsToObject(s.headers, null)));
    const sent = host.calls.filter((c) => c.op === 'associate').map(toWire);
    expect(sent).toHaveLength(2);
    sent.forEach((w, i) => {
      // The Code tab's @odata.id is absolute; the host takes the target's id.
      const odataId = (code[i].body as { '@odata.id': string })['@odata.id'];
      expect(odataId.endsWith((w.body as { '@odata.id': string })['@odata.id'])).toBe(true);
      expectSame({ ...code[i], body: w.body }, w, sentHeaders.relate(code[i].headers));
    });
  });

  it('Associate a single-valued lookup: PATCH the source, which must exist (#254)', async () => {
    const s = {
      table: 'account',
      sourceId: ACCOUNT_ID,
      navProperty: 'primarycontactid',
      targets: [CONTACT_ID],
      targetNames: {},
      headers: relatePane(),
      dirty: new Set(),
    } as unknown as AssociateState;
    const [code] = codeTab({
      method: 'PATCH',
      built: buildAssociate(s),
      headers: relateRequestHeaders('PATCH', headerItemsToObject(s.headers, null)),
      body: buildAssociateBody(s) as Record<string, unknown>,
    });
    await run(() => runtime.associate(s, headerItemsToObject(s.headers, null)));
    expect(host.calls).toHaveLength(1);
    expectSame(code, toWire(host.calls[0]), sentHeaders.relatePatch(code.headers));
    expect(code.headers['If-Match']).toBe('*');
  });

  it('Disassociate over a collection: one DELETE $ref per target', async () => {
    const s = {
      table: 'account',
      sourceId: ACCOUNT_ID,
      navProperty: 'contact_customer_accounts',
      targetIds: [CONTACT_ID, CONTACT_2],
      targetNames: {},
      headers: relatePane(),
      dirty: new Set(),
    } as unknown as DisassociateState;
    const requests = buildDisassociateRequests(s);
    const code = codeTab({
      method: 'DELETE',
      built: buildDisassociate(s),
      headers: headerItemsToObject(s.headers, null),
      multiRequests: requests.map((r) => ({ method: r.method, relativeUrl: r.relativeUrl })),
    });
    await run(() => runtime.disassociate(s, headerItemsToObject(s.headers, null)));
    const sent = host.calls.filter((c) => c.op === 'disassociate').map(toWire);
    expect(sent).toHaveLength(2);
    sent.forEach((w, i) => expectSame(code[i], w, sentHeaders.relate(code[i].headers)));
  });

  // All or nothing (#224): the same requests, as one executeTransaction call.
  it.each(['associate', 'disassociate'] as const)(
    '%s, all or nothing: one change set with the same requests',
    async (op) => {
      const assoc = op === 'associate';
      const s = {
        table: 'account',
        sourceId: ACCOUNT_ID,
        navProperty: 'contact_customer_accounts',
        [assoc ? 'targets' : 'targetIds']: [CONTACT_ID, CONTACT_2],
        targetNames: {},
        allOrNothing: true,
        headers: relatePane(),
        dirty: new Set(),
      } as unknown as AssociateState & DisassociateState;
      const requests = assoc ? buildAssociateRequests(s) : buildDisassociateRequests(s);
      const inputs: CodegenInputs = {
        method: assoc ? 'POST' : 'DELETE',
        built: assoc ? buildAssociate(s) : buildDisassociate(s),
        headers: headerItemsToObject(s.headers, null),
        multiRequests: requests.map((r) => ({
          method: r.method,
          relativeUrl: r.relativeUrl,
          body: r.body as Record<string, unknown> | undefined,
        })),
        changeSet: true,
      };
      expect(JSON.parse(generateCode('json', inputs)).mode).toBe('change-set');
      const code = codeTab(inputs);
      await run(() =>
        assoc
          ? runtime.associate(s, headerItemsToObject(s.headers, null))
          : runtime.disassociate(s, headerItemsToObject(s.headers, null)),
      );
      expect(host.calls.filter((c) => c.op === op)).toHaveLength(0);
      const tx = host.calls.filter((c) => c.op === 'executeTransaction');
      expect(tx).toHaveLength(1);
      const parts = tx[0].args[0] as {
        method: string;
        url: string;
        headers: Headers;
        body?: unknown;
      }[];
      expect(tx[0].args).toHaveLength(1);
      expect(parts).toHaveLength(2);
      parts.forEach((p, i) =>
        expectSame(
          code[i],
          { method: p.method, path: p.url, headers: p.headers, body: p.body },
          sentHeaders.relate(code[i].headers),
        ),
      );
    },
  );
});
