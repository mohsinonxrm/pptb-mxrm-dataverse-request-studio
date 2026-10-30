// The write executors' refusals and host calls (#235): requests the PPTB host
// can't send as authored are refused before any host call (PUT one column,
// alternate-key Upsert, single-property Delete), ids that would change the
// URL aren't sent (#221), bound operations need their record, and an older
// ToolBox that would drop impersonation doesn't run the write as you (#109).
import { beforeEach, describe, expect, it, vi } from 'vitest';
import {
  executeAction,
  executeCreate,
  executeDelete,
  executeFunction,
  executeMerge,
  executeUpdate,
  executeUpsert,
  executeWorkflow,
} from './dataverseExecutor';
import { buildCreateBody, buildMergeBody } from './urlBuilder';
import {
  ACCOUNT_ID,
  CONTACT_ID,
  account,
  contact,
  createState,
  registerFixtureTables,
} from '../test/fixtures';
import {
  defaultBypassOptions,
  type DeleteState,
  type MergeState,
  type UpdateState,
  type UpsertState,
} from '../state/writeState';
import type {
  ExecuteActionState,
  ExecuteFunctionState,
  ExecuteWorkflowState,
} from '../state/executeState';

// The bridge decides host capabilities once at module load; mock it so each
// test picks the host it runs against and sees every host call.
const host = vi.hoisted(() => ({
  additionalHeaders: true,
  create: vi.fn(),
  update: vi.fn(),
  delete: vi.fn(),
  retrieve: vi.fn(),
  execute: vi.fn(),
  queryData: vi.fn(),
}));
vi.mock('../host/pptbBridge', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/pptbBridge')>()),
  isEmbedded: () => true,
  hostSupportsAdditionalHeaders: () => host.additionalHeaders,
  dvHost: {
    create: host.create,
    update: host.update,
    delete: host.delete,
    retrieve: host.retrieve,
    execute: host.execute,
    queryData: host.queryData,
  },
}));
const tables = vi.hoisted(() => new Map<string, unknown>());
vi.mock('../host/metadataProvider', () => ({
  metadata: {
    peekTable: (n: string) => tables.get(n),
    getTable: async (n: string) => tables.get(n),
  },
}));
vi.mock('../host/csdlProvider', async (importOriginal) => {
  const ops: Record<string, unknown> = {
    new_Score: {
      name: 'new_Score',
      kind: 'Action',
      binding: { kind: 'entity', entityType: 'account' },
      parameters: [],
    },
    new_Rank: {
      name: 'new_Rank',
      kind: 'Function',
      binding: { kind: 'entity', entityType: 'account' },
      parameters: [],
    },
  };
  return {
    ...(await importOriginal<typeof import('../host/csdlProvider')>()),
    findActionSync: (n: string) => ops[n],
    actions: { find: async (n: string) => ops[n] },
  };
});

const CALLER = '00000000-0000-0000-0000-0000000000e1';
const BAD_ID = 'x)/contacts(y';
const PANE = {
  Accept: 'application/json',
  'Content-Type': 'application/json; charset=utf-8',
  'OData-Version': '4.0',
  'X-Custom': 'typed',
};

const anyHostCall = () =>
  [host.create, host.update, host.delete, host.execute, host.queryData].some(
    (f) => f.mock.calls.length > 0,
  );

const update = (over: Partial<UpdateState> = {}) =>
  ({
    table: 'account',
    recordId: ACCOUNT_ID,
    method: 'PATCH',
    putColumn: null,
    fieldValues: { name: 'Fabrikam' },
    nullFields: [],
    rawBody: null,
    prefer: { returnRepresentation: false, includeAnnotations: false, maxPageSize: null },
    headers: [],
    returnSelect: [],
    concurrency: { kind: 'none' },
    bypass: defaultBypassOptions(),
    dirty: new Set(),
    ...over,
  }) as unknown as UpdateState;

const upsert = (over: Partial<UpsertState> = {}) =>
  ({
    table: 'account',
    key: { kind: 'guid', recordId: ACCOUNT_ID },
    fieldValues: { name: 'Fabrikam' },
    nullFields: [],
    rawBody: null,
    prefer: { returnRepresentation: false, includeAnnotations: false, maxPageSize: null },
    headers: [],
    returnSelect: [],
    concurrency: { kind: 'none' },
    bypass: defaultBypassOptions(),
    dirty: new Set(),
    ...over,
  }) as unknown as UpsertState;

const del = (over: Partial<DeleteState> = {}) =>
  ({
    table: 'account',
    recordId: ACCOUNT_ID,
    scope: { kind: 'whole-row' },
    concurrency: { kind: 'none' },
    headers: [],
    bypassCustomPlugins: false,
    bypass: defaultBypassOptions(),
    confirmText: '',
    acknowledged: true,
    dirty: new Set(),
    ...over,
  }) as unknown as DeleteState;

const merge = (over: Partial<MergeState> = {}) =>
  ({
    table: 'account',
    targetId: ACCOUNT_ID,
    subordinateId: '00000000-0000-0000-0000-0000000000a2',
    fieldChoices: {},
    customValues: {},
    performParentingChecks: false,
    headers: [],
    bypass: defaultBypassOptions(),
    dirty: new Set(),
    ...over,
  }) as unknown as MergeState;

beforeEach(() => {
  registerFixtureTables();
  for (const t of [account, contact]) tables.set(t.logicalName, t);
  host.additionalHeaders = true;
  for (const f of [host.create, host.update, host.delete, host.execute, host.queryData]) {
    f.mockReset().mockResolvedValue(undefined);
  }
  host.retrieve.mockReset().mockResolvedValue({ accountid: ACCOUNT_ID });
});

describe('requests the host can only send in another shape are refused, not reshaped', () => {
  it('Update with PUT (one column) sends nothing: the host only PATCHes the row', async () => {
    const r = await executeUpdate(
      update({ method: 'PUT', putColumn: 'name' } as Partial<UpdateState>),
    );
    expect(r.ok).toBe(false);
    expect(r.status).toBe(400);
    expect(JSON.stringify(r.body)).toMatch(/PUT single-column .* isn't supported/);
    expect(anyHostCall()).toBe(false);
  });

  it('Upsert by alternate key sends nothing: the host only addresses rows by GUID', async () => {
    const r = await executeUpsert(
      upsert({
        key: {
          kind: 'alternate',
          keyName: 'akey_accountnumber',
          keyValues: { accountnumber: 'A-1' },
        },
      }),
    );
    expect(r.status).toBe(400);
    expect(JSON.stringify(r.body)).toMatch(/Alternate-key Upsert/);
    expect(anyHostCall()).toBe(false);
  });

  it('Delete of one column sends nothing, rather than deleting the whole row', async () => {
    const r = await executeDelete(del({ scope: { kind: 'single-property', column: 'name' } }));
    expect(r.status).toBe(400);
    expect(JSON.stringify(r.body)).toMatch(/Single-property DELETE/);
    expect(host.delete).not.toHaveBeenCalled();
    expect(anyHostCall()).toBe(false);
  });

  it('a whole-row Delete deletes the row', async () => {
    const r = await executeDelete(del(), { ...PANE, 'If-Match': '*' });
    expect(r.status).toBe(204);
    expect(host.delete).toHaveBeenCalledWith('account', ACCOUNT_ID, {
      'X-Custom': 'typed',
      'If-Match': '*',
    });
  });
});

describe('missing or unsafe ids', () => {
  it.each([
    ['Update', () => executeUpdate(update({ recordId: null }))],
    ['Delete', () => executeDelete(del({ recordId: null }))],
    ['Upsert', () => executeUpsert(upsert({ key: { kind: 'guid', recordId: null } }))],
  ])('%s without a record sends nothing', async (_name, exec) => {
    const r = await exec();
    expect(r.status).toBe(400);
    expect(anyHostCall()).toBe(false);
  });

  it.each([
    ['Update', () => executeUpdate(update({ recordId: BAD_ID }))],
    ['Delete', () => executeDelete(del({ recordId: BAD_ID }))],
    ['Upsert', () => executeUpsert(upsert({ key: { kind: 'guid', recordId: BAD_ID } }))],
    ['Merge (master)', () => executeMerge(merge({ targetId: BAD_ID }))],
    ['Merge (subordinate)', () => executeMerge(merge({ subordinateId: BAD_ID }))],
    [
      'Execute Workflow',
      () =>
        executeWorkflow({
          workflowId: BAD_ID,
          entityId: ACCOUNT_ID,
          headers: [],
          dirty: new Set(),
        } as ExecuteWorkflowState),
    ],
  ])('%s with an id that would change the URL sends nothing (#221)', async (_name, exec) => {
    const r = await exec();
    expect(r.ok).toBe(false);
    expect(JSON.stringify(r.body)).toMatch(/Not sent/);
    expect(anyHostCall()).toBe(false);
  });
});

describe('bound operations need their record', () => {
  it('a record-bound action without one sends nothing', async () => {
    const r = await executeAction({
      actionName: 'new_Score',
      boundRecordId: null,
      paramValues: {},
      headers: [],
      category: 'custom-api',
      dirty: new Set(),
    } as unknown as ExecuteActionState);
    expect(r.status).toBe(400);
    expect(JSON.stringify(r.body)).toMatch(/is bound — pick a source record/);
    expect(anyHostCall()).toBe(false);
  });

  it('a record-bound function without one sends nothing', async () => {
    const r = await executeFunction({
      functionName: 'new_Rank',
      boundRecordId: null,
      paramValues: {},
      useParamAliases: false,
      headers: [],
      category: 'custom-api',
      dirty: new Set(),
    } as unknown as ExecuteFunctionState);
    expect(r.status).toBe(400);
    expect(anyHostCall()).toBe(false);
  });

  it('an unknown operation sends nothing', async () => {
    const r = await executeAction({
      actionName: 'new_Gone',
      boundRecordId: null,
      paramValues: {},
      headers: [],
      category: 'custom-api',
      dirty: new Set(),
    } as unknown as ExecuteActionState);
    expect(r.status).toBe(404);
    expect(anyHostCall()).toBe(false);
  });
});

describe('host calls', () => {
  it('Create sends the encoded body and the typed rows, minus the host-built ones', async () => {
    const s = createState({
      fieldValues: {
        name: 'Contoso',
        sample_tags: [100, 200],
        primarycontactid: { id: CONTACT_ID, targetEntity: 'contact' },
      },
    });
    const r = await executeCreate(s, { ...PANE, MSCRMCallerID: CALLER });
    expect(r.status).toBe(201);
    expect(host.create).toHaveBeenCalledWith('account', buildCreateBody(s), {
      'X-Custom': 'typed',
      MSCRMCallerID: CALLER,
    });
    expect(host.create.mock.calls[0][1]).toEqual({
      name: 'Contoso',
      sample_tags: '100,200',
      'primarycontactid@odata.bind': `/contacts(${CONTACT_ID})`,
    });
  });

  it('Upsert by GUID PATCHes the record with the Create-shaped body', async () => {
    const r = await executeUpsert(
      upsert({ fieldValues: { primarycontactid: { id: CONTACT_ID, targetEntity: 'contact' } } }),
    );
    expect(r.status).toBe(204);
    expect(host.update).toHaveBeenCalledWith(
      'account',
      ACCOUNT_ID,
      { 'primarycontactid@odata.bind': `/contacts(${CONTACT_ID})` },
      undefined,
    );
  });

  it('Merge runs the Merge action with the built body', async () => {
    const s = merge();
    const r = await executeMerge(s, PANE);
    expect(r.status).toBe(204);
    expect(host.execute).toHaveBeenCalledWith(
      { operationName: 'Merge', operationType: 'action', parameters: buildMergeBody(s) },
      { 'X-Custom': 'typed' },
    );
  });

  it('a host error comes back as a failed result, not a throw', async () => {
    host.create.mockRejectedValue(
      Object.assign(new Error('0x80040237: duplicate record'), { statusCode: 412 }),
    );
    const r = await executeCreate(createState({ fieldValues: { name: 'x' } }));
    expect(r.ok).toBe(false);
    expect(r.status).toBe(412);
  });
});

describe('an older ToolBox that drops headers (before 1.2.7)', () => {
  beforeEach(() => {
    host.additionalHeaders = false;
  });

  it('sends the write without headers', async () => {
    const r = await executeCreate(createState({ fieldValues: { name: 'x' } }), PANE);
    expect(r.status).toBe(201);
    expect(host.create).toHaveBeenCalledWith('account', { name: 'x' }, undefined);
  });

  it.each([
    [
      'Create',
      () => executeCreate(createState({ fieldValues: { name: 'x' } }), { MSCRMCallerID: CALLER }),
    ],
    ['Update', () => executeUpdate(update(), { MSCRMCallerID: CALLER })],
    ['Delete', () => executeDelete(del(), { MSCRMCallerID: CALLER })],
    ['Merge', () => executeMerge(merge(), { MSCRMCallerID: CALLER })],
  ])(
    "%s doesn't run as you when the impersonation header would be dropped (#109)",
    async (_n, exec) => {
      const r = await exec();
      expect(r.ok).toBe(false);
      expect(JSON.stringify(r.body)).toMatch(/can't send MSCRMCallerID/);
      expect(anyHostCall()).toBe(false);
    },
  );
});
