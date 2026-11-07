// Execute Action / Execute Workflow send the Headers pane's rows on PPTB
// 1.2.7+ (impersonation, bypass, …) and keep the header-free call on older
// hosts (#266).
import { beforeEach, describe, expect, it, vi } from 'vitest';
import {
  HEADER_CONSENT_DECLINED_CODE,
  executeAction,
  executeFunction,
  executeWorkflow,
} from './dataverseExecutor';
import type {
  ExecuteActionState,
  ExecuteFunctionState,
  ExecuteWorkflowState,
} from '../state/executeState';

// The bridge decides host capabilities once at module load; mock it so each
// test picks the host it runs against and sees every host call.
const host = vi.hoisted(() => ({
  additionalHeaders: true,
  execute: vi.fn(),
  queryData: vi.fn(),
}));
vi.mock('../host/pptbBridge', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/pptbBridge')>()),
  isEmbedded: () => true,
  hostSupportsAdditionalHeaders: () => host.additionalHeaders,
  dvHost: { execute: host.execute, queryData: host.queryData },
}));
vi.mock('../host/csdlProvider', () => ({
  actions: {
    find: async (name: string) => ({ name, binding: { kind: 'unbound' }, parameters: [] }),
  },
}));
vi.mock('./executeBuilders', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./executeBuilders')>()),
  buildExecuteActionBody: () => ({}),
  buildExecuteFunctionWirePath: () => 'WhoAmI()',
}));

const CALLER = '00000000-0000-0000-0000-0000000000c1';
// What the modes hand over: the Headers pane as a map, host-owned rows included.
const PANE: Record<string, string> = {
  Accept: 'application/json',
  'Content-Type': 'application/json; charset=utf-8',
  'OData-MaxVersion': '4.0',
  'OData-Version': '4.0',
  MSCRMCallerID: CALLER,
};

const action = {
  actionName: 'new_DoThing',
  boundRecordId: null,
  paramValues: {},
  headers: [],
  category: 'custom-api',
  dirty: new Set(),
} as unknown as ExecuteActionState;

const workflow = {
  workflowId: '00000000-0000-0000-0000-0000000000f1',
  entityId: '00000000-0000-0000-0000-0000000000a1',
  headers: [],
  dirty: new Set(),
} as unknown as ExecuteWorkflowState;

beforeEach(() => {
  host.additionalHeaders = true;
  host.execute.mockReset().mockResolvedValue({});
  host.queryData.mockReset().mockResolvedValue({ value: [] });
});

describe.each([
  ['Execute Action', (h = PANE) => executeAction(action, h)],
  ['Execute Workflow', (h = PANE) => executeWorkflow(workflow, h)],
])('%s', (_, run) => {
  it('sends the pane’s headers, minus the ones the host builds (PPTB 1.2.7+)', async () => {
    await run();
    expect(host.execute).toHaveBeenCalledWith(expect.any(Object), { MSCRMCallerID: CALLER });
  });

  it('keeps the header-free call on a host that would drop them (PPTB < 1.2.7)', async () => {
    host.additionalHeaders = false;
    const { MSCRMCallerID: _caller, ...withoutCaller } = PANE;
    await run({ ...withoutCaller, 'MSCRM.BypassBusinessLogicExecution': 'CustomSync' });
    expect(host.execute).toHaveBeenCalledWith(expect.any(Object), undefined);
  });

  it("doesn't run as you when that host would drop an impersonation header (#109)", async () => {
    host.additionalHeaders = false;
    const r = await run();
    expect(host.execute).not.toHaveBeenCalled();
    expect(r.ok).toBe(false);
    expect(JSON.stringify(r.body)).toMatch(
      /can't send MSCRMCallerID, so the request would run as you/,
    );
  });

  it('reports a declined header prompt as not sent', async () => {
    host.execute.mockRejectedValue(new Error('Additional headers were rejected by the user'));
    const r = await run();
    expect(r.ok).toBe(false);
    expect((r.body as { error: { code: string } }).error.code).toBe(HEADER_CONSENT_DECLINED_CODE);
  });
});

it.each([
  ['Execute Action', (h: Record<string, string>) => executeAction(action, h)],
  ['Execute Workflow', (h: Record<string, string>) => executeWorkflow(workflow, h)],
])('%s without extra rows stays header-free, so it never prompts', async (_, run) => {
  const { MSCRMCallerID: _caller, ...builtins } = PANE;
  await run(builtins);
  expect(host.execute).toHaveBeenCalledWith(expect.any(Object), undefined);
});

it('never sends credential rows', async () => {
  await executeAction(action, { ...PANE, Cookie: 'a=b', Authorization: 'Bearer x' });
  expect(host.execute).toHaveBeenCalledWith(expect.any(Object), { MSCRMCallerID: CALLER });
});

describe('Execute Function (#109)', () => {
  const fn = {
    functionName: 'WhoAmI',
    boundRecordId: null,
    paramValues: {},
    headers: [],
    dirty: new Set(),
  } as unknown as ExecuteFunctionState;

  it('sends the impersonation header, and no other row', async () => {
    await executeFunction(fn, { ...PANE, 'MSCRM.BypassBusinessLogicExecution': 'CustomSync' });
    expect(host.queryData).toHaveBeenCalledWith('WhoAmI()', { MSCRMCallerID: CALLER });
  });

  it('sends nothing without one (no consent prompt)', async () => {
    await executeFunction(fn, { Accept: 'application/json' });
    expect(host.queryData).toHaveBeenCalledWith('WhoAmI()', undefined);
  });
});
