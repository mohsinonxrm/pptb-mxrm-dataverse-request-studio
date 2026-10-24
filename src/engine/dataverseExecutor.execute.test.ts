// Execute Action / Execute Workflow send the Headers pane's rows on PPTB
// 1.2.7+ (impersonation, bypass, …) and keep the header-free call on older
// hosts (#266).
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { HEADER_CONSENT_DECLINED_CODE, executeAction, executeWorkflow } from './dataverseExecutor';
import type { ExecuteActionState, ExecuteWorkflowState } from '../state/executeState';

// The bridge decides host capabilities once at module load; mock it so each
// test picks the host it runs against and sees every host call.
const host = vi.hoisted(() => ({
  additionalHeaders: true,
  execute: vi.fn(),
}));
vi.mock('../host/pptbBridge', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/pptbBridge')>()),
  isEmbedded: () => true,
  hostSupportsAdditionalHeaders: () => host.additionalHeaders,
  dvHost: { execute: host.execute },
}));
vi.mock('../host/csdlProvider', () => ({
  actions: {
    find: async (name: string) => ({ name, binding: { kind: 'unbound' }, parameters: [] }),
  },
}));
vi.mock('./executeBuilders', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./executeBuilders')>()),
  buildExecuteActionBody: () => ({}),
}));

const CALLER = '00000000-0000-0000-0000-0000000000c1';
// What the modes hand over: the Headers pane as a map, host-owned rows included.
const PANE = {
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
});

describe.each([
  ['Execute Action', () => executeAction(action, PANE)],
  ['Execute Workflow', () => executeWorkflow(workflow, PANE)],
])('%s', (_, run) => {
  it('sends the pane’s headers, minus the ones the host builds (PPTB 1.2.7+)', async () => {
    await run();
    expect(host.execute).toHaveBeenCalledWith(expect.any(Object), { MSCRMCallerID: CALLER });
  });

  it('keeps the header-free call on a host that would drop them (PPTB < 1.2.7)', async () => {
    host.additionalHeaders = false;
    await run();
    expect(host.execute).toHaveBeenCalledWith(expect.any(Object), undefined);
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
