// #109: collection-valued Associate / Disassociate send impersonation and
// bypass headers; a declined prompt is asked once and sends nothing; the
// single-valued fallback never drops the user's own headers.
import { beforeEach, describe, expect, it, vi } from 'vitest';
import {
  HEADER_CONSENT_DECLINED_CODE,
  executeAssociate,
  executeDisassociate,
} from './dataverseExecutor';
import type { AssociateState, DisassociateState } from '../state/relateState';

const host = vi.hoisted(() => ({
  single: false,
  additionalHeaders: true,
  queryData: vi.fn(),
  associate: vi.fn(),
  disassociate: vi.fn(),
  update: vi.fn(),
  retrieve: vi.fn(),
}));
vi.mock('../host/pptbBridge', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/pptbBridge')>()),
  isEmbedded: () => true,
  hostSupportsAdditionalHeaders: () => host.additionalHeaders,
  dvHost: {
    queryData: host.queryData,
    associate: host.associate,
    disassociate: host.disassociate,
    update: host.update,
    retrieve: host.retrieve,
  },
}));
vi.mock('../host/metadataProvider', () => ({
  metadata: {
    peekTable: () => ({
      logicalName: 'account',
      entitySetName: 'accounts',
      primaryKey: 'accountid',
      navigationProperties: [
        {
          name: 'nav',
          cardinality: host.single ? 'ManyToOne' : 'ManyToMany',
          targetEntity: 'account',
        },
      ],
    }),
    getTable: async () => undefined,
  },
}));
vi.mock('./urlBuilder', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./urlBuilder')>()),
  buildAssociateRequests: (s: AssociateState) =>
    s.targets.map((targetId) => ({ method: host.single ? 'PATCH' : 'POST', targetId, body: {} })),
  buildDisassociateRequests: (s: DisassociateState) =>
    host.single
      ? [{ method: 'PATCH', targetId: null, body: {} }]
      : s.targetIds.map((targetId) => ({ method: 'DELETE', targetId })),
}));

const SRC = '00000000-0000-0000-0000-00000000c0de';
const T1 = '00000000-0000-0000-0000-0000000000a1';
const T2 = '00000000-0000-0000-0000-0000000000a2';
const CALLER = '00000000-0000-0000-0000-0000000000c1';
const PANE = {
  Accept: 'application/json',
  MSCRMCallerID: CALLER,
  'MSCRM.BypassBusinessLogicExecution': 'CustomSync',
  'X-Custom': 'x',
};
const SENT = { MSCRMCallerID: CALLER, 'MSCRM.BypassBusinessLogicExecution': 'CustomSync' };
const assoc = {
  table: 'account',
  sourceId: SRC,
  navProperty: 'nav',
  targets: [T1, T2],
} as unknown as AssociateState;
const disassoc = {
  table: 'account',
  sourceId: SRC,
  navProperty: 'nav',
  targetIds: [T1, T2],
} as unknown as DisassociateState;
const declined = () => new Error('Additional headers were rejected by the user');

beforeEach(() => {
  host.single = false;
  host.additionalHeaders = true;
  host.queryData.mockReset().mockResolvedValue({ value: [] });
  host.associate.mockReset().mockResolvedValue(undefined);
  host.disassociate.mockReset().mockResolvedValue(undefined);
  host.update.mockReset().mockResolvedValue(undefined);
  host.retrieve.mockReset().mockResolvedValue({ accountid: SRC });
});

describe('collection-valued', () => {
  it('Associate sends impersonation and bypass on every target, not other rows', async () => {
    await executeAssociate(assoc, PANE);
    expect(host.associate).toHaveBeenCalledTimes(2);
    expect(host.associate).toHaveBeenCalledWith('account', SRC, 'nav', 'account', T1, SENT);
  });

  it('Disassociate does too', async () => {
    await executeDisassociate(disassoc, PANE);
    expect(host.disassociate).toHaveBeenCalledWith('account', SRC, 'nav', T2, SENT);
  });

  it('a declined prompt is reported once, and the other targets are not tried', async () => {
    host.associate.mockRejectedValue(declined());
    const r = await executeAssociate(assoc, PANE);
    expect(host.associate).toHaveBeenCalledTimes(1);
    expect((r.body as { error: { code: string } }).error.code).toBe(HEADER_CONSENT_DECLINED_CODE);
  });

  it('without allow-listed rows nothing extra is sent', async () => {
    await executeAssociate(assoc, { Accept: 'application/json' });
    expect(host.associate).toHaveBeenCalledWith('account', SRC, 'nav', 'account', T1, undefined);
  });
});

describe('single-valued', () => {
  it('adds the headers to If-Match: *', async () => {
    host.single = true;
    await executeAssociate({ ...assoc, targets: [T1] } as AssociateState, PANE);
    expect(host.update).toHaveBeenCalledWith('account', SRC, {}, { ...SENT, 'If-Match': '*' });
  });

  it("doesn't fall back to a header-free PATCH when the user's own headers were declined", async () => {
    host.single = true;
    host.update.mockRejectedValue(declined());
    const r = await executeDisassociate(disassoc, PANE);
    expect(host.update).toHaveBeenCalledTimes(1);
    expect((r.body as { error: { code: string } }).error.code).toBe(HEADER_CONSENT_DECLINED_CODE);
  });
});

describe('older ToolBox (no header pass-through)', () => {
  it("doesn't send a relate whose bypass header would be dropped", async () => {
    host.additionalHeaders = false;
    const r = await executeAssociate(assoc, {
      'MSCRM.BypassCustomPluginExecution': 'true',
    });
    expect(host.associate).not.toHaveBeenCalled();
    expect(JSON.stringify(r.body)).toMatch(/so plug-ins and flows would still run/);
  });

  it('sends one without allow-listed headers as before', async () => {
    host.additionalHeaders = false;
    await executeAssociate(assoc, { 'X-Custom': 'x' });
    expect(host.associate).toHaveBeenCalledWith('account', SRC, 'nav', 'account', T1, undefined);
  });
});
