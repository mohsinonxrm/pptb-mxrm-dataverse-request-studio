// #224: All or nothing — collection-valued Associate / Disassociate send
// their targets as one change set (`executeTransaction`) when the option is
// on and the host has it, map the host's results back per target, and keep
// the one-request-per-target path otherwise (option off, one target, or a
// ToolBox older than 1.2.7).
import { beforeEach, describe, expect, it, vi } from 'vitest';
import {
  HEADER_CONSENT_DECLINED_CODE,
  executeAssociate,
  executeDisassociate,
  relateAsChangeSet,
  relateDoneCount,
} from './dataverseExecutor';
import type { AssociateState, DisassociateState } from '../state/relateState';

const host = vi.hoisted(() => ({
  embedded: true,
  batch: true,
  linked: [] as string[],
  queryData: vi.fn(),
  associate: vi.fn(),
  disassociate: vi.fn(),
  executeTransaction: vi.fn(),
}));
vi.mock('../host/pptbBridge', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/pptbBridge')>()),
  isEmbedded: () => host.embedded,
  hostSupportsAdditionalHeaders: () => host.batch,
  hostSupportsBatch: () => host.batch,
  dvHost: {
    queryData: host.queryData,
    associate: host.associate,
    disassociate: host.disassociate,
    executeTransaction: host.executeTransaction,
  },
}));
vi.mock('../host/metadataProvider', () => ({
  metadata: {
    peekTable: () => ({
      logicalName: 'account',
      entitySetName: 'accounts',
      primaryKey: 'accountid',
      navigationProperties: [{ name: 'nav', cardinality: 'ManyToMany', targetEntity: 'account' }],
    }),
    getTable: async () => undefined,
  },
}));
const ORG = 'https://contoso.crm.dynamics.com/api/data/v9.2';
vi.mock('./urlBuilder', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./urlBuilder')>()),
  buildAssociateRequests: (s: AssociateState) =>
    s.targets.map((targetId) => ({
      method: 'POST',
      relativeUrl: `/api/data/v9.2/accounts(${s.sourceId})/nav/$ref`,
      relativeNoBase: `/accounts(${s.sourceId})/nav/$ref`,
      body: { '@odata.id': `${ORG}/accounts(${targetId})` },
      targetId,
    })),
  buildDisassociateRequests: (s: DisassociateState) =>
    s.targetIds.map((targetId) => ({
      method: 'DELETE',
      relativeUrl: `/api/data/v9.2/accounts(${s.sourceId})/nav(${targetId})/$ref`,
      relativeNoBase: `/accounts(${s.sourceId})/nav(${targetId})/$ref`,
      targetId,
    })),
}));

const SRC = '00000000-0000-0000-0000-00000000c0de';
const T1 = '00000000-0000-0000-0000-0000000000a1';
const T2 = '00000000-0000-0000-0000-0000000000a2';
const T3 = '00000000-0000-0000-0000-0000000000a3';
const CALLER = '00000000-0000-0000-0000-0000000000c1';
const PANE = { Accept: 'application/json', MSCRMCallerID: CALLER, 'X-Custom': 'x' };

const assoc = (targets: string[], allOrNothing = true) =>
  ({
    table: 'account',
    sourceId: SRC,
    navProperty: 'nav',
    targets,
    allOrNothing,
  }) as unknown as AssociateState;
const disassoc = (targetIds: string[], allOrNothing = true) =>
  ({
    table: 'account',
    sourceId: SRC,
    navProperty: 'nav',
    targetIds,
    allOrNothing,
  }) as unknown as DisassociateState;

const ok204 = (contentId: string) => ({
  contentId,
  status: 204,
  statusText: 'No Content',
  headers: {},
});

beforeEach(() => {
  host.embedded = true;
  host.batch = true;
  host.linked = [];
  host.queryData
    .mockReset()
    .mockImplementation(async () => ({ value: host.linked.map((accountid) => ({ accountid })) }));
  host.associate.mockReset().mockResolvedValue(undefined);
  host.disassociate.mockReset().mockResolvedValue(undefined);
  host.executeTransaction
    .mockReset()
    .mockImplementation(async (reqs: { contentId: string }[]) =>
      reqs.map((r) => ok204(r.contentId)),
    );
});

describe('relateAsChangeSet', () => {
  it('needs the option, two or more targets, and a host that can send a change set', () => {
    expect(relateAsChangeSet(true, 2)).toBe(true);
    expect(relateAsChangeSet(undefined, 2)).toBe(false);
    expect(relateAsChangeSet(true, 1)).toBe(false);
    host.batch = false;
    expect(relateAsChangeSet(true, 2)).toBe(false);
    // Standalone: the Code tab shows the change set (Execute reports 501).
    host.embedded = false;
    expect(relateAsChangeSet(true, 2)).toBe(true);
  });
});

describe('Associate, all or nothing', () => {
  it('sends every target as one change set, headers on each request', async () => {
    const r = await executeAssociate(assoc([T1, T2]), PANE);
    expect(host.associate).not.toHaveBeenCalled();
    expect(host.executeTransaction).toHaveBeenCalledTimes(1);
    expect(host.executeTransaction.mock.calls[0]).toEqual([
      [T1, T2].map((t, i) => ({
        method: 'POST',
        url: `accounts(${SRC})/nav/$ref`,
        body: { '@odata.id': `${ORG}/accounts(${t})` },
        headers: { MSCRMCallerID: CALLER },
        contentId: String(i + 1),
      })),
    ]);
    expect(r).toMatchObject({ ok: true, status: 204, changeSet: true });
    expect(r.targets).toEqual([
      { targetId: T1, state: 'done' },
      { targetId: T2, state: 'done' },
    ]);
    expect(relateDoneCount(r, 2)).toBe(2);
  });

  it('leaves already-linked targets out of the change set', async () => {
    host.linked = [T1];
    const r = await executeAssociate(assoc([T1, T2, T3]));
    const sent = host.executeTransaction.mock.calls[0][0] as { url: string; contentId: string }[];
    expect(sent.map((q) => q.contentId)).toEqual(['1', '2']);
    expect(r.ok).toBe(true);
    expect(r.targets!.map((t) => t.state)).toEqual(['skipped', 'done', 'done']);
  });

  it('sends nothing when every target is already linked', async () => {
    host.linked = [T1, T2];
    const r = await executeAssociate(assoc([T1, T2]));
    expect(host.executeTransaction).not.toHaveBeenCalled();
    expect(r.ok).toBe(true);
    expect(r.targets!.map((t) => t.state)).toEqual(['skipped', 'skipped']);
  });

  it('a failed request rolls the rest back: its status, nothing partial', async () => {
    host.executeTransaction.mockResolvedValue([
      {
        contentId: '2',
        status: 404,
        statusText: 'Not Found',
        headers: {},
        body: { error: { code: '0x80040217', message: 'account With Id = x Does Not Exist' } },
      },
    ]);
    const r = await executeAssociate(assoc([T1, T2, T3]));
    expect(r).toMatchObject({ ok: false, status: 404, outcome: '404', changeSet: true });
    expect(r.targets).toEqual([
      { targetId: T1, state: 'rolled-back' },
      {
        targetId: T2,
        state: 'failed',
        status: 404,
        statusText: 'Not Found',
        error: { code: '0x80040217', message: 'account With Id = x Does Not Exist' },
      },
      { targetId: T3, state: 'rolled-back' },
    ]);
    expect((r.body as { error: { code: string } }).error.code).toBe('0x80040217');
    expect(relateDoneCount(r, 3)).toBeNull();
  });

  it("marks every target rolled back when the failure can't be tied to one", async () => {
    host.executeTransaction.mockResolvedValue([
      { status: 400, statusText: 'Bad Request', headers: {}, body: { error: { message: 'Bad' } } },
    ]);
    const r = await executeAssociate(assoc([T1, T2]));
    expect(r).toMatchObject({ ok: false, status: 400 });
    expect(r.targets!.map((t) => t.state)).toEqual(['rolled-back', 'rolled-back']);
    expect((r.body as { error: { message: string } }).error.message).toBe('Bad');
  });

  it('a host error means nothing applied', async () => {
    host.executeTransaction.mockRejectedValue(new Error('HTTP 500: boom'));
    const r = await executeAssociate(assoc([T1, T2]));
    expect(r).toMatchObject({ ok: false, status: 500, changeSet: true });
    expect(r.targets!.map((t) => t.state)).toEqual(['rolled-back', 'rolled-back']);
  });

  it('a declined header prompt reads as not sent', async () => {
    host.executeTransaction.mockRejectedValue(
      new Error('Dataverse additional headers were rejected by the user'),
    );
    const r = await executeAssociate(assoc([T1, T2]), PANE);
    expect((r.body as { error: { code: string } }).error.code).toBe(HEADER_CONSENT_DECLINED_CODE);
    expect(r.targets).toBeUndefined();
  });

  it('standalone, reports 501 like every other mode', async () => {
    host.embedded = false;
    host.batch = false;
    host.executeTransaction.mockImplementation(async () => {
      const { HostNotAvailableError } = await import('../host/pptbBridge');
      throw new HostNotAvailableError('window.dataverseAPI');
    });
    const r = await executeAssociate(assoc([T1, T2]));
    expect(r.status).toBe(501);
  });
});

describe('Associate, one request per target', () => {
  it('is the default', async () => {
    const r = await executeAssociate(assoc([T1, T2], false));
    expect(host.executeTransaction).not.toHaveBeenCalled();
    expect(host.associate).toHaveBeenCalledTimes(2);
    expect(r.changeSet).toBeUndefined();
  });

  it('is used for a single target', async () => {
    await executeAssociate(assoc([T1]));
    expect(host.executeTransaction).not.toHaveBeenCalled();
    expect(host.associate).toHaveBeenCalledTimes(1);
  });

  it('is the fallback on a ToolBox without executeTransaction', async () => {
    host.batch = false;
    host.associate.mockResolvedValueOnce(undefined).mockRejectedValueOnce(new Error('HTTP 404: x'));
    const r = await executeAssociate(assoc([T1, T2]));
    expect(host.executeTransaction).not.toHaveBeenCalled();
    expect(host.associate).toHaveBeenCalledTimes(2);
    // The old behaviour: the first stays linked, the run is partial.
    expect(r.outcome).toBe('partial');
    expect(r.targets!.map((t) => t.state)).toEqual(['done', 'failed']);
  });
});

describe('Disassociate, all or nothing', () => {
  it('sends every DELETE as one change set', async () => {
    const r = await executeDisassociate(disassoc([T1, T2]), PANE);
    expect(host.disassociate).not.toHaveBeenCalled();
    expect(host.executeTransaction.mock.calls[0]).toEqual([
      [T1, T2].map((t, i) => ({
        method: 'DELETE',
        url: `accounts(${SRC})/nav(${t})/$ref`,
        headers: { MSCRMCallerID: CALLER },
        contentId: String(i + 1),
      })),
    ]);
    expect(r).toMatchObject({ ok: true, status: 204, changeSet: true });
    expect(r.targets!.map((t) => t.state)).toEqual(['done', 'done']);
  });

  it('a failed DELETE rolls the other back', async () => {
    host.executeTransaction.mockResolvedValue([
      { contentId: '1', status: 404, statusText: 'Not Found', headers: {}, body: undefined },
    ]);
    const r = await executeDisassociate(disassoc([T1, T2]));
    expect(r).toMatchObject({ ok: false, status: 404 });
    expect(r.targets!.map((t) => t.state)).toEqual(['failed', 'rolled-back']);
    // No error envelope: the status line stands in for the message.
    expect(r.targets![0].error).toEqual({ message: 'Not Found' });
  });

  it('falls back to one DELETE per target on an older ToolBox', async () => {
    host.batch = false;
    await executeDisassociate(disassoc([T1, T2]));
    expect(host.executeTransaction).not.toHaveBeenCalled();
    expect(host.disassociate).toHaveBeenCalledTimes(2);
  });
});
