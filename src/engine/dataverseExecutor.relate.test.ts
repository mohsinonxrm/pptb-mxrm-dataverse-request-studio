// Single-valued Associate / Disassociate can't create the source record (#254).
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { executeAssociate, executeDisassociate } from './dataverseExecutor';
import { relateRequestHeaders } from './urlBuilder';
import type { AssociateState, DisassociateState } from '../state/relateState';

const host = vi.hoisted(() => ({
  embedded: true,
  additionalHeaders: true,
  retrieve: vi.fn(),
  update: vi.fn(),
}));
vi.mock('../host/pptbBridge', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/pptbBridge')>()),
  isEmbedded: () => host.embedded,
  hostSupportsAdditionalHeaders: () => host.additionalHeaders,
  dvHost: { retrieve: host.retrieve, update: host.update },
}));
vi.mock('../host/metadataProvider', () => ({
  metadata: {
    peekTable: () => ({
      logicalName: 'contact',
      primaryKey: 'contactid',
      displayName: 'Contact',
      navigationProperties: [
        { name: 'parentcustomerid_account', cardinality: 'ManyToOne', targetEntity: 'account' },
      ],
    }),
    getTable: async () => undefined,
  },
}));
// The request builders read the live table registry; the executor only needs their output.
vi.mock('./urlBuilder', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./urlBuilder')>()),
  buildAssociateRequests: () => [
    {
      method: 'PATCH',
      relativeUrl: '/api/data/v9.2/contacts(src)',
      relativeNoBase: '/contacts(src)',
      body: { 'parentcustomerid_account@odata.bind': 'accounts(tgt)' },
      targetId: 'tgt',
    },
  ],
  buildDisassociateRequests: () => [
    {
      method: 'PATCH',
      relativeUrl: '/api/data/v9.2/contacts(src)',
      relativeNoBase: '/contacts(src)',
      targetId: null,
      body: { 'parentcustomerid_account@odata.bind': null },
    },
  ],
}));

const SRC = '00000000-0000-0000-0000-00000000c0de';
const assoc = {
  table: 'contact',
  sourceId: SRC,
  navProperty: 'parentcustomerid_account',
  targets: ['tgt'],
} as unknown as AssociateState;
const disassoc = {
  table: 'contact',
  sourceId: SRC,
  navProperty: 'parentcustomerid_account',
  targetIds: [],
} as unknown as DisassociateState;
const missing = () => new Error("0x80040217: Entity 'contact' With Id = x Does Not Exist");

beforeEach(() => {
  host.embedded = true;
  host.additionalHeaders = true;
  host.retrieve.mockReset().mockResolvedValue({ contactid: SRC });
  host.update.mockReset().mockResolvedValue(undefined);
});

describe.each([
  ['Associate', () => executeAssociate(assoc)],
  ['Disassociate', () => executeDisassociate(disassoc)],
])('single-valued %s', (_name, run) => {
  it('sends If-Match: * on a host that passes headers, with no extra read', async () => {
    const r = await run();
    expect(host.retrieve).not.toHaveBeenCalled();
    expect(host.update).toHaveBeenCalledWith('contact', SRC, expect.any(Object), {
      'If-Match': '*',
    });
    expect(r.status).toBe(204);
  });

  it("checks the source exists first on a host that can't, and doesn't PATCH when it's missing", async () => {
    host.additionalHeaders = false;
    host.retrieve.mockRejectedValue(missing());
    const r = await run();
    expect(host.retrieve).toHaveBeenCalledWith('contact', SRC, ['contactid']);
    expect(host.update).not.toHaveBeenCalled();
    expect(r.status).toBe(404);
  });

  it('updates an existing source on a host that can’t send headers', async () => {
    host.additionalHeaders = false;
    const r = await run();
    expect(host.update).toHaveBeenCalledWith('contact', SRC, expect.any(Object), undefined);
    expect(r.status).toBe(204);
  });
});

describe('relateRequestHeaders', () => {
  it('adds If-Match: * to the PATCH form only', () => {
    expect(relateRequestHeaders('PATCH', { Accept: 'x' })).toEqual({
      Accept: 'x',
      'If-Match': '*',
    });
    expect(relateRequestHeaders('POST', { Accept: 'x' })).toEqual({ Accept: 'x' });
    expect(relateRequestHeaders('DELETE', {})).toEqual({});
  });
});

describe('review follow-ups (#254)', () => {
  it('falls back to reading the source first when the header prompt is declined', async () => {
    host.update
      .mockReset()
      .mockRejectedValueOnce(new Error('Dataverse additional headers were rejected by the user'))
      .mockResolvedValueOnce(undefined);
    const r = await executeAssociate(assoc);
    expect(host.retrieve).toHaveBeenCalledWith('contact', SRC, ['contactid']);
    expect(host.update).toHaveBeenLastCalledWith('contact', SRC, expect.any(Object));
    expect(r.status).toBe(204);
  });

  it('declined + missing source: no PATCH without the header', async () => {
    host.update.mockReset().mockRejectedValueOnce(new Error('header consent declined'));
    host.retrieve.mockRejectedValue(missing());
    const r = await executeDisassociate(disassoc);
    expect(host.update).toHaveBeenCalledTimes(1);
    expect(r.status).toBe(404);
  });

  it("keeps the user's own If-Match, in any case", () => {
    expect(relateRequestHeaders('PATCH', { 'if-match': 'W/"5"' })).toEqual({ 'if-match': 'W/"5"' });
    expect(relateRequestHeaders('PATCH', { 'If-Match': 'W/"5"' })).toEqual({ 'If-Match': 'W/"5"' });
  });
});
