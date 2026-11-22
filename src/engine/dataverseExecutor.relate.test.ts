// Single-valued Associate / Disassociate can't create the source record (#254).
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { executeAssociate, executeDisassociate } from './dataverseExecutor';
import { relateRequestHeaders } from './urlBuilder';
import type { AssociateState, DisassociateState } from '../state/relateState';

const host = vi.hoisted(() => ({
  embedded: true,
  retrieve: vi.fn(),
  update: vi.fn(),
}));
vi.mock('../host/pptbBridge', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/pptbBridge')>()),
  isEmbedded: () => host.embedded,
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
  // A real GUID: the executor refuses ids that aren't (#221).
  targets: ['00000000-0000-0000-0000-0000000000a1'],
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
  host.retrieve.mockReset().mockResolvedValue({ contactid: SRC });
  host.update.mockReset().mockResolvedValue(undefined);
});

describe.each([
  ['Associate', () => executeAssociate(assoc)],
  ['Disassociate', () => executeDisassociate(disassoc)],
])('single-valued %s', (_name, run) => {
  it('sends If-Match: *, with no extra read', async () => {
    const r = await run();
    expect(host.retrieve).not.toHaveBeenCalled();
    expect(host.update).toHaveBeenCalledWith('contact', SRC, expect.any(Object), {
      'If-Match': '*',
    });
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
    expect(JSON.stringify(r.body)).toMatch(/No Contact record has the id .*nothing was updated/);
  });

  it.each([
    [
      'the desktop app',
      "Error invoking remote method 'dataverse.retrieve': Error: 0x80040217: Entity 'contact' With Id = x Does Not Exist",
    ],
    ['the PPTB VS Code extension', 'Request failed with status code 404'],
  ])('declined: reads 404 from the error text %s rejects with', async (_host, message) => {
    host.update.mockReset().mockRejectedValueOnce(new Error('header consent declined'));
    host.retrieve.mockRejectedValue(new Error(message));
    const r = await executeAssociate(assoc);
    expect(host.update).toHaveBeenCalledTimes(1);
    expect(r.status).toBe(404);
    expect(JSON.stringify(r.body)).toMatch(/nothing was updated/);
  });

  it("declined: doesn't PATCH when the check itself fails", async () => {
    host.update.mockReset().mockRejectedValueOnce(new Error('header consent declined'));
    host.retrieve.mockRejectedValue(Object.assign(new Error('Access denied'), { statusCode: 403 }));
    const r = await executeAssociate(assoc);
    expect(host.update).toHaveBeenCalledTimes(1);
    expect(r.status).toBe(403);
    expect(JSON.stringify(r.body)).toMatch(/Nothing was updated.*Access denied/);
  });

  it("keeps the user's own If-Match, in any case", () => {
    expect(relateRequestHeaders('PATCH', { 'if-match': 'W/"5"' })).toEqual({ 'if-match': 'W/"5"' });
    expect(relateRequestHeaders('PATCH', { 'If-Match': 'W/"5"' })).toEqual({ 'If-Match': 'W/"5"' });
  });
});

describe('ids and names that would change the URL (#221)', () => {
  it.each([
    ['a source id with a path', { ...assoc, sourceId: `${SRC})/contact_customer_accounts` }],
    ['a target id that is not a GUID', { ...assoc, targets: ['tgt'] }],
    ['a navigation property with a query', { ...assoc, navProperty: 'nav?$top=1' }],
  ])('Associate sends nothing for %s', async (_name, s) => {
    const r = await executeAssociate(s as AssociateState);
    expect(host.update).not.toHaveBeenCalled();
    expect(host.retrieve).not.toHaveBeenCalled();
    expect(r.ok).toBe(false);
    expect(JSON.stringify(r.body)).toMatch(/Not sent: /);
  });

  it('Disassociate sends nothing for a bad source id', async () => {
    const r = await executeDisassociate({ ...disassoc, sourceId: 'x' } as DisassociateState);
    expect(host.update).not.toHaveBeenCalled();
    expect(r.ok).toBe(false);
  });
});
