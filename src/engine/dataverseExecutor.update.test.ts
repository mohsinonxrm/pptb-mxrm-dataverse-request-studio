// executeUpdate's "Require existing record" on hosts with and without header
// pass-through (#218).
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { executeUpdate } from './dataverseExecutor';
import { defaultBypassOptions, type ConcurrencyMode, type UpdateState } from '../state/writeState';

// The bridge decides host capabilities once at module load; mock it so each
// test picks the host it runs against and sees every host call.
const host = vi.hoisted(() => ({
  embedded: true,
  additionalHeaders: false,
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
    peekTable: () => ({ logicalName: 'account', primaryKey: 'accountid', displayName: 'Account' }),
    getTable: async () => undefined,
  },
}));

const ID = '00000000-0000-0000-0000-0000000000aa';
const IF_MATCH = { 'If-Match': '*' };

function state(concurrency: ConcurrencyMode): UpdateState {
  return {
    table: 'account',
    recordId: ID,
    method: 'PATCH',
    putColumn: null,
    fieldValues: {},
    nullFields: [],
    rawBody: null,
    prefer: { returnRepresentation: false, includeAnnotations: false, maxPageSize: null },
    headers: [],
    returnSelect: [],
    concurrency,
    bypass: defaultBypassOptions(),
    dirty: new Set(),
  } as unknown as UpdateState;
}

const notFound = Object.assign(new Error("Entity 'account' With Id = … Does Not Exist"), {
  statusCode: 404,
});

beforeEach(() => {
  host.embedded = true;
  host.additionalHeaders = false;
  host.retrieve.mockReset().mockResolvedValue({ accountid: ID });
  host.update.mockReset().mockResolvedValue(undefined);
});

describe('executeUpdate on a host without header pass-through (PPTB < 1.2.7)', () => {
  it('reads the record first and updates it when it exists', async () => {
    const r = await executeUpdate(state({ kind: 'update-only' }), IF_MATCH);
    expect(host.retrieve).toHaveBeenCalledWith('account', ID, ['accountid']);
    expect(host.update).toHaveBeenCalledTimes(1);
    expect(r.status).toBe(204);
  });

  it("doesn't send the PATCH when the record is missing", async () => {
    host.retrieve.mockRejectedValue(notFound);
    const r = await executeUpdate(state({ kind: 'update-only' }), IF_MATCH);
    expect(host.update).not.toHaveBeenCalled();
    expect(r.status).toBe(404);
    expect(r.ok).toBe(false);
    expect(JSON.stringify(r.body)).toMatch(/No Account record has the id .*nothing was updated/);
  });

  it.each([
    [
      'PPTB < 1.2.7',
      "Error invoking remote method 'dataverse.retrieve': Error: 0x80040217: Entity 'account' With Id = x Does Not Exist",
    ],
    ['the PPTB VS Code extension', 'Request failed with status code 404'],
  ])('reads 404 from the error text %s rejects with', async (_host, message) => {
    host.retrieve.mockRejectedValue(new Error(message));
    const r = await executeUpdate(state({ kind: 'update-only' }), IF_MATCH);
    expect(host.update).not.toHaveBeenCalled();
    expect(r.status).toBe(404);
    expect(JSON.stringify(r.body)).toMatch(/nothing was updated/);
  });

  it("doesn't send the PATCH when the check itself fails", async () => {
    host.retrieve.mockRejectedValue(Object.assign(new Error('Access denied'), { statusCode: 403 }));
    const r = await executeUpdate(state({ kind: 'update-only' }), IF_MATCH);
    expect(host.update).not.toHaveBeenCalled();
    expect(r.status).toBe(403);
    expect(JSON.stringify(r.body)).toMatch(/Nothing was updated.*Access denied/);
  });

  it('skips the check when the precondition is None (upsert semantics chosen)', async () => {
    await executeUpdate(state({ kind: 'none' }));
    expect(host.retrieve).not.toHaveBeenCalled();
    expect(host.update).toHaveBeenCalledTimes(1);
  });
});

describe('executeUpdate on a host with header pass-through (PPTB 1.2.7+)', () => {
  it('sends If-Match: * with the PATCH and does no extra read', async () => {
    host.additionalHeaders = true;
    const r = await executeUpdate(state({ kind: 'update-only' }), IF_MATCH);
    expect(host.retrieve).not.toHaveBeenCalled();
    expect(host.update).toHaveBeenCalledWith('account', ID, expect.any(Object), IF_MATCH);
    expect(r.status).toBe(204);
  });
});
