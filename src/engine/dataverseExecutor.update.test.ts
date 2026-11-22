// executeUpdate's "Require existing record" sends If-Match: * with the PATCH
// and does no extra read (#218).
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { executeUpdate } from './dataverseExecutor';
import { defaultBypassOptions, type ConcurrencyMode, type UpdateState } from '../state/writeState';

// Mock the bridge so each test sees every host call.
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

beforeEach(() => {
  host.embedded = true;
  host.retrieve.mockReset().mockResolvedValue({ accountid: ID });
  host.update.mockReset().mockResolvedValue(undefined);
});

describe('executeUpdate', () => {
  it('sends If-Match: * with the PATCH and does no extra read', async () => {
    const r = await executeUpdate(state({ kind: 'update-only' }), IF_MATCH);
    expect(host.retrieve).not.toHaveBeenCalled();
    expect(host.update).toHaveBeenCalledWith('account', ID, expect.any(Object), IF_MATCH);
    expect(r.status).toBe(204);
    // The host passes no row back; no stand-in body either (#224).
    expect(r.body).toBeNull();
  });

  it('sends no precondition when it is None (upsert semantics chosen)', async () => {
    await executeUpdate(state({ kind: 'none' }));
    expect(host.retrieve).not.toHaveBeenCalled();
    expect(host.update).toHaveBeenCalledWith('account', ID, expect.any(Object), undefined);
  });
});
