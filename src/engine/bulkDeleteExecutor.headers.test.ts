// #109: the Bulk Delete Job runs as the impersonated user, when one is set.
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { executeBulkDelete } from './bulkDeleteExecutor';
import type { BulkDeleteState } from './bulkDelete';

const host = vi.hoisted(() => ({ additionalHeaders: true, execute: vi.fn() }));
vi.mock('../host/pptbBridge', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/pptbBridge')>()),
  isEmbedded: () => true,
  hostSupportsAdditionalHeaders: () => host.additionalHeaders,
  dvHost: { execute: host.execute },
}));
vi.mock('./bulkDelete', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./bulkDelete')>()),
  buildBulkDeleteBody: () => ({ body: { JobName: 'job' }, problems: [], deletesAll: false }),
}));

const CALLER = '00000000-0000-0000-0000-0000000000c1';
const state = {} as BulkDeleteState;

beforeEach(() => {
  host.additionalHeaders = true;
  host.execute.mockReset().mockResolvedValue({ JobId: 'j' });
});

describe('executeBulkDelete', () => {
  it('sends the impersonation header, and no other row', async () => {
    await executeBulkDelete(state, { MSCRMCallerID: CALLER, 'X-Custom': 'x' });
    expect(host.execute).toHaveBeenCalledWith(
      expect.objectContaining({ operationName: 'BulkDelete' }),
      {
        MSCRMCallerID: CALLER,
      },
    );
  });

  it('sends nothing extra without one', async () => {
    await executeBulkDelete(state, { Accept: 'application/json' });
    expect(host.execute).toHaveBeenCalledWith(expect.any(Object), undefined);
  });

  it("isn't sent on a ToolBox that would drop the impersonation header", async () => {
    host.additionalHeaders = false;
    const r = await executeBulkDelete(state, { MSCRMCallerID: CALLER });
    expect(host.execute).not.toHaveBeenCalled();
    expect(r.ok).toBe(false);
  });
});
