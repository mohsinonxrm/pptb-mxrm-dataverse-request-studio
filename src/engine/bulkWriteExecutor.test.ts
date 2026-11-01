// Bulk Write sends its headers on hosts that pass them, and maps errors like the other modes (#257).
import { beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';
import { __registerLiveTables, type TableMeta } from '../metadata/metadata';
import { defaultBypassOptions } from '../state/writeState';
import { newBulkRow, type BulkWriteState } from './bulkWrite';
import { executeBulkWrite } from './bulkWriteExecutor';
import { HEADER_CONSENT_DECLINED_CODE } from './dataverseExecutor';

const host = vi.hoisted(() => ({ embedded: true, additionalHeaders: true, execute: vi.fn() }));
vi.mock('../host/pptbBridge', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/pptbBridge')>()),
  isEmbedded: () => host.embedded,
  hostSupportsAdditionalHeaders: () => host.additionalHeaders,
  dvHost: { execute: host.execute },
}));

const account: TableMeta = {
  logicalName: 'account',
  entitySetName: 'accounts',
  displayName: 'Account',
  primaryKey: 'accountid',
  primaryName: 'name',
  columns: [
    { logicalName: 'accountid', displayName: 'Account', attributeType: 'Uniqueidentifier' },
    { logicalName: 'name', displayName: 'Account Name', attributeType: 'String', format: 'Text' },
  ],
  navigationProperties: [],
} as TableMeta;

beforeAll(() => __registerLiveTables([account]));

const state: BulkWriteState = {
  table: 'account',
  operation: 'CreateMultiple',
  rows: [newBulkRow({ fieldValues: { name: 'DRS E2E 1' } })],
  headers: [],
  bypass: defaultBypassOptions(),
  dirty: new Set(),
};
const HEADERS = {
  Accept: 'application/json',
  'Content-Type': 'application/json',
  'MSCRM.BypassBusinessLogicExecution': 'CustomSync',
};
const WIRE = { 'MSCRM.BypassBusinessLogicExecution': 'CustomSync' };

beforeEach(() => {
  host.additionalHeaders = true;
  host.execute.mockReset().mockResolvedValue({ Ids: ['00000000-0000-0000-0000-0000000000a1'] });
});

describe('executeBulkWrite headers (#257)', () => {
  it('passes the non-host-owned headers (bypass) on a host that sends them', async () => {
    const r = await executeBulkWrite(state, HEADERS);
    expect(host.execute).toHaveBeenCalledWith(
      expect.objectContaining({ operationName: 'CreateMultiple', entityName: 'account' }),
      WIRE,
    );
    expect(r.status).toBe(200);
  });

  it('sends nothing extra on a host that can’t (the advisory blocks bypass there)', async () => {
    host.additionalHeaders = false;
    await executeBulkWrite(state, HEADERS);
    expect(host.execute).toHaveBeenCalledWith(expect.any(Object), undefined);
  });

  it('reports a declined header prompt as not sent', async () => {
    host.execute.mockRejectedValue(
      new Error('Dataverse additional headers were rejected by the user'),
    );
    const r = await executeBulkWrite(state, HEADERS);
    expect(r.status).toBe(403);
    expect(JSON.stringify(r.body)).toContain(HEADER_CONSENT_DECLINED_CODE);
  });
});

describe('executeBulkWrite errors', () => {
  it('maps the real status and strips the IPC wrapper', async () => {
    host.execute.mockRejectedValue(
      new Error(
        "Error invoking remote method 'dataverse.execute': Error: 0x80040217: Entity 'account' With Id = x Does Not Exist",
      ),
    );
    const r = await executeBulkWrite(state);
    expect(r.status).toBe(404);
    expect(r.outcome).toBe('404');
    const message = (r.body as { error: { message: string } }).error.message;
    expect(message).not.toMatch(/Error invoking remote method/);
    expect(message).toMatch(/Does Not Exist/);
  });

  it("doesn't label an unknown failure as 404", async () => {
    host.execute.mockRejectedValue(new Error('socket hang up'));
    const r = await executeBulkWrite(state);
    expect(r.ok).toBe(false);
    expect(r.outcome).not.toBe('404');
  });
});
