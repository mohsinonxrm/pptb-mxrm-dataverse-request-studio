// #225 PR 2: the three requests ToolBox's typed calls can't send (PUT of one
// column, clearing one column, Upsert by alternate key) go as a one-part
// `$batch`. The real executor and bridge run against a recording
// `window.dataverseAPI.executeBatch`; each test checks the part that went
// out (method, URL, body, headers or none) and how the answer is mapped.
import { beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';

// An embedded ToolBox host, in place before any module reads it: the bridge
// detects embedding once, at load.
const host = vi.hoisted(() => {
  const dataverseAPI = {
    executeBatch: vi.fn(),
    update: vi.fn(),
    delete: vi.fn(),
    retrieve: vi.fn(),
  };
  const w = window as unknown as Record<string, unknown>;
  w.dataverseAPI = dataverseAPI;
  w.toolboxAPI = {
    utils: { getCurrentTheme: async () => 'light' },
    connections: { getActiveConnection: async () => null },
    events: { on: () => {} },
  };
  return dataverseAPI;
});

const tables = vi.hoisted(() => new Map<string, unknown>());
vi.mock('../host/metadataProvider', () => ({
  metadata: {
    peekTable: (n: string) => tables.get(n),
    getTable: async (n: string) => tables.get(n),
  },
}));

import {
  executeDelete,
  executeUpdate,
  executeUpsert,
  HEADER_CONSENT_DECLINED_CODE,
} from './dataverseExecutor';
import { composeHeaderItems } from './requestHeaders';
import { headerItemsToObject, type HeaderItem } from '../editors/HeadersEditor';
import { emptyPrefer } from '../editors/PreferEditor';
import { ACCOUNT_ID, account, contact, registerFixtureTables } from '../test/fixtures';
import {
  defaultBypassOptions,
  type ConcurrencyMode,
  type DeleteState,
  type UpdateState,
  type UpsertState,
} from '../state/writeState';

const CALLER = '00000000-0000-0000-0000-0000000000e1';

interface Part {
  method: string;
  url: string;
  contentId: string;
  headers?: Record<string, string>;
  body?: unknown;
}

/**
 * The one part of the one `executeBatch` call. No connection target, and
 * outer headers only when `outer` is given: impersonation, which Dataverse
 * ignores inside a part (live check, #225).
 */
function sentPart(outer?: Record<string, string>): Part {
  expect(host.executeBatch).toHaveBeenCalledTimes(1);
  const args = host.executeBatch.mock.calls[0];
  expect(args).toHaveLength(outer ? 3 : 2);
  expect(args[1]).toBeUndefined();
  if (outer) expect(args[2]).toEqual(outer);
  const parts = args[0] as Part[];
  expect(parts).toHaveLength(1);
  return parts[0];
}

/** What the part answers. */
const answer = (r: Record<string, unknown>) => host.executeBatch.mockResolvedValueOnce([r]);

let hid = 0;
/** The Headers pane's defaults, plus the given rows. */
const pane = (extra: Record<string, string> = {}): HeaderItem[] =>
  Object.entries({
    Accept: 'application/json',
    'OData-MaxVersion': '4.0',
    'OData-Version': '4.0',
    'Content-Type': 'application/json; charset=utf-8',
    ...extra,
  }).map(([name, value]) => ({ id: `h${++hid}`, name, value, enabled: true }));

/** The headers a mode hands Execute: its rows plus the composed panes (no Prefer). */
const composed = (s: {
  headers: HeaderItem[];
  concurrency?: ConcurrencyMode;
  bypass?: UpdateState['bypass'];
}) => headerItemsToObject(composeHeaderItems(s), null);

const put = (over: Partial<UpdateState> = {}) =>
  ({
    table: 'account',
    recordId: ACCOUNT_ID,
    method: 'PUT',
    putColumn: 'name',
    fieldValues: { name: 'Contoso' },
    nullFields: [],
    rawBody: null,
    prefer: emptyPrefer(),
    headers: pane(),
    returnSelect: [],
    concurrency: { kind: 'none' },
    bypass: defaultBypassOptions(),
    dirty: new Set(),
    ...over,
  }) as UpdateState;

const clear = (over: Partial<DeleteState> = {}) =>
  ({
    table: 'account',
    recordId: ACCOUNT_ID,
    scope: { kind: 'single-property', column: 'description' },
    concurrency: { kind: 'none' },
    headers: pane(),
    bypassCustomPlugins: false,
    bypass: defaultBypassOptions(),
    confirmText: '',
    acknowledged: false,
    dirty: new Set(),
    ...over,
  }) as DeleteState;

const byKey = (over: Partial<UpsertState> = {}) =>
  ({
    table: 'account',
    key: { kind: 'alternate', keyName: 'akey_accountnumber', keyValues: { accountnumber: 'A-1' } },
    fieldValues: { name: 'Fabrikam' },
    nullFields: [],
    rawBody: null,
    prefer: emptyPrefer(),
    headers: pane(),
    returnSelect: [],
    concurrency: { kind: 'none' },
    bypass: defaultBypassOptions(),
    dirty: new Set(),
    ...over,
  }) as UpsertState;

beforeAll(() => {
  registerFixtureTables();
  for (const t of [account, contact]) tables.set(t.logicalName, t);
});
beforeEach(() => {
  for (const f of Object.values(host)) f.mockReset();
  host.executeBatch.mockResolvedValue([{ status: 204, statusText: 'No Content', headers: {} }]);
});

/** The typed calls a raw shape must never fall back to. */
function expectNoTypedCall() {
  expect(host.update).not.toHaveBeenCalled();
  expect(host.delete).not.toHaveBeenCalled();
}

describe('Update: PUT of one column', () => {
  it('PUT <set>(<id>)/<column> with { value }, and no headers: nothing to approve', async () => {
    answer({
      status: 204,
      statusText: 'No Content',
      headers: { 'OData-Version': '4.0', 'x-ms-service-request-id': 'abc' },
      batchHeaders: { 'x-ms-ratelimit-burst-remaining-xrm-requests': '5999' },
    });
    const s = put();
    const r = await executeUpdate(s, composed(s));
    expect(sentPart()).toEqual({
      method: 'PUT',
      url: `accounts(${ACCOUNT_ID})/name`,
      contentId: '1',
      body: { value: 'Contoso' },
    });
    expectNoTypedCall();
    expect(r).toMatchObject({
      status: 204,
      statusText: 'No Content',
      ok: true,
      outcome: 'ok',
      body: null,
      transport: 'batch',
      headers: { 'odata-version': '4.0', 'x-ms-service-request-id': 'abc' },
      batchHeaders: { 'x-ms-ratelimit-burst-remaining-xrm-requests': '5999' },
    });
  });

  it('sends bypass in the part, impersonation on the outer $batch, and no If-Match', async () => {
    // Live (#225): Dataverse refuses any If-Match on a column's URL (405,
    // 0x80060888), so the Precondition pane's choice isn't sent.
    const s = put({
      concurrency: { kind: 'update-only' },
      headers: pane({ MSCRMCallerID: CALLER }),
      bypass: { ...defaultBypassOptions(), businessLogic: 'sync' },
    });
    await executeUpdate(s, composed(s));
    expect(sentPart({ MSCRMCallerID: CALLER }).headers).toEqual({
      'MSCRM.BypassBusinessLogicExecution': 'CustomSync',
    });
  });

  it('sends no Prefer: PUT has no Prefer pane, even with return=representation left on', async () => {
    const s = put({ prefer: { ...emptyPrefer(), returnRepresentation: true } });
    await executeUpdate(s, composed(s));
    expect(sentPart()).not.toHaveProperty('headers');
  });

  it('a multi-select choice goes as "1,2", not an array', async () => {
    const s = put({ putColumn: 'sample_tags', fieldValues: { sample_tags: [100, 200] } });
    await executeUpdate(s, composed(s));
    expect(sentPart()).toMatchObject({
      url: `accounts(${ACCOUNT_ID})/sample_tags`,
      body: { value: '100,200' },
    });
  });

  it("doesn't send an ETag either: there's no versioning at column level", async () => {
    const s = put({ concurrency: { kind: 'etag', etag: 'W/"1"' } });
    await executeUpdate(s, composed(s));
    expect(sentPart()).not.toHaveProperty('headers');
  });

  it('a column name that would change the URL is refused before anything is sent', async () => {
    const r = await executeUpdate(
      put({ putColumn: 'name/../x', fieldValues: { 'name/../x': 'a' } }),
    );
    expect(r.status).toBe(400);
    expect(host.executeBatch).not.toHaveBeenCalled();
  });
});

describe('Delete: clear one column', () => {
  it('DELETE <set>(<id>)/<column>, no body, no headers; never the whole-row delete', async () => {
    const s = clear();
    const r = await executeDelete(s, composed(s));
    expect(sentPart()).toEqual({
      method: 'DELETE',
      url: `accounts(${ACCOUNT_ID})/description`,
      contentId: '1',
    });
    expectNoTypedCall();
    expect(r).toMatchObject({ status: 204, ok: true, transport: 'batch' });
  });

  it('a missing record is a 404 carrying the Dataverse error', async () => {
    answer({
      status: 404,
      statusText: 'Not Found',
      headers: {},
      body: {
        error: { code: '0x80040217', message: `account With Id = ${ACCOUNT_ID} Does Not Exist` },
      },
    });
    const s = clear({ headers: pane({ MSCRMCallerID: CALLER }) });
    const r = await executeDelete(s, composed(s));
    expect(r).toMatchObject({ status: 404, outcome: '404', ok: false, transport: 'batch' });
    // The record is missing, not the impersonated user (#431).
    expect(r.impersonatedUserNotFound).toBeUndefined();
  });

  it('a 404 that names the impersonated caller says so (#431)', async () => {
    answer({
      status: 404,
      statusText: 'Not Found',
      headers: {},
      body: {
        error: { code: '0x80040217', message: `systemuser With Id = ${CALLER} Does Not Exist` },
      },
    });
    const s = clear({ headers: pane({ MSCRMCallerID: CALLER }) });
    const r = await executeDelete(s, composed(s));
    expect(sentPart({ MSCRMCallerID: CALLER })).not.toHaveProperty('headers');
    expect(r.impersonatedUserNotFound).toBe('MSCRMCallerID');
  });

  it('declining the header prompt is "not sent", as for the typed calls', async () => {
    host.executeBatch
      .mockReset()
      .mockRejectedValue(
        new Error(
          "Error invoking remote method 'dataverse.executeBatch': Error: Dataverse additional headers were rejected",
        ),
      );
    const s = clear({ bypass: { ...defaultBypassOptions(), businessLogic: 'sync' } });
    const r = await executeDelete(s, composed(s));
    expect(r.status).toBe(403);
    expect((r.body as { error: { code: string; message: string } }).error).toMatchObject({
      code: HEADER_CONSENT_DECLINED_CODE,
      message: expect.stringContaining('(MSCRM.BypassBusinessLogicExecution)'),
    });
  });

  it("doesn't send an ETag: Dataverse refuses one on a column's URL", async () => {
    const s = clear({ concurrency: { kind: 'etag', etag: 'W/"7"' } });
    await executeDelete(s, composed(s));
    expect(sentPart()).not.toHaveProperty('headers');
  });

  it('the whole-row delete still goes through the typed call', async () => {
    host.delete.mockResolvedValue(undefined);
    const r = await executeDelete(clear({ scope: { kind: 'whole-row' } }));
    expect(host.delete).toHaveBeenCalledTimes(1);
    expect(host.executeBatch).not.toHaveBeenCalled();
    expect(r.transport).toBeUndefined();
  });
});

describe('Upsert by alternate key', () => {
  it('PATCH <set>(<key>=value) with the body, and no headers by default', async () => {
    answer({
      status: 204,
      headers: {
        'odata-entityid':
          "https://contoso.crm.dynamics.com/api/data/v9.2/accounts(accountnumber='A-1')",
      },
    });
    const s = byKey();
    const r = await executeUpsert(s, composed(s));
    expect(sentPart()).toEqual({
      method: 'PATCH',
      url: "accounts(accountnumber='A-1')",
      contentId: '1',
      body: { name: 'Fabrikam' },
    });
    expectNoTypedCall();
    expect(r).toMatchObject({ status: 204, ok: true, transport: 'batch' });
    expect(r.headers['odata-entityid']).toContain("accountnumber='A-1'");
  });

  it('encodes key values as the Code tab does, and keeps a ? from being cut off', async () => {
    const s = byKey({
      key: {
        kind: 'alternate',
        keyName: 'akey_composite',
        keyValues: { name: "O'Brien & Co? #1 é", numberofemployees: '12' },
      },
    });
    await executeUpsert(s, composed(s));
    expect(sentPart().url).toBe(
      "accounts(name='O''Brien%20%26%20Co%3F%20%231%20%C3%A9',numberofemployees=12)",
    );
  });

  it('create-only sends If-None-Match: *, update-only If-Match: *', async () => {
    const createOnly = byKey({ concurrency: { kind: 'create-only' } });
    await executeUpsert(createOnly, composed(createOnly));
    expect(sentPart().headers).toEqual({ 'If-None-Match': '*' });

    host.executeBatch.mockClear();
    const updateOnly = byKey({ concurrency: { kind: 'update-only' } });
    await executeUpsert(updateOnly, composed(updateOnly));
    expect(sentPart().headers).toEqual({ 'If-Match': '*' });
  });

  it('return=representation sends Prefer and $select, and a create is a 201 with the row', async () => {
    const row = { '@odata.etag': 'W/"9"', accountid: ACCOUNT_ID, name: 'Fabrikam' };
    answer({ status: 201, statusText: 'Created', headers: { etag: 'W/"9"' }, body: row });
    const s = byKey({
      prefer: { ...emptyPrefer(), returnRepresentation: true },
      returnSelect: ['name'],
    });
    const r = await executeUpsert(s, composed(s));
    // The Prefer the pane composes, annotations included, as the Code tab shows it.
    expect(sentPart()).toMatchObject({
      url: "accounts(accountnumber='A-1')?$select=name",
      headers: {
        Prefer:
          'odata.include-annotations="OData.Community.Display.V1.FormattedValue", return=representation',
      },
    });
    expect(r).toMatchObject({
      status: 201,
      statusText: 'Created',
      ok: true,
      body: row,
      headers: { etag: 'W/"9"' },
    });
  });

  it("annotations alone don't send a Prefer: they change nothing on a 204", async () => {
    const s = byKey({ prefer: { ...emptyPrefer(), allAnnotations: true } });
    await executeUpsert(s, composed(s));
    expect(sentPart()).not.toHaveProperty('headers');
  });

  it('an update with return=representation is a 200 with the row', async () => {
    answer({ status: 200, statusText: 'OK', headers: {}, body: { accountid: ACCOUNT_ID } });
    const s = byKey({ prefer: { ...emptyPrefer(), returnRepresentation: true } });
    const r = await executeUpsert(s, composed(s));
    expect(r).toMatchObject({ status: 200, ok: true, body: { accountid: ACCOUNT_ID } });
  });

  it('a stale If-Match ETag is a real 412, with the Dataverse error, nothing thrown', async () => {
    answer({
      status: 412,
      statusText: 'Precondition Failed',
      headers: {},
      body: {
        error: { code: '0x80060882', message: "The version of the existing record doesn't match." },
      },
    });
    const s = byKey({ concurrency: { kind: 'etag', etag: 'W/"1"' } });
    const r = await executeUpsert(s, composed(s));
    expect(sentPart().headers).toEqual({ 'If-Match': 'W/"1"' });
    expect(r).toMatchObject({
      status: 412,
      statusText: 'Precondition Failed',
      ok: false,
      outcome: 'error',
      transport: 'batch',
      body: { error: { code: '0x80060882' } },
    });
  });

  it('create-only on an existing row is a 412', async () => {
    answer({
      status: 412,
      headers: {},
      body: {
        error: { code: '0x80040237', message: 'A record with matching key values already exists.' },
      },
    });
    const s = byKey({ concurrency: { kind: 'create-only' } });
    const r = await executeUpsert(s, composed(s));
    expect(r).toMatchObject({ status: 412, statusText: 'Precondition Failed', ok: false });
  });

  it('an unfilled key column is refused before anything is sent', async () => {
    const r = await executeUpsert(
      byKey({ key: { kind: 'alternate', keyName: 'akey_composite', keyValues: { name: 'x' } } }),
    );
    expect(r.status).toBe(400);
    expect(JSON.stringify(r.body)).toContain('numberofemployees');
    expect(host.executeBatch).not.toHaveBeenCalled();
  });

  it('by GUID it stays on the typed update', async () => {
    host.update.mockResolvedValue(undefined);
    const r = await executeUpsert(byKey({ key: { kind: 'guid', recordId: ACCOUNT_ID } }));
    expect(host.update).toHaveBeenCalledTimes(1);
    expect(host.executeBatch).not.toHaveBeenCalled();
    expect(r.transport).toBeUndefined();
  });
});
