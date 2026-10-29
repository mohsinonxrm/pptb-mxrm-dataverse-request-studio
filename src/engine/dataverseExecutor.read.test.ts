// What the read modes hand the host for `$top` and `Prefer: odata.maxpagesize`
// on hosts with and without header pass-through (#219).
import { beforeEach, describe, expect, it, vi } from 'vitest';
import {
  HEADER_CONSENT_DECLINED_CODE,
  executeAbsoluteUrl,
  executeNextLink,
  executePredefined,
  executeRetrieveMultiple,
  readPreferHeaders,
} from './dataverseExecutor';
import { makeExpandSpec } from '../editors/ExpandEditor';
import { emptyPrefer } from '../editors/PreferEditor';
import type { PredefinedQueryState, RetrieveNextLinkState } from '../state/readState';
import { registerFixtureTables, retrieveMultipleState } from '../test/fixtures';

// The bridge decides host capabilities once at module load; mock it so each
// test picks the host it runs against and sees every host call.
const host = vi.hoisted(() => ({
  additionalHeaders: true,
  queryData: vi.fn(),
}));
vi.mock('../host/pptbBridge', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/pptbBridge')>()),
  isEmbedded: () => true,
  hostSupportsAdditionalHeaders: () => host.additionalHeaders,
  dvHost: { queryData: host.queryData },
}));
vi.mock('../host/metadataProvider', () => ({
  metadata: {
    peekTable: () => ({ logicalName: 'account', entitySetName: 'accounts' }),
    getTable: async () => undefined,
  },
}));

const QUERY_ID = '00000000-0000-0000-0000-0000000000b1';
const PAGE_50 = { Prefer: 'odata.include-annotations="*",odata.maxpagesize=50' };

const paged = (top: number | null, maxpagesize: number | null) =>
  retrieveMultipleState({ top, prefer: { ...emptyPrefer(), maxpagesize } });

function predefined(top: number | null, maxpagesize: number | null): PredefinedQueryState {
  return {
    table: 'account',
    queryId: QUERY_ID,
    queryType: 'savedQuery',
    top,
    prefer: { ...emptyPrefer(), maxpagesize },
    headers: [],
    dirty: new Set(),
  } as unknown as PredefinedQueryState;
}

beforeEach(() => {
  registerFixtureTables();
  host.additionalHeaders = true;
  host.queryData.mockReset().mockResolvedValue({ value: [] });
});

describe('readPreferHeaders', () => {
  it('asks for the page size plus every annotation, since it replaces the host Prefer', () => {
    expect(readPreferHeaders(50)).toEqual(PAGE_50);
  });

  it('sends nothing without a page size, so the read never prompts', () => {
    expect(readPreferHeaders(null)).toBeUndefined();
    expect(readPreferHeaders(0)).toBeUndefined();
  });

  it('sends nothing on a host that would drop it (PPTB < 1.2.7)', () => {
    host.additionalHeaders = false;
    expect(readPreferHeaders(50)).toBeUndefined();
  });
});

describe('executeRetrieveMultiple', () => {
  it('sends the page size and leaves $top out, as the Code tab does (PPTB 1.2.7+)', async () => {
    await executeRetrieveMultiple(paged(10, 50));
    expect(host.queryData).toHaveBeenCalledWith('/accounts', PAGE_50);
  });

  it('keeps $top when the host can’t send the page size (PPTB < 1.2.7)', async () => {
    host.additionalHeaders = false;
    await executeRetrieveMultiple(paged(10, 50));
    expect(host.queryData).toHaveBeenCalledWith('/accounts?$top=10', undefined);
  });

  it('still drops $top where Dataverse rejects it (nested $expand over a collection)', async () => {
    host.additionalHeaders = false;
    const nested = {
      ...makeExpandSpec('contact_customer_accounts'),
      nestedExpand: [makeExpandSpec('parentcustomerid_contact')],
    };
    await executeRetrieveMultiple({ ...paged(10, 50), expand: [nested] });
    expect(host.queryData).toHaveBeenCalledWith(expect.not.stringContaining('$top'), undefined);
  });

  it('stays header-free without a page size', async () => {
    await executeRetrieveMultiple(paged(10, null));
    expect(host.queryData).toHaveBeenCalledWith('/accounts?$top=10', undefined);
  });

  it('reports a declined Prefer prompt as not sent', async () => {
    host.queryData.mockRejectedValue(new Error('Additional headers were rejected by the user'));
    const r = await executeRetrieveMultiple(paged(null, 50));
    expect(r.ok).toBe(false);
    expect((r.body as { error: { code: string; message: string } }).error).toMatchObject({
      code: HEADER_CONSENT_DECLINED_CODE,
      message: expect.stringContaining('(Prefer)'),
    });
  });
});

describe('next pages', () => {
  const NEXT =
    'https://contoso.crm.dynamics.com/api/data/v9.2/accounts?$skiptoken=%3Ccookie%20pagenumber=%222%22%20/%3E';

  it('asks for the first page’s page size again', async () => {
    await executeAbsoluteUrl(NEXT, 50);
    expect(host.queryData).toHaveBeenCalledWith(
      'accounts?$skiptoken=%3Ccookie%20pagenumber=%222%22%20/%3E',
      PAGE_50,
    );
  });

  it('sends nothing when the first page had no page size', async () => {
    await executeAbsoluteUrl(NEXT, null);
    expect(host.queryData).toHaveBeenCalledWith(expect.any(String), undefined);
  });

  it('Retrieve NextLink sends the Prefer pane’s page size', async () => {
    const s = { url: NEXT, prefer: { ...emptyPrefer(), maxpagesize: 50 } };
    await executeNextLink(s as unknown as RetrieveNextLinkState);
    expect(host.queryData).toHaveBeenCalledWith(expect.any(String), PAGE_50);
  });

  it('Retrieve NextLink sends no page size by default (no prompt, the link’s own size)', async () => {
    const s = { url: NEXT, prefer: emptyPrefer() };
    await executeNextLink(s as unknown as RetrieveNextLinkState);
    expect(host.queryData).toHaveBeenCalledWith(expect.any(String), undefined);
  });
});

describe('executePredefined', () => {
  it('sends the page size and leaves $top out (PPTB 1.2.7+)', async () => {
    await executePredefined(predefined(10, 50));
    expect(host.queryData).toHaveBeenCalledWith(`accounts?savedQuery=${QUERY_ID}`, PAGE_50);
  });

  it('keeps $top when the host can’t send the page size (PPTB < 1.2.7)', async () => {
    host.additionalHeaders = false;
    await executePredefined(predefined(10, 50));
    expect(host.queryData).toHaveBeenCalledWith(
      `accounts?savedQuery=${QUERY_ID}&$top=10`,
      undefined,
    );
  });
});

describe('read headers (#109)', () => {
  const CALLER = '00000000-0000-0000-0000-0000000000c1';
  // What the read modes hand over: the Headers pane as a map.
  const PANE = {
    Accept: 'application/json',
    'OData-MaxVersion': '4.0',
    'OData-Version': '4.0',
    MSCRMCallerID: CALLER,
    'X-Custom': 'kept in the Code tab',
  };

  it('Retrieve Multiple sends impersonation, not other rows, and keeps $top without a page size', async () => {
    await executeRetrieveMultiple(paged(10, null), PANE);
    expect(host.queryData).toHaveBeenCalledWith('/accounts?$top=10', { MSCRMCallerID: CALLER });
  });

  it('with a page size, sends both, and $top stays out', async () => {
    await executeRetrieveMultiple(paged(10, 50), PANE);
    expect(host.queryData).toHaveBeenCalledWith('/accounts', { MSCRMCallerID: CALLER, ...PAGE_50 });
  });

  it('later pages run as the same caller', async () => {
    await executeAbsoluteUrl(
      'https://org.crm.dynamics.com/api/data/v9.2/accounts?$skiptoken=x',
      null,
      PANE,
    );
    expect(host.queryData).toHaveBeenCalledWith('accounts?$skiptoken=x', { MSCRMCallerID: CALLER });
  });

  it('Retrieve NextLink and Predefined Query send them too', async () => {
    await executeNextLink(
      {
        url: 'https://org.crm.dynamics.com/api/data/v9.2/accounts?$skiptoken=y',
        prefer: emptyPrefer(),
        headers: [],
        dirty: new Set(),
      } as unknown as RetrieveNextLinkState,
      PANE,
    );
    expect(host.queryData).toHaveBeenLastCalledWith('accounts?$skiptoken=y', {
      MSCRMCallerID: CALLER,
    });
    await executePredefined(predefined(null, null), PANE);
    expect(host.queryData).toHaveBeenLastCalledWith(`accounts?savedQuery=${QUERY_ID}`, {
      MSCRMCallerID: CALLER,
    });
  });

  it("doesn't run as you on a host that would drop the header (PPTB < 1.2.7)", async () => {
    host.additionalHeaders = false;
    const r = await executeRetrieveMultiple(paged(10, null), PANE);
    expect(host.queryData).not.toHaveBeenCalled();
    expect(r.ok).toBe(false);
  });

  it('a plain read still sends nothing (no consent prompt)', async () => {
    const { MSCRMCallerID: _c, 'X-Custom': _x, ...plain } = PANE;
    await executeRetrieveMultiple(paged(10, null), plain);
    expect(host.queryData).toHaveBeenCalledWith('/accounts?$top=10', undefined);
  });
});
