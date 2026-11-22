// #225: `dvHost.rawRequest` sends one request as a one-part `$batch` through
// ToolBox's `executeBatch` and maps the operation's result.
import { afterEach, describe, expect, it, vi } from 'vitest';

const win = window as unknown as { toolboxAPI?: unknown; dataverseAPI?: unknown };

/** Load the bridge afresh with a fake host: it decides once, at load, whether it's embedded. */
async function loadBridge(executeBatch: (...args: unknown[]) => Promise<unknown>) {
  vi.resetModules();
  win.toolboxAPI = {
    utils: { getCurrentTheme: async () => 'light' },
    connections: { getActiveConnection: async () => null },
    events: { on: () => {} },
  };
  win.dataverseAPI = { executeBatch };
  return import('./pptbBridge');
}

afterEach(() => {
  delete win.toolboxAPI;
  delete win.dataverseAPI;
  vi.restoreAllMocks();
});

const ID = '00000000-0000-0000-0000-0000000000a1';

describe('dvHost.rawRequest', () => {
  it('sends one part with contentId 1, and no headers key when there are none', async () => {
    const executeBatch = vi.fn(async (..._args: unknown[]) => [{ status: 204, headers: {} }]);
    const { dvHost } = await loadBridge(executeBatch);
    await dvHost.rawRequest({
      method: 'PUT',
      path: `/accounts(${ID})/name`,
      headers: {},
      body: { value: 'Contoso' },
    });
    expect(executeBatch).toHaveBeenCalledTimes(1);
    const [parts, target, outer] = executeBatch.mock.calls[0] as unknown as [
      Record<string, unknown>[],
      unknown,
      unknown,
    ];
    expect(parts).toEqual([
      { method: 'PUT', url: `accounts(${ID})/name`, contentId: '1', body: { value: 'Contoso' } },
    ]);
    expect(parts[0]).not.toHaveProperty('headers');
    // Never outer headers: they don't reach the part.
    expect(target).toBeUndefined();
    expect(outer).toBeUndefined();
    expect(executeBatch.mock.calls[0]).toHaveLength(2);
  });

  it('leaves the body key off when there is no body, and sends headers when given', async () => {
    const executeBatch = vi.fn(async (..._args: unknown[]) => [{ status: 204, headers: {} }]);
    const { dvHost } = await loadBridge(executeBatch);
    await dvHost.rawRequest({
      method: 'DELETE',
      path: `accounts(${ID})/description`,
      headers: { 'If-Match': '*' },
    });
    const [parts] = executeBatch.mock.calls[0] as unknown as [Record<string, unknown>[]];
    expect(parts[0]).toEqual({
      method: 'DELETE',
      url: `accounts(${ID})/description`,
      contentId: '1',
      headers: { 'If-Match': '*' },
    });
    expect(parts[0]).not.toHaveProperty('body');
  });

  it('passes connectionTarget through', async () => {
    const executeBatch = vi.fn(async (..._args: unknown[]) => [
      { status: 200, headers: {}, body: { value: [] } },
    ]);
    const { dvHost } = await loadBridge(executeBatch);
    await dvHost.rawRequest({ method: 'GET', path: 'accounts', connectionTarget: 'secondary' });
    expect(executeBatch.mock.calls[0]?.[1]).toBe('secondary');
    await dvHost.rawRequest({ method: 'GET', path: 'accounts', connectionTarget: 2 });
    expect(executeBatch.mock.calls[1]?.[1]).toBe(2);
  });

  it('maps statusText, lowercased headers, batchHeaders and body', async () => {
    const executeBatch = vi.fn(async (..._args: unknown[]) => [
      {
        contentId: '1',
        status: 201,
        statusText: 'Created',
        headers: { ETag: 'W/"1"', 'OData-EntityId': 'https://x/api/data/v9.2/accounts(1)' },
        batchHeaders: { 'x-ms-ratelimit-burst-remaining-xrm-requests': '5999' },
        body: { name: 'Contoso' },
      },
    ]);
    const { dvHost } = await loadBridge(executeBatch);
    const r = await dvHost.rawRequest({ method: 'POST', path: 'accounts', body: { name: 'C' } });
    expect(r).toEqual({
      status: 201,
      statusText: 'Created',
      headers: { etag: 'W/"1"', 'odata-entityid': 'https://x/api/data/v9.2/accounts(1)' },
      batchHeaders: { 'x-ms-ratelimit-burst-remaining-xrm-requests': '5999' },
      body: { name: 'Contoso' },
    });
  });

  it('throws when ToolBox returns no operation result', async () => {
    const { dvHost } = await loadBridge(vi.fn(async (..._args: unknown[]) => []));
    await expect(dvHost.rawRequest({ method: 'GET', path: 'accounts' })).rejects.toThrow(
      'Malformed Dataverse batch response: no operation result',
    );
  });

  it('refuses a path the host would reject, without calling it', async () => {
    const executeBatch = vi.fn(async (..._args: unknown[]) => [{ status: 204, headers: {} }]);
    const { dvHost } = await loadBridge(executeBatch);
    await expect(dvHost.rawRequest({ method: 'GET', path: '$batch' })).rejects.toThrow(/Not sent/);
    expect(executeBatch).not.toHaveBeenCalled();
  });

  it('throws HostNotAvailableError outside PPTB', async () => {
    vi.resetModules();
    const { dvHost, HostNotAvailableError } = await import('./pptbBridge');
    await expect(dvHost.rawRequest({ method: 'GET', path: 'accounts' })).rejects.toBeInstanceOf(
      HostNotAvailableError,
    );
  });
});
