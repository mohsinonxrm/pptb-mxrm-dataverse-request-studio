// Host events and feature detection in the bridge (#235): theme from the host
// and `settings:updated`, connection events reaching subscribers (incl. the
// `connection:deleted` id check), PPTB 1.2.7 header pass-through detection
// (#219, #109) and which headers a dvHost call actually sends.
// The startup connection handshake is covered in pptbBridge.connection.test.ts.
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { HostSession } from './pptbBridge';

type Handler = (e: unknown, payload: unknown) => void;
type Conn = { id?: string; url?: string; environment?: string; name?: string } | null;

const CONN = { id: 'c1', url: 'https://org1.crm.dynamics.com/', environment: 'Dev', name: 'Org 1' };
const OTHER = {
  id: 'c2',
  url: 'https://org2.crm.dynamics.com/',
  environment: 'Test',
  name: 'Org 2',
};

/** A fake toolboxAPI whose answers resolve at once; the test fires host events. */
function fakeHost(opts: { theme?: unknown; conn?: Conn; noEvents?: boolean } = {}) {
  let handler: Handler | null = null;
  const api = {
    utils: { getCurrentTheme: vi.fn(async () => opts.theme ?? 'light') },
    connections: { getActiveConnection: vi.fn(async (): Promise<Conn> => opts.conn ?? null) },
    events: opts.noEvents ? {} : { on: (h: Handler) => (handler = h) },
  };
  return {
    api,
    fire: (event: string, data?: unknown) => handler?.(null, { event, data }),
    raw: (payload: unknown) => handler?.(null, payload),
  };
}

/** A fake dataverseAPI; `batch` adds `executeBatch`, the PPTB 1.2.7 probe. */
function fakeDataverse(batch: boolean) {
  const methods = {
    create: vi.fn(async () => ({ id: '1' })),
    update: vi.fn(async () => undefined),
    delete: vi.fn(async () => undefined),
    queryData: vi.fn(async () => ({ value: [] })),
    execute: vi.fn(async () => ({})),
    associate: vi.fn(async () => undefined),
    disassociate: vi.fn(async () => undefined),
    executeBatch: vi.fn(async () => []),
    executeTransaction: vi.fn(async () => []),
  };
  if (!batch) delete (methods as Partial<typeof methods>).executeBatch;
  return methods;
}

// Two macrotasks: the handshake chains more than one await.
const flush = () =>
  new Promise((r) => setTimeout(r, 0)).then(() => new Promise((r) => setTimeout(r, 0)));

const win = window as unknown as { toolboxAPI?: unknown; dataverseAPI?: unknown };

async function loadBridge(toolbox?: unknown, dataverse?: unknown) {
  vi.resetModules();
  if (toolbox) win.toolboxAPI = toolbox;
  if (dataverse) win.dataverseAPI = dataverse;
  const bridge = await import('./pptbBridge');
  await flush();
  return bridge;
}

/** Record every session the bridge emits from now on. */
function record(bridge: { subscribe: (l: (s: HostSession) => void) => () => void }) {
  const seen: HostSession[] = [];
  bridge.subscribe((s) => seen.push(s));
  seen.length = 0; // drop the immediate replay
  return seen;
}

beforeEach(() => {
  vi.spyOn(console, 'warn').mockImplementation(() => {});
  vi.spyOn(console, 'error').mockImplementation(() => {});
});

afterEach(() => {
  delete win.toolboxAPI;
  delete win.dataverseAPI;
  window.history.replaceState(null, '', '/');
  vi.restoreAllMocks();
});

describe('theme', () => {
  it('takes the theme from getCurrentTheme(), whatever its case', async () => {
    const bridge = await loadBridge(fakeHost({ theme: 'Dark' }).api);
    expect(bridge.getSession().theme).toBe('dark');
  });

  it('falls back to light for an unknown theme or a failed read', async () => {
    expect((await loadBridge(fakeHost({ theme: 'sepia' }).api)).getSession().theme).toBe('light');
    const host = fakeHost({ conn: CONN });
    host.api.utils.getCurrentTheme.mockRejectedValueOnce(new Error('IPC'));
    const bridge = await loadBridge(host.api);
    expect(bridge.getSession()).toMatchObject({ theme: 'light', connected: true });
  });

  it('settings:updated with a theme applies it at once and tells subscribers', async () => {
    const host = fakeHost();
    const bridge = await loadBridge(host.api);
    const seen = record(bridge);
    host.fire('settings:updated', { theme: 'DARK' });
    expect(bridge.getSession().theme).toBe('dark');
    expect(seen.map((s) => s.theme)).toEqual(['dark']);
    host.fire('settings:updated', { theme: 'light' });
    expect(bridge.getSession().theme).toBe('light');
    expect(host.api.utils.getCurrentTheme).toHaveBeenCalledTimes(1); // the handshake only
  });

  it('settings:updated without a theme re-reads it from the host', async () => {
    const host = fakeHost();
    const bridge = await loadBridge(host.api);
    host.api.utils.getCurrentTheme.mockResolvedValueOnce('dark');
    host.fire('settings:updated', { language: 'fr' });
    await flush();
    expect(host.api.utils.getCurrentTheme).toHaveBeenCalledTimes(2);
    expect(bridge.getSession().theme).toBe('dark');
  });

  it('a theme change keeps the connection', async () => {
    const host = fakeHost({ conn: CONN });
    const bridge = await loadBridge(host.api);
    host.fire('settings:updated', { theme: 'dark' });
    expect(bridge.getSession()).toMatchObject({
      theme: 'dark',
      connected: true,
      environmentUrl: CONN.url,
    });
  });

  it('paints with ?theme=dark before the host answers, standalone too', async () => {
    window.history.replaceState(null, '', '/?theme=dark');
    expect((await loadBridge()).getSession().theme).toBe('dark');

    const host = fakeHost({ theme: 'light' });
    let answer!: (t: string) => void;
    host.api.utils.getCurrentTheme.mockReturnValueOnce(new Promise((r) => (answer = r)));
    const bridge = await loadBridge(host.api);
    expect(bridge.getSession().theme).toBe('dark');
    answer('light');
    await flush();
    expect(bridge.getSession().theme).toBe('light');
  });
});

describe('connection events', () => {
  it('connection:created and connection:updated re-read the connection for subscribers', async () => {
    const host = fakeHost();
    const bridge = await loadBridge(host.api);
    expect(bridge.getSession()).toMatchObject({ connected: false, handshakeDone: true });
    const seen = record(bridge);

    host.api.connections.getActiveConnection.mockResolvedValueOnce(CONN);
    host.fire('connection:created');
    await flush();
    expect(seen[seen.length - 1]).toMatchObject({
      connected: true,
      environmentUrl: CONN.url,
      environment: 'Dev',
      connectionName: 'Org 1',
    });

    host.api.connections.getActiveConnection.mockResolvedValueOnce(OTHER);
    host.fire('connection:updated');
    await flush();
    expect(seen[seen.length - 1]).toMatchObject({
      connected: true,
      environmentUrl: OTHER.url,
      environment: 'Test',
      connectionName: 'Org 2',
    });
  });

  it('connection:deleted for another connection is ignored', async () => {
    const host = fakeHost({ conn: CONN });
    const bridge = await loadBridge(host.api);
    const seen = record(bridge);
    host.fire('connection:deleted', { id: 'someone-else' });
    await flush();
    expect(host.api.connections.getActiveConnection).toHaveBeenCalledTimes(1);
    expect(seen).toEqual([]);
    expect(bridge.getSession().connected).toBe(true);
  });

  it('connection:deleted for our connection disconnects when the host says none', async () => {
    const host = fakeHost({ conn: CONN });
    const bridge = await loadBridge(host.api);
    host.api.connections.getActiveConnection.mockResolvedValueOnce(null);
    host.fire('connection:deleted', { id: 'c1' });
    await flush();
    expect(bridge.getSession()).toMatchObject({
      connected: false,
      environmentUrl: '',
      environment: '',
      connectionName: '',
    });
  });

  it('connection:deleted for our connection disconnects even when the re-read fails', async () => {
    const host = fakeHost({ conn: CONN });
    const bridge = await loadBridge(host.api);
    host.api.connections.getActiveConnection.mockRejectedValueOnce(new Error('IPC'));
    host.fire('connection:deleted', { id: 'c1' });
    await flush();
    expect(bridge.getSession()).toMatchObject({ connected: false, environmentUrl: '' });
  });

  it('connection:deleted without an id re-reads and keeps the connection if the read fails', async () => {
    const host = fakeHost({ conn: CONN });
    const bridge = await loadBridge(host.api);
    host.api.connections.getActiveConnection.mockRejectedValueOnce(new Error('IPC'));
    host.fire('connection:deleted');
    await flush();
    expect(host.api.connections.getActiveConnection).toHaveBeenCalledTimes(2);
    expect(bridge.getSession()).toMatchObject({ connected: true, environmentUrl: CONN.url });
  });

  it('connection:deleted re-reads when our connection id is unknown', async () => {
    const host = fakeHost({ conn: { url: CONN.url } }); // the host gave no id
    const bridge = await loadBridge(host.api);
    host.api.connections.getActiveConnection.mockResolvedValueOnce(null);
    host.fire('connection:deleted', { id: 'c9' });
    await flush();
    expect(bridge.getSession().connected).toBe(false);
  });

  it('ignores unknown events and malformed payloads', async () => {
    const host = fakeHost({ conn: CONN });
    const bridge = await loadBridge(host.api);
    const seen = record(bridge);
    host.fire('terminal:output', { text: 'hi' });
    host.raw(null);
    host.raw({ event: 42 });
    host.raw('connection:updated');
    await flush();
    expect(seen).toEqual([]);
    expect(host.api.connections.getActiveConnection).toHaveBeenCalledTimes(1);
  });

  it('still handshakes when the host has no event bus', async () => {
    const host = fakeHost({ conn: CONN, noEvents: true });
    const bridge = await loadBridge(host.api);
    expect(bridge.getSession()).toMatchObject({ connected: true, handshakeDone: true });
    expect(console.warn).toHaveBeenCalledWith(
      expect.stringContaining('toolboxAPI.events.on missing'),
    );
  });
});

describe('subscribe', () => {
  it('replays the current session at once, then each change, until unsubscribed', async () => {
    const host = fakeHost({ conn: CONN });
    const bridge = await loadBridge(host.api);
    const listener = vi.fn();
    const off = bridge.subscribe(listener);
    expect(listener).toHaveBeenCalledTimes(1);
    expect(listener).toHaveBeenLastCalledWith(bridge.getSession());
    host.fire('settings:updated', { theme: 'dark' });
    expect(listener).toHaveBeenCalledTimes(2);
    off();
    host.fire('settings:updated', { theme: 'light' });
    expect(listener).toHaveBeenCalledTimes(2);
  });

  it('a throwing subscriber does not stop the others', async () => {
    const host = fakeHost();
    const bridge = await loadBridge(host.api);
    bridge.subscribe((s) => {
      if (s.theme === 'dark') throw new Error('boom');
    });
    const after = vi.fn();
    bridge.subscribe(after);
    host.fire('settings:updated', { theme: 'dark' });
    expect(after).toHaveBeenLastCalledWith(expect.objectContaining({ theme: 'dark' }));
    expect(console.error).toHaveBeenCalled();
  });
});

describe('standalone (no host)', () => {
  it('is not embedded, is done handshaking, and is not connected', async () => {
    const bridge = await loadBridge();
    expect(bridge.isEmbedded()).toBe(false);
    expect(bridge.getSession()).toMatchObject({
      embedded: false,
      connected: false,
      handshakeDone: true,
      environmentUrl: '',
    });
    await expect(bridge.recheckConnection()).resolves.toBeUndefined();
  });

  it('dvHost and tbHost calls throw HostNotAvailableError', async () => {
    const bridge = await loadBridge();
    expect(() => bridge.dvHost.create('accounts', {})).toThrow(bridge.HostNotAvailableError);
    expect(() => bridge.tbHost.utils.copyToClipboard('x')).toThrow(bridge.HostNotAvailableError);
    await expect(bridge.dataverseQuery('accounts')).rejects.toThrow('running standalone');
  });
});

describe('feature detection', () => {
  it('PPTB 1.2.7+ (executeBatch present) sends headers and batches', async () => {
    const bridge = await loadBridge(fakeHost().api, fakeDataverse(true));
    expect(bridge.hostSupportsAdditionalHeaders()).toBe(true);
    expect(bridge.hostSupportsBatch()).toBe(true);
    expect(bridge.pageSizeIsSent()).toBe(true);
  });

  it('an older host (no executeBatch) sends neither, and drops the page size', async () => {
    const bridge = await loadBridge(fakeHost().api, fakeDataverse(false));
    expect(bridge.hostSupportsAdditionalHeaders()).toBe(false);
    expect(bridge.hostSupportsBatch()).toBe(false);
    expect(bridge.pageSizeIsSent()).toBe(false);
  });

  it('standalone: no host headers, but the Code tab request carries the page size', async () => {
    const bridge = await loadBridge(undefined, fakeDataverse(true));
    expect(bridge.hostSupportsAdditionalHeaders()).toBe(false);
    expect(bridge.hostSupportsBatch()).toBe(false);
    expect(bridge.pageSizeIsSent()).toBe(true);
  });

  it('is decided once, when the bridge loads', async () => {
    const dataverse = fakeDataverse(false);
    const bridge = await loadBridge(fakeHost().api, dataverse);
    (dataverse as Record<string, unknown>).executeBatch = vi.fn();
    expect(bridge.hostSupportsAdditionalHeaders()).toBe(false);
  });
});

describe('dvHost header pass-through', () => {
  const H = { MSCRM_CallerID: '00000000-0000-0000-0000-000000000001' };
  const PARAMS = { operationName: 'WhoAmI', operationType: 'function' as const };

  /** Every header-taking dvHost call, with the arguments it passes on before the headers. */
  function calls(
    dvHost: Awaited<ReturnType<typeof loadBridge>>['dvHost'],
    headers?: Record<string, string>,
  ) {
    return [
      [
        'create',
        () => dvHost.create('account', { name: 'A' }, headers),
        ['account', { name: 'A' }],
      ],
      [
        'update',
        () => dvHost.update('account', 'id1', { name: 'B' }, headers),
        ['account', 'id1', { name: 'B' }],
      ],
      ['delete', () => dvHost.delete('account', 'id1', headers), ['account', 'id1']],
      ['queryData', () => dvHost.queryData('accounts?$top=1', headers), ['accounts?$top=1']],
      ['execute', () => dvHost.execute(PARAMS, headers), [PARAMS]],
      [
        'associate',
        () => dvHost.associate('account', 'id1', 'rel', 'contact', 'id2', headers),
        ['account', 'id1', 'rel', 'contact', 'id2'],
      ],
      [
        'disassociate',
        () => dvHost.disassociate('account', 'id1', 'rel', 'id2', headers),
        ['account', 'id1', 'rel', 'id2'],
      ],
    ] as const;
  }

  it('PPTB 1.2.7+: headers go last, after an undefined options slot', async () => {
    const dataverse = fakeDataverse(true);
    const bridge = await loadBridge(fakeHost().api, dataverse);
    for (const [name, call, args] of calls(bridge.dvHost, H)) {
      await call();
      expect(dataverse[name].mock.calls, name).toEqual([[...args, undefined, H]]);
    }
  });

  it('PPTB 1.2.7+: no headers or an empty map keeps the old call shape', async () => {
    const dataverse = fakeDataverse(true);
    const bridge = await loadBridge(fakeHost().api, dataverse);
    for (const headers of [undefined, {}]) {
      for (const [name, call, args] of calls(bridge.dvHost, headers)) {
        dataverse[name].mockClear();
        await call();
        expect(dataverse[name].mock.calls, name).toEqual([[...args]]);
      }
    }
  });

  it('an older host drops the headers', async () => {
    const dataverse = fakeDataverse(false);
    const bridge = await loadBridge(fakeHost().api, dataverse);
    for (const [name, call, args] of calls(bridge.dvHost, H)) {
      await call();
      expect(dataverse[name].mock.calls, name).toEqual([[...args]]);
    }
  });

  it('$batch passes non-empty outer headers and drops an empty map', async () => {
    const dataverse = fakeDataverse(true);
    const bridge = await loadBridge(fakeHost().api, dataverse);
    const requests = [{ method: 'GET', url: 'accounts' }] as Parameters<
      typeof bridge.dvHost.executeBatch
    >[0];
    await bridge.dvHost.executeBatch(requests, H);
    await bridge.dvHost.executeBatch(requests, {});
    await bridge.dvHost.executeTransaction(requests, H);
    await bridge.dvHost.executeTransaction(requests);
    expect(dataverse.executeBatch.mock.calls).toEqual([[requests, undefined, H], [requests]]);
    expect(dataverse.executeTransaction.mock.calls).toEqual([[requests, undefined, H], [requests]]);
  });
});
