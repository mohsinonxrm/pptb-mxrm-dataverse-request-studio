// The startup handshake vs a connection ToolBox binds while it's in flight
// (#281): the tool must end up connected, and Retry must re-read it.
import { afterEach, describe, expect, it, vi } from 'vitest';

type Handler = (e: unknown, payload: { event: string; data?: unknown }) => void;

/** A fake toolboxAPI whose getActiveConnection answers are released by the test. */
function fakeHost() {
  let handler: Handler | null = null;
  const reads: Array<(c: unknown) => void> = [];
  const api = {
    utils: { getCurrentTheme: async () => 'light' },
    connections: {
      getActiveConnection: vi.fn(() => new Promise((resolve) => reads.push(resolve))),
    },
    events: { on: (h: Handler) => (handler = h) },
  };
  return {
    api,
    reads,
    fire: (event: string) => handler?.(null, { event, data: { id: 'c1' } }),
  };
}

const CONN = { id: 'c1', url: 'https://org.crm.dynamics.com/', environment: 'Dev' };
const flush = () => new Promise((r) => setTimeout(r, 0));

async function loadBridge(api: unknown) {
  vi.resetModules();
  (window as unknown as { toolboxAPI: unknown }).toolboxAPI = api;
  return import('./pptbBridge');
}

afterEach(() => {
  delete (window as unknown as { toolboxAPI?: unknown }).toolboxAPI;
});

describe('connection handshake (#281)', () => {
  it('hears a connection bound while the first read is in flight', async () => {
    const host = fakeHost();
    const bridge = await loadBridge(host.api);
    await flush();
    expect(host.reads).toHaveLength(1); // the handshake's read, still pending
    host.fire('connection:updated'); // ToolBox binds the connection
    await flush();
    expect(host.reads).toHaveLength(2);
    host.reads[1](CONN); // the event's read: connected
    await flush();
    host.reads[0](null); // the stale handshake read: "no connection"
    await flush();
    expect(bridge.getSession()).toMatchObject({
      connected: true,
      environmentUrl: CONN.url,
      handshakeDone: true,
    });
  });

  it('Retry (recheckConnection) picks up a connection the tool missed', async () => {
    const host = fakeHost();
    const bridge = await loadBridge(host.api);
    await flush();
    host.reads[0](null);
    await flush();
    expect(bridge.getSession()).toMatchObject({ connected: false, handshakeDone: true });
    const done = bridge.recheckConnection();
    await flush();
    host.reads[1](CONN);
    await done;
    expect(bridge.getSession()).toMatchObject({ connected: true, environmentUrl: CONN.url });
  });
});

describe('out-of-order connection reads (#281 review)', () => {
  it('a stale "not connected" never shows the no-connection state while a newer read is pending', async () => {
    const host = fakeHost();
    const bridge = await loadBridge(host.api);
    await flush();
    host.fire('connection:updated');
    await flush();
    host.reads[0](null); // stale handshake answer lands first
    await flush();
    expect(bridge.getSession().handshakeDone).toBe(false); // no "connected: false" flash
    host.reads[1](CONN);
    await flush();
    expect(bridge.getSession()).toMatchObject({ connected: true, handshakeDone: true });
  });

  it('keeps a connection an older read found when the newer read fails', async () => {
    const host = fakeHost();
    const bridge = await loadBridge(host.api);
    await flush();
    host.fire('connection:updated');
    await flush();
    host.reads[0](CONN); // older read: connected
    await flush();
    expect(bridge.getSession()).toMatchObject({ connected: true, handshakeDone: true });
    // The newer read fails at the IPC level: keep the last known connection.
    host.api.connections.getActiveConnection.mockImplementationOnce(() =>
      Promise.reject(new Error('IPC')),
    );
    host.fire('connection:updated');
    await flush();
    expect(bridge.getSession()).toMatchObject({ connected: true });
  });

  it('a failed first read ends the handshake as not connected', async () => {
    const host = fakeHost();
    host.api.connections.getActiveConnection.mockImplementationOnce(() =>
      Promise.reject(new Error('IPC')),
    );
    const bridge = await loadBridge(host.api);
    await flush();
    await flush();
    expect(bridge.getSession()).toMatchObject({ connected: false, handshakeDone: true });
  });
});
