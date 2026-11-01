import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { createWebWorker } from 'monaco-editor/esm/vs/base/browser/webWorkerFactory';
import { EditorWorkerHost } from 'monaco-editor/esm/vs/editor/common/services/editorWorkerHost';
import {
  createInPageWorker,
  InPageWorker,
  resetWorkersBlocked,
  workerOrInPage,
} from './inPageWorker';

// The in-page workers (#255) driven by Monaco's own worker client, the way
// the editor and its JSON mode use them.

// Monaco's WebWorker checks `instanceof Worker`, which jsdom doesn't have.
beforeAll(() => {
  if (typeof Worker === 'undefined') vi.stubGlobal('Worker', class {});
});
afterAll(() => {
  vi.unstubAllGlobals();
});
afterEach(() => {
  resetWorkersBlocked();
  vi.restoreAllMocks();
});

const MODEL = 'inmemory://model/1';

const CREATE_DATA = {
  languageSettings: { validate: true, allowComments: false, schemas: [] },
  languageId: 'json',
  enableSchemaRequest: false,
};

/** JSON's worker, set up as Monaco's JSON mode does (jsonMode.js `createWebWorker`). */
function jsonClient() {
  const worker = createInPageWorker('json').then((w) => {
    w.postMessage('ignore');
    w.postMessage(CREATE_DATA);
    return w;
  });
  const client = createWebWorker(worker);
  EditorWorkerHost.setChannel(client, { $fhr: () => Promise.resolve(undefined) });
  return client;
}

async function validate(client: ReturnType<typeof jsonClient>, text: string) {
  await client.proxy.$acceptNewModel({ url: MODEL, versionId: 1, lines: [text], EOL: '\n' });
  return (await client.proxy.$fmr('doValidation', [MODEL])) as { message: string }[];
}

/**
 * Starts a JSON-style in-page worker over the raw protocol and sends it one
 * `$fmr` request; resolves with the reply.
 */
async function rawJsonRequest(worker: InPageWorker, method: string) {
  const replies: { seq?: string; err?: unknown; res?: unknown }[] = [];
  worker.onmessage = (ev) => replies.push(ev.data);
  const req = (id: string, m: string, args: unknown[]) =>
    worker.postMessage({ vsWorker: 1, req: id, channel: 'default', method: m, args, type: 0 });
  worker.postMessage('ignore');
  worker.postMessage(CREATE_DATA);
  worker.postMessage('-please-ignore-');
  req('1', '$initialize', [1]);
  req('2', '$fmr', [method, [MODEL]]);
  await vi.waitFor(() => expect(replies.find((r) => r.seq === '2')).toBeDefined());
  return replies.find((r) => r.seq === '2')!;
}

describe('in-page workers (#255)', () => {
  it('runs the editor worker', async () => {
    const client = createWebWorker(createInPageWorker('editorWorkerService'));
    EditorWorkerHost.setChannel(client, { $fhr: () => Promise.resolve(undefined) });
    await expect(client.proxy.$ping()).resolves.toBe('pong');
    client.dispose();
  });

  it("runs JSON's language service: validation", async () => {
    const client = jsonClient();
    expect((await validate(client, '{ "name": "x" }oops')).length).toBeGreaterThan(0);
    expect(await validate(client, '{ "name": "x" }')).toEqual([]);
    client.dispose();
  });

  it('starts a fresh JSON worker after the last one stopped (idle restart)', async () => {
    const first = jsonClient();
    expect((await validate(first, '{ oops')).length).toBeGreaterThan(0);
    first.dispose();
    const second = jsonClient();
    expect((await validate(second, '{ oops')).length).toBeGreaterThan(0);
    second.dispose();
  });

  it("runs JSON's language service: formatting", async () => {
    const client = jsonClient();
    await client.proxy.$acceptNewModel({
      url: MODEL,
      versionId: 1,
      lines: ['{"a":1}'],
      EOL: '\n',
    });
    const edits = (await client.proxy.$fmr('format', [
      MODEL,
      null,
      { tabSize: 2, insertSpaces: true },
    ])) as unknown[];
    expect(edits.length).toBeGreaterThan(0);
    client.dispose();
  });

  it('delivers messages asynchronously', async () => {
    const worker = new InPageWorker(() => ({}), false);
    const replies: unknown[] = [];
    worker.onmessage = (ev) => replies.push(ev.data);
    const request = {
      vsWorker: 1,
      req: '1',
      channel: 'default',
      method: '$initialize',
      args: [0],
      type: 0,
    };
    worker.postMessage('-please-ignore-');
    worker.postMessage(request);
    expect(replies).toEqual([]);
    await vi.waitFor(() => expect(replies).toHaveLength(1));
    expect(replies[0]).toMatchObject({ seq: '1', type: 1 });
    worker.terminate();
  });

  it('stops after terminate()', async () => {
    const worker = new InPageWorker(() => ({}), false);
    const replies: unknown[] = [];
    worker.onmessage = (ev) => replies.push(ev.data);
    worker.terminate();
    worker.postMessage('-please-ignore-');
    worker.postMessage({
      vsWorker: 1,
      req: '1',
      channel: 'default',
      method: '$initialize',
      args: [0],
      type: 0,
    });
    await new Promise((r) => setTimeout(r, 20));
    expect(replies).toEqual([]);
  });

  it('reports a failure to start as an error event', async () => {
    const worker = new InPageWorker(() => {
      throw new Error('boom');
    }, false);
    const errors: string[] = [];
    worker.addEventListener('error', (ev) => errors.push((ev as ErrorEvent).message));
    worker.postMessage('-please-ignore-');
    await vi.waitFor(() => expect(errors).toEqual(['boom']));
    worker.terminate();
  });

  it("answers at once when JSON's service can't start, instead of hanging", async () => {
    const worker = new InPageWorker(() => {
      throw new Error('bad settings');
    }, true);
    const errors: string[] = [];
    worker.addEventListener('error', (ev) => errors.push((ev as ErrorEvent).message));
    const reply = await rawJsonRequest(worker, 'doValidation');
    expect(reply.err).toBeDefined();
    expect(errors).toEqual(['bad settings']);
    worker.terminate();
  });

  it("answers at once when the JSON service's chunk can't load", async () => {
    vi.resetModules();
    vi.doMock('./jsonWorkerInPage', () => {
      throw new Error('Failed to fetch dynamically imported module');
    });
    try {
      const fresh = await import('./inPageWorker');
      const worker = (await fresh.createInPageWorker('json')) as unknown as InPageWorker;
      const errors: string[] = [];
      worker.addEventListener('error', (ev) => errors.push((ev as ErrorEvent).message));
      const reply = await rawJsonRequest(worker, 'doValidation');
      expect(reply.err).toBeDefined();
      // The failed import, reported as the worker's error.
      expect(errors).toHaveLength(1);
      worker.terminate();
    } finally {
      vi.doUnmock('./jsonWorkerInPage');
      vi.resetModules();
    }
  });

  it('runs a real JSON request over the raw protocol', async () => {
    const worker = (await createInPageWorker('json')) as unknown as InPageWorker;
    const reply = await rawJsonRequest(worker, 'doValidation');
    expect(reply.err).toBeUndefined();
    worker.terminate();
  });
});

describe('workerOrInPage', () => {
  it('uses the real worker when it starts', () => {
    const real = {} as Worker;
    expect(workerOrInPage('json', () => real)).toBe(real);
  });

  it('falls back to an in-page worker, and stops trying real ones', async () => {
    vi.spyOn(console, 'info').mockImplementation(() => {});
    const start = vi.fn(() => {
      throw new DOMException('cross-origin', 'SecurityError');
    });
    const first = workerOrInPage('editorWorkerService', start);
    expect(await first).toBeInstanceOf(InPageWorker);
    const second = workerOrInPage('json', start);
    expect(await second).toBeInstanceOf(InPageWorker);
    expect(start).toHaveBeenCalledTimes(1);
    expect(console.info).toHaveBeenCalledTimes(1);
    (await first).terminate();
    (await second).terminate();
  });
});
