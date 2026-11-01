import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { createWebWorker } from 'monaco-editor/esm/vs/base/browser/webWorkerFactory';
import { EditorWorkerHost } from 'monaco-editor/esm/vs/editor/common/services/editorWorkerHost';
import {
  createInPageWorker,
  InPageWorker,
  resetWorkersBlocked,
  WorkerWithFallback,
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
function jsonClient(start: Promise<Worker> | Worker = createInPageWorker('json')) {
  const worker = Promise.resolve(start).then((w) => {
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

/** A Worker stand-in that can answer, or fail the way a blocked or missing worker script does. */
class FakeWorker extends EventTarget {
  onmessage: ((ev: MessageEvent) => void) | null = null;
  onmessageerror: ((ev: MessageEvent) => void) | null = null;
  posted: unknown[] = [];
  terminated = false;
  postMessage(message: unknown) {
    this.posted.push(message);
  }
  terminate() {
    this.terminated = true;
  }
  reply(data: unknown) {
    this.onmessage?.(new MessageEvent('message', { data }));
  }
  // A CSP-blocked or missing worker script: a plain `error` Event, no throw.
  fail() {
    this.dispatchEvent(new Event('error', { cancelable: true }));
  }
}

describe('workerOrInPage', () => {
  it('uses the real worker when it starts, and passes its messages and errors through', () => {
    const real = new FakeWorker();
    const worker = workerOrInPage(
      'json',
      () => real as unknown as Worker,
    ) as unknown as WorkerWithFallback;
    expect(worker).toBeInstanceOf(WorkerWithFallback);
    const replies: unknown[] = [];
    const errors: Event[] = [];
    worker.onmessage = (ev) => replies.push(ev.data);
    worker.addEventListener('error', (ev) => errors.push(ev));
    worker.postMessage('hello');
    expect(real.posted).toEqual(['hello']);
    real.reply('answer');
    expect(replies).toEqual(['answer']);
    // Once it has answered, a failure is the worker's own: passed on, no switch.
    real.fail();
    expect(errors).toHaveLength(1);
    expect(real.terminated).toBe(false);
    worker.terminate();
    expect(real.terminated).toBe(true);
  });

  it("switches to an in-page worker when JSON's worker fails before answering", async () => {
    vi.spyOn(console, 'info').mockImplementation(() => {});
    const real = new FakeWorker();
    const start = vi.fn(() => real as unknown as Worker);
    const client = jsonClient(workerOrInPage('json', start));
    setTimeout(() => real.fail(), 20);
    // The handshake and model sent to the dead worker are replayed in-page.
    expect((await validate(client, '{ oops')).length).toBeGreaterThan(0);
    expect(real.terminated).toBe(true);
    expect(console.info).toHaveBeenCalledTimes(1);
    // Later workers go in-page straight away.
    expect(await workerOrInPage('json', start)).toBeInstanceOf(InPageWorker);
    expect(start).toHaveBeenCalledTimes(1);
    client.dispose();
  });

  it('switches once, even if the worker reports more than one load failure', async () => {
    vi.spyOn(console, 'info').mockImplementation(() => {});
    const warn = vi.spyOn(console, 'warn');
    const real = new FakeWorker();
    const client = jsonClient(workerOrInPage('json', () => real as unknown as Worker));
    setTimeout(() => {
      real.fail();
      real.fail();
    }, 20);
    expect((await validate(client, '{ oops')).length).toBeGreaterThan(0);
    expect((await validate(client, '{}')).length).toBe(0);
    expect(console.info).toHaveBeenCalledTimes(1);
    // A second in-page worker would answer the replayed requests twice.
    expect(warn).not.toHaveBeenCalled();
    client.dispose();
  });

  it("passes on the worker's own runtime error before it answered, without switching", () => {
    vi.spyOn(console, 'info').mockImplementation(() => {});
    const real = new FakeWorker();
    const worker = workerOrInPage(
      'json',
      () => real as unknown as Worker,
    ) as unknown as WorkerWithFallback;
    const errors: string[] = [];
    worker.addEventListener('error', (ev) => errors.push((ev as ErrorEvent).message));
    worker.postMessage('ignore');
    real.dispatchEvent(new ErrorEvent('error', { message: 'boom', cancelable: true }));
    expect(errors).toEqual(['boom']);
    expect(real.terminated).toBe(false);
    expect(console.info).not.toHaveBeenCalled();
    worker.terminate();
  });

  it("doesn't switch once a message transferred buffers (they can't be replayed)", () => {
    vi.spyOn(console, 'info').mockImplementation(() => {});
    const real = new FakeWorker();
    const worker = workerOrInPage(
      'json',
      () => real as unknown as Worker,
    ) as unknown as WorkerWithFallback;
    const errors: Event[] = [];
    worker.addEventListener('error', (ev) => errors.push(ev));
    worker.postMessage({ data: new ArrayBuffer(8) }, [new ArrayBuffer(8)]);
    real.fail();
    expect(errors).toHaveLength(1);
    expect(console.info).not.toHaveBeenCalled();
    worker.terminate();
  });

  it("doesn't switch after terminate(), nor finish a switch it was in", async () => {
    vi.spyOn(console, 'info').mockImplementation(() => {});
    const first = new FakeWorker();
    const a = workerOrInPage(
      'json',
      () => first as unknown as Worker,
    ) as unknown as WorkerWithFallback;
    a.terminate();
    first.fail();
    expect(console.info).not.toHaveBeenCalled();

    const second = new FakeWorker();
    const b = workerOrInPage(
      'json',
      () => second as unknown as Worker,
    ) as unknown as WorkerWithFallback;
    const replies: unknown[] = [];
    b.onmessage = (ev) => replies.push(ev.data);
    b.postMessage('ignore');
    second.fail();
    b.terminate();
    await new Promise((r) => setTimeout(r, 50));
    expect(replies).toEqual([]);
  });

  it('switches to an in-page worker when the editor worker fails before answering', async () => {
    vi.spyOn(console, 'info').mockImplementation(() => {});
    const real = new FakeWorker();
    const client = createWebWorker(
      Promise.resolve(workerOrInPage('editorWorkerService', () => real as unknown as Worker)),
    );
    EditorWorkerHost.setChannel(client, { $fhr: () => Promise.resolve(undefined) });
    setTimeout(() => real.fail(), 20);
    await expect(client.proxy.$ping()).resolves.toBe('pong');
    client.dispose();
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
