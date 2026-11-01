// Monaco's workers, run on the page when a Worker can't start (#255).
//
// The PPTB VS Code extension loads DRS into a `vscode-webview://` document but
// serves dist/ from a `vscode-cdn.net` URL, so `new Worker(url)` throws a
// SecurityError (a worker script must be same-origin), and the extension's
// CSP allows no `blob:` or `data:` worker either. Monaco falls back to the
// page for its editor worker, but its JSON language service has none: JSON
// bodies lost their error squiggles, completion and formatting, and every
// keystroke threw.
//
// InPageWorker stands in for the Worker: the same worker code (Monaco's
// EditorWorker and, for JSON, its JSONWorker) behind Monaco's own
// WebWorkerServer, mirroring editor.worker.start.js's `start()`. A
// MessageChannel carries the messages, so they're cloned and delivered
// asynchronously as a worker's are. The work runs on the main thread, which is
// fine at request-body sizes.
import { WebWorkerServer } from 'monaco-editor/esm/vs/base/common/worker/webWorker';
import { EditorWorker } from 'monaco-editor/esm/vs/editor/common/services/editorWebWorker';
import { EditorWorkerHost } from 'monaco-editor/esm/vs/editor/common/services/editorWorkerHost';

/** What a worker's foreign module gets: the page's host methods and the synced models. */
export interface WorkerContext {
  host: Record<string, (...args: unknown[]) => Promise<unknown>>;
  getMirrorModels: () => unknown[];
}

/** Creates the worker's foreign module: JSONWorker for JSON, nothing for the editor worker. */
export type CreateClient = (ctx: WorkerContext, createData: unknown) => object;

export class InPageWorker extends EventTarget {
  onmessage: ((ev: MessageEvent) => void) | null = null;
  onmessageerror: ((ev: MessageEvent) => void) | null = null;
  private readonly channel = new MessageChannel();
  private server: WebWorkerServer<EditorWorker> | null = null;
  private skippedFirst = false;
  private terminated = false;

  /**
   * `withCreateData`: the worker reads a second message as its create data
   * before starting, as json.worker.js does (Monaco's JSON mode posts
   * "ignore", then the language settings). The editor worker starts on its
   * first message.
   */
  constructor(
    private readonly createClient: CreateClient,
    private readonly withCreateData: boolean,
  ) {
    super();
    this.channel.port1.onmessage = (ev) => this.onmessage?.(ev);
    this.channel.port1.onmessageerror = (ev) => this.onmessageerror?.(ev);
    this.channel.port2.onmessage = (ev) => this.receive(ev.data);
  }

  postMessage(message: unknown, transfer: Transferable[] = []) {
    if (!this.terminated) this.channel.port1.postMessage(message, transfer);
  }

  terminate() {
    this.terminated = true;
    this.server = null;
    this.channel.port1.close();
    this.channel.port2.close();
  }

  private receive(data: unknown) {
    if (this.terminated) return;
    try {
      if (this.server) this.server.onmessage(data);
      else if (this.withCreateData && !this.skippedFirst) this.skippedFirst = true;
      else this.server = this.start(this.withCreateData ? data : undefined);
    } catch (error) {
      this.reportError(error);
    }
  }

  // What a worker's uncaught error does: Monaco listens for it.
  private reportError(error: unknown) {
    this.dispatchEvent(
      new ErrorEvent('error', {
        error,
        message: error instanceof Error ? error.message : String(error),
      }),
    );
  }

  // The worker's foreign module. If it can't be created, the worker still
  // starts without one, so Monaco's requests fail at once ("Missing method")
  // instead of waiting forever for a worker that never answers.
  private createForeignModule(ctx: WorkerContext, createData: unknown): object {
    try {
      return this.createClient(ctx, createData);
    } catch (error) {
      this.reportError(error);
      return {};
    }
  }

  private start(createData: unknown) {
    const server: WebWorkerServer<EditorWorker> = new WebWorkerServer(
      // A real worker's replies go without a transfer list too.
      (msg) => this.channel.port2.postMessage(msg),
      (workerServer) => {
        const editorWorkerHost = EditorWorkerHost.getChannel(workerServer);
        const host = new Proxy(
          {},
          {
            get(_target, prop) {
              // Not thenable, so returning it from an async function works.
              if (prop === 'then') return undefined;
              if (typeof prop !== 'string') throw new Error('Not supported');
              return (...args: unknown[]) => editorWorkerHost.$fhr(prop, args);
            },
          },
        );
        const ctx: WorkerContext = {
          host,
          getMirrorModels: () => server.requestHandler.getModels(),
        };
        return new EditorWorker(this.createForeignModule(ctx, createData));
      },
    );
    return server;
  }
}

/** An in-page stand-in for the worker Monaco asks for by `label`. */
export function createInPageWorker(label: string): Promise<Worker> {
  // The JSON language service (~130 kB) loads only when it's needed here. If
  // its chunk can't load (DRS updated while open, #187), the worker starts
  // without it: JSON requests fail at once rather than hang.
  const worker =
    label === 'json'
      ? import('./jsonWorkerInPage').then(
          (m) => new InPageWorker(m.createJsonWorker, true),
          (error: unknown) =>
            new InPageWorker(() => {
              throw error;
            }, true),
        )
      : Promise.resolve(new InPageWorker(() => ({}), false));
  // Monaco uses only the part of Worker that InPageWorker has.
  return worker as unknown as Promise<Worker>;
}

let workersBlocked = false;

function blockWorkers(reason: unknown) {
  if (workersBlocked) return;
  workersBlocked = true;
  console.info(
    `Monaco's workers can't start here, so they run on the page (#255): ${
      reason instanceof Error ? reason.message : String(reason)
    }`,
  );
}

/**
 * A real Worker that switches to an in-page one if it fails before it ever
 * answers. A worker can also fail after `new Worker()` returns: a CSP
 * `worker-src` that forbids it, or a worker file that's gone (DRS updated
 * while open, #187), fire an `error` event instead of throwing. Monaco would
 * then wait forever for its first reply, so the messages sent so far (its
 * handshake) are replayed to an in-page worker instead.
 *
 * Only a load failure switches: it fires a plain Event. An ErrorEvent is the
 * worker's own runtime error, passed on to Monaco as before.
 */
export class WorkerWithFallback extends EventTarget {
  onmessage: ((ev: MessageEvent) => void) | null = null;
  onmessageerror: ((ev: MessageEvent) => void) | null = null;
  private current: Worker | InPageWorker;
  // Sent before the worker first answered: replayed if it fails to load.
  // null once it answered, or once a message transferred buffers (those
  // can't be sent twice), or once it switched.
  private sent: unknown[] | null = [];
  private switching = false;
  private terminated = false;

  constructor(
    real: Worker,
    private readonly label: string,
  ) {
    super();
    this.current = real;
    this.listen(real);
    real.addEventListener('error', (ev: Event) => {
      if (this.terminated || this.switching) return;
      if (this.sent === null || ev instanceof ErrorEvent) {
        this.forwardError(ev);
        return;
      }
      this.switching = true;
      ev.preventDefault();
      real.terminate();
      blockWorkers(`the ${label} worker failed to load`);
      void this.switchToInPage();
    });
  }

  postMessage(message: unknown, transfer: Transferable[] = []) {
    if (this.terminated) return;
    if (this.sent) {
      if (transfer.length) this.sent = null;
      else this.sent.push(message);
    }
    this.current.postMessage(message, transfer);
  }

  terminate() {
    this.terminated = true;
    this.current.terminate();
  }

  private listen(worker: Worker | InPageWorker) {
    worker.onmessage = (ev: MessageEvent) => {
      this.sent = null; // It answered, so it started.
      this.onmessage?.(ev);
    };
    worker.onmessageerror = (ev: MessageEvent) => this.onmessageerror?.(ev);
  }

  private forwardError(ev: Event) {
    const { error, message, filename, lineno, colno } = ev as ErrorEvent;
    this.dispatchEvent(new ErrorEvent('error', { error, message, filename, lineno, colno }));
  }

  private async switchToInPage() {
    const worker = (await createInPageWorker(this.label)) as unknown as InPageWorker;
    if (this.terminated) {
      worker.terminate();
      return;
    }
    // Messages posted meanwhile went to the dead worker and are in `sent`.
    const replay = this.sent ?? [];
    this.sent = null;
    this.current = worker;
    this.listen(worker);
    worker.addEventListener('error', (ev) => this.forwardError(ev));
    try {
      for (const message of replay) worker.postMessage(message);
    } catch (error) {
      this.forwardError(new ErrorEvent('error', { error, message: String(error) }));
    }
  }
}

/**
 * `start()`'s Worker, falling back to an in-page one if Workers can't start
 * here: `new Worker()` throws (a cross-origin worker URL, as in the VS Code
 * extension) or the worker fails before it answers. After the first failure,
 * every worker runs in-page.
 */
export function workerOrInPage(label: string, start: () => Worker): Worker | Promise<Worker> {
  if (!workersBlocked) {
    try {
      // Monaco uses only the part of Worker that WorkerWithFallback has.
      return new WorkerWithFallback(start(), label) as unknown as Worker;
    } catch (error) {
      blockWorkers(error);
    }
  }
  return createInPageWorker(label);
}

/** For tests: forget an earlier failure. */
export function resetWorkersBlocked() {
  workersBlocked = false;
}
