// Types for the Monaco internals the in-page worker uses (#255,
// inPageWorker.ts). monaco-editor ships them as plain ESM without .d.ts; these
// are their shapes in 0.54. inPageWorker.test.ts runs them for real.

declare module 'monaco-editor/esm/vs/base/common/worker/webWorker' {
  export class WebWorkerServer<H extends object = object> {
    constructor(
      postMessage: (msg: unknown, transfer?: Transferable[]) => void,
      requestHandlerFactory: (server: WebWorkerServer<H>) => H,
    );
    readonly requestHandler: H;
    onmessage(msg: unknown): void;
  }
}

declare module 'monaco-editor/esm/vs/editor/common/services/editorWebWorker' {
  export class EditorWorker {
    constructor(foreignModule?: object | null);
    getModels(): unknown[];
  }
}

declare module 'monaco-editor/esm/vs/editor/common/services/editorWorkerHost' {
  export class EditorWorkerHost {
    static getChannel(server: unknown): {
      $fhr(method: string, args: unknown[]): Promise<unknown>;
    };
    static setChannel(client: unknown, host: object): void;
  }
}

declare module 'monaco-editor/esm/vs/base/browser/webWorkerFactory' {
  export function createWebWorker(worker: Promise<Worker>): {
    proxy: Record<string, (...args: unknown[]) => Promise<unknown>>;
    dispose(): void;
  };
}

// json.worker.js through the `?in-page` transform in vite.config.ts.
declare module 'monaco-editor/esm/vs/language/json/json.worker?in-page' {
  export class JSONWorker {
    constructor(ctx: unknown, createData: unknown);
  }
}
