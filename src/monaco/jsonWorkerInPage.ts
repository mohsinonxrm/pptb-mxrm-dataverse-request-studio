// The JSON language service for the in-page worker (#255), in its own chunk:
// it loads only where a real worker can't start. `?in-page` is json.worker.js
// with its worker bootstrap swapped for an export (vite.config.ts).
import { JSONWorker } from 'monaco-editor/esm/vs/language/json/json.worker?in-page';
import type { CreateClient } from './inPageWorker';

export const createJsonWorker: CreateClient = (ctx, createData) => new JSONWorker(ctx, createData);
