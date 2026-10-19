// Monaco bootstrap — loaded on demand by `MonacoEditor` (React.lazy), so the
// ~3.6 MB monaco-editor chunk stays off DRS's startup path (#22). Everything
// that must run before the first <Editor> mounts lives here, in one place.
import Editor, { loader } from '@monaco-editor/react';
import * as monaco from 'monaco-editor';
// Monaco language workers — bundled by Vite as separate same-origin chunks
// via the `?worker` suffix. The CSP allows worker-src from the iframe's own
// origin, so these load cleanly inside PPTB (the `cdn.jsdelivr.net` path
// that Monaco defaults to is the one that was blocked).
//
// Without this MonacoEnvironment setup, language modes that ship a worker
// (TypeScript/JavaScript validation, JSON parsing, HTML/CSS analysis) throw
// `You must define a function MonacoEnvironment.getWorkerUrl or
// MonacoEnvironment.getWorker` — visible in the PPTB devtools when any
// Code-tab format other than plain text is selected.
import editorWorker from 'monaco-editor/esm/vs/editor/editor.worker?worker';
import jsonWorker from 'monaco-editor/esm/vs/language/json/json.worker?worker';
import tsWorker from 'monaco-editor/esm/vs/language/typescript/ts.worker?worker';
import cssWorker from 'monaco-editor/esm/vs/language/css/css.worker?worker';
import htmlWorker from 'monaco-editor/esm/vs/language/html/html.worker?worker';

// Bundle Monaco locally — PPTB's CSP forbids fetching from cdn.jsdelivr.net
// ("script-src 'self' 'unsafe-inline' pptb-webview:"). `loader.config({ monaco })`
// tells `@monaco-editor/react` to use the already-imported package instead
// of its default CDN-AMD loader path.
loader.config({ monaco });

// Route language workers to Vite's same-origin bundled chunks (see the
// `?worker` imports above). The fallback `editorWorker` covers plain-text
// modes that don't have a dedicated language worker (curl, csharp, etc.).
self.MonacoEnvironment = {
  getWorker(_workerId: string, label: string) {
    if (label === 'json') return new jsonWorker();
    if (label === 'typescript' || label === 'javascript') return new tsWorker();
    if (label === 'css' || label === 'scss' || label === 'less') return new cssWorker();
    if (label === 'html' || label === 'handlebars' || label === 'razor') return new htmlWorker();
    return new editorWorker();
  },
};

export default Editor;
