// Monaco bootstrap — loaded on demand by `MonacoEditor` (React.lazy), so the
// monaco-editor chunk stays off DRS's startup path (#22). Everything that
// must run before the first <Editor> mounts lives here, in one place.
import { loader } from '@monaco-editor/react';
// The editor core with its standard features (find, folding, bracket
// matching, …), without the languages `import 'monaco-editor'` brings: that
// entry registers ~80 basic languages and the TypeScript, CSS and HTML
// language services, whose workers alone were ~7.6 MB of dist (#236).
import 'monaco-editor/esm/vs/editor/edcore.main';
import * as monaco from 'monaco-editor/esm/vs/editor/editor.api';
// The languages DRS shows: JSON bodies (with its language service, so
// malformed JSON is underlined), FetchXML, and the Code tab's formats
// (FORMAT_LANG in engine/codeGenerators.ts). These are tokenizers only, which
// is all a read-only snippet or the Editor playground needs.
import 'monaco-editor/esm/vs/language/json/monaco.contribution';
import 'monaco-editor/esm/vs/basic-languages/javascript/javascript.contribution';
import 'monaco-editor/esm/vs/basic-languages/csharp/csharp.contribution';
import 'monaco-editor/esm/vs/basic-languages/python/python.contribution';
import 'monaco-editor/esm/vs/basic-languages/powershell/powershell.contribution';
import 'monaco-editor/esm/vs/basic-languages/shell/shell.contribution';
import 'monaco-editor/esm/vs/basic-languages/xml/xml.contribution';
// Monaco's workers — bundled by Vite as separate same-origin chunks via the
// `?worker` suffix. The CSP allows worker-src from the iframe's own origin,
// so these load cleanly inside PPTB (the `cdn.jsdelivr.net` path that Monaco
// defaults to is the one that was blocked). JSON's language service runs in
// its own worker; everything else uses the editor worker.
import editorWorker from 'monaco-editor/esm/vs/editor/editor.worker?worker';
import jsonWorker from 'monaco-editor/esm/vs/language/json/json.worker?worker';
import { workerOrInPage } from './inPageWorker';
import SchemaEditor from './SchemaEditor';

// Bundle Monaco locally — PPTB's CSP forbids fetching from cdn.jsdelivr.net
// ("script-src 'self' 'unsafe-inline' pptb-webview:"). `loader.config({ monaco })`
// tells `@monaco-editor/react` to use the already-imported package instead
// of its default CDN-AMD loader path.
loader.config({ monaco });

// Without this, Monaco throws `You must define a function
// MonacoEnvironment.getWorkerUrl or MonacoEnvironment.getWorker` when the
// JSON service (or the editor's own diffing / link detection) starts. Where
// a worker can't start (the PPTB VS Code extension serves the worker scripts
// from another origin), the same code runs on the page instead (#255).
self.MonacoEnvironment = {
  getWorker(_workerId: string, label: string) {
    return workerOrInPage(label, () => (label === 'json' ? new jsonWorker() : new editorWorker()));
  },
};

// `@monaco-editor/react`'s Editor, with per-model JSON schemas (#268)
export default SchemaEditor;
