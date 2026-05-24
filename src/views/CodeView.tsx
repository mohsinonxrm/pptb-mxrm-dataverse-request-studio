import { useEffect, useRef, useState } from 'react';
import Editor, { loader, type OnMount } from '@monaco-editor/react';
import * as monaco from 'monaco-editor';
// Type-only namespace import — keeps `monacoNs.editor.IStandaloneEditorConstructionOptions`
// dotted access working without forcing every type ref to dig through `typeof monaco`.
import type * as monacoNs from 'monaco-editor';
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
import {
  TabList, Tab, Button, Tooltip, tokens, Caption1, MessageBar, MessageBarBody, Badge,
} from '@fluentui/react-components';
import {
  Copy20Regular, ArrowReset20Regular, Play20Filled, Code20Filled, Edit20Filled,
} from '@fluentui/react-icons';
import { PaneHead } from '../editors/PaneHead';
import {
  generateCode, FORMAT_LABELS, FORMAT_LANG,
  generatePowerAutomateFields,
  type CodeFormat, type CodegenInputs,
} from '../engine/codeGenerators';
import { PowerAutomatePane } from './PowerAutomatePane';
import type { ThemeMode } from '../theme/theme';
import type { RequestGroup } from '../registry/requestTypes';

// Two tabs (top-level): "Code" (read-only Monaco, multi-format) and "Editor"
// (writable Monaco — playground starter code per format, with copy/reset).
// The Editor tab matches the Dataverse REST Builder UX: pick a starter, tweak
// it, copy or save as a snippet.
type ViewTab = 'code' | 'editor';
const FORMATS: CodeFormat[] = ['fetch', 'xrm', 'xrm-batch', 'xhr', 'powerautomate', 'csharp', 'powershell', 'curl', 'json'];
// Power Automate is a field-list form, not source code — skip it in the Editor tab.
const EDITOR_FORMATS: CodeFormat[] = FORMATS.filter(f => f !== 'powerautomate');

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

export interface CodeViewProps {
  themeMode: ThemeMode;
  inputs: CodegenInputs;
  group?: RequestGroup;
}

export function CodeView({ themeMode, inputs, group = 'read' }: CodeViewProps) {
  const [viewTab, setViewTab] = useState<ViewTab>('code');
  const [fmt, setFmt] = useState<CodeFormat>('fetch');
  const [code, setCode] = useState('');

  // Editor-tab state — separate so the user's edits aren't blown away when
  // they hop to the Code tab and back.
  const [editorFmt, setEditorFmt] = useState<CodeFormat>('fetch');
  const [editorCode, setEditorCode] = useState('');
  /** When the request state changes, the Editor tab keeps the user's edits
   *  unless they explicitly Reset. We track whether the editor was ever
   *  dirty-edited so the Reset button can be highlighted appropriately. */
  const [editorDirty, setEditorDirty] = useState(false);

  // Refresh the read-only Code monaco whenever the inputs or format change
  useEffect(() => {
    setCode(generateCode(fmt, inputs));
  }, [fmt, inputs]);

  // For the Editor tab, regenerate when the user picks a NEW starter (or first
  // mount). We deliberately don't re-generate on every `inputs` change there —
  // that would clobber their edits.
  useEffect(() => {
    setEditorCode(generateCode(editorFmt, inputs));
    setEditorDirty(false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editorFmt]);

  const monacoOptions: monacoNs.editor.IStandaloneEditorConstructionOptions = {
    lineNumbers: 'on',
    renderLineHighlight: 'line',
    minimap: { enabled: false },
    scrollBeyondLastLine: false,
    fontFamily: 'Cascadia Mono, Consolas, monospace',
    fontSize: 13,
    tabSize: 2,
    insertSpaces: true,
    wordWrap: 'on',
    automaticLayout: true,
    padding: { top: 12, bottom: 12 },
    formatOnPaste: true,
    formatOnType: true,
    bracketPairColorization: { enabled: true },
    guides: { bracketPairs: true, indentation: true },
    smoothScrolling: true,
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', minHeight: 0, height: '100%' }}>
      <PaneHead
        icon={Code20Filled}
        title="Generated code"
        sub={viewTab === 'code'
          ? '8 formats — read-only, regenerated on every builder change.'
          : 'Writable playground — edit the starter, copy or run it (mock).'}
        group={group}
      />

      <TabList
        selectedValue={viewTab}
        onTabSelect={(_, d) => setViewTab(d.value as ViewTab)}
        size="small"
        appearance="subtle"
        style={{ marginBottom: 10 }}
      >
        <Tab value="code" icon={<Code20Filled />}>Code</Tab>
        <Tab value="editor" icon={<Edit20Filled />}>Editor</Tab>
      </TabList>

      {viewTab === 'code' ? (
        <CodePane
          fmt={fmt} setFmt={setFmt}
          code={code}
          inputs={inputs}
          themeMode={themeMode}
          monacoOptions={monacoOptions}
        />
      ) : (
        <EditorPane
          fmt={editorFmt}
          setFmt={(f) => { setEditorFmt(f); /* effect resets code */ }}
          code={editorCode}
          setCode={(s) => { setEditorCode(s); setEditorDirty(true); }}
          onReset={() => { setEditorCode(generateCode(editorFmt, inputs)); setEditorDirty(false); }}
          dirty={editorDirty}
          themeMode={themeMode}
          monacoOptions={{ ...monacoOptions, readOnly: false }}
        />
      )}
    </div>
  );
}

// ────────────────────────────────────────────────────────────
// Code pane (read-only)
// ────────────────────────────────────────────────────────────
function CodePane({
  fmt, setFmt, code, inputs, themeMode, monacoOptions,
}: {
  fmt: CodeFormat;
  setFmt: (f: CodeFormat) => void;
  code: string;
  inputs: CodegenInputs;
  themeMode: ThemeMode;
  monacoOptions: monacoNs.editor.IStandaloneEditorConstructionOptions;
}) {
  // Power Automate gets the connector-style form view, not a Monaco editor —
  // each field is its own labeled input with a per-row Copy button so users
  // can paste one value at a time into the connector's UI.
  const isPA = fmt === 'powerautomate';
  const paSpec = isPA ? generatePowerAutomateFields(inputs) : null;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', minHeight: 0, flexGrow: 1 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10, flexWrap: 'wrap' }}>
        <TabList
          selectedValue={fmt}
          onTabSelect={(_, d) => setFmt(d.value as CodeFormat)}
          appearance="subtle"
          size="small"
        >
          {FORMATS.map(f => (
            <Tab key={f} value={f}>{FORMAT_LABELS[f]}</Tab>
          ))}
        </TabList>
        <span style={{ flexGrow: 1 }} />
        {!isPA && (
          <>
            <Caption1 style={{ color: tokens.colorNeutralForeground3, fontFamily: tokens.fontFamilyMonospace }}>
              {code.length.toLocaleString()} chars · {Math.round(new Blob([code]).size / 1024 * 10) / 10} KB
            </Caption1>
            <Tooltip content="Copy code" relationship="label">
              <Button icon={<Copy20Regular />} appearance="subtle" size="small"
                onClick={() => navigator.clipboard?.writeText(code)}>Copy</Button>
            </Tooltip>
          </>
        )}
      </div>

      {isPA && paSpec ? (
        <div style={{ flexGrow: 1, overflowY: 'auto', paddingRight: 4 }}>
          <PowerAutomatePane spec={paSpec} />
        </div>
      ) : (
        <>
          <MonacoFrame>
            <Editor
              height="100%"
              language={FORMAT_LANG[fmt]}
              value={code}
              theme={themeMode === 'dark' ? 'vs-dark' : 'light'}
              options={{ ...monacoOptions, readOnly: true, domReadOnly: true }}
            />
          </MonacoFrame>

          <Caption1 style={{ marginTop: 8, color: tokens.colorNeutralForeground3 }}>
            Tip: replace <code>&lt;access-token&gt;</code> with a real bearer token (e.g. from <code>Get-AzAccessToken</code>, MSAL, or the connector's auth step) before running.
          </Caption1>
        </>
      )}
    </div>
  );
}

// ────────────────────────────────────────────────────────────
// Editor pane (writable — playground)
// ────────────────────────────────────────────────────────────
function EditorPane({
  fmt, setFmt, code, setCode, onReset, dirty, themeMode, monacoOptions,
}: {
  fmt: CodeFormat;
  setFmt: (f: CodeFormat) => void;
  code: string;
  setCode: (s: string) => void;
  onReset: () => void;
  dirty: boolean;
  themeMode: ThemeMode;
  monacoOptions: monacoNs.editor.IStandaloneEditorConstructionOptions;
}) {
  const [output, setOutput] = useState<string>('');
  const [running, setRunning] = useState(false);
  const editorRef = useRef<monacoNs.editor.IStandaloneCodeEditor | null>(null);
  const onMount: OnMount = (editor) => { editorRef.current = editor; };

  const onMockRun = async () => {
    setRunning(true);
    setOutput('');
    try {
      // We can't safely eval arbitrary JS against a real Dataverse, so the
      // mock Run simulates by intercepting `fetch` and `console.log` and
      // routing the request to a deterministic mock response. For now we just
      // show a structured echo of what the code would attempt + a stub result.
      // (Hooking up the real mock executor would require parsing the URL out
      // of the user-edited code — left as a follow-up.)
      await new Promise(r => setTimeout(r, 350));
      const sampleResp = {
        status: 200,
        ok: true,
        body: {
          '@odata.context': '$metadata#accounts(name,revenue)',
          value: [
            { name: 'Contoso Ltd.',  revenue: 18_500_000 },
            { name: 'Fabrikam, Inc.', revenue: 9_120_000 },
          ],
        },
      };
      setOutput(`✓ Mock run finished\n\n${JSON.stringify(sampleResp, null, 2)}`);
    } catch (e) {
      setOutput(String(e));
    } finally {
      setRunning(false);
    }
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', minHeight: 0, flexGrow: 1 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10, flexWrap: 'wrap' }}>
        <Caption1 style={{ color: tokens.colorNeutralForeground3, marginRight: 4 }}>Starter:</Caption1>
        <TabList
          selectedValue={fmt}
          onTabSelect={(_, d) => setFmt(d.value as CodeFormat)}
          appearance="subtle"
          size="small"
        >
          {EDITOR_FORMATS.map(f => (
            <Tab key={f} value={f}>{FORMAT_LABELS[f]}</Tab>
          ))}
        </TabList>
        <span style={{ flexGrow: 1 }} />
        {dirty && <Badge size="small" appearance="ghost" color="warning">edited</Badge>}
        <Tooltip content="Reset to the freshly-generated starter for this format" relationship="description">
          <Button icon={<ArrowReset20Regular />} appearance="subtle" size="small" onClick={onReset}>
            Reset
          </Button>
        </Tooltip>
        <Tooltip content="Copy current editor contents" relationship="label">
          <Button icon={<Copy20Regular />} appearance="subtle" size="small"
            onClick={() => navigator.clipboard?.writeText(code)}>Copy</Button>
        </Tooltip>
        <Tooltip content="Run against the mock simulator — returns sample data without touching a real environment" relationship="description">
          <Button icon={<Play20Filled />} appearance="primary" size="small" onClick={onMockRun} disabled={running}>
            {running ? 'Running…' : 'Run (mock)'}
          </Button>
        </Tooltip>
      </div>

      <MessageBar layout="multiline" intent="info" style={{ marginBottom: 10 }}>
        <MessageBarBody>
          <strong>Playground.</strong> Edit the starter however you like, copy it out, or hit <strong>Run (mock)</strong> to see a simulated response below. Reset re-generates the starter from the current builder state.
        </MessageBarBody>
      </MessageBar>

      <div style={{ display: 'grid', gridTemplateRows: '1fr auto', gap: 10, minHeight: 0, flexGrow: 1 }}>
        <MonacoFrame>
          <Editor
            height="100%"
            language={FORMAT_LANG[fmt]}
            value={code}
            onChange={(v) => setCode(v ?? '')}
            theme={themeMode === 'dark' ? 'vs-dark' : 'light'}
            options={monacoOptions}
            onMount={onMount}
          />
        </MonacoFrame>

        {output && (
          <div style={{
            borderRadius: tokens.borderRadiusMedium,
            border: `1px solid ${tokens.colorNeutralStroke2}`,
            backgroundColor: tokens.colorNeutralBackground1,
            padding: 12,
            maxHeight: 220,
            overflow: 'auto',
          }}>
            <Caption1 style={{
              display: 'block',
              fontWeight: 600,
              color: tokens.colorNeutralForeground2,
              marginBottom: 6,
            }}>
              Mock output
            </Caption1>
            <pre style={{
              margin: 0,
              fontFamily: tokens.fontFamilyMonospace,
              fontSize: 12,
              color: tokens.colorNeutralForeground1,
              whiteSpace: 'pre-wrap',
              wordBreak: 'break-word',
            }}>{output}</pre>
          </div>
        )}
      </div>
    </div>
  );
}

// Shared Monaco container — bordered + radius + min-height
function MonacoFrame({ children }: { children: React.ReactNode }) {
  return (
    <div style={{
      flexGrow: 1,
      border: `1px solid ${tokens.colorNeutralStroke2}`,
      borderRadius: tokens.borderRadiusMedium,
      overflow: 'hidden',
      minHeight: 320,
    }}>
      {children}
    </div>
  );
}
