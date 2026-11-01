import { lazy, Suspense } from 'react';
import type { SchemaEditorProps } from './SchemaEditor';
import { Spinner } from '@fluentui/react-components';

// Drop-in replacement for `@monaco-editor/react`'s default `Editor` export.
// Monaco (and its loader/worker setup in ./monacoSetup) is fetched the first
// time any editor renders — the Code tab, the JSON body editor, the action
// parameter JSON view or the predefined-query FetchXML view — instead of at
// startup (#22). Later editors render immediately from the cached chunk.
const LazyEditor = lazy(() => import('./monacoSetup'));

export default function MonacoEditor(props: SchemaEditorProps) {
  return (
    <Suspense
      fallback={
        <div
          style={{
            height: props.height ?? '100%',
            width: props.width ?? '100%',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
          }}
        >
          <Spinner size="small" label="Loading editor…" />
        </div>
      }
    >
      <LazyEditor {...props} />
    </Suspense>
  );
}
