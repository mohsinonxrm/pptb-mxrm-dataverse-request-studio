// The editor `MonacoEditor` renders: `@monaco-editor/react`'s Editor, plus
// an optional JSON Schema for its model (#268). Lives in the lazy Monaco
// chunk (monacoSetup re-exports it), so callers pass a plain object and
// never import Monaco themselves.
//
// Monaco's JSON language service takes every schema in one global option,
// matched to models by file name. So each JSON editor gets its own model
// path, and the schemas of the mounted editors are collected here and
// applied together.
import { useEffect, useId } from 'react';
import Editor, { type EditorProps } from '@monaco-editor/react';
import * as monaco from 'monaco-editor/esm/vs/editor/editor.api';
import type { BodyJsonSchema } from '../engine/bodySchema';

export interface SchemaEditorProps extends EditorProps {
  /** JSON Schema for a `language="json"` editor: completions, hover, type checks. */
  jsonSchema?: BodyJsonSchema;
}

// Model file name → schema, as JSON text (compared by value)
const schemas = new Map<string, string>();
let appliedKey = '[]';
let applyQueued = false;

// Changing the JSON options restarts the JSON worker, so apply once per tick
// (a table switch removes the old schema and adds the new one) and only when
// the schemas changed.
function queueApply(): void {
  if (applyQueued) return;
  applyQueued = true;
  queueMicrotask(() => {
    applyQueued = false;
    const entries = [...schemas].sort(([a], [b]) => a.localeCompare(b));
    const key = JSON.stringify(entries);
    if (key === appliedKey) return;
    appliedKey = key;
    const json = monaco.languages.json.jsonDefaults;
    json.setDiagnosticsOptions({
      ...json.diagnosticsOptions,
      schemas: entries.map(([file, text]) => ({
        uri: `inmemory://drs-schema/${file}`,
        fileMatch: [file],
        schema: JSON.parse(text),
      })),
    });
  });
}

export default function SchemaEditor({ jsonSchema, ...props }: SchemaEditorProps) {
  const id = useId();
  const file = `drs-${id.replace(/[^A-Za-z0-9]/g, '')}.json`;
  const schemaText = jsonSchema ? JSON.stringify(jsonSchema) : undefined;

  useEffect(() => {
    if (!schemaText) return;
    schemas.set(file, schemaText);
    queueApply();
    return () => {
      schemas.delete(file);
      queueApply();
    };
  }, [file, schemaText]);

  // A JSON editor always gets its own path, so a schema that arrives later
  // (metadata still loading) doesn't swap the model under the user
  const path = props.path ?? (props.language === 'json' ? file : undefined);
  return <Editor {...props} path={path} />;
}
