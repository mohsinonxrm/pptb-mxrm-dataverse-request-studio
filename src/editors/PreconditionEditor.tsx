// PreconditionEditor — unified HTTP-precondition picker for write modes.
//
// Replaces (Update mode) IfMatchCardEditor, (Delete mode) ConcurrencyEditor,
// and (Upsert mode) UpsertModeEditor + a second ConcurrencyEditor. All three
// modes were rendering the SAME `state.concurrency: ConcurrencyMode` field
// through different UIs with different labels, causing semantic drift across
// modes that share the same underlying setting.
//
// Source-of-truth terminology aligns with Microsoft's Web API docs which call
// `If-Match` / `If-None-Match` "preconditions". The four kinds:
//
//   • none         → no header emitted; server proceeds regardless
//   • update-only  → If-Match: *           → 412 if record doesn't exist
//   • create-only  → If-None-Match: *      → 412 if record exists (Upsert-only)
//   • etag         → If-Match: "<etag>"    → 412 unless current version matches
//
// Caller-supplied `available` array constrains which kinds the radio list
// surfaces. Mode-specific rules:
//
//   Delete  : ['none', 'update-only', 'etag']                (create-only doesn't apply)
//   Update  : ['none', 'update-only', 'etag']                (same — no "update if absent")
//   Upsert  : ['none', 'create-only', 'update-only', 'etag']  (all four)
//
// Visual: stacked radio cards with title + emitted header (mono) + one-line
// description. The etag card expands inline when selected — no separate
// "Advanced" sub-pane like Upsert previously had.

import {
  Field,
  Input,
  Button,
  Caption1,
  Radio,
  RadioGroup,
  tokens,
  mergeClasses,
} from '@fluentui/react-components';
import { ShieldLock20Filled } from '@fluentui/react-icons';
import { PaneHead } from './PaneHead';
import { useStudioStyles } from '../primitives/styles';
import type { ConcurrencyMode } from '../state/writeState';
import type { RequestGroup } from '../registry/requestTypes';

export type PreconditionKind = ConcurrencyMode['kind'];

export interface PreconditionEditorProps {
  /** Current value (matches state.concurrency on Delete/Update/Upsert). */
  mode: ConcurrencyMode;
  setMode: (m: ConcurrencyMode) => void;
  /**
   * Which kinds to surface. Each write mode picks a subset; the component
   * doesn't decide policy. Default is all four. Order in the array is
   * preserved in the UI.
   */
  available?: PreconditionKind[];
  /**
   * Live ETag of the target record (when known). Drives the "Use current"
   * button next to the etag input. Update mode supplies this from the
   * row-preview fetch; Delete + Upsert can wire the same.
   */
  currentEtag?: string | null;
  /** Theming hook for the pane header — same convention other editors use. */
  group?: RequestGroup;
}

/** Static label + description for each kind. Stays the same across modes; the
 *  caller only controls availability. */
const KIND_META: Record<PreconditionKind, { title: string; header: string; body: string }> = {
  none: {
    title: 'None',
    header: '(no header)',
    body: "Proceed regardless of the row's state. Server decides what to do (default for Upsert: create if missing, update if present).",
  },
  'update-only': {
    title: 'Require existing record',
    header: 'If-Match: *',
    body: "Fail with 412 Precondition Failed if no record matches the addressed key. Use to guarantee you're updating an existing row.",
  },
  'create-only': {
    title: 'Require absent record',
    header: 'If-None-Match: *',
    body: 'Fail with 412 Precondition Failed if a record already matches the addressed key. Upsert "insert-only" — never overwrite.',
  },
  etag: {
    title: 'Match specific version (ETag)',
    header: 'If-Match: "<etag>"',
    body: 'Fail with 412 Precondition Failed if the record has been modified since the supplied ETag was issued. Classic optimistic concurrency.',
  },
};

const ALL_KINDS: PreconditionKind[] = ['none', 'update-only', 'create-only', 'etag'];

export function PreconditionEditor({
  mode,
  setMode,
  available = ALL_KINDS,
  currentEtag,
  group = 'write',
}: PreconditionEditorProps) {
  const s = useStudioStyles();
  const selected = mode.kind;

  const onPick = (kind: PreconditionKind) => {
    if (kind === 'etag') {
      // Preserve any existing etag value when re-selecting; otherwise empty.
      const existing = mode.kind === 'etag' ? mode.etag : '';
      setMode({ kind: 'etag', etag: existing });
    } else {
      setMode({ kind });
    }
  };

  return (
    <div>
      <PaneHead
        icon={ShieldLock20Filled}
        title="Precondition"
        sub="Optimistic concurrency & existence checks — controls the If-Match / If-None-Match header."
        group={group}
      />

      <div style={{ maxWidth: 720 }}>
        <RadioGroup value={selected} onChange={(_, d) => onPick(d.value as PreconditionKind)}>
          {available.map((kind) => {
            const meta = KIND_META[kind];
            const isSelected = selected === kind;
            return (
              <div
                key={kind}
                // Reuse the existing inlineCard primitive used by every
                // other write editor — keeps the visual weight consistent
                // across the mode group.
                className={mergeClasses(s.inlineCard)}
                style={{
                  padding: '10px 12px',
                  marginBottom: 8,
                  borderColor: isSelected ? tokens.colorBrandStroke1 : undefined,
                  background: isSelected ? tokens.colorBrandBackground2 : undefined,
                }}
              >
                <Radio
                  value={kind}
                  label={
                    <span style={{ display: 'inline-flex', flexDirection: 'column', gap: 4 }}>
                      <span
                        style={{
                          display: 'inline-flex',
                          alignItems: 'center',
                          gap: 10,
                          flexWrap: 'wrap',
                        }}
                      >
                        <strong style={{ fontSize: 13 }}>{meta.title}</strong>
                        <code
                          style={{
                            fontFamily: tokens.fontFamilyMonospace,
                            fontSize: 11,
                            color: tokens.colorBrandForeground2,
                            background: tokens.colorBrandBackground2Hover,
                            padding: '2px 6px',
                            borderRadius: tokens.borderRadiusSmall,
                          }}
                        >
                          {meta.header}
                        </code>
                      </span>
                      <Caption1 style={{ color: tokens.colorNeutralForeground2, lineHeight: 1.45 }}>
                        {meta.body}
                      </Caption1>
                    </span>
                  }
                />

                {/* ETag input renders inline under the radio when "etag" is selected.
                    Replaces the old separate "Advanced" sub-pane in Upsert. */}
                {kind === 'etag' && isSelected && (
                  <div
                    style={{
                      marginTop: 10,
                      marginLeft: 26,
                      display: 'flex',
                      flexDirection: 'column',
                      gap: 6,
                      maxWidth: 520,
                    }}
                  >
                    <Field
                      label="ETag value"
                      hint='Pulled from a previous response&apos;s @odata.etag — looks like `W/"7281965"`.'
                    >
                      <span style={{ display: 'flex', gap: 6 }}>
                        <Input
                          value={mode.kind === 'etag' ? mode.etag : ''}
                          onChange={(_, d) => setMode({ kind: 'etag', etag: d.value })}
                          placeholder='W/"12345678"'
                          style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 12, flex: 1 }}
                        />
                        {currentEtag && (
                          <Button
                            size="small"
                            appearance="outline"
                            onClick={() => setMode({ kind: 'etag', etag: currentEtag })}
                            title={`Populate from the current record's @odata.etag (${currentEtag})`}
                          >
                            Use current
                          </Button>
                        )}
                      </span>
                    </Field>
                  </div>
                )}
              </div>
            );
          })}
        </RadioGroup>
      </div>
    </div>
  );
}

/**
 * Convert a precondition to the header it emits on the wire. Returns
 * `null` when no header should be sent (kind === 'none' or etag has no
 * value). Pure function — same logic used to live inline in each mode's
 * effectiveHeaders memo and in `concurrencyToHeader` in ConcurrencyEditor.
 *
 * Re-exported here so consumers can import everything from one module.
 */
export function preconditionToHeader(m: ConcurrencyMode): { name: string; value: string } | null {
  switch (m.kind) {
    case 'update-only':
      return { name: 'If-Match', value: '*' };
    case 'create-only':
      return { name: 'If-None-Match', value: '*' };
    case 'etag':
      if (!m.etag.trim()) return null;
      return { name: 'If-Match', value: m.etag };
    case 'none':
    default:
      return null;
  }
}
