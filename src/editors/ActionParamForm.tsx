// ActionParamForm — metadata-driven form for an OData action/function's
// parameters. Each param renders as the appropriate Fluent v9 control based
// on its CSDL EDM type:
//
//   • Edm.String                       → Input
//   • Edm.Int32/Int64/Decimal/Double   → SpinButton
//   • Edm.Boolean                      → Switch
//   • Edm.Guid                         → Input (mono)
//   • Edm.DateTimeOffset               → DatePicker + TimePicker
//   • OptionSetValue                   → Combobox over the per-param option set
//   • Collection(Edm.String|Int32|Guid)→ Tag editor
//   • EntityReference                  → entity-type picker + record picker
//   • EntitySpecific                   → record picker scoped to the CSDL type
//   • EntityCollection                 → multi-target record picker
//   • ComplexType                      → recursive nested form
//
// Reference: https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-web-api-actions

import { useEffect, useMemo, useState } from 'react';
import {
  Input,
  SpinButton,
  Switch,
  Combobox,
  Option,
  Persona,
  Field,
  Caption1,
  Badge,
  Tag,
  TagPicker,
  TagPickerControl,
  TagPickerGroup,
  TagPickerInput,
  TagPickerList,
  TagPickerOption,
  useTagPickerFilter,
  MessageBar,
  MessageBarBody,
  tokens,
  mergeClasses,
  Spinner,
  ToggleButton,
  Link,
  Button,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
} from '@fluentui/react-components';
import { DatePicker } from '@fluentui/react-datepicker-compat';
import { TimePicker } from '@fluentui/react-timepicker-compat';
import {
  Settings20Filled,
  Document20Regular,
  FormNew20Regular,
  FormNew20Filled,
  Code20Regular,
  Open20Regular,
} from '@fluentui/react-icons';
import Editor from '@monaco-editor/react';
import { SegmentedToggle } from '../primitives/SegmentedToggle';
import type { ThemeMode } from '../theme/theme';
import { FieldSetEditor } from './FieldSetEditor';
import type { CreateFieldValue } from '../state/writeState';
import { useStudioStyles } from '../primitives/styles';
import { PaneHead } from './PaneHead';
import { findTable } from '../mock/metadata';
import { RecordPicker } from '../primitives/RecordPicker';
import { useScopedEntities } from '../host/useScopedEntities';
import type { ActionParam, CsdlAction } from '../mock/actionsCsdl';
import type { ExecParamValue } from '../state/executeState';
import type { RequestGroup } from '../registry/requestTypes';
// Live-metadata warmer — kicks off the fetch for the entity's
// AttributeMetadata so findTable() returns a populated TableMeta on the
// next render. Without this, EntitySpecific params (e.g. WinOpportunity's
// OpportunityClose: mscrm.opportunityclose) would never resolve their
// typed FieldSetEditor and stay stuck on the JSON-fallback input.
import { useLiveTable as warmTableMetadata } from '../host/useLiveMetadata';

// Per design feedback: no "smart" Status/StatusCode inference. Showing a
// labeled optionset for some Edm.Int32 params and a SpinButton for others
// creates UX inconsistency the user can't predict. All Edm.Int32 params
// render as plain SpinButton; if the user needs an enum value they consult
// the docs or use the JSON view.

export interface ActionParamFormProps {
  action: CsdlAction;
  values: Record<string, ExecParamValue>;
  setValue: (paramName: string, value: ExecParamValue) => void;
  /**
   * Optional bulk setter used by the JSON view to replace the entire param
   * map at once. When omitted, the JSON toggle is hidden and only the form
   * view is available. Modes that wire it up (Execute Action / Function /
   * etc.) get a "build the request as raw JSON" escape hatch.
   */
  setValues?: (next: Record<string, ExecParamValue>) => void;
  group?: RequestGroup;
  /** Theme mode for the Monaco JSON editor — drives vs-dark vs light. */
  themeMode?: ThemeMode;
}

/**
 * MS Learn reference index URL. Per design feedback — not every action /
 * function has a canonical docs page (custom APIs, custom actions, and
 * some niche OOB items aren't documented individually). Linking to the
 * INDEX page lets the user search from there reliably, instead of risking
 * 404s on per-op deep links.
 *
 *   Actions:   /webapi/reference/actions?view=dataverse-latest
 *   Functions: /webapi/reference/functions?view=dataverse-latest
 */
function learnUrlFor(action: CsdlAction): string {
  const subpath = action.kind === 'Action' ? 'actions' : 'functions';
  return `https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/${subpath}?view=dataverse-latest`;
}

export function ActionParamForm({
  action,
  values,
  setValue,
  setValues,
  group = 'execute',
  themeMode = 'light',
}: ActionParamFormProps) {
  const s = useStudioStyles();

  // ── Form / JSON toggle state ──
  //
  // Mirrors FieldSetEditor's pattern: form mode is the default typed-input
  // view; JSON mode is the escape hatch for cases CSDL can't express
  // (Edm.Int32 status codes, EntitySpecific reference vs. instance choice,
  // domain-specific Custom API params with no metadata clues). Only
  // enabled when the parent passes `setValues` — otherwise we can't write
  // back the parsed body, so the toggle hides.
  const [bodyMode, setBodyMode] = useState<'form' | 'json'>('form');
  const [jsonText, setJsonText] = useState<string>('');
  const [jsonError, setJsonError] = useState<string | null>(null);
  // Snapshot values into jsonText when entering JSON mode, OR when values
  // change externally (e.g. after a Save load). Form-mode edits keep
  // jsonText in sync via the same effect.
  useEffect(() => {
    if (bodyMode !== 'json') {
      // Pre-warm the snapshot too so a future Form→JSON switch lands fast.
      try {
        setJsonText(JSON.stringify(values, null, 2));
      } catch {
        /* ignore */
      }
      return;
    }
    try {
      setJsonText(JSON.stringify(values, null, 2));
      setJsonError(null);
    } catch (e) {
      setJsonError(e instanceof Error ? e.message : String(e));
    }
  }, [bodyMode, values]);

  const learnUrl = learnUrlFor(action);

  if (action.parameters.length === 0) {
    return (
      <div>
        <PaneHead
          icon={Settings20Filled}
          title="Parameters"
          sub="No parameters — this operation can be invoked as-is."
          group={group}
        >
          {learnUrl && <DocsLink href={learnUrl} />}
        </PaneHead>
        <MessageBar layout="multiline" intent="info" style={{ maxWidth: 720 }}>
          <MessageBarBody>
            <code>{action.name}</code> takes no parameters. Click <strong>Execute</strong> to run
            it.
          </MessageBarBody>
        </MessageBar>
      </div>
    );
  }

  const missingRequired = action.parameters
    .filter((p) => p.required)
    .filter((p) => isEmpty(values[p.name]));

  return (
    <div>
      <PaneHead
        icon={Settings20Filled}
        title="Parameters"
        sub={
          <>
            Required parameters first (marked{' '}
            <Badge appearance="tint" color="danger" size="extra-small">
              req
            </Badge>
            ). Optional parameters live in the accordion below.
            {missingRequired.length > 0 && <> · {missingRequired.length} required unset.</>}
          </>
        }
        group={group}
      >
        <Badge appearance="ghost">
          {action.parameters.filter((p) => !isEmpty(values[p.name])).length} of{' '}
          {action.parameters.length} set
        </Badge>
        {missingRequired.length > 0 && (
          <Badge appearance="tint" color="danger">
            {missingRequired.length} required unset
          </Badge>
        )}
        {learnUrl && <DocsLink href={learnUrl} />}
        {/* Form / JSON toggle — only rendered when the parent gave us a
            bulk setter. Same SegmentedToggle the FieldSetEditor uses so the
            visual language is consistent across the studio. */}
        {setValues && (
          <SegmentedToggle ariaLabel="Parameter editor mode">
            <ToggleButton
              checked={bodyMode === 'form'}
              icon={<FormNew20Regular />}
              onClick={() => setBodyMode('form')}
            >
              Form
            </ToggleButton>
            <ToggleButton
              checked={bodyMode === 'json'}
              icon={<Code20Regular />}
              onClick={() => setBodyMode('json')}
            >
              JSON
            </ToggleButton>
          </SegmentedToggle>
        )}
      </PaneHead>

      {bodyMode === 'form' ? (
        (() => {
          // Split parameters into required + optional. Required render first
          // (the user must fill these to enable Execute); optional collapse
          // into an accordion at the bottom so the main pane stays focused
          // on what's blocking the request.
          const requiredParams = action.parameters.filter((p) => p.required);
          const optionalParams = action.parameters.filter((p) => !p.required);
          const optionalSet = optionalParams.filter((p) => !isEmpty(values[p.name])).length;
          return (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 12, maxWidth: 880 }}>
              {/* Required parameters — always visible */}
              {requiredParams.map((param) => (
                <ParamRow
                  key={param.name}
                  action={action}
                  param={param}
                  value={values[param.name]}
                  onChange={(v) => setValue(param.name, v)}
                  styles={s}
                />
              ))}
              {requiredParams.length === 0 && optionalParams.length > 0 && (
                <Caption1
                  style={{
                    color: tokens.colorNeutralForeground3,
                    fontStyle: 'italic',
                    padding: '4px 2px',
                  }}
                >
                  This operation has no required parameters — all inputs below are optional.
                </Caption1>
              )}

              {/* Optional parameters — collapsed into an accordion. The
                  user expands when they need to override a default. The
                  header reports "N of M set" so the user can see at-a-
                  glance whether any optional values are already populated. */}
              {optionalParams.length > 0 && (
                <Accordion
                  collapsible
                  multiple
                  // Auto-expand when any optional has a value (e.g. after a
                  // saved-request load) — keeps populated fields visible.
                  defaultOpenItems={optionalSet > 0 ? ['optional'] : []}
                >
                  <AccordionItem value="optional">
                    <AccordionHeader>
                      <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
                        <span style={{ fontWeight: 600 }}>Optional parameters</span>
                        <Badge appearance="ghost" size="small">
                          {optionalSet} of {optionalParams.length} set
                        </Badge>
                      </span>
                    </AccordionHeader>
                    <AccordionPanel>
                      <div
                        style={{ display: 'flex', flexDirection: 'column', gap: 12, paddingTop: 6 }}
                      >
                        {optionalParams.map((param) => (
                          <ParamRow
                            key={param.name}
                            action={action}
                            param={param}
                            value={values[param.name]}
                            onChange={(v) => setValue(param.name, v)}
                            styles={s}
                          />
                        ))}
                      </div>
                    </AccordionPanel>
                  </AccordionItem>
                </Accordion>
              )}
            </div>
          );
        })()
      ) : (
        // JSON view — Monaco editor with the body shape. Edits parse back
        // to `values` on every keystroke; invalid JSON keeps the prior
        // values until the user fixes it (jsonError is surfaced inline).
        <JsonEditPane
          text={jsonText}
          error={jsonError}
          themeMode={themeMode}
          onChange={(t) => {
            setJsonText(t);
            try {
              const parsed = JSON.parse(t || '{}');
              if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
                setJsonError(null);
                setValues?.(parsed as Record<string, ExecParamValue>);
              } else {
                setJsonError('Top-level must be a JSON object.');
              }
            } catch (e) {
              setJsonError(e instanceof Error ? e.message : String(e));
            }
          }}
        />
      )}
    </div>
  );
}

/** External-link badge that points at the action's MS Learn reference. */
function DocsLink({ href }: { href: string }) {
  return (
    <Link href={href} target="_blank" rel="noopener noreferrer" style={{ fontSize: 11 }}>
      <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
        MS Learn <Open20Regular style={{ width: 12, height: 12 }} />
      </span>
    </Link>
  );
}

/** Monaco-backed JSON editor with inline parse-error indicator. */
function JsonEditPane({
  text,
  error,
  themeMode,
  onChange,
}: {
  text: string;
  error: string | null;
  themeMode: ThemeMode;
  onChange: (next: string) => void;
}) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8, maxWidth: 880 }}>
      <MessageBar layout="multiline" intent="info">
        <MessageBarBody>
          Hand-craft the request body as JSON. Useful when the metadata-driven form can't express
          the shape — e.g. Int32 params whose values map to documented enums, EntitySpecific params
          that should be <code>{'{ "@odata.id": "<set>(<id>)" }'}</code> references instead of
          inline instances, or Custom API params with domain-specific shapes. Refer to the MS Learn
          page for the canonical body.
        </MessageBarBody>
      </MessageBar>
      <div
        style={{
          border: `1px solid ${error ? tokens.colorPaletteRedBorder2 : tokens.colorNeutralStroke2}`,
          borderRadius: tokens.borderRadiusMedium,
          overflow: 'hidden',
          height: 360,
        }}
      >
        <Editor
          height="100%"
          language="json"
          value={text}
          onChange={(v) => onChange(v ?? '')}
          theme={themeMode === 'dark' ? 'vs-dark' : 'light'}
          options={{
            lineNumbers: 'on',
            minimap: { enabled: false },
            scrollBeyondLastLine: false,
            fontFamily: 'Cascadia Mono, Consolas, monospace',
            fontSize: 13,
            tabSize: 2,
            wordWrap: 'on',
            automaticLayout: true,
            padding: { top: 12, bottom: 12 },
            bracketPairColorization: { enabled: true },
            guides: { bracketPairs: true, indentation: true },
            smoothScrolling: true,
            formatOnPaste: true,
          }}
        />
      </div>
      {error && (
        <Caption1 style={{ color: tokens.colorPaletteRedForeground1 }}>
          Parse error: {error}
        </Caption1>
      )}
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// One parameter row
// ──────────────────────────────────────────────────────────────
function ParamRow({
  action,
  param,
  value,
  onChange,
  styles,
}: {
  /** Owning action — needed by ParamInput to resolve action-level context
   *  (e.g. the Status param's entity statuscode optionset). */
  action: CsdlAction;
  param: ActionParam;
  value: ExecParamValue;
  onChange: (v: ExecParamValue) => void;
  styles: ReturnType<typeof useStudioStyles>;
}) {
  const missing = param.required && isEmpty(value);
  return (
    <div
      className={mergeClasses(styles.inlineCard)}
      style={{
        padding: 12,
        display: 'grid',
        gridTemplateColumns: '220px 1fr',
        gap: 14,
        alignItems: 'start',
        borderColor: missing ? tokens.colorPaletteRedBorder2 : undefined,
      }}
    >
      <div style={{ display: 'flex', flexDirection: 'column', gap: 2, paddingTop: 4 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <span style={{ fontFamily: tokens.fontFamilyMonospace, fontWeight: 600 }}>
            {param.name}
          </span>
          {param.required && (
            <Badge appearance="tint" color="danger" size="extra-small">
              req
            </Badge>
          )}
        </div>
        <Caption1 style={{ color: tokens.colorNeutralForeground3, fontSize: 10 }}>
          <span
            style={{ color: tokens.colorBrandForeground2, fontFamily: tokens.fontFamilyMonospace }}
          >
            {param.type}
          </span>
          {param.entityType && <> · {param.entityType}</>}
        </Caption1>
      </div>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 4, minWidth: 0 }}>
        <ParamInput action={action} param={param} value={value} onChange={onChange} />
        {param.description && (
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>{param.description}</Caption1>
        )}
      </div>
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Per-type input dispatch
// ──────────────────────────────────────────────────────────────
function ParamInput({
  action,
  param,
  value,
  onChange,
}: {
  /**
   * Owning action — used to derive Status-context for Int32 params at the
   * ACTION parameter level. Optional because ComplexTypeInput recursively
   * renders nested fields via ParamInput, and a nested complex-type field
   * named "Status" should NOT inherit the parent action's statuscode
   * context (it's a property of the complex type, not the action).
   */
  action?: CsdlAction;
  param: ActionParam;
  value: ExecParamValue;
  onChange: (v: ExecParamValue) => void;
}) {
  switch (param.type) {
    case 'Edm.String':
      return (
        <Input
          size="small"
          value={(value as string) ?? ''}
          onChange={(_, d) => onChange(d.value)}
          placeholder="value"
        />
      );
    case 'Edm.Guid':
      return (
        <Input
          size="small"
          value={(value as string) ?? ''}
          onChange={(_, d) => onChange(d.value)}
          placeholder="00000000-0000-0000-0000-000000000000"
          style={{ fontFamily: tokens.fontFamilyMonospace }}
        />
      );
    case 'Edm.Int32':
    case 'Edm.Int64':
      // Integer SpinButton — step=1 + precision=0 to lock out decimal
      // entry. SpinButton's `onChange` can emit either `value` (when
      // typed-then-blurred) or `displayValue` (when keypress); we coerce
      // to a whole number via Math.trunc so a user typing "3.5" still
      // commits as 3, matching Edm.Int32 semantics.
      return (
        <SpinButton
          size="small"
          value={Number.isFinite(value as number) ? (value as number) : 0}
          step={1}
          precision={0}
          onChange={(_, d) => {
            const raw =
              d.value != null ? d.value : d.displayValue != null ? Number(d.displayValue) : NaN;
            if (Number.isFinite(raw)) onChange(Math.trunc(raw));
          }}
        />
      );
    case 'Edm.Decimal':
    case 'Edm.Double':
      return (
        <SpinButton
          size="small"
          value={Number(value ?? 0)}
          step={0.01}
          onChange={(_, d) => onChange(d.value ?? Number(d.displayValue ?? 0))}
        />
      );
    case 'Edm.Boolean':
      return (
        <Switch
          checked={value === true}
          onChange={(_, d) => onChange(d.checked)}
          label={value === true ? 'true' : 'false'}
        />
      );
    case 'Edm.DateTimeOffset':
      return <DateTimeInput value={(value as string) ?? ''} onChange={onChange} />;
    case 'OptionSetValue':
      return (
        <OptionSetInput param={param} value={value as number | undefined} onChange={onChange} />
      );
    case 'Collection(Edm.String)':
      return <StringArrayInput value={(value as string[]) ?? []} onChange={onChange} />;
    case 'Collection(Edm.Int32)':
      return <NumberArrayInput value={(value as number[]) ?? []} onChange={onChange} />;
    case 'Collection(Edm.Guid)':
      return <StringArrayInput value={(value as string[]) ?? []} onChange={onChange} mono />;
    case 'EntityReference':
      // EntityReference → reference to an EXISTING row (record picker).
      return (
        <EntityReferenceInput
          value={value as { id: string; entityType: string } | undefined}
          onChange={onChange}
          param={param}
          polymorphic
        />
      );
    case 'EntitySpecific':
      // EntitySpecific → NEW entity instance constructed inline. Renders a
      // metadata-driven entity form (e.g. WinOpportunity.OpportunityClose).
      return (
        <EntitySpecificInput
          param={param}
          value={value as Record<string, unknown> | undefined}
          onChange={onChange}
        />
      );
    case 'EntityCollection':
      return (
        <EntityCollectionInput
          value={(value as Array<{ id: string; entityType: string }>) ?? []}
          onChange={onChange}
          param={param}
        />
      );
    case 'ComplexType':
      return (
        <ComplexTypeInput
          param={param}
          value={value as Record<string, unknown> | undefined}
          onChange={onChange}
        />
      );
    case 'Edm.Binary':
      return (
        <Input
          size="small"
          value={(value as string) ?? ''}
          onChange={(_, d) => onChange(d.value)}
          placeholder="base64-encoded bytes"
          style={{ fontFamily: tokens.fontFamilyMonospace }}
        />
      );
  }
}

// ── DateTime composite picker ──
function DateTimeInput({ value, onChange }: { value: string; onChange: (v: string) => void }) {
  const parsed = value ? new Date(value) : null;
  const valid = parsed && !Number.isNaN(parsed.getTime()) ? parsed : null;
  const write = (d: Date | null | undefined) => {
    if (!d) {
      onChange('');
      return;
    }
    const next = new Date(d);
    if (valid) next.setHours(valid.getHours(), valid.getMinutes(), 0, 0);
    onChange(next.toISOString());
  };
  const writeTime = (d: Date | null | undefined) => {
    if (!d || !valid) return;
    const next = new Date(valid);
    next.setHours(d.getHours(), d.getMinutes(), 0, 0);
    onChange(next.toISOString());
  };
  return (
    <span style={{ display: 'inline-flex', gap: 6, alignItems: 'center', flexWrap: 'wrap' }}>
      <DatePicker
        size="small"
        value={valid}
        onSelectDate={write}
        placeholder="Date…"
        formatDate={(d) => (d ? d.toLocaleDateString() : '')}
        allowTextInput
        style={{ minWidth: 130 }}
      />
      <TimePicker
        size="small"
        freeform
        selectedTime={valid}
        onTimeChange={(_, d) => writeTime(d.selectedTime)}
        placeholder="Time…"
        style={{ minWidth: 100 }}
        hourCycle="h23"
      />
    </span>
  );
}

// ── OptionSet combobox ──
function OptionSetInput({
  param,
  value,
  onChange,
}: {
  param: ActionParam;
  value: number | undefined;
  onChange: (v: number) => void;
}) {
  const options = param.optionSet ?? [];
  const cur = options.find((o) => o.value === value);
  return (
    <Combobox
      size="small"
      value={cur?.label ?? (value != null ? String(value) : '')}
      selectedOptions={value != null ? [String(value)] : []}
      onOptionSelect={(_, d) => d.optionValue && onChange(Number(d.optionValue))}
      placeholder="Pick a value…"
    >
      {options.map((o) => (
        <Option key={o.value} value={String(o.value)} text={o.label}>
          {o.label}
          <span
            style={{
              color: tokens.colorNeutralForeground3,
              fontFamily: tokens.fontFamilyMonospace,
              fontSize: 10,
              marginLeft: 6,
            }}
          >
            · {o.value}
          </span>
        </Option>
      ))}
    </Combobox>
  );
}

// ── String/Number array editor via TagPicker pattern ──
function StringArrayInput({
  value,
  onChange,
  mono,
}: {
  value: string[];
  onChange: (v: string[]) => void;
  mono?: boolean;
}) {
  return (
    <span style={{ display: 'flex', gap: 6, alignItems: 'center', flexWrap: 'wrap' }}>
      {value.map((v, i) => (
        <Tag
          key={i}
          shape="rounded"
          dismissible
          onClick={() => onChange(value.filter((_, j) => j !== i))}
          style={{ fontFamily: mono ? tokens.fontFamilyMonospace : tokens.fontFamilyBase }}
        >
          {v}
        </Tag>
      ))}
      <Input
        size="small"
        placeholder="add… (Enter)"
        style={{
          width: 160,
          fontFamily: mono ? tokens.fontFamilyMonospace : tokens.fontFamilyBase,
        }}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            const v = (e.target as HTMLInputElement).value.trim();
            if (v) {
              onChange([...value, v]);
              (e.target as HTMLInputElement).value = '';
            }
            e.preventDefault();
          }
        }}
      />
    </span>
  );
}

function NumberArrayInput({
  value,
  onChange,
}: {
  value: number[];
  onChange: (v: number[]) => void;
}) {
  return (
    <span style={{ display: 'flex', gap: 6, alignItems: 'center', flexWrap: 'wrap' }}>
      {value.map((v, i) => (
        <Tag
          key={i}
          shape="rounded"
          dismissible
          onClick={() => onChange(value.filter((_, j) => j !== i))}
        >
          {String(v)}
        </Tag>
      ))}
      <Input
        size="small"
        type="number"
        placeholder="add… (Enter)"
        style={{ width: 120 }}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            const v = Number((e.target as HTMLInputElement).value);
            if (Number.isFinite(v)) {
              onChange([...value, v]);
              (e.target as HTMLInputElement).value = '';
            }
            e.preventDefault();
          }
        }}
      />
    </span>
  );
}

// ── Entity reference picker (polymorphic = pick entity type first) ──
//
// Live-data only: the entity-type combobox lists scoped entities from
// `useScopedEntities()`; the record picker uses the same `RecordPicker`
// primitive Merge / Delete / Update use — typeahead against live Dataverse,
// no fixture catalog.
function EntityReferenceInput({
  value,
  onChange,
  param,
  polymorphic,
}: {
  value: { id: string; entityType: string } | undefined;
  onChange: (v: { id: string; entityType: string } | null) => void;
  param: ActionParam;
  polymorphic: boolean;
}) {
  const { entities } = useScopedEntities();
  const entityType = value?.entityType ?? param.entityType ?? entities[0]?.logicalName ?? '';
  // Warm the chosen entity's metadata so RecordPicker can display the
  // primaryName column when results arrive.
  warmTableMetadata(entityType || null);
  const tbl = findTable(entityType);

  return (
    <span style={{ display: 'flex', flexDirection: 'column', gap: 6, minWidth: 0 }}>
      {polymorphic && (
        <Combobox
          size="small"
          value={tbl?.displayName ?? entityType}
          selectedOptions={[entityType]}
          onOptionSelect={(_, d) => {
            if (!d.optionValue) return;
            // Switching the target entity invalidates the previously-picked
            // record id (different GUID space), so we reset the id.
            onChange({ id: '', entityType: d.optionValue });
          }}
          placeholder="Pick target entity…"
        >
          {entities.map((t) => (
            <Option key={t.logicalName} value={t.logicalName} text={t.displayName}>
              <div style={{ display: 'flex', flexDirection: 'column' }}>
                <span>{t.displayName}</span>
                <Caption1
                  style={{
                    color: tokens.colorNeutralForeground3,
                    fontFamily: tokens.fontFamilyMonospace,
                  }}
                >
                  {t.logicalName}
                </Caption1>
              </div>
            </Option>
          ))}
        </Combobox>
      )}
      {!polymorphic && (
        <Badge
          appearance="ghost"
          style={{ fontFamily: tokens.fontFamilyMonospace, alignSelf: 'flex-start' }}
        >
          → {tbl?.entitySetName ?? entityType}
        </Badge>
      )}
      {entityType && (
        <RecordPicker
          table={entityType}
          selectedId={value?.id || null}
          onPick={(r) => onChange(r ? { id: r.id, entityType } : null)}
          placeholder={`Search ${tbl?.displayName ?? entityType} records…`}
        />
      )}
      <Caption1 style={{ color: tokens.colorNeutralForeground3, fontSize: 10 }}>
        Emitted as{' '}
        <code style={{ fontFamily: tokens.fontFamilyMonospace }}>
          @odata.type=Microsoft.Dynamics.CRM.{entityType}
        </code>
      </Caption1>
    </span>
  );
}

// ── Entity collection picker (multi-target) ──
function EntityCollectionInput({
  value,
  onChange,
  param,
}: {
  value: Array<{ id: string; entityType: string }>;
  onChange: (v: Array<{ id: string; entityType: string }>) => void;
  param: ActionParam;
}) {
  const entityType = param.entityType ?? '';
  // Warm the chosen entity's metadata so RecordPicker can show primary names.
  warmTableMetadata(entityType || null);
  const tbl = findTable(entityType);

  // Live-data multi-record picker — mirrors AssociateTargetsEditor's pattern.
  // Each pick appends to the value array; the inline RecordPicker remounts
  // via a key bump so the typeahead clears for the next pick.
  const [pickerKey, setPickerKey] = useState(0);
  const onAdd = (id: string | null) => {
    if (!id) return;
    if (value.some((v) => v.id === id)) return; // dedupe
    onChange([...value, { id, entityType }]);
    setPickerKey((k) => k + 1);
  };
  const onRemove = (id: string) => onChange(value.filter((v) => v.id !== id));

  if (!entityType) {
    return (
      <MessageBar layout="multiline" intent="warning">
        <MessageBarBody>
          This collection parameter has no entity type declared in CSDL.
        </MessageBarBody>
      </MessageBar>
    );
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 6, minWidth: 0 }}>
      <Badge
        appearance="ghost"
        style={{ fontFamily: tokens.fontFamilyMonospace, alignSelf: 'flex-start' }}
      >
        → Collection({entityType})
      </Badge>
      {value.length > 0 && (
        <div
          style={{
            display: 'flex',
            flexDirection: 'column',
            gap: 4,
            padding: 6,
            border: `1px solid ${tokens.colorNeutralStroke2}`,
            borderRadius: tokens.borderRadiusMedium,
          }}
        >
          {value.map((v) => (
            <div key={v.id} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <Persona size="extra-small" name={v.id} avatar={{ color: 'colorful' }} />
              <code
                style={{
                  flex: 1,
                  minWidth: 0,
                  fontFamily: tokens.fontFamilyMonospace,
                  fontSize: 11,
                  color: tokens.colorNeutralForeground3,
                }}
              >
                /{tbl?.entitySetName ?? entityType}({v.id})
              </code>
              <Button size="small" appearance="subtle" onClick={() => onRemove(v.id)}>
                Remove
              </Button>
            </div>
          ))}
        </div>
      )}
      <RecordPicker
        key={pickerKey}
        table={entityType}
        selectedId={null}
        onPick={(r) => onAdd(r?.id ?? null)}
        placeholder={`Add ${tbl?.displayName ?? entityType} record…`}
      />
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// EntitySpecific — a NEW entity instance constructed inline.
//
// Per docs (e.g. WinOpportunity.OpportunityClose), the user supplies a
// metadata-driven body for the target entity — not a reference to an
// existing row. We render a slim nested FieldSetEditor scoped to the
// target entity's writable columns. The wire encoder prepends the
// @odata.type discriminator at body-build time.
//
// Reference:
//   https://learn.microsoft.com/en-us/dotnet/api/microsoft.crm.sdk.messages.winopportunityrequest
// ──────────────────────────────────────────────────────────────
function EntitySpecificInput({
  param,
  value,
  onChange,
}: {
  param: ActionParam;
  value: Record<string, unknown> | undefined;
  onChange: (v: Record<string, CreateFieldValue>) => void;
}) {
  const entityType = param.entityType ?? '';
  // Warm the entity's live AttributeMetadata. The fetch runs in the
  // background via the live-metadata registry; once it lands findTable()
  // returns a populated TableMeta and this row re-renders with the typed
  // FieldSetEditor. While loading we show a spinner + reassurance text.
  const { loading: warmingTable } = warmTableMetadata(entityType || null);
  const tbl = findTable(entityType);
  const v = (value as Record<string, CreateFieldValue>) ?? {};

  if (!tbl) {
    // Two reasons we'd hit this:
    //   1. The entity's metadata hasn't loaded YET (warmingTable=true) →
    //      show a non-alarming spinner, the form will appear shortly.
    //   2. The entity genuinely isn't in this environment (warmingTable
    //      false AND tbl undefined) → fall back to the JSON input.
    if (warmingTable) {
      return (
        <div
          style={{
            padding: 10,
            background: tokens.colorNeutralBackground2,
            border: `1px dashed ${tokens.colorBrandStroke2}`,
            borderRadius: tokens.borderRadiusMedium,
            display: 'flex',
            alignItems: 'center',
            gap: 8,
          }}
        >
          <Spinner size="extra-tiny" />
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            Loading <code>{entityType}</code> metadata from Dataverse — typed form will appear
            shortly…
          </Caption1>
        </div>
      );
    }
    return (
      <MessageBar layout="multiline" intent="warning">
        <MessageBarBody>
          No metadata for entity type <code>{entityType}</code> — the entity isn't available in this
          environment. For now, edit the JSON inline:
        </MessageBarBody>
        <Input
          size="small"
          value={JSON.stringify(v)}
          onChange={(_, d) => {
            try {
              const parsed = JSON.parse(d.value || '{}');
              if (parsed && typeof parsed === 'object') onChange(parsed);
            } catch {
              /* ignore — user is mid-typing */
            }
          }}
          style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11, marginTop: 6 }}
        />
      </MessageBar>
    );
  }

  return (
    <div
      style={{
        padding: 10,
        background: tokens.colorNeutralBackground2,
        border: `1px dashed ${tokens.colorBrandStroke2}`,
        borderRadius: tokens.borderRadiusMedium,
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 10 }}>
        <FormNew20Regular style={{ width: 14, height: 14, color: tokens.colorBrandForeground1 }} />
        <strong style={{ fontSize: 11 }}>{tbl.displayName}</strong>
        <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
          new entity instance · @odata.type discriminator added automatically
        </Caption1>
      </div>
      <FieldSetEditor
        table={entityType}
        values={v}
        setValues={(next) => onChange(next)}
        group="execute"
      />
    </div>
  );
}

// ── Complex type — recursive nested form ──
function ComplexTypeInput({
  param,
  value,
  onChange,
}: {
  param: ActionParam;
  value: Record<string, unknown> | undefined;
  onChange: (v: Record<string, unknown>) => void;
}) {
  const fields = param.complexType?.fields ?? [];
  const v = value ?? {};
  if (fields.length === 0) {
    // Fallback when the CSDL didn't ship a `fields` definition for this
    // complex type — let the user edit raw JSON. Parse on each keystroke;
    // invalid JSON silently keeps the prior value until the user fixes it.
    return (
      <div
        style={{
          padding: 10,
          background: tokens.colorNeutralBackground2,
          border: `1px dashed ${tokens.colorNeutralStroke2}`,
          borderRadius: tokens.borderRadiusMedium,
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 6 }}>
          <Document20Regular
            style={{ width: 14, height: 14, color: tokens.colorNeutralForeground3 }}
          />
          <strong style={{ fontSize: 11 }}>{param.complexType?.name ?? param.name}</strong>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            complex type · no field schema in mock — edit JSON
          </Caption1>
        </div>
        <Input
          size="small"
          value={JSON.stringify(v)}
          onChange={(_, d) => {
            try {
              const parsed = JSON.parse(d.value || '{}');
              if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) onChange(parsed);
            } catch {
              /* mid-typing — drop until valid */
            }
          }}
          style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11, width: '100%' }}
        />
      </div>
    );
  }
  return (
    <div
      style={{
        padding: 10,
        background: tokens.colorNeutralBackground2,
        border: `1px dashed ${tokens.colorNeutralStroke2}`,
        borderRadius: tokens.borderRadiusMedium,
        display: 'flex',
        flexDirection: 'column',
        gap: 10,
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
        <Document20Regular
          style={{ width: 14, height: 14, color: tokens.colorNeutralForeground3 }}
        />
        <strong style={{ fontSize: 11 }}>{param.complexType?.name ?? param.name}</strong>
        <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>nested complex type</Caption1>
      </div>
      {fields.map((f) => (
        <Field
          key={f.name}
          label={
            <span style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>
              {f.name} <span style={{ color: tokens.colorNeutralForeground3 }}>· {f.type}</span>
              {f.required && (
                <Badge
                  appearance="tint"
                  color="danger"
                  size="extra-small"
                  style={{ marginLeft: 6 }}
                >
                  req
                </Badge>
              )}
            </span>
          }
          hint={f.description}
        >
          <ParamInput
            param={f}
            value={v[f.name]}
            onChange={(nv) => onChange({ ...v, [f.name]: nv })}
          />
        </Field>
      ))}
    </div>
  );
}

function isEmpty(v: unknown): boolean {
  if (v == null) return true;
  if (typeof v === 'string' && v === '') return true;
  if (Array.isArray(v) && v.length === 0) return true;
  if (typeof v === 'object' && v !== null && 'id' in (v as Record<string, unknown>)) {
    return !(v as { id?: string }).id;
  }
  return false;
}
