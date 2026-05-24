// Per-AttributeTypeCode value input — dispatches to the appropriate Fluent
// v9 control based on the column's Dataverse attribute type.
//
// Renders the right Fluent v9 control based on:
//   - the column's AttributeTypeCode + per-type metadata (Format, MaxLength,
//     Precision, Targets, DateTimeBehavior, OptionSet)
//   - the operator's arity (0 / 1 / 2 / 'n')
//   - the rule's `valKind` (literal | column) — column mode swaps the input
//     for a same-type column picker (column-vs-column comparison).

import {
  Input, Switch, SpinButton, Combobox, Option, Tooltip, Button, tokens, Textarea,
  Persona, MessageBar, MessageBarBody, ToggleButton,
  InteractionTag, InteractionTagPrimary, InteractionTagSecondary, TagGroup,
  TagPicker, TagPickerControl, TagPickerGroup, TagPickerInput, TagPickerList,
  TagPickerOption, useTagPickerFilter, Tag,
} from '@fluentui/react-components';
import { useEffect, useMemo, useState } from 'react';
import { DatePicker } from '@fluentui/react-datepicker-compat';
import { TimePicker } from '@fluentui/react-timepicker-compat';
import { Add20Regular, Delete20Regular, Database20Regular } from '@fluentui/react-icons';
import type { OperatorDef } from './operators';
import {
  columnOptions, findTable, type ColumnMeta, type TableMeta, type DateTimeColumnMeta,
  type StringColumnMeta, type MemoColumnMeta, type IntegerColumnMeta,
} from '../../mock/metadata';
import { useLookupRecords } from '../../host/useLookupRecords';
import { useColumnDetail } from '../../host/useColumnDetail';
import type { JSX } from 'react';

// ============================================================
// Shared value-source shape (same as before, plus valKind)
// ============================================================
export interface ValueSource {
  id: string;
  col: string;
  op: string;
  valKind?: 'literal' | 'column';
  val?: string;
  vals?: [string, string];
  values?: string[];
}
export type ValueSourcePatch = Partial<Pick<ValueSource, 'val' | 'vals' | 'values' | 'valKind'>>;

export interface ValueInputProps {
  rule: ValueSource;
  op: OperatorDef;
  col: ColumnMeta | undefined;
  /** Used for column-mode RHS — picks from the same entity's other columns. */
  parentTable?: TableMeta;
  onChange: (patch: ValueSourcePatch) => void;
}

// ============================================================
// Entry point — dispatches by arity, then by AttributeTypeCode
// ============================================================
export function FilterValueInput(props: ValueInputProps): JSX.Element {
  const { rule, op, col, parentTable } = props;

  // Lazy enrichment — fires the right $expand=OptionSet or detailed-metadata
  // fetch for THIS column the first time the editor renders. No-op when
  // the column is already enriched or doesn't need extras.
  useColumnDetail(parentTable?.logicalName ?? null, col?.logicalName ?? null);

  if (op.arity === 0) {
    return (
      <span style={{ fontSize: 11, color: tokens.colorNeutralForeground3, fontStyle: 'italic' }}>
        — no value
      </span>
    );
  }
  if (op.arity === 2) return <TwoValueInput {...props} />;
  if (op.arity === 'n') return <ArrayValueInput {...props} />;

  // arity 1 — column-mode is only allowed when op.supportsColumnRhs
  if (rule.valKind === 'column' && op.supportsColumnRhs) {
    return <ColumnRhsPicker {...props} />;
  }
  return <SingleLiteralInput {...props} />;
}

// ============================================================
// Single literal input — operator-typed first, then column-typed
// ============================================================
//
// Dataverse query functions impose their OWN value type regardless of the
// column's attribute type. For example `LastXDays(PropertyName='createdon',
// PropertyValue=Edm.Int64)` — the column is DateTime but the value must be
// an integer (number of days). Same shape for every `LastX*` / `NextX*` /
// `OlderThanX*` / `InFiscalPeriod` / `InFiscalYear` and the fiscal-X-units
// functions. The operator catalogue marks all of these `intValue: true`.
//
// `Above` / `Under` / `AboveOrEqual` / `UnderOrEqual` / `NotUnder` take a
// GUID. Those are gated to Lookup-family columns by `allowedTypes`, so the
// Lookup typeahead below already produces a GUID-typed value.
//
// We therefore handle operator-typed cases FIRST and only fall back to the
// column-type-based dispatch for plain comparisons and OData string functions.
function SingleLiteralInput({ rule, op, col, onChange }: ValueInputProps): JSX.Element {
  // ── Operator-typed overrides ──────────────────────────────
  // Integer-valued Dataverse query functions (LastXDays, LastXFiscalYears,
  // InFiscalPeriod, …). The column may be DateTime/Integer/whatever — the
  // value is always Edm.Int64.
  if (op.intValue) {
    return (
      <NumericInputCtl
        val={rule.val ?? ''}
        onChange={v => onChange({ val: v })}
        integer
        step={1}
        min={0}
      />
    );
  }
  if (!col) {
    return <Input size="small" value={rule.val ?? ''} onChange={(_, d) => onChange({ val: d.value })} placeholder="value" />;
  }
  switch (col.attributeType) {
    case 'Boolean':
      return <BooleanInput col={col} val={rule.val} onChange={v => onChange({ val: v })} />;
    case 'Picklist':
    case 'State':
    case 'Status':
    case 'EntityName':
      return <ChoiceComboInput col={col} val={rule.val} onChange={v => onChange({ val: v })} />;
    case 'DateTime':
      return <DateTimeInputCtl col={col} val={rule.val ?? ''} onChange={v => onChange({ val: v })} />;
    case 'String':
      return <StringInputCtl col={col} val={rule.val ?? ''} onChange={v => onChange({ val: v })} />;
    case 'Memo':
      return <MemoInputCtl col={col} val={rule.val ?? ''} onChange={v => onChange({ val: v })} />;
    case 'Integer':
      return <IntegerInputCtl col={col} op={op} val={rule.val ?? ''} onChange={v => onChange({ val: v })} />;
    case 'BigInt':
      return <NumericInputCtl
        val={rule.val ?? ''} onChange={v => onChange({ val: v })}
        min={col.minValue} max={col.maxValue} integer step={1} />;
    case 'Decimal':
      return <NumericInputCtl
        val={rule.val ?? ''} onChange={v => onChange({ val: v })}
        min={col.minValue} max={col.maxValue} step={Math.pow(10, -col.precision)} />;
    case 'Double':
      return <NumericInputCtl
        val={rule.val ?? ''} onChange={v => onChange({ val: v })}
        min={col.minValue} max={col.maxValue} step={Math.pow(10, -col.precision)} />;
    case 'Money':
      return <MoneyInputCtl precision={col.precision} val={rule.val ?? ''} onChange={v => onChange({ val: v })}
        min={col.minValue} max={col.maxValue} />;
    case 'Lookup':
    case 'Customer':
    case 'Owner':
      return <LookupTypeaheadCtl col={col} val={rule.val ?? ''} onChange={v => onChange({ val: v })} />;
    case 'Uniqueidentifier':
      return <GuidInputCtl val={rule.val ?? ''} onChange={v => onChange({ val: v })} />;
    case 'MultiSelectPicklist':
      // Single-value compare is rarely meaningful — show a Combobox of the options
      return <ChoiceComboInput col={col} val={rule.val} onChange={v => onChange({ val: v })} />;
    default:
      return <Input size="small" value={rule.val ?? ''} onChange={(_, d) => onChange({ val: d.value })} placeholder="value" />;
  }
}

// ============================================================
// Per-type controls
// ============================================================
function StringInputCtl({ col, val, onChange }: { col: StringColumnMeta; val: string; onChange: (v: string) => void }) {
  const isLong = col.format === 'TextArea';
  if (isLong) {
    return (
      <Textarea
        rows={2}
        value={val}
        maxLength={col.maxLength}
        onChange={(_, d) => onChange(d.value)}
        style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 12 }}
      />
    );
  }
  const inputType =
    col.format === 'Email' ? 'email' :
    col.format === 'Url'   ? 'url' :
    col.format === 'Phone' ? 'tel' : 'text';
  return (
    <Input
      size="small"
      type={inputType}
      value={val}
      maxLength={col.maxLength}
      placeholder={
        col.format === 'Email' ? 'name@example.com' :
        col.format === 'Url'   ? 'https://example.com' :
        col.format === 'Phone' ? '+1 (555) 555-0000' :
        col.format === 'TickerSymbol' ? 'MSFT' :
        'value'
      }
      onChange={(_, d) => onChange(d.value)}
    />
  );
}

function MemoInputCtl({ col, val, onChange }: { col: MemoColumnMeta; val: string; onChange: (v: string) => void }) {
  const [focused, setFocused] = useState(false);
  // Compact-by-default: matches the height of a regular `<Input>` when
  // blurred so the filter row's visual baseline stays uniform across all
  // value-input types. Click into the box → expands to 3 rows; click out
  // → collapses back to single-line. Width fills the value cell.
  //
  // Sizing rationale: Fluent v9 `<Input>` renders at ~32px tall by
  // default — `Textarea` with `rows={1}` is slightly taller due to inner
  // padding, so we lock `height: 32px` when collapsed via inline style.
  // The placeholder is shortened to "Up to N chars…" so it fits on one
  // visible line at the narrower height.
  return (
    <Textarea
      rows={focused ? 3 : 1}
      value={val}
      maxLength={col.maxLength}
      placeholder={focused
        ? `Up to ${col.maxLength.toLocaleString()} characters`
        : 'Click to expand…'}
      onChange={(_, d) => onChange(d.value)}
      onFocus={() => setFocused(true)}
      onBlur={() => setFocused(false)}
      style={{
        fontFamily: tokens.fontFamilyMonospace,
        fontSize: 12,
        width: '100%',
        height: focused ? 'auto' : '32px',
        minHeight: focused ? undefined : '32px',
        resize: focused ? 'vertical' : 'none',
        overflow: focused ? 'auto' : 'hidden',
      }}
    />
  );
}

function DateTimeInputCtl({ col, val, onChange }: { col: DateTimeColumnMeta; val: string; onChange: (v: string) => void }) {
  // DateOnly behavior or format → DatePicker alone; DateAndTime → DatePicker + TimePicker
  const isDateOnly = col.format === 'DateOnly' || col.dateTimeBehavior === 'DateOnly';
  const behaviorHint =
    col.dateTimeBehavior === 'UserLocal'  ? 'stored as UTC, displayed in user TZ' :
    col.dateTimeBehavior === 'DateOnly'   ? 'midnight UTC, no TZ conversion' :
                                            'wall-clock as UTC (no TZ math)';

  // Parse the stored ISO string back to a Date for the pickers
  const parsed: Date | null = val ? new Date(val) : null;
  const dateOnly: Date | null = parsed && !Number.isNaN(parsed.getTime()) ? parsed : null;

  const writeDate = (d: Date | null | undefined) => {
    if (!d) { onChange(''); return; }
    if (isDateOnly) {
      // ISO yyyy-mm-dd (DateOnly columns)
      const yyyy = d.getFullYear();
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      const dd = String(d.getDate()).padStart(2, '0');
      onChange(`${yyyy}-${mm}-${dd}`);
    } else {
      // Preserve existing time portion if any, default 00:00
      const prior = parsed && !Number.isNaN(parsed.getTime()) ? parsed : new Date(d);
      const next = new Date(d);
      next.setHours(prior.getHours(), prior.getMinutes(), 0, 0);
      onChange(next.toISOString());
    }
  };
  const writeTime = (d: Date | null | undefined) => {
    if (!d || !dateOnly) { return; }
    const next = new Date(dateOnly);
    next.setHours(d.getHours(), d.getMinutes(), 0, 0);
    onChange(next.toISOString());
  };

  if (isDateOnly) {
    return (
      <Tooltip content={`Behavior: ${col.dateTimeBehavior} — ${behaviorHint}`} relationship="description">
        <DatePicker
          size="small"
          value={dateOnly}
          onSelectDate={writeDate}
          placeholder="Pick a date…"
          formatDate={(d) => d ? d.toLocaleDateString() : ''}
          allowTextInput
        />
      </Tooltip>
    );
  }
  return (
    <Tooltip content={`Behavior: ${col.dateTimeBehavior} — ${behaviorHint}`} relationship="description">
      <span style={{ display: 'inline-flex', gap: 6, alignItems: 'center', flexWrap: 'wrap' }}>
        <DatePicker
          size="small"
          value={dateOnly}
          onSelectDate={writeDate}
          placeholder="Date…"
          formatDate={(d) => d ? d.toLocaleDateString() : ''}
          style={{ minWidth: 130 }}
          allowTextInput
        />
        <TimePicker
          size="small"
          freeform
          selectedTime={dateOnly}
          onTimeChange={(_, d) => writeTime(d.selectedTime)}
          placeholder="Time…"
          style={{ minWidth: 100 }}
          hourCycle="h23"
        />
      </span>
    </Tooltip>
  );
}

function IntegerInputCtl({ col, val, onChange }: { col: IntegerColumnMeta; op: OperatorDef; val: string; onChange: (v: string) => void }) {
  // Format-driven: Duration → minute presets; Language/Locale/TimeZone → combobox stubs
  if (col.format === 'Duration') {
    const presets = [1, 15, 30, 60, 240, 480, 1440, 10080]; // 1m..1w in minutes
    return (
      <span style={{ display: 'inline-flex', gap: 4, alignItems: 'center', flexWrap: 'wrap' }}>
        <SpinButton
          size="small"
          value={Number(val || 0)}
          min={col.minValue ?? 0}
          max={col.maxValue ?? 525_600}
          step={1}
          onChange={(_, d) => {
            const v = d.value ?? Number(d.displayValue ?? 0);
            onChange(String(v));
          }}
        />
        <span style={{ fontSize: 10, color: tokens.colorNeutralForeground3 }}>min</span>
        {presets.slice(0, 5).map(p => {
          const label = p < 60 ? `${p}m` : p < 1440 ? `${p / 60}h` : `${p / 1440}d`;
          const isActive = String(p) === val;
          return (
            <ToggleButton
              key={p}
              size="small"
              shape="circular"
              checked={isActive}
              onClick={() => onChange(String(p))}
            >
              {label}
            </ToggleButton>
          );
        })}
      </span>
    );
  }
  if (col.format === 'Language' || col.format === 'Locale') {
    return (
      <Combobox
        size="small"
        value={val}
        selectedOptions={val ? [val] : []}
        onOptionSelect={(_, d) => d.optionValue && onChange(d.optionValue)}
        placeholder="LCID"
      >
        <Option value="1033" text="1033">English (US) · 1033</Option>
        <Option value="1031" text="1031">German · 1031</Option>
        <Option value="1036" text="1036">French · 1036</Option>
        <Option value="1041" text="1041">Japanese · 1041</Option>
        <Option value="1043" text="1043">Dutch · 1043</Option>
      </Combobox>
    );
  }
  if (col.format === 'TimeZone') {
    return (
      <Combobox
        size="small"
        value={val}
        selectedOptions={val ? [val] : []}
        onOptionSelect={(_, d) => d.optionValue && onChange(d.optionValue)}
        placeholder="time zone code"
      >
        <Option value="4" text="4">Pacific (US&amp;Canada) · 4</Option>
        <Option value="10" text="10">Eastern (US&amp;Canada) · 10</Option>
        <Option value="85" text="85">UTC · 85</Option>
        <Option value="105" text="105">W. Europe · 105</Option>
      </Combobox>
    );
  }
  return (
    <NumericInputCtl
      val={val} onChange={onChange}
      min={col.minValue} max={col.maxValue} integer step={1}
    />
  );
}

function NumericInputCtl({ val, onChange, min, max, step, integer }: {
  val: string; onChange: (v: string) => void;
  min?: number; max?: number; step?: number; integer?: boolean;
}) {
  return (
    <SpinButton
      size="small"
      value={Number(val || 0)}
      min={min}
      max={max}
      step={step ?? (integer ? 1 : 0.01)}
      onChange={(_, d) => {
        const v = d.value ?? Number(d.displayValue ?? 0);
        onChange(String(v));
      }}
    />
  );
}

function MoneyInputCtl({ precision, val, onChange, min, max }: {
  precision: number; val: string; onChange: (v: string) => void;
  min?: number; max?: number;
}) {
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
      <span style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 12, color: tokens.colorNeutralForeground3 }}>$</span>
      <NumericInputCtl
        val={val} onChange={onChange}
        min={min} max={max}
        step={Math.pow(10, -precision)}
      />
    </span>
  );
}

function BooleanInput({ col, val, onChange }: { col: import('../../mock/metadata').BooleanColumnMeta; val?: string; onChange: (v: string) => void }) {
  // val is '0' / '1' / 'true' / 'false'
  const isTrue = val === '1' || val === 'true';
  return (
    <Switch
      checked={isTrue}
      onChange={(_, d) => onChange(d.checked ? '1' : '0')}
      label={isTrue ? col.trueOption.label : col.falseOption.label}
    />
  );
}

function ChoiceComboInput({ col, val, onChange }: { col: ColumnMeta; val?: string; onChange: (v: string) => void }) {
  const options = columnOptions(col) ?? [];
  const cur = options.find(o => String(o.value) === val);
  return (
    <Combobox
      size="small"
      value={cur?.label ?? val ?? ''}
      selectedOptions={val ? [val] : []}
      onOptionSelect={(_, d) => d.optionValue && onChange(d.optionValue)}
    >
      {options.map(o => (
        <Option key={o.value} value={String(o.value)} text={o.label}>
          {o.label}
          <span style={{ color: tokens.colorNeutralForeground3, fontFamily: tokens.fontFamilyMonospace, fontSize: 10, marginLeft: 6 }}>
            · {o.value}
          </span>
        </Option>
      ))}
    </Combobox>
  );
}

function LookupTypeaheadCtl({ col, val, onChange }: {
  col: import('../../mock/metadata').LookupColumnMeta | import('../../mock/metadata').CustomerColumnMeta | import('../../mock/metadata').OwnerColumnMeta;
  val: string; onChange: (v: string) => void;
}) {
  // Polymorphic-aware typeahead.
  //
  // Single-target lookups (`primarycontactid` → contact) just need a record
  // typeahead. Polymorphic lookups (Customer = account|contact, Owner =
  // systemuser|team, multi-target Lookup like `regardingobjectid`) need TWO
  // steps: pick the target table FIRST, then search that table's records.
  //
  // Wire-side note: $filter on a polymorphic lookup operates on
  // `_<attr>_value` which is just the GUID — Dataverse doesn't care which
  // target the GUID resolves to for an `eq` comparison (the GUID space is
  // unique across the org). The target switcher is purely a UX scoping
  // device so the user finds the right record in the typeahead. The wire
  // value stays a bare GUID.
  const polymorphic = col.targets.length > 1;
  const [target, setTarget] = useState<string>(col.targets[0] ?? '');
  // Sync target if the column's targets array becomes populated later
  // (useColumnDetail enrichment lands after first render for Lookup cols).
  useEffect(() => {
    if (!target && col.targets.length > 0) setTarget(col.targets[0]);
  }, [col.targets, target]);

  const [query, setQuery] = useState('');
  const { rows, loading, error } = useLookupRecords(target || null, query);
  const selected = rows.find(r => r.id === val);
  const tbl = target ? findTable(target) : undefined;
  const targetLabel = tbl?.displayName ?? target ?? 'record';

  if (!target) {
    // Targets not loaded yet OR column truly has no targets — fall back
    // to a free-form GUID input so the user can paste one manually.
    return (
      <Input
        size="small"
        value={val}
        placeholder="enter target GUID…"
        onChange={(_, d) => onChange(d.value)}
        style={{ fontFamily: tokens.fontFamilyMonospace }}
      />
    );
  }

  // Layout: single-target stays inline (compact "→ contact" chip + one
  // combobox fits the rule row cleanly). Polymorphic stacks vertically —
  // an inline target combobox + record combobox is too wide for the rule
  // row and overflows the FilterEditor panel (reported visually as the
  // record combobox extending past the right edge of the container).
  // Stacking keeps both controls within the value column's width.
  return (
    <span
      style={{
        display: 'flex',
        flexDirection: polymorphic ? 'column' : 'row',
        alignItems: polymorphic ? 'stretch' : 'center',
        gap: polymorphic ? 4 : 6,
        minWidth: 0,
        // Polymorphic uses the full value-column width since both controls
        // share it; single-target uses its natural intrinsic width.
        width: polymorphic ? '100%' : undefined,
      }}
    >
      {polymorphic ? (
        // Polymorphic — target picker first (its own row), scoped record
        // search second. Switching target doesn't auto-clear val because
        // the user may have pasted a GUID they want to keep.
        <Combobox
          size="small"
          value={tbl?.displayName ?? target}
          selectedOptions={[target]}
          onOptionSelect={(_, d) => {
            if (!d.optionValue) return;
            setTarget(d.optionValue);
            setQuery('');
          }}
          // No fixed width — fills the value column.
        >
          {col.targets.map(t => {
            const tt = findTable(t);
            return (
              <Option key={t} value={t} text={tt?.displayName ?? t}>
                {tt?.displayName ?? t}
              </Option>
            );
          })}
        </Combobox>
      ) : (
        // Single-target — just show the target as a non-interactive chip
        // for context. The user can't change it.
        <span style={{ fontSize: 10, color: tokens.colorNeutralForeground3, fontFamily: tokens.fontFamilyMonospace }}>
          → {target}
        </span>
      )}
      <Combobox
        size="small"
        value={selected?.name ?? query}
        selectedOptions={val ? [val] : []}
        onOptionSelect={(_, d) => { if (d.optionValue) { onChange(d.optionValue); setQuery(''); } }}
        onChange={(e) => setQuery((e.target as HTMLInputElement).value)}
        placeholder={`Search ${targetLabel}…`}
      >
        {loading && (
          <Option value="__loading" text="" disabled>
            <span style={{ color: tokens.colorNeutralForeground3, fontSize: 11 }}>Searching…</span>
          </Option>
        )}
        {error && (
          <Option value="__err" text="" disabled>
            <span style={{ color: tokens.colorPaletteRedForeground1, fontSize: 11 }}>{error}</span>
          </Option>
        )}
        {!loading && !error && rows.length === 0 && (
          <Option value="__none" text="" disabled>
            <span style={{ color: tokens.colorNeutralForeground3, fontSize: 11 }}>No matches</span>
          </Option>
        )}
        {rows.map(r => (
          <Option key={r.id} value={r.id} text={r.name}>
            <Persona size="extra-small" name={r.name} secondaryText={r.id} avatar={{ color: 'colorful' }} />
          </Option>
        ))}
      </Combobox>
    </span>
  );
}

function GuidInputCtl({ val, onChange }: { val: string; onChange: (v: string) => void }) {
  const valid = !val || /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(val);
  return (
    <Input
      size="small"
      value={val}
      placeholder="00000000-0000-0000-0000-000000000000"
      onChange={(_, d) => onChange(d.value)}
      style={{ fontFamily: tokens.fontFamilyMonospace }}
      contentAfter={!valid ? <span style={{ color: tokens.colorPaletteRedForeground1, fontSize: 10 }}>invalid GUID</span> : undefined}
    />
  );
}

// ============================================================
// Two-value range input (Between / NotBetween / InFiscalPeriodAndYear)
// ============================================================
function TwoValueInput({ rule, op, col, onChange }: ValueInputProps): JSX.Element {
  const [a, b] = rule.vals ?? ['', ''];
  const placeholderA = op.id.startsWith('InFiscal') ? 'period' : 'low';
  const placeholderB = op.id.startsWith('InFiscal') ? 'year' : 'high';

  // ── Operator-typed override ──────────────────────────────
  // `InFiscalPeriodAndYear` / `InOrAfterFiscalPeriodAndYear` /
  // `InOrBeforeFiscalPeriodAndYear` all take TWO Edm.Int64 values
  // (period + year) even though the column is DateTime. Without this
  // check we'd fall into the DateTime branch below and render two
  // DatePickers — wrong shape.
  if (op.intValue) {
    return (
      <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
        <SpinButton
          size="small"
          value={Number(a || 0)}
          min={0}
          step={1}
          onChange={(_, d) => onChange({ vals: [String(d.value ?? d.displayValue ?? a), b] })}
        />
        <span style={{ fontSize: 11, color: tokens.colorNeutralForeground3 }}>and</span>
        <SpinButton
          size="small"
          value={Number(b || 0)}
          min={0}
          step={1}
          onChange={(_, d) => onChange({ vals: [a, String(d.value ?? d.displayValue ?? b)] })}
        />
      </div>
    );
  }

  // Date types — render two DatePickers (covers Between / NotBetween on
  // a DateTime column, where both PropertyValues are dates).
  if (col?.attributeType === 'DateTime') {
    const parse = (s: string) => s ? new Date(s) : null;
    const isDateOnly = col.format === 'DateOnly';
    const writeISO = (d: Date | null | undefined) => {
      if (!d) return '';
      if (isDateOnly) {
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        return `${yyyy}-${mm}-${dd}`;
      }
      return d.toISOString();
    };
    return (
      <div style={{ display: 'flex', gap: 6, alignItems: 'center', flexWrap: 'wrap' }}>
        <DatePicker
          size="small"
          value={parse(a)}
          onSelectDate={(d) => onChange({ vals: [writeISO(d), b] })}
          placeholder={placeholderA}
          allowTextInput
          formatDate={(d) => d ? d.toLocaleDateString() : ''}
        />
        <span style={{ fontSize: 11, color: tokens.colorNeutralForeground3 }}>and</span>
        <DatePicker
          size="small"
          value={parse(b)}
          onSelectDate={(d) => onChange({ vals: [a, writeISO(d)] })}
          placeholder={placeholderB}
          allowTextInput
          formatDate={(d) => d ? d.toLocaleDateString() : ''}
        />
      </div>
    );
  }

  // Numeric / fiscal — SpinButtons (preferred v9 numeric control)
  if (isNumericLike(col)) {
    return (
      <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
        <SpinButton
          size="small"
          value={Number(a || 0)}
          onChange={(_, d) => onChange({ vals: [String(d.value ?? d.displayValue ?? a), b] })}
        />
        <span style={{ fontSize: 11, color: tokens.colorNeutralForeground3 }}>and</span>
        <SpinButton
          size="small"
          value={Number(b || 0)}
          onChange={(_, d) => onChange({ vals: [a, String(d.value ?? d.displayValue ?? b)] })}
        />
      </div>
    );
  }

  // Text fallback (shouldn't normally hit — Between is restricted to numbers/dates)
  return (
    <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
      <Input size="small" value={a} placeholder={placeholderA}
        onChange={(_, d) => onChange({ vals: [d.value, b] })} />
      <span style={{ fontSize: 11, color: tokens.colorNeutralForeground3 }}>and</span>
      <Input size="small" value={b} placeholder={placeholderB}
        onChange={(_, d) => onChange({ vals: [a, d.value] })} />
    </div>
  );
}

// ============================================================
// Array value input (In / NotIn / Between via array form, ContainValues, …)
// ============================================================
function ArrayValueInput({ rule, op, col, onChange }: ValueInputProps): JSX.Element {
  const values = rule.values ?? [];
  // Multi-select choice fns — filterable TagPicker over the option set.
  if (op.kind === 'dv-fn-array' && col && columnOptions(col) && (op.id === 'ContainValues' || op.id === 'DoesNotContainValues')) {
    return <ChoiceTagPicker col={col} values={values} onChange={(v) => onChange({ values: v })} placeholder="filter & pick option(s)" />;
  }
  // In / NotIn on Picklist — same TagPicker treatment, single source of truth.
  if (col && columnOptions(col) && (op.id === 'In' || op.id === 'NotIn')) {
    return <ChoiceTagPicker col={col} values={values} onChange={(v) => onChange({ values: v })} placeholder="filter & pick value(s)" />;
  }
  // Generic tag input — free-form scalars (numbers, strings, GUIDs)
  return (
    <div style={{ display: 'flex', gap: 6, alignItems: 'center', flexWrap: 'wrap' }}>
      <TagGroup
        size="small"
        onDismiss={(_, d) => onChange({ values: values.filter(v => v !== d.value) })}
      >
        {values.map((v) => (
          <InteractionTag key={v} size="small" shape="circular" appearance="outline" value={v}>
            <InteractionTagPrimary hasSecondaryAction>
              <span style={{ fontFamily: tokens.fontFamilyMonospace }}>{v}</span>
            </InteractionTagPrimary>
            <InteractionTagSecondary aria-label="Remove" />
          </InteractionTag>
        ))}
      </TagGroup>
      <ArrayAddInput onAdd={v => onChange({ values: [...values, v] })} intValue={!!op.intValue} />
    </div>
  );
}

function ArrayAddInput({ onAdd, intValue }: { onAdd: (v: string) => void; intValue: boolean }) {
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
      <Input
        size="small"
        type={intValue ? 'number' : 'text'}
        placeholder="+ add"
        style={{ width: 100 }}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            const v = (e.target as HTMLInputElement).value.trim();
            if (v) {
              onAdd(v);
              (e.target as HTMLInputElement).value = '';
            }
            e.preventDefault();
          }
        }}
        onBlur={(e) => {
          const v = (e.target as HTMLInputElement).value.trim();
          if (v) {
            onAdd(v);
            (e.target as HTMLInputElement).value = '';
          }
        }}
      />
      <Tooltip content="Press Enter to add" relationship="description">
        <Add20Regular style={{ width: 14, height: 14, color: tokens.colorNeutralForeground3 }} />
      </Tooltip>
    </span>
  );
}

// ============================================================
// ChoiceTagPicker — filterable multi-select for picklist-backed
// filter operators (In / NotIn / ContainValues / DoesNotContainValues).
//
// Mirrors the FieldSetEditor MultiChoiceTagPicker — string-keyed under
// the hood, integer-array on the wire. Tag chips render the option's
// display label; typing in the input filters by label.
// ============================================================
function ChoiceTagPicker({
  col, values, onChange, placeholder,
}: {
  col: ColumnMeta;
  values: string[];
  onChange: (v: string[]) => void;
  placeholder?: string;
}) {
  const options = columnOptions(col) ?? [];
  const [query, setQuery] = useState('');

  const labelByKey = useMemo(() => {
    const m = new Map<string, string>();
    for (const o of options) m.set(String(o.value), o.label);
    return m;
  }, [options]);
  const allKeys = useMemo(() => options.map(o => String(o.value)), [options]);

  const children = useTagPickerFilter({
    query,
    options: allKeys,
    noOptionsElement: <TagPickerOption value="__no-matches">No matches</TagPickerOption>,
    renderOption: (key) => (
      <TagPickerOption key={key} value={key} text={labelByKey.get(key) ?? key}>
        <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
          {labelByKey.get(key) ?? key}
          <span style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 10, color: tokens.colorNeutralForeground3 }}>
            · {key}
          </span>
        </span>
      </TagPickerOption>
    ),
    filter: (key) => {
      if (values.includes(key)) return false;
      const label = labelByKey.get(key) ?? key;
      return label.toLowerCase().includes(query.toLowerCase());
    },
  });

  return (
    <TagPicker
      selectedOptions={values}
      onOptionSelect={(_, d) => {
        if (d.value === '__no-matches') return;
        onChange(d.selectedOptions);
        setQuery('');
      }}
      disableAutoFocus={query.length === 0}
    >
      <TagPickerControl style={{ minWidth: 220, maxWidth: 480 }}>
        <TagPickerGroup aria-label="Selected values">
          {values.map(k => (
            <Tag key={k} value={k} shape="rounded" size="small">
              {labelByKey.get(k) ?? k}
            </Tag>
          ))}
        </TagPickerGroup>
        <TagPickerInput
          aria-label="Pick option"
          placeholder={values.length === 0 ? (placeholder ?? 'filter & pick…') : ''}
          value={query}
          onChange={(e) => setQuery(e.target.value)}
        />
      </TagPickerControl>
      <TagPickerList>{children}</TagPickerList>
    </TagPicker>
  );
}

// ============================================================
// Column-mode RHS picker — same-type columns of the same entity
// ============================================================
function ColumnRhsPicker({ rule, col, parentTable, onChange }: ValueInputProps): JSX.Element {
  if (!parentTable || !col) {
    return <Input size="small" value={rule.val ?? ''} disabled placeholder="(column compare unavailable)" />;
  }
  const sameType = parentTable.columns.filter(
    c => c.attributeType === col.attributeType && c.logicalName !== col.logicalName,
  );
  const cur = sameType.find(c => c.logicalName === rule.val);

  if (sameType.length === 0) {
    return (
      <MessageBar layout="multiline" intent="warning" style={{ padding: '4px 8px' }}>
        <MessageBarBody>
          No other <code style={{ fontFamily: tokens.fontFamilyMonospace }}>{col.attributeType}</code> columns on this table to compare against. Switch back to <strong>literal</strong>.
        </MessageBarBody>
      </MessageBar>
    );
  }
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
      <Database20Regular style={{ width: 14, height: 14, color: tokens.colorBrandForeground1 }} />
      <span style={{ fontSize: 10, color: tokens.colorNeutralForeground3, fontFamily: tokens.fontFamilyMonospace }}>same-row</span>
      <Combobox
        size="small"
        value={cur?.displayName ?? rule.val ?? ''}
        selectedOptions={rule.val ? [rule.val] : []}
        onOptionSelect={(_, d) => d.optionValue && onChange({ val: d.optionValue })}
        placeholder="pick a column"
      >
        {sameType.map(c => (
          <Option key={c.logicalName} value={c.logicalName} text={c.displayName}>
            <span style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              {c.displayName}
              <span style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 10, color: tokens.colorNeutralForeground3 }}>
                {c.logicalName}
              </span>
            </span>
          </Option>
        ))}
      </Combobox>
    </span>
  );
}

// ============================================================
// Helpers
// ============================================================
const isNumericLike = (c?: ColumnMeta) =>
  c?.attributeType === 'Integer' || c?.attributeType === 'BigInt' ||
  c?.attributeType === 'Decimal' || c?.attributeType === 'Double' || c?.attributeType === 'Money';
