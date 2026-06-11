// FieldSetEditor — metadata-driven write-side form.
//
// "Field set" pane: every writable column on the target table is enumerable,
// but only the ones the user explicitly enables land in the POST body.
// Required columns surface a red "req" badge and a warning when left empty.
//
// Each row renders the appropriate Fluent v9 control for its AttributeTypeCode:
//
//   String          → Input  (type=email/url/tel from format)
//   Memo            → Textarea (chars-left counter)
//   Integer/BigInt  → SpinButton (Duration/Language/Locale/TimeZone variants)
//   Decimal/Double  → SpinButton (step driven by precision)
//   Money           → SpinButton with leading $ glyph
//   Boolean         → Switch (true/false labels from metadata)
//   DateTime        → DatePicker (+ TimePicker when DateAndTime)
//   Picklist/State/Status/EntityName → Combobox of options
//   MultiSelectPicklist → TagPicker with filterable typeahead (Tag-based chips)
//   Lookup/Customer/Owner → live RecordPicker typeahead over the target
//                            entity with the target surfaced as a chip
//   Uniqueidentifier → masked Input with GUID validation
//
// The pane carries metadata chrome — type chip, logical name, required/system
// indicators, precision/maxLength hints — to keep the user oriented while
// they're working with raw schema names. This is the same v9 design language
// used in SelectEditor and FilterValueInput.

import { useMemo, useState, useEffect } from 'react';
import {
  Input,
  Textarea,
  SpinButton,
  Switch,
  Combobox,
  Option,
  Button,
  Caption1,
  Tooltip,
  Badge,
  Persona,
  Field,
  ToggleButton,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  mergeClasses,
  tokens,
  TagPicker,
  TagPickerControl,
  TagPickerGroup,
  TagPickerInput,
  TagPickerList,
  TagPickerOption,
  useTagPickerFilter,
  Tag,
} from '@fluentui/react-components';
import { DatePicker } from '@fluentui/react-datepicker-compat';
import { TimePicker } from '@fluentui/react-timepicker-compat';
import {
  Add20Regular,
  Delete20Regular,
  Search20Regular,
  FormNew20Filled,
  Database20Regular,
  Key20Regular,
  Code20Regular,
  FormNew20Regular,
  Copy20Regular,
  TextNumberFormat20Regular,
  Link20Regular,
  ListBar20Regular,
} from '@fluentui/react-icons';
import { useStudioStyles } from '../primitives/styles';
import { SegmentedToggle } from '../primitives/SegmentedToggle';
import { PaneHead } from './PaneHead';
import {
  findTable,
  findColumn,
  isLookupLike,
  columnOptions,
  type ColumnMeta,
  type TableMeta,
  type StringColumnMeta,
  type MemoColumnMeta,
  type IntegerColumnMeta,
  type DateTimeColumnMeta,
  type MoneyColumnMeta,
  type DecimalColumnMeta,
  type DoubleColumnMeta,
  type BigIntColumnMeta,
  type BooleanColumnMeta,
  type LookupColumnMeta,
  type CustomerColumnMeta,
  type OwnerColumnMeta,
} from '../mock/metadata';
import { useLookupRecords } from '../host/useLookupRecords';
import { useColumnDetail } from '../host/useColumnDetail';
// Monaco — same instance + worker setup as CodeView. The `loader.config`
// side-effect runs when CodeView is first imported (every mode file
// imports it), so by the time FieldSetEditor renders, Monaco is ready.
import Editor from '@monaco-editor/react';
import type { ThemeMode } from '../theme/theme';
import type { CreateFieldValue, LookupFieldValue } from '../state/writeState';
import type { RequestGroup } from '../registry/requestTypes';

// ──────────────────────────────────────────────────────────────
// System-managed columns we hide from the Create form.
// (Primary keys, audit columns, computed totals, etc. Users *can* set
// some of these on create per docs, but it's the path of pain — keep
// them out of the default form and rely on the "Add field" picker
// for the rare power-user case.)
// ──────────────────────────────────────────────────────────────
const SYSTEM_MANAGED = new Set([
  'createdon',
  'modifiedon',
  'createdby',
  'modifiedby',
  'createdonbehalfby',
  'modifiedonbehalfby',
  'versionnumber',
  'importsequencenumber',
  'overriddencreatedon',
  'utcconversiontimezonecode',
  'timezoneruleversionnumber',
]);

const isSystemManaged = (c: ColumnMeta): boolean =>
  SYSTEM_MANAGED.has(c.logicalName) ||
  c.attributeType === 'File' ||
  c.attributeType === 'Image' ||
  (c.attributeType === 'Uniqueidentifier' && c.isPrimaryKey === true);

// ──────────────────────────────────────────────────────────────
// Props
// ──────────────────────────────────────────────────────────────
export interface FieldSetEditorProps {
  table: string;
  values: Record<string, CreateFieldValue>;
  setValues: (next: Record<string, CreateFieldValue>) => void;
  /**
   * Columns the user wants to send as EXPLICIT NULL in the body. Optional —
   * when omitted the "Set null" affordance is hidden and the editor behaves
   * as before. CreateMode/UpdateMode/UpsertMode pass these in to enable
   * the affordance and persist the flagged columns in state.
   */
  nullFields?: string[];
  setNullFields?: (next: string[]) => void;
  group?: RequestGroup;
  /** Theme mode — drives the Monaco JSON editor's vs-dark vs light. Caller-
   *  optional because not every consumer (e.g. nested ActionParamForm rows)
   *  has a theme handy. Defaults to 'light'. */
  themeMode?: ThemeMode;
  /**
   * Which write operation this editor is being used for. Drives the
   * column-eligibility filter:
   *
   *   - 'create' → only columns where `isValidForCreate !== false`
   *   - 'update' → only columns where `isValidForUpdate !== false`
   *   - 'upsert' → INTERSECTION of create + update (safest — the same
   *     body has to be acceptable on either path because Upsert decides
   *     at execute time whether the row exists)
   *   - undefined → permissive (every non-system-managed column), used by
   *     ActionParamForm's nested entity-param renderer where the user is
   *     authoring an arbitrary payload, not a CRUD body.
   *
   * Defaults to 'create' because that's the most restrictive surface and
   * Create mode is the primary caller. Update / Upsert pass the right
   * value explicitly.
   */
  purpose?: 'create' | 'update' | 'upsert';
}

type Filter = 'all' | 'required' | 'set' | 'common';

// ──────────────────────────────────────────────────────────────
// Public component
// ──────────────────────────────────────────────────────────────
export function FieldSetEditor({
  table,
  values,
  setValues,
  nullFields: nullFieldsProp,
  setNullFields: setNullFieldsProp,
  group = 'write',
  themeMode = 'light',
  purpose = 'create',
}: FieldSetEditorProps) {
  // Internal fallback when the caller doesn't wire up explicit-null tracking.
  // Keeps the affordance hidden in those modes (the Set-null button checks
  // for the controlled handler before rendering itself).
  const nullFields = nullFieldsProp ?? [];
  const nullFieldsSupported = !!setNullFieldsProp;
  const isNull = (k: string) => nullFields.includes(k);
  const addNull = (k: string) => {
    if (!setNullFieldsProp) return;
    if (nullFields.includes(k)) return;
    setNullFieldsProp([...nullFields, k]);
    // When the user flips a field to explicit-null, clear any typed value
    // so the body matches what they see (only the null entry remains).
    if (values[k] !== undefined) {
      const next = { ...values };
      delete next[k];
      setValues(next);
    }
  };
  const removeNull = (k: string) => {
    if (!setNullFieldsProp) return;
    setNullFieldsProp(nullFields.filter((n) => n !== k));
  };
  const tbl = findTable(table);
  const [search, setSearch] = useState('');
  const [filter, setFilter] = useState<Filter>('set');
  const [addPickerOpen, setAddPickerOpen] = useState(false);

  // Form ↔ JSON toggle (per v2.2 Mode_Create design). When 'json', the user
  // edits the body directly in a Textarea; when valid JSON, it parses back to
  // `values`. When 'form', the existing per-field controls are the source of
  // truth. The toggle holds its own draft text so editing doesn't ping-pong
  // through every keystroke.
  const [bodyMode, setBodyMode] = useState<'form' | 'json'>('form');
  const [jsonText, setJsonText] = useState<string>('');
  const [jsonError, setJsonError] = useState<string | null>(null);

  // Snapshot values into jsonText whenever we enter JSON mode
  useEffect(() => {
    if (bodyMode !== 'json') return;
    setJsonText(JSON.stringify(valuesToJsonBody(tbl, values), null, 2));
    setJsonError(null);
  }, [bodyMode, tbl, values]);

  // Set / Update / Remove helpers
  const setField = (k: string, v: CreateFieldValue) => setValues({ ...values, [k]: v });
  const removeField = (k: string) => {
    const next = { ...values };
    delete next[k];
    setValues(next);
  };

  const writable = useMemo<ColumnMeta[]>(
    () =>
      (tbl?.columns ?? [])
        .filter((c) => !isSystemManaged(c))
        // Per-purpose write-eligibility — backed by AttributeMetadata.
        //   • Create surfaces columns where IsValidForCreate is true.
        //   • Update surfaces columns where IsValidForUpdate is true.
        // We treat `undefined` as "true" so the filter doesn't accidentally
        // hide everything on entities that haven't been enriched yet, or
        // on columns whose host build doesn't return the flag. The real
        // server response is the source of truth — sending an attribute
        // that the server doesn't accept gets a 400 with a clear message.
        .filter((c) => {
          if (purpose === 'create') return c.isValidForCreate !== false;
          if (purpose === 'update') return c.isValidForUpdate !== false;
          if (purpose === 'upsert')
            return c.isValidForCreate !== false && c.isValidForUpdate !== false;
          return true;
        }),
    [tbl, purpose],
  );

  const requiredFields = useMemo(() => writable.filter((c) => c.required), [writable]);
  // "Populated" includes both columns with a value AND columns flagged as
  // explicit-null — both produce a body entry, so both count toward
  // "the user has expressed an intent for this column".
  const populatedFields = useMemo(
    () => writable.filter((c) => c.logicalName in values || nullFields.includes(c.logicalName)),
    [writable, values, nullFields],
  );
  // Required-field check: a column marked explicit-null still counts as
  // "missing" since the server will reject a null on a required column —
  // this matches Dataverse's runtime behavior and surfaces the issue.
  const missingRequired = useMemo(
    () =>
      requiredFields.filter(
        (c) => !(c.logicalName in values) || nullFields.includes(c.logicalName),
      ),
    [requiredFields, values, nullFields],
  );

  // What to render in the form area.
  //
  // Order rule (matches the user's mental model when adding fields):
  //   1. Required columns first, in metadata order — required fields are
  //      always present, always at the top.
  //   2. User-added (non-required) columns in the order the user added
  //      them — JS object iteration preserves insertion order, so the
  //      most-recently-added field naturally appears at the bottom of
  //      the list and the user doesn't have to fish for it.
  //
  // We do NOT alphabetize. Alphabetical sort makes "Add field" feel
  // unpredictable (a new field can slot into the middle), which is what
  // the user flagged.
  const visible = useMemo(() => {
    const writableById = new Map(writable.map((c) => [c.logicalName, c]));

    // Required columns in metadata order (regardless of whether they're populated)
    const requiredCols = writable.filter((c) => c.required);
    const requiredIds = new Set(requiredCols.map((c) => c.logicalName));

    // Then user-added non-required fields in insertion order. Sourced
    // from BOTH `values` keys (typed entries) AND `nullFields` (explicit
    // null entries) — both produce a body entry the user wants to see.
    const userAddedCols: ColumnMeta[] = [];
    const seenAdded = new Set<string>();
    const consider = (k: string) => {
      if (requiredIds.has(k) || seenAdded.has(k)) return;
      const col = writableById.get(k);
      if (col) {
        userAddedCols.push(col);
        seenAdded.add(k);
      }
    };
    for (const k of Object.keys(values)) consider(k);
    for (const k of nullFields) consider(k);

    let pool: ColumnMeta[] = [...requiredCols, ...userAddedCols];

    if (filter === 'required') pool = pool.filter((c) => c.required);
    if (filter === 'set')
      pool = pool.filter((c) => c.logicalName in values || nullFields.includes(c.logicalName));
    if (filter === 'common') {
      pool = pool.filter(
        (c) =>
          c.required ||
          c.logicalName === tbl?.primaryName ||
          c.logicalName === 'statecode' ||
          c.logicalName === 'statuscode' ||
          isLookupLike(c),
      );
    }
    if (search) {
      const q = search.toLowerCase();
      pool = pool.filter(
        (c) =>
          c.displayName.toLowerCase().includes(q) ||
          c.logicalName.toLowerCase().includes(q) ||
          c.attributeType.toLowerCase().includes(q),
      );
    }
    return pool;
  }, [writable, values, filter, search, tbl]);

  // Pool of fields available to add (writable, not currently visible)
  const addable = useMemo(() => {
    const visibleIds = new Set(visible.map((c) => c.logicalName));
    return writable.filter((c) => !visibleIds.has(c.logicalName));
  }, [writable, visible]);

  if (!tbl) {
    return (
      <MessageBar layout="multiline" intent="error">
        <MessageBarBody>
          Unknown table <code>{table}</code>.
        </MessageBarBody>
      </MessageBar>
    );
  }

  return (
    <div>
      <PaneHead
        icon={FormNew20Filled}
        title="Field set"
        sub={
          <span>
            Pick the columns this <code>POST</code> should set.{' '}
            <span style={{ color: tokens.colorNeutralForeground3 }}>
              Only fields below are emitted in the body — empty columns stay out.
            </span>
          </span>
        }
        group={group}
      >
        {bodyMode === 'form' && (
          <Input
            contentBefore={<Search20Regular />}
            placeholder="Filter fields…"
            value={search}
            onChange={(_, d) => setSearch(d.value)}
            size="small"
            style={{ width: 220 }}
          />
        )}
        <Badge appearance="ghost">{populatedFields.length} populated</Badge>
        {requiredFields.length > 0 && (
          <Badge appearance="tint" color={missingRequired.length ? 'danger' : 'success'}>
            {requiredFields.length - missingRequired.length}/{requiredFields.length} required
          </Badge>
        )}
        {/* Form ↔ JSON segmented toggle. Uses the shared SegmentedToggle
            primitive so it picks up the same hairline + radius treatment
            as FilterEditor (AND/OR) and OrderbyEditor (Asc/Desc). */}
        <SegmentedToggle ariaLabel="Body editor mode">
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
      </PaneHead>

      {missingRequired.length > 0 && (
        <MessageBar layout="multiline" intent="warning" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            <MessageBarTitle>
              {missingRequired.length} required field
              {missingRequired.length === 1 ? ' is' : 's are'} unset.
            </MessageBarTitle>
            Dataverse will reject the create with{' '}
            <code>0x80040217 · A required field is missing</code> unless every required column has a
            value. <strong>Missing:</strong>{' '}
            {missingRequired.map((c) => (
              <code key={c.logicalName} style={{ marginRight: 6 }}>
                {c.logicalName}
              </code>
            ))}
          </MessageBarBody>
        </MessageBar>
      )}

      {bodyMode === 'form' ? (
        <>
          {/* Filter chip row */}
          <div style={{ display: 'flex', gap: 6, marginBottom: 14, flexWrap: 'wrap' }}>
            <FilterChip
              label="Set"
              count={populatedFields.length}
              active={filter === 'set'}
              onClick={() => setFilter('set')}
            />
            <FilterChip
              label="Required"
              count={requiredFields.length}
              active={filter === 'required'}
              onClick={() => setFilter('required')}
            />
            <FilterChip
              label="Common"
              active={filter === 'common'}
              onClick={() => setFilter('common')}
            />
            <FilterChip
              label="All"
              count={writable.length}
              active={filter === 'all'}
              onClick={() => setFilter('all')}
            />
          </div>

          {/* Sectioned field rows — General / Choices / Lookups / System */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 14, maxWidth: 880 }}>
            {visible.length === 0 && (
              <Caption1
                style={{
                  padding: '24px 4px',
                  color: tokens.colorNeutralForeground3,
                  textAlign: 'center',
                  fontStyle: 'italic',
                }}
              >
                No fields match. Switch the filter or use “Add field”.
              </Caption1>
            )}
            {sectionsOf(visible).map((sec) => (
              <FieldSection
                key={sec.id}
                title={sec.title}
                icon={sec.icon}
                tagText={sec.tagText}
                tagAppearance={sec.tagAppearance}
              >
                {sec.cols.map((col) => (
                  <FieldRow
                    key={col.logicalName}
                    table={tbl}
                    col={col}
                    value={values[col.logicalName]}
                    populated={col.logicalName in values || nullFields.includes(col.logicalName)}
                    isNull={isNull(col.logicalName)}
                    nullSupported={nullFieldsSupported}
                    onChange={(v) => setField(col.logicalName, v)}
                    onRemove={() => {
                      // Remove from both value map AND null flag — a single
                      // delete from the user's POV.
                      removeField(col.logicalName);
                      if (nullFields.includes(col.logicalName)) removeNull(col.logicalName);
                    }}
                    onSetNull={() => addNull(col.logicalName)}
                    onClearNull={() => removeNull(col.logicalName)}
                  />
                ))}
              </FieldSection>
            ))}
          </div>
        </>
      ) : (
        // JSON edit mode — Monaco editor with json language. Parses back
        // to fieldValues on every keystroke so the Form view stays in
        // sync.
        <JsonEditPane
          tbl={tbl}
          text={jsonText}
          error={jsonError}
          themeMode={themeMode}
          onChange={(t) => {
            setJsonText(t);
            try {
              const parsed = JSON.parse(t);
              if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
                setValues(jsonBodyToValues(tbl, parsed));
                setJsonError(null);
              } else {
                setJsonError('Body must be a JSON object.');
              }
            } catch (e) {
              setJsonError(e instanceof Error ? e.message : 'Invalid JSON.');
            }
          }}
        />
      )}

      {/* Add field — only if there's something left to add */}
      {addable.length > 0 && (
        <div
          style={{ marginTop: 14, display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}
        >
          {!addPickerOpen ? (
            <Button
              icon={<Add20Regular />}
              appearance="outline"
              size="small"
              onClick={() => setAddPickerOpen(true)}
            >
              Add field
            </Button>
          ) : (
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
              <Combobox
                size="small"
                placeholder={`Add one of ${addable.length} fields…`}
                style={{ minWidth: 320 }}
                onOptionSelect={(_, d) => {
                  if (!d.optionValue) return;
                  const col = findColumn(tbl, d.optionValue);
                  if (!col) return;
                  setField(col.logicalName, defaultValueFor(col));
                  setAddPickerOpen(false);
                }}
              >
                {addable.map((c) => (
                  <Option key={c.logicalName} value={c.logicalName} text={c.displayName}>
                    <div style={{ display: 'flex', flexDirection: 'column' }}>
                      <span>{c.displayName}</span>
                      <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                        {c.logicalName} · <code>{c.attributeType}</code>
                        {c.required && (
                          <span style={{ marginLeft: 6, color: tokens.colorPaletteRedForeground1 }}>
                            required
                          </span>
                        )}
                      </Caption1>
                    </div>
                  </Option>
                ))}
              </Combobox>
              <Button size="small" appearance="subtle" onClick={() => setAddPickerOpen(false)}>
                Cancel
              </Button>
            </span>
          )}
          <Caption1 style={{ color: tokens.colorNeutralForeground3, marginLeft: 'auto' }}>
            {addable.length} more available on <code>{tbl.logicalName}</code>
          </Caption1>
        </div>
      )}
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// FieldRow — one column's label + value + remove
// ──────────────────────────────────────────────────────────────
function FieldRow({
  table,
  col,
  value,
  populated,
  isNull,
  nullSupported,
  onChange,
  onRemove,
  onSetNull,
  onClearNull,
}: {
  table: TableMeta;
  col: ColumnMeta;
  value: CreateFieldValue | undefined;
  populated: boolean;
  /** True when the column is flagged for explicit-null in the body. */
  isNull: boolean;
  /** True when the parent supplied set-null handlers — drives whether the
   *  "Set null" button renders. Modes that haven't opted into the feature
   *  still get a clean row without the affordance. */
  nullSupported: boolean;
  onChange: (v: CreateFieldValue) => void;
  onRemove: () => void;
  onSetNull: () => void;
  onClearNull: () => void;
}) {
  const s = useStudioStyles();
  const isMissing = col.required && !populated;

  return (
    <div
      className={mergeClasses(s.inlineCard)}
      style={{
        padding: '10px 12px',
        display: 'grid',
        gridTemplateColumns: '220px 1fr 30px',
        gap: 12,
        alignItems: 'start',
        borderColor: isMissing ? tokens.colorPaletteRedBorder2 : undefined,
      }}
    >
      {/* Label cell */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 2, paddingTop: 4 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <span style={{ fontWeight: 500, color: tokens.colorNeutralForeground1 }}>
            {col.displayName}
          </span>
          {col.required && (
            <Tooltip
              content="Required — Dataverse rejects the create when this is empty."
              relationship="description"
            >
              <Badge appearance="tint" color="danger" size="extra-small">
                req
              </Badge>
            </Tooltip>
          )}
          {isLookupLike(col) && (
            <Tooltip
              content={`Lookup → emits ${col.logicalName}@odata.bind`}
              relationship="description"
            >
              <Database20Regular
                style={{ width: 12, height: 12, color: tokens.colorBrandForeground1 }}
              />
            </Tooltip>
          )}
          {col.attributeType === 'Uniqueidentifier' && (
            <Tooltip
              content="Primary key — Dataverse will auto-generate this. Override only if you know what you're doing."
              relationship="description"
            >
              <Key20Regular
                style={{ width: 12, height: 12, color: tokens.colorPaletteGoldBorderActive }}
              />
            </Tooltip>
          )}
        </div>
        <Caption1
          style={{
            color: tokens.colorNeutralForeground3,
            fontFamily: tokens.fontFamilyMonospace,
            fontSize: 10,
          }}
        >
          {col.logicalName} ·{' '}
          <span style={{ color: tokens.colorBrandForeground2 }}>{col.attributeType}</span>
          {typeHint(col)}
        </Caption1>
      </div>

      {/* Value cell — single row: [input (or null preview)] [Switch at end].
          Input is the PRIMARY affordance (most of the row width); the
          "Send as null" Switch sits at the trailing edge as a modifier.
          When the user toggles null on, the input slot swaps to a body-
          fragment preview but the switch stays in the same position so
          the layout doesn't shift. Description (if any) renders below. */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 4, minWidth: 0 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 0 }}>
          <div style={{ flex: 1, minWidth: 0 }}>
            {isNull ? (
              // Null-mode preview. Same row footprint as the input, so the
              // Switch's position doesn't jump when the user toggles. Shows
              // the exact body fragment we'll emit so there's no doubt:
              //   regular column → `"col": null`
              //   lookup column  → `"col@odata.bind": null`
              <div
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: 6,
                  padding: '4px 8px',
                  background: tokens.colorPaletteDarkOrangeBackground1,
                  border: `1px dashed ${tokens.colorPaletteDarkOrangeBorder2}`,
                  borderRadius: tokens.borderRadiusSmall,
                  minHeight: 28,
                  overflow: 'hidden',
                }}
                title="This field will be sent as null. Toggle the Switch off to enter a value."
              >
                <Badge
                  appearance="filled"
                  color="warning"
                  size="extra-small"
                  style={{ fontFamily: tokens.fontFamilyMonospace, fontWeight: 700 }}
                >
                  null
                </Badge>
                <code
                  style={{
                    fontFamily: tokens.fontFamilyMonospace,
                    fontSize: 11,
                    color: tokens.colorNeutralForeground2,
                    whiteSpace: 'nowrap',
                    overflow: 'hidden',
                    textOverflow: 'ellipsis',
                    minWidth: 0,
                  }}
                >
                  {isLookupLike(col)
                    ? `"${col.logicalName}@odata.bind": null`
                    : `"${col.logicalName}": null`}
                </code>
              </div>
            ) : (
              <FieldInput table={table} col={col} value={value} onChange={onChange} />
            )}
          </div>
          {nullSupported && (
            // Trailing Switch — compact, label sits to the right so the
            // gesture feels like "enable this modifier" rather than
            // "this is the primary affordance".
            <Switch
              checked={isNull}
              onChange={(_, d) => {
                if (d.checked) onSetNull();
                else onClearNull();
              }}
              label={
                <span
                  style={{
                    fontSize: 11,
                    color: tokens.colorNeutralForeground3,
                    whiteSpace: 'nowrap',
                  }}
                >
                  Send as{' '}
                  <code style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>null</code>
                </span>
              }
              labelPosition="after"
            />
          )}
        </div>
        {col.description && !isNull && (
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>{col.description}</Caption1>
        )}
      </div>

      {/* Remove */}
      {!col.required ? (
        <Tooltip content="Remove field from the body" relationship="label">
          <Button
            size="small"
            appearance="subtle"
            icon={<Delete20Regular />}
            onClick={onRemove}
            aria-label={`Remove ${col.displayName}`}
          />
        </Tooltip>
      ) : (
        <span style={{ width: 30 }} aria-hidden="true" />
      )}
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// FieldInput — dispatch to the right control for the column's type
// ──────────────────────────────────────────────────────────────
function FieldInput({
  table,
  col,
  value,
  onChange,
}: {
  table: TableMeta;
  col: ColumnMeta;
  value: CreateFieldValue | undefined;
  onChange: (v: CreateFieldValue) => void;
}) {
  // Lazy per-column enrichment. The basic Attributes projection that
  // `useLiveTable` fetches is intentionally slim — no OptionSet for
  // picklists, no Targets for lookups, no MaxLength/Format/Precision for
  // strings/numerics. Each individual field row fires this hook when it
  // mounts, which on first render triggers the typed-cast fetch via
  // pptbClient.getAttributeWithOptionSet / getAttributeDetailedMetadata.
  // The result patches the column in the live registry, which re-renders
  // this row with full data — picklist options light up, lookups gain
  // their target list, integers gain their min/max, etc.
  //
  // Same pattern FilterValueInput uses in Retrieve Multiple. Cached per
  // (entity, attribute) so re-renders don't refetch.
  useColumnDetail(table.logicalName, col.logicalName);
  switch (col.attributeType) {
    case 'String':
      return <StringFieldInput col={col} value={value as string | undefined} onChange={onChange} />;
    case 'Memo':
      return <MemoFieldInput col={col} value={value as string | undefined} onChange={onChange} />;
    case 'Integer':
      return (
        <IntegerFieldInput col={col} value={value as number | undefined} onChange={onChange} />
      );
    case 'BigInt':
      return (
        <NumericFieldInput
          value={value as number | undefined}
          onChange={onChange}
          integer
          min={col.minValue}
          max={col.maxValue}
        />
      );
    case 'Decimal':
      return (
        <DecimalFieldInput col={col} value={value as number | undefined} onChange={onChange} />
      );
    case 'Double':
      return (
        <DecimalFieldInput col={col} value={value as number | undefined} onChange={onChange} />
      );
    case 'Money':
      return <MoneyFieldInput col={col} value={value as number | undefined} onChange={onChange} />;
    case 'Boolean':
      return (
        <BooleanFieldInput col={col} value={value as boolean | undefined} onChange={onChange} />
      );
    case 'DateTime':
      return (
        <DateTimeFieldInput col={col} value={value as string | undefined} onChange={onChange} />
      );
    case 'Picklist':
    case 'State':
    case 'Status':
    case 'EntityName':
      return <ChoiceFieldInput col={col} value={value as number | undefined} onChange={onChange} />;
    case 'MultiSelectPicklist':
      return (
        <MultiChoiceFieldInput
          col={col}
          value={value as number[] | undefined}
          onChange={onChange}
        />
      );
    case 'Lookup':
    case 'Customer':
    case 'Owner':
      return (
        <LookupFieldInput
          col={col}
          value={value as LookupFieldValue | undefined}
          onChange={onChange}
        />
      );
    case 'Uniqueidentifier':
      return <GuidFieldInput value={value as string | undefined} onChange={onChange} />;
    default:
      return (
        <Input
          size="small"
          value={String(value ?? '')}
          onChange={(_, d) => onChange(d.value)}
          placeholder="value"
        />
      );
  }
}

// ──────────────────────────────────────────────────────────────
// Per-type controls (mirrors FilterValueInput patterns)
// ──────────────────────────────────────────────────────────────
function StringFieldInput({
  col,
  value,
  onChange,
}: {
  col: StringColumnMeta;
  value: string | undefined;
  onChange: (v: string) => void;
}) {
  const isLong = col.format === 'TextArea';
  if (isLong) {
    return (
      <Textarea
        rows={3}
        value={value ?? ''}
        maxLength={col.maxLength}
        onChange={(_, d) => onChange(d.value)}
        style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 12 }}
      />
    );
  }
  const type =
    col.format === 'Email'
      ? 'email'
      : col.format === 'Url'
        ? 'url'
        : col.format === 'Phone'
          ? 'tel'
          : 'text';
  const placeholder =
    col.format === 'Email'
      ? 'name@example.com'
      : col.format === 'Url'
        ? 'https://example.com'
        : col.format === 'Phone'
          ? '+1 (555) 555-0000'
          : col.format === 'TickerSymbol'
            ? 'MSFT'
            : `up to ${col.maxLength} chars`;
  return (
    <Input
      size="small"
      type={type}
      value={value ?? ''}
      maxLength={col.maxLength}
      placeholder={placeholder}
      onChange={(_, d) => onChange(d.value)}
    />
  );
}

function MemoFieldInput({
  col,
  value,
  onChange,
}: {
  col: MemoColumnMeta;
  value: string | undefined;
  onChange: (v: string) => void;
}) {
  const v = value ?? '';
  return (
    <Field
      hint={
        <span style={{ color: tokens.colorNeutralForeground3 }}>
          {v.length.toLocaleString()} / {col.maxLength.toLocaleString()}
        </span>
      }
    >
      <Textarea
        rows={4}
        value={v}
        maxLength={col.maxLength}
        placeholder={`up to ${col.maxLength.toLocaleString()} chars`}
        onChange={(_, d) => onChange(d.value)}
        style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 12 }}
      />
    </Field>
  );
}

function IntegerFieldInput({
  col,
  value,
  onChange,
}: {
  col: IntegerColumnMeta;
  value: number | undefined;
  onChange: (v: number) => void;
}) {
  // Format-driven presets (mirrors FilterValueInput)
  if (col.format === 'Language' || col.format === 'Locale') {
    return (
      <Combobox
        size="small"
        value={value != null ? String(value) : ''}
        selectedOptions={value != null ? [String(value)] : []}
        onOptionSelect={(_, d) => d.optionValue && onChange(Number(d.optionValue))}
        placeholder="LCID"
      >
        <Option value="1033">English (US) · 1033</Option>
        <Option value="1031">German · 1031</Option>
        <Option value="1036">French · 1036</Option>
        <Option value="1041">Japanese · 1041</Option>
        <Option value="1043">Dutch · 1043</Option>
      </Combobox>
    );
  }
  if (col.format === 'Duration') {
    return (
      <span style={{ display: 'inline-flex', gap: 4, alignItems: 'center' }}>
        <NumericFieldInput
          value={value}
          onChange={onChange}
          min={col.minValue ?? 0}
          max={col.maxValue ?? 525_600}
          step={1}
          integer
        />
        <span style={{ fontSize: 10, color: tokens.colorNeutralForeground3 }}>min</span>
      </span>
    );
  }
  return (
    <NumericFieldInput
      value={value}
      onChange={onChange}
      min={col.minValue}
      max={col.maxValue}
      step={1}
      integer
    />
  );
}

function DecimalFieldInput({
  col,
  value,
  onChange,
}: {
  col: DecimalColumnMeta | DoubleColumnMeta;
  value: number | undefined;
  onChange: (v: number) => void;
}) {
  return (
    <NumericFieldInput
      value={value}
      onChange={onChange}
      min={col.minValue}
      max={col.maxValue}
      step={Math.pow(10, -col.precision)}
    />
  );
}

function NumericFieldInput({
  value,
  onChange,
  min,
  max,
  step,
  integer,
}: {
  value: number | undefined;
  onChange: (v: number) => void;
  min?: number;
  max?: number;
  step?: number;
  integer?: boolean;
}) {
  return (
    <SpinButton
      size="small"
      value={value ?? 0}
      min={min}
      max={max}
      step={step ?? (integer ? 1 : 0.01)}
      onChange={(_, d) => {
        const n = d.value ?? Number(d.displayValue ?? 0);
        if (Number.isFinite(n)) onChange(integer ? Math.round(n) : n);
      }}
    />
  );
}

function MoneyFieldInput({
  col,
  value,
  onChange,
}: {
  col: MoneyColumnMeta;
  value: number | undefined;
  onChange: (v: number) => void;
}) {
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
      <span
        style={{
          fontFamily: tokens.fontFamilyMonospace,
          fontSize: 12,
          color: tokens.colorNeutralForeground3,
        }}
      >
        $
      </span>
      <NumericFieldInput
        value={value}
        onChange={onChange}
        min={col.minValue}
        max={col.maxValue}
        step={Math.pow(10, -col.precision)}
      />
      <Caption1 style={{ color: tokens.colorNeutralForeground3, fontSize: 10 }}>
        currency = active org transactioncurrency
      </Caption1>
    </span>
  );
}

function BooleanFieldInput({
  col,
  value,
  onChange,
}: {
  col: BooleanColumnMeta;
  value: boolean | undefined;
  onChange: (v: boolean) => void;
}) {
  const checked = value === true;
  return (
    <Switch
      checked={checked}
      onChange={(_, d) => onChange(d.checked)}
      label={checked ? col.trueOption.label : col.falseOption.label}
    />
  );
}

function DateTimeFieldInput({
  col,
  value,
  onChange,
}: {
  col: DateTimeColumnMeta;
  value: string | undefined;
  onChange: (v: string) => void;
}) {
  const isDateOnly = col.format === 'DateOnly' || col.dateTimeBehavior === 'DateOnly';
  const behaviorHint =
    col.dateTimeBehavior === 'UserLocal'
      ? 'stored as UTC, displayed in user TZ'
      : col.dateTimeBehavior === 'DateOnly'
        ? 'midnight UTC, no TZ conversion'
        : 'wall-clock as UTC (no TZ math)';

  const parsed = value ? new Date(value) : null;
  const dateOnly = parsed && !Number.isNaN(parsed.getTime()) ? parsed : null;

  const writeDate = (d: Date | null | undefined) => {
    if (!d) {
      onChange('');
      return;
    }
    if (isDateOnly) {
      const yyyy = d.getFullYear();
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      const dd = String(d.getDate()).padStart(2, '0');
      onChange(`${yyyy}-${mm}-${dd}`);
    } else {
      const prior = parsed && !Number.isNaN(parsed.getTime()) ? parsed : new Date(d);
      const next = new Date(d);
      next.setHours(prior.getHours(), prior.getMinutes(), 0, 0);
      onChange(next.toISOString());
    }
  };
  const writeTime = (d: Date | null | undefined) => {
    if (!d || !dateOnly) return;
    const next = new Date(dateOnly);
    next.setHours(d.getHours(), d.getMinutes(), 0, 0);
    onChange(next.toISOString());
  };

  if (isDateOnly) {
    return (
      <Tooltip
        content={`Behavior: ${col.dateTimeBehavior} — ${behaviorHint}`}
        relationship="description"
      >
        <DatePicker
          size="small"
          value={dateOnly}
          onSelectDate={writeDate}
          placeholder="Pick a date…"
          formatDate={(d) => (d ? d.toLocaleDateString() : '')}
          allowTextInput
        />
      </Tooltip>
    );
  }
  return (
    <Tooltip
      content={`Behavior: ${col.dateTimeBehavior} — ${behaviorHint}`}
      relationship="description"
    >
      <span style={{ display: 'inline-flex', gap: 6, alignItems: 'center', flexWrap: 'wrap' }}>
        <DatePicker
          size="small"
          value={dateOnly}
          onSelectDate={writeDate}
          placeholder="Date…"
          formatDate={(d) => (d ? d.toLocaleDateString() : '')}
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

function ChoiceFieldInput({
  col,
  value,
  onChange,
}: {
  col: ColumnMeta;
  value: number | undefined;
  onChange: (v: number) => void;
}) {
  const options = columnOptions(col) ?? [];
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

function MultiChoiceFieldInput({
  col,
  value,
  onChange,
}: {
  col: ColumnMeta;
  value: number[] | undefined;
  onChange: (v: number[]) => void;
}) {
  return (
    <div>
      <MultiChoiceTagPicker col={col} value={value} onChange={onChange} />
      <Caption1
        style={{
          display: 'block',
          color: tokens.colorNeutralForeground3,
          marginTop: 4,
          fontSize: 10,
        }}
      >
        Emitted as comma-separated integers (e.g. <code>"1,2,3"</code>) per Dataverse Web API
        contract.
      </Caption1>
    </div>
  );
}

/**
 * Filterable multi-select via Fluent v9 TagPicker.
 *
 * The TagPicker works in string-keyed terms, so we use `String(option.value)`
 * as the option key and translate to the integer array at the boundary. This
 * lets the picker handle filtering / typeahead natively while we keep the
 * canonical OData-friendly representation in state.
 */
function MultiChoiceTagPicker({
  col,
  value,
  onChange,
}: {
  col: ColumnMeta;
  value: number[] | undefined;
  onChange: (v: number[]) => void;
}) {
  const options = columnOptions(col) ?? [];
  const [query, setQuery] = useState('');
  const selected = (value ?? []).map(String);

  const labelByKey = useMemo(() => {
    const m = new Map<string, string>();
    for (const o of options) m.set(String(o.value), o.label);
    return m;
  }, [options]);
  const allKeys = useMemo(() => options.map((o) => String(o.value)), [options]);

  const children = useTagPickerFilter({
    query,
    options: allKeys,
    noOptionsElement: <TagPickerOption value="__no-matches">No matches</TagPickerOption>,
    renderOption: (key) => (
      <TagPickerOption key={key} value={key} text={labelByKey.get(key) ?? key}>
        <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
          {labelByKey.get(key) ?? key}
          <span
            style={{
              fontFamily: tokens.fontFamilyMonospace,
              fontSize: 10,
              color: tokens.colorNeutralForeground3,
            }}
          >
            · {key}
          </span>
        </span>
      </TagPickerOption>
    ),
    filter: (key) => {
      if (selected.includes(key)) return false;
      const label = labelByKey.get(key) ?? key;
      return label.toLowerCase().includes(query.toLowerCase());
    },
  });

  return (
    <TagPicker
      selectedOptions={selected}
      onOptionSelect={(_, d) => {
        if (d.value === '__no-matches') return;
        const next = d.selectedOptions
          .map((s) => Number(s))
          .filter((n) => Number.isFinite(n))
          .sort((a, b) => a - b);
        onChange(next);
        setQuery('');
      }}
      disableAutoFocus={query.length === 0}
    >
      <TagPickerControl style={{ maxWidth: 520 }}>
        <TagPickerGroup aria-label="Selected choices">
          {selected.map((k) => (
            <Tag key={k} value={k} shape="rounded">
              {labelByKey.get(k) ?? k}
            </Tag>
          ))}
        </TagPickerGroup>
        <TagPickerInput
          aria-label="Choose options"
          placeholder={selected.length === 0 ? 'Type to filter options…' : ''}
          value={query}
          onChange={(e) => setQuery(e.target.value)}
        />
      </TagPickerControl>
      <TagPickerList>{children}</TagPickerList>
    </TagPicker>
  );
}

// Exported so MergeFieldDiff can reuse it for the custom-override path
// on lookup columns. Same live-typeahead semantics as in the field set.
export function LookupFieldInput({
  col,
  value,
  onChange,
}: {
  col: LookupColumnMeta | CustomerColumnMeta | OwnerColumnMeta;
  value: LookupFieldValue | undefined;
  onChange: (v: LookupFieldValue) => void;
}) {
  // Polymorphic lookups (Customer / Owner) — let user pick the target table first.
  // Customer: account | contact. Owner: systemuser | team.
  //
  // Important: `col.targets` is EMPTY initially. The basic Attributes
  // projection doesn't include `Targets`; `useColumnDetail` (called from
  // the parent FieldInput) drills in via getAttributeDetailedMetadata and
  // patches the column in the live registry — which re-renders us with
  // a populated `col.targets`. So we can't capture `col.targets[0]` once
  // in useState init; we have to keep syncing as targets become
  // available.
  const [target, setTarget] = useState<string>(value?.targetEntity ?? col.targets[0] ?? '');
  const polymorphic = col.targets.length > 1;

  // Sync the local `target` state when col.targets becomes populated
  // (enrichment fetch completed) or when the parent rewrites `value`
  // with a different target (saved-request load, polymorphic switch
  // from outside, etc.).
  useEffect(() => {
    if (value?.targetEntity && value.targetEntity !== target) {
      setTarget(value.targetEntity);
    } else if (!target && col.targets.length > 0) {
      setTarget(col.targets[0]);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [col.targets, value?.targetEntity]);

  const targetTbl = findTable(target);

  // Live typeahead — same hook RecordPicker uses for the Retrieve Single
  // record selector. Resolves real records from the connected Dataverse.
  const [search, setSearch] = useState('');
  // True while the user is typing — distinguishes "show search text" from
  // "show resolved name of the picked record".
  const [userIsTyping, setUserIsTyping] = useState(false);
  const { rows, loading, error } = useLookupRecords(target, userIsTyping ? search : '');

  // Reset typing state when target changes (polymorphic switch) or the
  // parent clears the selection.
  useEffect(() => {
    if (!value?.id) {
      setSearch('');
      setUserIsTyping(false);
    }
  }, [value?.id, target]);

  const selectedRow = value?.id
    ? rows.find((r) => r.id.toLowerCase() === value.id.toLowerCase())
    : undefined;
  const selectedLabel = selectedRow?.name ?? (value?.id ? `(${value.id.slice(0, 8)}…)` : '');
  const displayValue = userIsTyping ? search : selectedLabel;

  // Before targets enrichment lands, `target` is empty and `col.targets`
  // is []. Show a "loading targets…" placeholder rather than the previous
  // "Search undefined records…" which read as a bug.
  if (!target) {
    return (
      <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
        <Badge appearance="ghost" style={{ fontFamily: tokens.fontFamilyMonospace }}>
          → loading targets…
        </Badge>
        <Combobox size="small" disabled placeholder="Loading…" style={{ minWidth: 240 }}>
          <Option value="__loading" text="" disabled>
            —
          </Option>
        </Combobox>
      </span>
    );
  }

  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
      {polymorphic ? (
        <Combobox
          size="small"
          value={targetTbl?.displayName ?? target}
          selectedOptions={[target]}
          onOptionSelect={(_, d) => {
            if (d.optionValue) {
              setTarget(d.optionValue);
              // Reset id when target changes — the GUID space differs per entity.
              onChange({ id: '', targetEntity: d.optionValue });
              setSearch('');
              setUserIsTyping(false);
            }
          }}
          style={{ width: 130 }}
        >
          {col.targets.map((t) => {
            const tt = findTable(t);
            return (
              <Option key={t} value={t} text={tt?.displayName ?? t}>
                {tt?.displayName ?? t}
              </Option>
            );
          })}
        </Combobox>
      ) : (
        <Badge appearance="ghost" style={{ fontFamily: tokens.fontFamilyMonospace }}>
          → {targetTbl?.entitySetName ?? target}
        </Badge>
      )}

      <Combobox
        size="small"
        freeform
        clearable
        value={displayValue}
        selectedOptions={value?.id ? [value.id] : []}
        onChange={(e) => {
          const next = (e.target as HTMLInputElement).value;
          setSearch(next);
          setUserIsTyping(true);
          if (!next && value?.id) {
            onChange({ id: '', targetEntity: target });
          }
        }}
        onOptionSelect={(_, d) => {
          if (d.optionValue) {
            onChange({ id: d.optionValue, targetEntity: target });
            setSearch('');
            setUserIsTyping(false);
          }
        }}
        placeholder={`Search ${targetTbl?.displayName ?? target} records…`}
        style={{ minWidth: 240 }}
        listbox={{ style: { maxHeight: 320 } }}
      >
        {loading && (
          <Option value="__loading" text="" disabled>
            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Searching…</Caption1>
          </Option>
        )}
        {!loading && error && (
          <Option value="__error" text="" disabled>
            <Caption1 style={{ color: tokens.colorPaletteRedForeground1 }}>{error}</Caption1>
          </Option>
        )}
        {!loading && !error && rows.length === 0 && (
          <Option value="__none" text="" disabled>
            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
              {search ? `No ${target} records match "${search}"` : `Type to search ${target}`}
            </Caption1>
          </Option>
        )}
        {!loading &&
          !error &&
          rows.map((r) => (
            <Option key={r.id} value={r.id} text={r.name}>
              <Persona
                size="extra-small"
                name={r.name}
                secondaryText={r.id}
                avatar={{ color: 'colorful' }}
              />
            </Option>
          ))}
      </Combobox>
    </span>
  );
}

function GuidFieldInput({
  value,
  onChange,
}: {
  value: string | undefined;
  onChange: (v: string) => void;
}) {
  const v = value ?? '';
  const valid = !v || /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(v);
  return (
    <Input
      size="small"
      value={v}
      placeholder="00000000-0000-0000-0000-000000000000"
      onChange={(_, d) => onChange(d.value)}
      style={{ fontFamily: tokens.fontFamilyMonospace }}
      contentAfter={
        !valid ? (
          <span style={{ color: tokens.colorPaletteRedForeground1, fontSize: 10 }}>
            invalid GUID
          </span>
        ) : undefined
      }
    />
  );
}

// ──────────────────────────────────────────────────────────────
// Helpers
// ──────────────────────────────────────────────────────────────
function FilterChip({
  label,
  count,
  active,
  onClick,
}: {
  label: string;
  count?: number;
  active: boolean;
  onClick: () => void;
}) {
  return (
    <ToggleButton size="small" shape="circular" checked={active} onClick={onClick}>
      {label}
      {count != null ? ` · ${count}` : ''}
    </ToggleButton>
  );
}

function typeHint(c: ColumnMeta): string {
  switch (c.attributeType) {
    case 'String':
    case 'Memo':
      return ` · max ${c.maxLength}`;
    case 'Decimal':
    case 'Double':
      return ` · ${c.precision} dec`;
    case 'Money':
      return ` · ${c.precision} dec`;
    case 'Integer':
      return c.minValue != null || c.maxValue != null
        ? ` · [${c.minValue ?? '−∞'}…${c.maxValue ?? '∞'}]`
        : '';
    case 'Lookup':
    case 'Customer':
    case 'Owner':
      return ` · → ${c.targets.join(' | ')}`;
    case 'Picklist':
    case 'MultiSelectPicklist':
    case 'State':
    case 'Status':
    case 'EntityName':
      return ` · ${('options' in c ? c.options : []).length} options`;
    case 'DateTime':
      return ` · ${c.format} (${c.dateTimeBehavior})`;
    default:
      return '';
  }
}

/** Default value when the user adds (or auto-includes) a column. */
export function defaultValueFor(col: ColumnMeta): CreateFieldValue {
  switch (col.attributeType) {
    case 'String':
    case 'Memo':
      return '';
    case 'Integer':
    case 'BigInt':
    case 'Decimal':
    case 'Double':
    case 'Money':
      return 0;
    case 'Boolean':
      return col.defaultValue ?? false;
    case 'DateTime':
      return '';
    case 'Picklist':
      return col.defaultFormValue ?? col.options[0]?.value ?? 0;
    case 'State':
      return col.options[0]?.value ?? 0;
    case 'Status':
      return col.options[0]?.value ?? 1;
    case 'EntityName':
      return col.options[0]?.value ?? 0;
    case 'MultiSelectPicklist':
      return [];
    case 'Lookup':
    case 'Customer':
    case 'Owner':
      return { id: '', targetEntity: col.targets[0] };
    case 'Uniqueidentifier':
      return '';
    default:
      return '';
  }
}

// ──────────────────────────────────────────────────────────────
// Section grouping — General / Choices / Lookups / System
// Matches the v2.2 Mode_Create design (modes-write.jsx §Mode_Create):
//   "General"  → scalar fields (string, number, money, boolean, datetime)
//   "Choices"  → option-set fields (picklist, multi, state, status, entityname)
//   "Lookups"  → relationship fields (lookup, customer, owner) — labeled with
//                an @odata.bind hint per v2.2 design
//   "System"   → uniqueidentifier (rare — primary key override)
// ──────────────────────────────────────────────────────────────
type SectionId = 'general' | 'choices' | 'lookups' | 'system';

interface FieldSection {
  id: SectionId;
  title: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  icon: React.FC<any>;
  tagText?: string;
  tagAppearance?: 'brand' | 'subtle';
  cols: ColumnMeta[];
}

function sectionOf(col: ColumnMeta): SectionId {
  if (isLookupLike(col)) return 'lookups';
  switch (col.attributeType) {
    case 'Picklist':
    case 'MultiSelectPicklist':
    case 'State':
    case 'Status':
    case 'EntityName':
      return 'choices';
    case 'Uniqueidentifier':
      return 'system';
    default:
      return 'general';
  }
}

function sectionsOf(cols: ColumnMeta[]): FieldSection[] {
  const grouped: Record<SectionId, ColumnMeta[]> = {
    general: [],
    choices: [],
    lookups: [],
    system: [],
  };
  for (const c of cols) grouped[sectionOf(c)].push(c);
  const out: FieldSection[] = [];
  if (grouped.general.length)
    out.push({
      id: 'general',
      title: 'General',
      icon: TextNumberFormat20Regular,
      cols: grouped.general,
    });
  if (grouped.choices.length)
    out.push({
      id: 'choices',
      title: 'Choices',
      icon: ListBar20Regular,
      tagText: 'option sets',
      tagAppearance: 'subtle',
      cols: grouped.choices,
    });
  if (grouped.lookups.length)
    out.push({
      id: 'lookups',
      title: 'Lookups',
      icon: Link20Regular,
      tagText: '@odata.bind',
      tagAppearance: 'brand',
      cols: grouped.lookups,
    });
  if (grouped.system.length)
    out.push({
      id: 'system',
      title: 'System',
      icon: Key20Regular,
      cols: grouped.system,
    });
  return out;
}

function FieldSection({
  title,
  icon: Icon,
  tagText,
  tagAppearance,
  children,
}: {
  title: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  icon: React.FC<any>;
  tagText?: string;
  tagAppearance?: 'brand' | 'subtle';
  children: React.ReactNode;
}) {
  return (
    <div
      style={{
        border: `1px solid ${tokens.colorNeutralStroke2}`,
        borderRadius: tokens.borderRadiusMedium,
        background: tokens.colorNeutralBackground1,
        padding: 12,
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
        <Icon style={{ width: 16, height: 16, color: tokens.colorNeutralForeground2 }} />
        <strong style={{ fontSize: 12, color: tokens.colorNeutralForeground1 }}>{title}</strong>
        {tagText && (
          <Badge
            appearance={tagAppearance === 'brand' ? 'tint' : 'ghost'}
            color={tagAppearance === 'brand' ? 'brand' : undefined}
          >
            {tagText}
          </Badge>
        )}
      </div>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>{children}</div>
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// JSON edit pane — raw body editing for power users
// ──────────────────────────────────────────────────────────────
function JsonEditPane({
  tbl,
  text,
  error,
  themeMode,
  onChange,
}: {
  tbl: TableMeta;
  text: string;
  error: string | null;
  themeMode: ThemeMode;
  onChange: (text: string) => void;
}) {
  void tbl;
  return (
    <div style={{ maxWidth: 880, display: 'flex', flexDirection: 'column', gap: 10 }}>
      <MessageBar layout="multiline" intent="info">
        <MessageBarBody>
          Raw JSON edit mode — what you type here becomes the request body verbatim. Lookups use{' '}
          <code style={{ fontFamily: tokens.fontFamilyMonospace }}>
            {'"<col>@odata.bind": "/<entitySet>(<guid>)"'}
          </code>{' '}
          and multi-select choices are comma-separated integer strings. Switch back to{' '}
          <strong>Form</strong> to use the metadata-driven controls.
        </MessageBarBody>
      </MessageBar>
      {error && (
        <MessageBar layout="multiline" intent="error">
          <MessageBarBody>
            <strong>Invalid JSON.</strong> {error}
          </MessageBarBody>
        </MessageBar>
      )}
      <div style={{ position: 'relative' }}>
        {/* Monaco editor with `json` language — gives us syntax coloring,
            bracket matching, line numbers, and inline schema-style error
            squiggles for free. Same instance + worker setup as CodeView
            (the side-effect `loader.config({monaco})` there pre-loads
            the workers globally). Height is fixed at 360 so the surrounding
            "Form / JSON" toggle + info bar don't get pushed off-screen on
            small viewports; the editor's internal scrollbar handles tall
            bodies. */}
        <div
          style={{
            height: 360,
            border: `1px solid ${tokens.colorNeutralStroke2}`,
            borderRadius: tokens.borderRadiusMedium,
            overflow: 'hidden',
            background:
              themeMode === 'dark'
                ? '#1e1e1e' // matches Monaco vs-dark background to avoid border flash on mount
                : tokens.colorNeutralBackground1,
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
              renderLineHighlight: 'line',
              minimap: { enabled: false },
              scrollBeyondLastLine: false,
              fontFamily: 'Cascadia Mono, Consolas, monospace',
              fontSize: 13,
              tabSize: 2,
              insertSpaces: true,
              wordWrap: 'on',
              automaticLayout: true,
              padding: { top: 10, bottom: 10 },
              formatOnPaste: true,
              formatOnType: true,
              bracketPairColorization: { enabled: true },
              guides: { bracketPairs: true, indentation: true },
            }}
          />
        </div>
        <Tooltip content="Copy body JSON" relationship="label">
          <Button
            icon={<Copy20Regular />}
            size="small"
            appearance="subtle"
            // z-index above Monaco's chrome so the copy button stays clickable.
            style={{ position: 'absolute', top: 6, right: 6, zIndex: 2 }}
            onClick={() => navigator.clipboard?.writeText(text)}
          />
        </Tooltip>
      </div>
    </div>
  );
}

/**
 * Convert in-memory fieldValues into the wire-format JSON body — same encoding
 * rules as engine/urlBuilder.buildCreateBody but lives here too so the JSON
 * edit pane can stay self-contained (no engine import).
 */
/**
 * Resolve the `@odata.bind` property name for a lookup column. See
 * engine/urlBuilder.bindPropertyFor for the full rationale. We keep a
 * twin copy here because FieldSetEditor's JSON preview is intentionally
 * engine-free (no urlBuilder import — the JSON pane needs to render even
 * before the request would otherwise execute).
 */
function bindPropertyFor(
  parentTbl: TableMeta,
  columnLogical: string,
  targetEntity: string,
): string {
  const nav = parentTbl.navigationProperties.find(
    (n) =>
      n.cardinality === 'ManyToOne' &&
      n.referencingAttribute === columnLogical &&
      n.targetEntity === targetEntity,
  );
  return `${nav?.name ?? columnLogical}@odata.bind`;
}

function valuesToJsonBody(
  tbl: TableMeta | undefined,
  values: Record<string, CreateFieldValue>,
): Record<string, unknown> {
  if (!tbl) return {};
  const body: Record<string, unknown> = {};
  for (const [field, raw] of Object.entries(values)) {
    if (raw == null) continue;
    const col = findColumn(tbl, field);
    if (!col) continue;
    if (isLookupLike(col)) {
      const lk = raw as { id: string; targetEntity: string };
      if (!lk?.id) continue;
      const target = findTable(lk.targetEntity);
      if (!target) continue;
      // Polymorphic-safe — picks `customerid_account` / `regardingobjectid_account_task`
      // when the column is multi-target, or the bare `<col>` when single-target.
      const bindProp = bindPropertyFor(tbl, col.logicalName, lk.targetEntity);
      body[bindProp] = `/${target.entitySetName}(${lk.id})`;
      continue;
    }
    if (col.attributeType === 'MultiSelectPicklist') {
      const arr = raw as number[];
      if (!Array.isArray(arr) || arr.length === 0) continue;
      body[col.logicalName] = arr.join(',');
      continue;
    }
    if (typeof raw === 'string' && raw === '') continue;
    body[col.logicalName] = raw;
  }
  return body;
}

/**
 * Reverse of valuesToJsonBody — parse a user-edited JSON body back into
 * in-memory fieldValues. Lookup-bind entries become LookupFieldValue, multi-
 * select comma-strings become number[]. Unknown columns are dropped.
 */
function jsonBodyToValues(
  tbl: TableMeta | undefined,
  body: Record<string, unknown>,
): Record<string, CreateFieldValue> {
  if (!tbl) return {};
  const values: Record<string, CreateFieldValue> = {};
  for (const [k, v] of Object.entries(body)) {
    if (k.startsWith('@')) continue;
    if (k.endsWith('@odata.bind')) {
      // The key before `@odata.bind` is the property name. Two shapes:
      //   1. Bare attribute name (single-target lookups): `primarycontactid`
      //   2. Target-disambiguated nav-property (polymorphic): `customerid_account`,
      //      `regardingobjectid_account_task`
      // We try (1) first via findColumn; if no lookup column matches, fall
      // back to matching nav-properties by name. The nav's `referencingAttribute`
      // gives us the underlying column, and `targetEntity` gives the target.
      const propName = k.replace('@odata.bind', '');
      const m = typeof v === 'string' ? v.match(/\/([^/(]+)\(([^)]+)\)/) : null;
      if (!m) continue;
      const [, entitySet, id] = m;

      // (1) Try as a bare column name first.
      const directCol = findColumn(tbl, propName);
      if (directCol && isLookupLike(directCol)) {
        // Resolve target entity from the URL's entity-set segment by
        // walking the column's known targets. If none match, fall back
        // to the first target. (Most plain Lookups have one target.)
        const targetLogical =
          directCol.targets.find((t) => findTable(t)?.entitySetName === entitySet) ??
          directCol.targets[0];
        values[propName] = { id, targetEntity: targetLogical };
        continue;
      }

      // (2) Try as a disambiguated nav-property name.
      const nav = tbl.navigationProperties.find(
        (n) => n.cardinality === 'ManyToOne' && n.name === propName,
      );
      if (nav && nav.referencingAttribute) {
        values[nav.referencingAttribute] = {
          id,
          targetEntity: nav.targetEntity,
        };
        continue;
      }
      // Unknown property — drop. Keeps the parser tolerant of stray
      // annotations the user might have copied from elsewhere.
      continue;
    }
    const col = findColumn(tbl, k);
    if (!col) continue;
    if (col.attributeType === 'MultiSelectPicklist') {
      if (typeof v === 'string') {
        const arr = v
          .split(',')
          .map((n) => Number(n.trim()))
          .filter((n) => Number.isFinite(n));
        values[k] = arr;
      } else if (Array.isArray(v)) {
        values[k] = v.map(Number).filter((n) => Number.isFinite(n));
      }
      continue;
    }
    values[k] = v as CreateFieldValue;
  }
  return values;
}

/** Seed values for a new Create against a given table — populates required fields. */
export function seedRequiredFieldValues(table: TableMeta): Record<string, CreateFieldValue> {
  const out: Record<string, CreateFieldValue> = {};
  for (const c of table.columns) {
    if (!c.required) continue;
    if (isSystemManaged(c)) continue;
    out[c.logicalName] = defaultValueFor(c);
  }
  return out;
}
