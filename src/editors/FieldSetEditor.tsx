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

import { createContext, useCallback, useContext, useMemo, useState, useEffect } from 'react';
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
  MessageBarActions,
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
  makeStyles,
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
import { AddFieldPicker } from './AddFieldPicker';
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
  type BooleanColumnMeta,
  type LookupColumnMeta,
  type CustomerColumnMeta,
  type OwnerColumnMeta,
} from '../metadata/metadata';
import { useLookupRecords } from '../host/useLookupRecords';
import {
  decodeFieldBody,
  encodeFieldBody,
  isRequiredColumn,
  mustSupplyOnCreate,
  parseRawBody,
  suppliedColumnTest,
  type DecodedFieldBody,
  type UnmappedBodyKey,
} from '../engine/fieldBody';
import { useColumnDetail } from '../host/useColumnDetail';
import {
  behaviorHint,
  behaviorOf,
  isDateOnlyColumn,
  localZoneLabel,
  parseDateTimeValue,
  withPickedDate,
  withPickedTime,
} from './dateTimeValue';
// Monaco — same instance + worker setup as CodeView. The `loader.config`
// side-effect lives in src/monaco/monacoSetup.ts, which the MonacoEditor
// wrapper lazy-loads on the first editor render (#22).
import Editor from '../monaco/MonacoEditor';
import type { ThemeMode } from '../theme/theme';
import type { CreateFieldValue, LookupFieldValue } from '../state/writeState';
import type { RequestGroup } from '../registry/requestTypes';
import { isColumnInScope, scopedColumns, useAdvancedFindOnly } from '../state/metadataScope';
import { buildBodySchema } from '../engine/bodySchema';
import { useLiveEntities } from '../host/useLiveMetadata';
import { copyText } from '../shell/copy';

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
  /**
   * JSON-mode body text, owned by the caller's state. When `setRawBody` is
   * passed, JSON mode is controlled: the typed text is the request body
   * (sent as parsed — nulls, annotations and unknown keys included), the
   * editor never rewrites it, and switching back to Form maps what it can
   * and reports the rest. `null` = Form mode. Without `setRawBody` (nested
   * entity params, the PUT single-column editor) JSON edits are parsed into
   * `values` as you type and unmapped keys are listed as not sent.
   */
  rawBody?: string | null;
  setRawBody?: (next: string | null) => void;
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
  /**
   * HTTP verb named in the pane header. Defaults from `purpose`: POST for
   * create, PATCH for update / upsert. Update's single-column PUT passes
   * 'PUT'.
   */
  method?: 'POST' | 'PATCH' | 'PUT';
}

type Filter = 'all' | 'required' | 'set' | 'common';

// ──────────────────────────────────────────────────────────────
// Public component
// ──────────────────────────────────────────────────────────────
export function FieldSetEditor({
  table,
  values: valuesProp,
  setValues,
  nullFields: nullFieldsProp,
  setNullFields: setNullFieldsProp,
  rawBody,
  setRawBody,
  group = 'write',
  themeMode = 'light',
  purpose = 'create',
  method,
}: FieldSetEditorProps) {
  const verb = method ?? (purpose === 'create' ? 'POST' : 'PATCH');
  // Create (and Upsert, which may create) must carry every required column.
  // Update is a partial PATCH: columns left out of the body keep their
  // value, so an unset required column is fine — only clearing one to null
  // is a problem (#80).
  const requiresAllRequired = purpose !== 'update';
  // Internal fallback when the caller doesn't wire up explicit-null tracking.
  // Keeps the affordance hidden in those modes (the Set-null button checks
  // for the controlled handler before rendering itself).
  const storedNullFields = nullFieldsProp ?? [];
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
  // edits the body directly in Monaco. The text is only ever written by the
  // user (and seeded once on entering JSON mode) — never re-serialized from
  // `values`, which used to reformat it mid-edit (#91).
  //
  // Controlled (caller passes setRawBody): the text lives in the mode state
  // and IS the body; the mode is derived from it, so it survives this pane
  // unmounting. Uncontrolled: local text, parsed into `values` per keystroke.
  const jsonControlled = !!setRawBody;
  const [localMode, setLocalMode] = useState<'form' | 'json'>('form');
  const [localText, setLocalText] = useState<string>('');
  const bodyMode: 'form' | 'json' = jsonControlled
    ? rawBody != null
      ? 'json'
      : 'form'
    : localMode;
  const jsonText = jsonControlled ? (rawBody ?? '') : localText;
  // Pending JSON → Form switch that needs the user's call (invalid JSON, or
  // keys the form can't show). Cleared on every edit.
  const [switchIssue, setSwitchIssue] = useState<
    { kind: 'invalid'; error: string } | { kind: 'unmapped'; decoded: DecodedFieldBody } | null
  >(null);

  // Set / Update / Remove helpers. (`values` / `nullFields` are declared
  // below, after the JSON preview; these only run in Form mode.)
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

  // JSON mode: parse the typed body and preview what it maps to, so the
  // header counts and required warnings describe the JSON, not the stale
  // form values. In Form mode these are the caller's values as-is.
  const jsonParse = useMemo(
    () => (bodyMode === 'json' ? parseRawBody(jsonText) : null),
    [bodyMode, jsonText],
  );
  const decodeOpts = useMemo(() => {
    const ids = new Set(writable.map((c) => c.logicalName));
    return { editable: (k: string) => ids.has(k), nullSupported: nullFieldsSupported };
  }, [writable, nullFieldsSupported]);
  const jsonDecoded = useMemo(
    () => (tbl && jsonParse?.ok ? decodeFieldBody(tbl, jsonParse.body, decodeOpts) : null),
    [tbl, jsonParse, decodeOpts],
  );
  const values = jsonDecoded ? jsonDecoded.values : valuesProp;
  const nullFields = jsonDecoded ? jsonDecoded.nullFields : storedNullFields;

  const enterJson = () => {
    if (bodyMode === 'json' || !tbl) return;
    const text = JSON.stringify(encodeFieldBody(tbl, valuesProp, storedNullFields), null, 2);
    setSwitchIssue(null);
    if (setRawBody) setRawBody(text);
    else {
      setLocalText(text);
      setLocalMode('json');
    }
  };
  const leaveJson = (decoded: DecodedFieldBody | null) => {
    if (decoded) {
      setValues(decoded.values);
      if (setNullFieldsProp) setNullFieldsProp(decoded.nullFields);
    }
    setSwitchIssue(null);
    if (setRawBody) setRawBody(null);
    else setLocalMode('form');
  };
  const goForm = () => {
    if (bodyMode === 'form' || !tbl) return;
    // Uncontrolled: values already track the JSON keystroke by keystroke.
    if (!jsonControlled) return leaveJson(null);
    const p = parseRawBody(jsonText);
    if (!p.ok) return setSwitchIssue({ kind: 'invalid', error: p.error });
    const decoded = decodeFieldBody(tbl, p.body, decodeOpts);
    if (decoded.unmapped.length > 0) return setSwitchIssue({ kind: 'unmapped', decoded });
    leaveJson(decoded);
  };
  const onJsonChange = (t: string) => {
    setSwitchIssue(null);
    if (setRawBody) {
      setRawBody(t);
      return;
    }
    setLocalText(t);
    const p = parseRawBody(t);
    if (p.ok && tbl) setValues(decodeFieldBody(tbl, p.body, decodeOpts).values);
  };

  // Same required rule as the mode's Execute gate, in Form and JSON mode
  // (#124): only ApplicationRequired columns count; SystemRequired ones such
  // as ownerid are defaulted by Dataverse. They still sit at the top of the
  // form (see `visible`) so they can be overridden.
  const requiredFields = useMemo(() => writable.filter(mustSupplyOnCreate), [writable]);
  // A required column counts as set when what will be sent gives it a value:
  // the typed body in JSON mode, a non-empty form value otherwise — the
  // test the mode's Execute gate uses. A seeded but empty value (an Owner
  // lookup with no record) is dropped from the body, so it isn't set (#124).
  const isSupplied = useMemo(
    () =>
      tbl
        ? suppliedColumnTest(tbl, bodyMode === 'json' ? jsonText : null, valuesProp)
        : () => false,
    [tbl, bodyMode, jsonText, valuesProp],
  );
  // "Populated" = the column produces a body entry: a supplied value or an
  // explicit null. Seeded empty values (required columns pre-filled when the
  // table's metadata is cached) aren't sent, so they don't count (#210).
  // Drives the "N populated" badge and the per-row missing marker.
  const isPopulated = useCallback(
    (c: ColumnMeta) => isSupplied(c) || nullFields.includes(c.logicalName),
    [isSupplied, nullFields],
  );
  const populatedFields = useMemo(() => writable.filter(isPopulated), [writable, isPopulated]);
  // The working set: every column added to the form (seeded, picked with
  // "Add field", or being edited) plus explicit nulls. The Set filter (the
  // default view) shows these, so an empty row stays visible to be filled in.
  const inWorkingSet = useCallback(
    (c: ColumnMeta) => c.logicalName in values || nullFields.includes(c.logicalName),
    [values, nullFields],
  );
  const workingSetCount = useMemo(
    () => writable.filter(inWorkingSet).length,
    [writable, inWorkingSet],
  );

  // Required-field check: a column marked explicit-null still counts as
  // "missing" since the server will reject a null on a required column —
  // this matches Dataverse's runtime behavior and surfaces the issue.
  // Update only flags required columns explicitly set to null.
  const missingRequired = useMemo(
    () =>
      requiredFields.filter(
        (c) => (requiresAllRequired && !isSupplied(c)) || nullFields.includes(c.logicalName),
      ),
    [requiredFields, isSupplied, nullFields, requiresAllRequired],
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
    const requiredCols = writable.filter(isRequiredColumn);
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

    // Same set the "Required · N" chip counts (#199): the columns a create
    // must supply. SystemRequired ones (ownerid, statecode) are defaulted
    // by Dataverse and stay out of this view.
    if (filter === 'required') pool = pool.filter(mustSupplyOnCreate);
    if (filter === 'set') pool = pool.filter(inWorkingSet);
    if (filter === 'common') {
      pool = pool.filter(
        (c) =>
          isRequiredColumn(c) ||
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
  }, [writable, values, nullFields, filter, search, tbl, inWorkingSet]);

  // Pool of fields available to add (writable, not currently visible). It
  // follows "Advanced Find only" (#44); required and filled-in fields always
  // show, since they come from `writable`.
  const advancedFindOnly = useAdvancedFindOnly();
  const addable = useMemo(() => {
    const visibleIds = new Set(visible.map((c) => c.logicalName));
    return scopedColumns(writable, advancedFindOnly).filter((c) => !visibleIds.has(c.logicalName));
  }, [writable, visible, advancedFindOnly]);

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
            Pick the columns this <code>{verb}</code> should set.{' '}
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
            aria-label="Filter fields"
            placeholder="Filter fields…"
            value={search}
            onChange={(_, d) => setSearch(d.value)}
            size="small"
            style={{ width: 220 }}
          />
        )}
        <Badge appearance="ghost">{populatedFields.length} populated</Badge>
        {requiresAllRequired && requiredFields.length > 0 && (
          <Badge appearance="tint" color={missingRequired.length ? 'danger' : 'success'}>
            {requiredFields.length - missingRequired.length}/{requiredFields.length} required
          </Badge>
        )}
        {/* Form ↔ JSON segmented toggle. Uses the shared SegmentedToggle
            primitive so it picks up the same hairline + radius treatment
            as FilterEditor (AND/OR) and OrderbyEditor (Asc/Desc). */}
        <SegmentedToggle ariaLabel="Body editor mode">
          <ToggleButton checked={bodyMode === 'form'} icon={<FormNew20Regular />} onClick={goForm}>
            Form
          </ToggleButton>
          <ToggleButton checked={bodyMode === 'json'} icon={<Code20Regular />} onClick={enterJson}>
            JSON
          </ToggleButton>
        </SegmentedToggle>
      </PaneHead>

      {missingRequired.length > 0 &&
        (requiresAllRequired ? (
          <MessageBar layout="multiline" intent="warning" style={{ marginBottom: 12 }}>
            <MessageBarBody>
              <MessageBarTitle>
                {missingRequired.length} required field
                {missingRequired.length === 1 ? ' is' : 's are'} unset.
              </MessageBarTitle>
              Dataverse will reject the{' '}
              {purpose === 'upsert' ? 'upsert when it creates the row' : 'create'} with{' '}
              <code>0x80040217 · A required field is missing</code> unless every required column has
              a value. <strong>Missing:</strong>{' '}
              {missingRequired.map((c) => (
                <code key={c.logicalName} style={{ marginRight: 6 }}>
                  {c.logicalName}
                </code>
              ))}
            </MessageBarBody>
          </MessageBar>
        ) : (
          <MessageBar layout="multiline" intent="warning" style={{ marginBottom: 12 }}>
            <MessageBarBody>
              <MessageBarTitle>
                {missingRequired.length} required field
                {missingRequired.length === 1 ? ' is' : 's are'} set to null.
              </MessageBarTitle>
              Dataverse may reject the update for clearing a required column. Columns left out of
              the body keep their current value. <strong>Cleared:</strong>{' '}
              {missingRequired.map((c) => (
                <code key={c.logicalName} style={{ marginRight: 6 }}>
                  {c.logicalName}
                </code>
              ))}
            </MessageBarBody>
          </MessageBar>
        ))}

      {bodyMode === 'form' ? (
        <>
          {/* Filter chip row */}
          <div style={{ display: 'flex', gap: 6, marginBottom: 14, flexWrap: 'wrap' }}>
            <FilterChip
              label="Set"
              count={workingSetCount}
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
                    populated={isPopulated(col)}
                    isNull={isNull(col.logicalName)}
                    nullSupported={nullFieldsSupported}
                    requiredOnWrite={requiresAllRequired}
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
        // JSON edit mode — Monaco editor with json language. The text is
        // never rewritten by us; see the bodyMode comment above.
        <JsonEditPane
          text={jsonText}
          error={jsonParse && !jsonParse.ok ? jsonParse.error : null}
          unmapped={jsonDecoded?.unmapped ?? []}
          verbatim={jsonControlled}
          switchIssue={switchIssue}
          onDiscard={() => leaveJson(null)}
          onSwitchAnyway={() => {
            if (switchIssue?.kind === 'unmapped') leaveJson(switchIssue.decoded);
          }}
          onStay={() => setSwitchIssue(null)}
          themeMode={themeMode}
          onChange={onJsonChange}
          table={tbl}
          writable={writable}
          nullSupported={nullFieldsSupported}
          advancedFindOnly={advancedFindOnly}
        />
      )}

      {/* Add field — only if there's something left to add */}
      {bodyMode === 'form' && addable.length > 0 && (
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
              <AddFieldPicker
                addable={addable}
                onPick={(logicalName) => {
                  const col = findColumn(tbl, logicalName);
                  if (!col) return;
                  setField(col.logicalName, defaultValueFor(col));
                  setAddPickerOpen(false);
                }}
              />
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

// Row layout has to hold at PPTB's default tool width (~840 px), where the
// value cell is only ~250 px wide (#112). The label column gives up width
// first (35 % of the row, 120–220 px), the value input keeps a 160 px floor,
// and the "Send as null" switch wraps onto its own line below the input
// instead of overlapping it. Fluent's Combobox / Dropdown / DatePicker carry
// intrinsic min-widths (250 px and up) that would otherwise push past the
// cell, so they're released to the cell width here.
const useFieldRowStyles = makeStyles({
  valueLine: {
    display: 'flex',
    flexWrap: 'wrap',
    alignItems: 'center',
    columnGap: '8px',
    rowGap: '4px',
    minWidth: 0,
  },
  valueInput: {
    flex: '1 1 160px',
    minWidth: 0,
    '& .fui-Input, & .fui-Combobox, & .fui-Dropdown, & .fui-SpinButton, & .fui-Textarea, & .fui-TagPickerControl':
      {
        minWidth: 0,
        maxWidth: '100%',
      },
    // The lookup search box sets an inline 240 px min-width.
    '& .fui-Combobox': {
      minWidth: '0 !important',
    },
  },
  nullSwitch: {
    flexShrink: 0,
  },
});

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
  requiredOnWrite,
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
  /** True when the body must carry every required column (create / upsert).
   *  False for update, where an unset required column is left unchanged. */
  requiredOnWrite: boolean;
  onChange: (v: CreateFieldValue) => void;
  onRemove: () => void;
  onSetNull: () => void;
  onClearNull: () => void;
}) {
  const s = useStudioStyles();
  const isMissing = isRequiredColumn(col) && (requiredOnWrite ? !populated : isNull);

  const rs = useFieldRowStyles();
  return (
    <div
      className={mergeClasses(s.inlineCard)}
      style={{
        padding: '10px 12px',
        display: 'grid',
        gridTemplateColumns: 'minmax(120px, min(220px, 35%)) minmax(0, 1fr) 30px',
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
          {isRequiredColumn(col) && (
            <Tooltip
              content={
                requiredOnWrite
                  ? 'Required — Dataverse rejects the create when this is empty.'
                  : 'Required — leave it out to keep its value; clearing it to null may be rejected.'
              }
              relationship="description"
            >
              <Badge appearance="tint" color="danger" size="small">
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
        <div className={rs.valueLine}>
          <div className={rs.valueInput}>
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
                  size="small"
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
              className={rs.nullSwitch}
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
      {!isRequiredColumn(col) ? (
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
// Accessible name for value controls that don't get `col` (#120).
const FieldLabelContext = createContext('Field value');

function FieldInput(props: Parameters<typeof FieldInputBody>[0]) {
  return (
    <FieldLabelContext.Provider value={props.col.displayName}>
      <FieldInputBody {...props} />
    </FieldLabelContext.Provider>
  );
}

function FieldInputBody({
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
  const detail = useColumnDetail(table.logicalName, col.logicalName);
  const input = renderFieldInput(col, value, onChange);
  // Surface a failed option / detail load instead of an empty list (#199).
  if (!detail.error) return input;
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
      {input}
      <Caption1 style={{ color: tokens.colorPaletteRedForeground1 }}>
        Couldn’t load column details: {detail.error}
      </Caption1>
    </div>
  );
}

function renderFieldInput(
  col: ColumnMeta,
  value: CreateFieldValue | undefined,
  onChange: (v: CreateFieldValue) => void,
) {
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
          aria-label={col.displayName}
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
        aria-label={col.displayName}
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
      aria-label={col.displayName}
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
        aria-label={col.displayName}
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
        aria-label={col.displayName}
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
  const fieldLabel = useContext(FieldLabelContext);
  return (
    <SpinButton
      aria-label={fieldLabel}
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
  // Parsing / serializing is per behavior — see editors/dateTimeValue.ts.
  // `current`'s LOCAL fields are what the pickers show.
  const isDateOnly = isDateOnlyColumn(col);
  const current = parseDateTimeValue(value, col);
  const hint = behaviorHint(col, current);

  const writeDate = (d: Date | null | undefined) => {
    if (!d) {
      onChange('');
      return;
    }
    // Keep the time already picked when the date changes (as stored, #337).
    onChange(withPickedDate(value, d, col));
  };
  const writeTime = (d: Date | null | undefined) => {
    if (!d) return;
    // Null when it's the time already shown: the picker reports on blur too.
    const next = withPickedTime(value, d, col);
    if (next !== null) onChange(next);
  };

  // What's on the wire — shown under the pickers so the zone is never a
  // guess (#92).
  const sent = value ? (
    <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
      Sends <code style={{ fontFamily: tokens.fontFamilyMonospace }}>{value}</code>
      {behaviorOf(col) === 'UserLocal' && (
        <> (UTC) · shown in {localZoneLabel(current ?? undefined)}</>
      )}
    </Caption1>
  ) : null;

  if (isDateOnly) {
    return (
      <span style={{ display: 'inline-flex', flexDirection: 'column', gap: 2 }}>
        <Tooltip content={`Behavior: ${behaviorOf(col)} — ${hint}`} relationship="description">
          <DatePicker
            size="small"
            value={current}
            onSelectDate={writeDate}
            placeholder="Pick a date…"
            formatDate={(d) => (d ? d.toLocaleDateString() : '')}
            allowTextInput
          />
        </Tooltip>
        {sent}
      </span>
    );
  }
  return (
    <span style={{ display: 'inline-flex', flexDirection: 'column', gap: 2 }}>
      <Tooltip content={`Behavior: ${behaviorOf(col)} — ${hint}`} relationship="description">
        <span style={{ display: 'inline-flex', gap: 6, alignItems: 'center', flexWrap: 'wrap' }}>
          <DatePicker
            size="small"
            value={current}
            onSelectDate={writeDate}
            placeholder="Date…"
            formatDate={(d) => (d ? d.toLocaleDateString() : '')}
            style={{ minWidth: 130 }}
            allowTextInput
          />
          <TimePicker
            size="small"
            freeform
            selectedTime={current}
            onTimeChange={(_, d) => writeTime(d.selectedTime)}
            placeholder="Time…"
            style={{ minWidth: 100 }}
            hourCycle="h23"
          />
        </span>
      </Tooltip>
      {sent}
    </span>
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
      aria-label={col.displayName}
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
  // eslint-disable-next-line react-hooks/exhaustive-deps -- columnOptions() is a new array for Boolean / option-less columns; these small memos just recompute then
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
        <Combobox
          aria-label={col.displayName}
          size="small"
          disabled
          placeholder="Loading…"
          style={{ minWidth: 240 }}
        >
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
          aria-label={`${col.displayName} (target table)`}
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
        aria-label={col.displayName}
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
  const fieldLabel = useContext(FieldLabelContext);
  const v = value ?? '';
  const valid = !v || /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(v);
  return (
    <Input
      aria-label={fieldLabel}
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
    case 'Picklist': {
      // No default (-1 / missing) leaves the choice unset — not sent until
      // the user picks one. Never invent 0 or the first option (#203).
      const d = validDefaultOption(col);
      return d ?? null;
    }
    case 'State':
      return col.options[0]?.value ?? 0;
    case 'Status':
      return col.options[0]?.value ?? 1;
    case 'EntityName':
      return col.options[0]?.value ?? 0;
    case 'MultiSelectPicklist': {
      const d = validDefaultOption(col);
      return d != null ? [d] : [];
    }
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

/**
 * The column's metadata default (`DefaultFormValue`) when it names one of
 * its options. -1 means "no default"; a value the option list doesn't
 * contain is ignored too. Undefined when there's nothing valid to pre-fill.
 */
function validDefaultOption(col: {
  defaultFormValue?: number;
  options: { value: number }[];
}): number | undefined {
  const d = col.defaultFormValue;
  if (typeof d !== 'number' || d === -1) return undefined;
  if (col.options.length > 0 && !col.options.some((o) => o.value === d)) return undefined;
  return d;
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
  text,
  error,
  unmapped,
  verbatim,
  switchIssue,
  onDiscard,
  onSwitchAnyway,
  onStay,
  themeMode,
  onChange,
  table,
  writable,
  nullSupported,
  advancedFindOnly,
}: {
  text: string;
  error: string | null;
  /** Keys in the typed body the Form view can't represent. */
  unmapped: UnmappedBodyKey[];
  /** True when the text is the request body (controlled JSON mode). */
  verbatim: boolean;
  switchIssue: { kind: 'invalid'; error: string } | { kind: 'unmapped' } | null;
  onDiscard: () => void;
  onSwitchAnyway: () => void;
  onStay: () => void;
  themeMode: ThemeMode;
  onChange: (text: string) => void;
  table: TableMeta;
  /** The columns the form can set; the schema offers the same (#268). */
  writable: ColumnMeta[];
  nullSupported: boolean;
  advancedFindOnly: boolean;
}) {
  // Completions, hover and type checks for the body (#268). Lookup targets'
  // entity sets come from their metadata when loaded, else the table list.
  const { entities } = useLiveEntities();
  const jsonSchema = useMemo(() => {
    const sets = new Map(entities.map((e) => [e.logicalName, e.entitySetName]));
    return buildBodySchema(table, writable, {
      nullSupported,
      suggest: (c) => isColumnInScope(c, advancedFindOnly),
      entitySetOf: (t) => findTable(t)?.entitySetName ?? sets.get(t),
    });
  }, [table, writable, nullSupported, advancedFindOnly, entities]);

  const keyList = (
    <ul style={{ margin: '4px 0 0', paddingLeft: 18 }}>
      {unmapped.map((u) => (
        <li key={u.key}>
          <code style={{ fontFamily: tokens.fontFamilyMonospace }}>{u.key}</code> — {u.reason}
        </li>
      ))}
    </ul>
  );
  return (
    <div style={{ maxWidth: 880, display: 'flex', flexDirection: 'column', gap: 10 }}>
      <MessageBar layout="multiline" intent="info">
        <MessageBarBody>
          {verbatim ? (
            <>
              Raw JSON edit mode — the request body is exactly this JSON, including{' '}
              <code style={{ fontFamily: tokens.fontFamilyMonospace }}>null</code> values and keys
              the form doesn&apos;t know.{' '}
            </>
          ) : (
            <>
              Raw JSON edit mode — edits are applied to the fields as you type. Keys that don&apos;t
              map to a field aren&apos;t sent.{' '}
            </>
          )}
          Lookups use{' '}
          <code style={{ fontFamily: tokens.fontFamilyMonospace }}>
            {'"<col>@odata.bind": "/<entitySet>(<guid>)"'}
          </code>{' '}
          and multi-select choices are comma-separated integer strings. Switch back to{' '}
          <strong>Form</strong> to use the metadata-driven controls.
        </MessageBarBody>
      </MessageBar>
      {switchIssue?.kind === 'invalid' && (
        <MessageBar layout="multiline" intent="error">
          <MessageBarBody>
            <MessageBarTitle>
              Can&apos;t switch to Form — the JSON doesn&apos;t parse.
            </MessageBarTitle>
            {switchIssue.error} Fix it, or discard your JSON edits and go back to the form as it
            was.
          </MessageBarBody>
          <MessageBarActions>
            <Button size="small" onClick={onDiscard}>
              Discard JSON edits
            </Button>
            <Button size="small" appearance="primary" onClick={onStay}>
              Keep editing
            </Button>
          </MessageBarActions>
        </MessageBar>
      )}
      {switchIssue?.kind === 'unmapped' && (
        <MessageBar layout="multiline" intent="warning">
          <MessageBarBody>
            <MessageBarTitle>
              {unmapped.length} key{unmapped.length === 1 ? '' : 's'} can&apos;t be shown in the
              Form view.
            </MessageBarTitle>
            Switching to Form drops {unmapped.length === 1 ? 'it' : 'them'} from the body:
            {keyList}
          </MessageBarBody>
          <MessageBarActions>
            <Button size="small" onClick={onSwitchAnyway}>
              Switch and drop {unmapped.length === 1 ? 'it' : 'them'}
            </Button>
            <Button size="small" appearance="primary" onClick={onStay}>
              Stay in JSON
            </Button>
          </MessageBarActions>
        </MessageBar>
      )}
      {error && !switchIssue && (
        <MessageBar layout="multiline" intent="error">
          <MessageBarBody>
            <strong>Invalid JSON.</strong> {error}
            {verbatim && ' Execute is disabled until the body parses.'}
          </MessageBarBody>
        </MessageBar>
      )}
      {!error && !switchIssue && unmapped.length > 0 && (
        <MessageBar layout="multiline" intent={verbatim ? 'info' : 'warning'}>
          <MessageBarBody>
            {verbatim ? (
              <>
                Sent as typed, but the Form view can&apos;t show{' '}
                {unmapped.length === 1 ? 'this key' : 'these keys'}:
              </>
            ) : (
              <>
                Not sent — {unmapped.length === 1 ? 'this key doesn’t' : 'these keys don’t'} map to
                a field:
              </>
            )}
            {keyList}
          </MessageBarBody>
        </MessageBar>
      )}
      <div style={{ position: 'relative' }}>
        {/* Monaco editor with `json` language — gives us syntax coloring,
            bracket matching, line numbers, and inline schema-style error
            squiggles for free. Same lazily loaded instance + worker setup
            as CodeView (src/monaco/monacoSetup.ts). Height is fixed at 360 so the surrounding
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
            jsonSchema={jsonSchema}
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
            onClick={() => void copyText(text, 'the body JSON')}
          />
        </Tooltip>
      </div>
    </div>
  );
}

/**
 * Seed values for a new Create against a given table — populates the
 * columns a create must supply (ApplicationRequired). SystemRequired
 * columns such as ownerid / statecode are defaulted by Dataverse, so they
 * aren't seeded: a seeded `statecode: 0` was sent without ever showing in
 * the form (#199).
 */
export function seedRequiredFieldValues(table: TableMeta): Record<string, CreateFieldValue> {
  const out: Record<string, CreateFieldValue> = {};
  for (const c of table.columns) {
    if (!mustSupplyOnCreate(c)) continue;
    if (isSystemManaged(c)) continue;
    out[c.logicalName] = defaultValueFor(c);
  }
  return out;
}
