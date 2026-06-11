// MergeFieldDiff — side-by-side field-level comparison + three-way choice.
//
// Modeled on the Power Apps "Merge duplicate records" UX:
//   https://learn.microsoft.com/en-us/power-apps/user/merge-duplicate-records
//
// For every column that differs between Target and Subordinate, the user
// picks one of:
//   • Keep Target          — default. Field NOT emitted in UpdateContent.
//   • Use Subordinate      — Field emitted with Subordinate's value.
//   • Custom override      — Field emitted with a free-form value the user types.
//
// The header summarizes how many overrides are pending (= the size of
// UpdateContent on the wire).
//
// Per docs (merge-entity-using-web-api):
//   "Merging moves any useful data from the Subordinate record to the Target
//    record. Any existing data in the Target record aren't overwritten."
// → Empty-Target/populated-Subordinate fields are implicitly filled by the
//   server; we surface them with a `auto-fill` annotation so the user knows
//   what's happening, but they DON'T need explicit overrides.

import { useMemo, useState } from 'react';
import {
  Caption1,
  Badge,
  Tooltip,
  Input,
  Textarea,
  SpinButton,
  Switch,
  Combobox,
  Option,
  Button,
  MessageBar,
  MessageBarBody,
  RadioGroup,
  Radio,
  tokens,
  mergeClasses,
} from '@fluentui/react-components';
import {
  ArrowLeft20Regular,
  Search20Regular,
  Eye20Regular,
  EyeOff20Regular,
  ChevronRight20Regular,
  BranchFork20Filled,
} from '@fluentui/react-icons';
import { useStudioStyles } from '../primitives/styles';
import { PaneHead } from './PaneHead';
import {
  findTable,
  findColumn,
  columnOptions,
  isLookupLike,
  isCompanionLogicalReadOnly,
  type ColumnMeta,
  type TableMeta,
  type LookupColumnMeta,
  type CustomerColumnMeta,
  type OwnerColumnMeta,
} from '../mock/metadata';
// Live snapshots are passed in as props (see MergeFieldDiffProps.target /
// .sub). We no longer read from the mock store.
import type { CreateFieldValue, LookupFieldValue, MergeFieldChoice } from '../state/writeState';
import { LookupFieldInput } from './FieldSetEditor';
import { useColumnDetail } from '../host/useColumnDetail';

// ──────────────────────────────────────────────────────────────
// Columns we hide from the merge diff (audit + system-managed + binary)
// ──────────────────────────────────────────────────────────────
const HIDDEN = new Set([
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
  'statecode',
  'statuscode', // merge owns these — Subordinate gets deactivated
]);

const isHidden = (c: ColumnMeta): boolean =>
  HIDDEN.has(c.logicalName) ||
  c.attributeType === 'File' ||
  c.attributeType === 'Image' ||
  (c.attributeType === 'Uniqueidentifier' && c.isPrimaryKey === true);

// ──────────────────────────────────────────────────────────────
// Component
// ──────────────────────────────────────────────────────────────
export interface MergeFieldDiffProps {
  table: string;
  /** Pre-fetched Target row. Owner mode (MergeMode) fetches via dataverseAPI
   *  and passes it in. `null` while loading OR if not yet picked. */
  target: Record<string, unknown> | null;
  /** Pre-fetched Subordinate row. Same lifecycle as target. */
  sub: Record<string, unknown> | null;
  fieldChoices: Record<string, MergeFieldChoice>;
  customValues: Record<string, CreateFieldValue>;
  setFieldChoice: (field: string, choice: MergeFieldChoice) => void;
  setCustomValue: (field: string, v: CreateFieldValue) => void;
}

export function MergeFieldDiff(props: MergeFieldDiffProps) {
  const { table, target, sub, fieldChoices, customValues, setFieldChoice, setCustomValue } = props;
  const tbl = findTable(table);
  const [search, setSearch] = useState('');
  const [showOnlyDiff, setShowOnlyDiff] = useState(true);

  // Classify each column row
  interface Row {
    col: ColumnMeta;
    targetVal: unknown;
    subVal: unknown;
    /** 'differs' = both have non-empty different values · 'auto-fill' = target empty, sub has value · 'sub-empty' = sub empty, target has value · 'same' = same · 'both-empty' */
    state: 'differs' | 'auto-fill' | 'sub-empty' | 'same' | 'both-empty';
  }

  const rows = useMemo<Row[]>(() => {
    if (!tbl || !target || !sub) return [];
    // Merge composes `UpdateContent` into the body, which is an update
    // semantically. Filter out:
    //   • System-hidden audit / state columns (HIDDEN set above)
    //   • Columns where IsValidForUpdate=false — read-only platform fields
    //     (createdon, audit-internal, computed/rollup outputs, the `_base`
    //     Money partner columns, etc.). The server will reject these in
    //     UpdateContent so showing them in the picker is misleading.
    //   • Companion `*name` / `*yominame` logical read-only columns —
    //     same convention as the read pickers (Select/Filter/Orderby).
    const visible = tbl.columns.filter(
      (c) => !isHidden(c) && c.isValidForUpdate !== false && !isCompanionLogicalReadOnly(c),
    );
    const out: Row[] = visible.map((col) => {
      const t = target[col.logicalName];
      const u = sub[col.logicalName];
      const tEmpty = t == null || t === '';
      const uEmpty = u == null || u === '';
      let state: Row['state'];
      if (tEmpty && uEmpty) state = 'both-empty';
      else if (tEmpty && !uEmpty) state = 'auto-fill';
      else if (!tEmpty && uEmpty) state = 'sub-empty';
      else state = String(t) === String(u) ? 'same' : 'differs';
      return { col, targetVal: t, subVal: u, state };
    });
    return out;
  }, [tbl, target, sub]);

  const filtered = useMemo(() => {
    let pool = rows;
    if (showOnlyDiff) pool = pool.filter((r) => r.state === 'differs' || r.state === 'auto-fill');
    if (search) {
      const q = search.toLowerCase();
      pool = pool.filter(
        (r) =>
          r.col.displayName.toLowerCase().includes(q) ||
          r.col.logicalName.toLowerCase().includes(q),
      );
    }
    return pool;
  }, [rows, showOnlyDiff, search]);

  const counts = useMemo(() => {
    const overrides = Object.values(fieldChoices).filter((c) => c !== 'target').length;
    const differs = rows.filter((r) => r.state === 'differs').length;
    const autoFill = rows.filter((r) => r.state === 'auto-fill').length;
    return { overrides, differs, autoFill };
  }, [fieldChoices, rows]);

  if (!tbl) {
    return (
      <MessageBar layout="multiline" intent="error">
        <MessageBarBody>
          Unknown table <code>{table}</code>.
        </MessageBarBody>
      </MessageBar>
    );
  }
  if (!target || !sub) {
    return (
      <div>
        <PaneHead
          icon={BranchFork20Filled}
          title="Field comparison"
          sub="Pick both Target and Subordinate to see a side-by-side diff."
          group="write"
        />
        <MessageBar layout="multiline" intent="info" style={{ maxWidth: 720 }}>
          <MessageBarBody>
            {!target && !sub
              ? 'Pick a Target (the record that survives) and a Subordinate (the duplicate to merge away) to start.'
              : !target
                ? 'Pick a Target record on the left rail.'
                : 'Pick a Subordinate record on the left rail.'}
          </MessageBarBody>
        </MessageBar>
      </div>
    );
  }

  return (
    <div>
      <PaneHead
        icon={BranchFork20Filled}
        title="Field comparison"
        sub={
          <span>
            For each differing field, pick what the merged row should carry — the Target's value,
            the Subordinate's value, or your own.{' '}
            <strong>Empty Target fields get auto-filled</strong> from the Subordinate by the server.
          </span>
        }
        group="write"
      >
        <Input
          contentBefore={<Search20Regular />}
          placeholder="Filter fields…"
          value={search}
          onChange={(_, d) => setSearch(d.value)}
          size="small"
          style={{ width: 220 }}
        />
        <Tooltip
          content={
            showOnlyDiff ? 'Showing only differing / auto-fill rows' : 'Showing every column'
          }
          relationship="description"
        >
          <Button
            size="small"
            appearance={showOnlyDiff ? 'primary' : 'outline'}
            icon={showOnlyDiff ? <Eye20Regular /> : <EyeOff20Regular />}
            onClick={() => setShowOnlyDiff((v) => !v)}
          >
            {showOnlyDiff ? 'Diff only' : 'All fields'}
          </Button>
        </Tooltip>
        <Badge appearance="tint" color={counts.overrides > 0 ? 'brand' : 'subtle'}>
          {counts.overrides} override{counts.overrides === 1 ? '' : 's'}
        </Badge>
        <Badge appearance="ghost">
          {counts.differs} differ · {counts.autoFill} auto-fill
        </Badge>
      </PaneHead>

      <div style={{ maxWidth: 1200, overflowX: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0 }}>
          <thead>
            <tr>
              <Th style={{ width: '20%' }}>Field</Th>
              <Th tone="target">Target (survives)</Th>
              <Th tone="sub">Subordinate</Th>
              <Th style={{ width: '34%' }}>Final value</Th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr>
                <td
                  colSpan={4}
                  style={{
                    padding: '24px 8px',
                    textAlign: 'center',
                    color: tokens.colorNeutralForeground3,
                    fontStyle: 'italic',
                  }}
                >
                  {showOnlyDiff
                    ? 'No differing fields. Toggle "All fields" to see the full row.'
                    : 'No fields match the search.'}
                </td>
              </tr>
            )}
            {filtered.map((r) => (
              <MergeRow
                key={r.col.logicalName}
                tbl={tbl}
                row={r}
                choice={fieldChoices[r.col.logicalName] ?? 'target'}
                customValue={customValues[r.col.logicalName]}
                onChoice={(c) => setFieldChoice(r.col.logicalName, c)}
                onCustom={(v) => setCustomValue(r.col.logicalName, v)}
              />
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// One field-comparison row
// ──────────────────────────────────────────────────────────────
function MergeRow({
  tbl,
  row,
  choice,
  customValue,
  onChoice,
  onCustom,
}: {
  tbl: TableMeta;
  row: {
    col: ColumnMeta;
    targetVal: unknown;
    subVal: unknown;
    state: 'differs' | 'auto-fill' | 'sub-empty' | 'same' | 'both-empty';
  };
  choice: MergeFieldChoice;
  customValue: CreateFieldValue | undefined;
  onChoice: (c: MergeFieldChoice) => void;
  onCustom: (v: CreateFieldValue) => void;
}) {
  const s = useStudioStyles();
  const { col, targetVal, subVal, state } = row;
  const isDiff = state === 'differs';
  const isAutoFill = state === 'auto-fill';

  // Subordinate value is unavailable to pick when subVal is empty
  const subSelectable = !(subVal == null || subVal === '');

  return (
    <tr>
      <Td>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontWeight: 500, color: tokens.colorNeutralForeground1, fontSize: 12 }}>
            {col.displayName}
          </span>
          <Caption1
            style={{
              color: tokens.colorNeutralForeground3,
              fontFamily: tokens.fontFamilyMonospace,
              fontSize: 10,
            }}
          >
            {col.logicalName} ·{' '}
            <span style={{ color: tokens.colorBrandForeground2 }}>{col.attributeType}</span>
          </Caption1>
          {isAutoFill && (
            <Tooltip
              content="Target is empty — server will auto-fill from Subordinate. No override needed."
              relationship="description"
            >
              <Badge
                appearance="tint"
                color="success"
                size="extra-small"
                style={{ alignSelf: 'flex-start' }}
              >
                auto-fill
              </Badge>
            </Tooltip>
          )}
        </div>
      </Td>
      <Td tone="target">
        <ValueCell col={col} value={targetVal} tbl={tbl} />
      </Td>
      <Td tone="sub">
        <ValueCell col={col} value={subVal} tbl={tbl} />
      </Td>
      <Td>
        {/* Choice picker */}
        <div className={mergeClasses(s.inlineCard)} style={{ padding: '6px 8px' }}>
          <RadioGroup
            layout="horizontal"
            value={choice}
            onChange={(_, d) => onChoice(d.value as MergeFieldChoice)}
          >
            <Radio
              value="target"
              label={
                <span style={{ fontSize: 11 }}>
                  Target
                  {!isDiff && !isAutoFill && (
                    <span style={{ color: tokens.colorNeutralForeground3 }}> · default</span>
                  )}
                </span>
              }
            />
            <Radio
              value="subordinate"
              disabled={!subSelectable}
              label={
                <span style={{ fontSize: 11 }}>
                  Subordinate
                  {!subSelectable && (
                    <span style={{ color: tokens.colorNeutralForeground3 }}> · empty</span>
                  )}
                </span>
              }
            />
            <Radio value="custom" label={<span style={{ fontSize: 11 }}>Custom</span>} />
          </RadioGroup>

          {choice === 'custom' && (
            <div style={{ marginTop: 6 }}>
              <CustomValueInput
                col={col}
                value={customValue}
                onChange={onCustom}
                fallback={String(targetVal ?? '')}
                parentTable={tbl.logicalName}
              />
            </div>
          )}

          {choice === 'subordinate' && (
            <div style={{ marginTop: 6, display: 'flex', alignItems: 'center', gap: 6 }}>
              <ArrowLeft20Regular
                style={{ width: 14, height: 14, color: tokens.colorBrandForeground1 }}
              />
              <Caption1 style={{ color: tokens.colorBrandForeground2, fontWeight: 500 }}>
                Will overwrite Target with: <strong>{formatValue(col, subVal)}</strong>
              </Caption1>
            </div>
          )}
        </div>
      </Td>
    </tr>
  );
}

// ──────────────────────────────────────────────────────────────
// Header cell + body cell helpers
// ──────────────────────────────────────────────────────────────
function Th({
  children,
  style,
  tone,
}: {
  children: React.ReactNode;
  style?: React.CSSProperties;
  tone?: 'target' | 'sub';
}) {
  const accent =
    tone === 'target'
      ? { background: 'rgba(31,127,56,.08)', color: tokens.colorPaletteGreenForeground1 }
      : tone === 'sub'
        ? { background: 'rgba(242,176,80,.08)', color: tokens.colorPaletteDarkOrangeForeground1 }
        : {};
  return (
    <th
      style={{
        textAlign: 'left',
        padding: '8px 10px',
        fontSize: 11,
        fontWeight: 600,
        color: tokens.colorNeutralForeground2,
        borderBottom: `2px solid ${tokens.colorNeutralStroke2}`,
        whiteSpace: 'nowrap',
        ...accent,
        ...style,
      }}
    >
      {children}
    </th>
  );
}

function Td({ children, tone }: { children: React.ReactNode; tone?: 'target' | 'sub' }) {
  const accent =
    tone === 'target'
      ? {
          background: 'rgba(31,127,56,.025)',
          borderLeft: `2px solid ${tokens.colorPaletteGreenBorderActive}`,
        }
      : tone === 'sub'
        ? {
            background: 'rgba(242,176,80,.025)',
            borderLeft: `2px solid ${tokens.colorPaletteDarkOrangeBorderActive}`,
          }
        : {};
  return (
    <td
      style={{
        padding: '8px 10px',
        verticalAlign: 'top',
        borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
        fontSize: 12,
        ...accent,
      }}
    >
      {children}
    </td>
  );
}

function ValueCell({ col, value, tbl }: { col: ColumnMeta; value: unknown; tbl: TableMeta }) {
  void tbl;
  if (value == null || value === '') {
    return (
      <span style={{ color: tokens.colorNeutralForeground3, fontStyle: 'italic', fontSize: 11 }}>
        (empty)
      </span>
    );
  }
  const rendered = formatValue(col, value);
  return (
    <span
      style={{
        fontFamily:
          col.attributeType === 'Uniqueidentifier' || col.attributeType === 'Lookup'
            ? tokens.fontFamilyMonospace
            : tokens.fontFamilyBase,
        fontSize: 12,
        color: tokens.colorNeutralForeground1,
        wordBreak: 'break-word',
      }}
    >
      {rendered}
    </span>
  );
}

function formatValue(col: ColumnMeta, value: unknown): string {
  if (value == null || value === '') return '(empty)';
  if (col.attributeType === 'Boolean') {
    return value === true || value === '1' || value === 1
      ? col.trueOption.label
      : col.falseOption.label;
  }
  const opts = columnOptions(col);
  if (opts) {
    const n = Number(value);
    if (Number.isFinite(n)) {
      const o = opts.find((x) => x.value === n);
      if (o) return `${o.label} · ${n}`;
    }
    if (col.attributeType === 'MultiSelectPicklist' && typeof value === 'string') {
      return value
        .split(',')
        .map((v) => {
          const o = opts.find((x) => String(x.value) === v);
          return o ? o.label : v;
        })
        .join(', ');
    }
  }
  if (col.attributeType === 'Money') {
    const n = Number(value);
    if (Number.isFinite(n))
      return n.toLocaleString(undefined, { style: 'currency', currency: 'USD' });
  }
  if (col.attributeType === 'DateTime' && typeof value === 'string') {
    try {
      const d = new Date(value);
      if (Number.isFinite(d.getTime())) return d.toLocaleString();
    } catch {
      /* fall through */
    }
  }
  return String(value);
}

// ──────────────────────────────────────────────────────────────
// Custom override input — small, inline, type-aware
// ──────────────────────────────────────────────────────────────
function CustomValueInput({
  col,
  value,
  onChange,
  fallback,
  parentTable,
}: {
  col: ColumnMeta;
  value: CreateFieldValue | undefined;
  onChange: (v: CreateFieldValue) => void;
  fallback: string;
  parentTable: string;
}) {
  // Lazy enrichment so picklist/lookup details (options, targets) land.
  // Same pattern as FieldSetEditor.FieldInput — fires once per (entity,
  // attr) and caches.
  useColumnDetail(parentTable, col.logicalName);
  switch (col.attributeType) {
    case 'Memo':
      return (
        <Textarea
          rows={2}
          size="small"
          value={(value as string) ?? fallback}
          onChange={(_, d) => onChange(d.value)}
          style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 12 }}
        />
      );
    case 'Integer':
    case 'BigInt':
    case 'Decimal':
    case 'Double':
    case 'Money':
      return (
        <SpinButton
          size="small"
          value={Number(value ?? 0)}
          onChange={(_, d) => {
            const n = d.value ?? Number(d.displayValue ?? 0);
            if (Number.isFinite(n)) onChange(n);
          }}
        />
      );
    case 'Boolean':
      return (
        <Switch
          checked={value === true}
          onChange={(_, d) => onChange(d.checked)}
          label={value === true ? col.trueOption.label : col.falseOption.label}
        />
      );
    case 'Picklist':
    case 'State':
    case 'Status':
    case 'EntityName': {
      const opts = columnOptions(col) ?? [];
      const cur = opts.find((o) => o.value === value);
      return (
        <Combobox
          size="small"
          value={cur?.label ?? (value != null ? String(value) : '')}
          selectedOptions={value != null ? [String(value)] : []}
          onOptionSelect={(_, d) => d.optionValue && onChange(Number(d.optionValue))}
          placeholder="Pick a value…"
        >
          {opts.map((o) => (
            <Option key={o.value} value={String(o.value)} text={o.label}>
              {o.label}
            </Option>
          ))}
        </Combobox>
      );
    }
    case 'Lookup':
    case 'Customer':
    case 'Owner':
      // Live typeahead — reuses the same `LookupFieldInput` the field-set
      // editor uses for Create/Update. Polymorphic targets (Customer /
      // Owner) get the target switcher first; single-target Lookups get
      // a single typeahead. Old code was a bare GUID Input which was
      // unusable in practice.
      return (
        <LookupFieldInput
          col={col as LookupColumnMeta | CustomerColumnMeta | OwnerColumnMeta}
          value={value as LookupFieldValue | undefined}
          onChange={onChange}
        />
      );
    case 'MultiSelectPicklist': {
      // Comma-separated integer string on the wire; render as a multi-pick
      // checkbox list inline. Same encoding rule the create body uses.
      const opts = columnOptions(col) ?? [];
      const currentArr: number[] = Array.isArray(value)
        ? (value as number[])
        : typeof value === 'string' && value
          ? value
              .split(',')
              .map((n) => Number(n.trim()))
              .filter((n) => Number.isFinite(n))
          : [];
      return (
        <Combobox
          size="small"
          multiselect
          value={currentArr
            .map((n) => opts.find((o) => o.value === n)?.label ?? String(n))
            .join(', ')}
          selectedOptions={currentArr.map(String)}
          onOptionSelect={(_, d) => {
            const next = (d.selectedOptions ?? [])
              .map((s) => Number(s))
              .filter((n) => Number.isFinite(n));
            onChange(next);
          }}
          placeholder="Pick one or more…"
        >
          {opts.map((o) => (
            <Option key={o.value} value={String(o.value)} text={o.label}>
              {o.label}
            </Option>
          ))}
        </Combobox>
      );
    }
    case 'DateTime':
      // ISO string input. Native datetime-local picker for the simplest path;
      // server accepts ISO 8601 either way.
      return (
        <Input
          size="small"
          type={col.format === 'DateOnly' ? 'date' : 'datetime-local'}
          value={typeof value === 'string' ? value.slice(0, 16) : ''}
          onChange={(_, d) => onChange(d.value)}
          style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 12 }}
        />
      );
    default:
      return (
        <Input
          size="small"
          value={(value as string) ?? fallback}
          onChange={(_, d) => onChange(d.value)}
        />
      );
  }
}

// Suppress unused imports kept for parity
void ChevronRight20Regular;
