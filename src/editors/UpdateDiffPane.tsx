// UpdateDiffPane — "Diff — before → after" card for the Update mode.
//
// Takes the original record state as a PROP (fetched live by the parent
// mode — UpdateMode owns the fetch, not us). Compares against the user's
// in-progress fieldValues. Renders:
//   • Changed rows: before → after, color-coded
//   • Unchanged rows: faded (collapsed under "Show unchanged" toggle)
//   • Per-row drill-down → switches to PUT single-column mode
//
// Layout:
//   * Diff           [3 changed]
//     - name             "Old"      → "New"
//     - revenue          $1.0M      → $2.5M
//   * Field set      [12 unchanged, hidden by default]   (collapsed)
//
// "Update single column": Drilling into a column flips the URL to
// /accounts(id)/name and the method to PUT.

import { useMemo, useState } from 'react';
import {
  Caption1,
  Badge,
  Tooltip,
  Button,
  Spinner,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  tokens,
} from '@fluentui/react-components';
import {
  BranchFork20Filled,
  ArrowSwap20Regular,
  Eye20Regular,
  EyeOff20Regular,
  ChevronRight20Filled,
} from '@fluentui/react-icons';
import { PaneHead } from './PaneHead';
import { findTable, columnOptions, isLookupLike, type ColumnMeta } from '../mock/metadata';
import type { CreateFieldValue, LookupFieldValue } from '../state/writeState';

// ──────────────────────────────────────────────────────────────
// Columns we exclude from the diff (system / binary)
// ──────────────────────────────────────────────────────────────
const HIDDEN_FROM_DIFF = new Set([
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

interface DiffRow {
  col: ColumnMeta;
  before: unknown;
  after: unknown;
  changed: boolean;
}

export interface UpdateDiffPaneProps {
  table: string;
  recordId: string | null;
  fieldValues: Record<string, CreateFieldValue>;
  /**
   * The current persisted state of the target record. Fetched live by
   * the parent UpdateMode (one round-trip per pick) and passed in as a
   * prop. `null` means "no record picked OR fetch still in flight".
   * `originalLoading` distinguishes those cases.
   */
  original: Record<string, unknown> | null;
  originalLoading: boolean;
  /** Drill into a single column → triggers PUT mode in the parent. */
  onDrillColumn: (column: string) => void;
}

export function UpdateDiffPane({
  table,
  recordId,
  fieldValues,
  original,
  originalLoading,
  onDrillColumn,
}: UpdateDiffPaneProps) {
  const tbl = findTable(table);
  const [showUnchanged, setShowUnchanged] = useState(false);

  const rows = useMemo<DiffRow[]>(() => {
    if (!tbl || !original) return [];
    const out: DiffRow[] = [];
    for (const col of tbl.columns) {
      if (HIDDEN_FROM_DIFF.has(col.logicalName)) continue;
      if (col.attributeType === 'File' || col.attributeType === 'Image') continue;
      const before = original[col.logicalName];
      const inDiff = col.logicalName in fieldValues;
      const raw = fieldValues[col.logicalName];
      const after = inDiff ? toComparable(col, raw) : before;
      const changed = inDiff && !sameAsBefore(col, before, raw);
      out.push({ col, before, after, changed });
    }
    return out;
  }, [tbl, original, fieldValues]);

  const changedCount = rows.filter((r) => r.changed).length;
  const unchangedCount = rows.length - changedCount;
  const displayRows = showUnchanged ? rows : rows.filter((r) => r.changed);

  if (!tbl) {
    return (
      <MessageBar layout="multiline" intent="error">
        <MessageBarBody>
          Unknown table <code>{table}</code>.
        </MessageBarBody>
      </MessageBar>
    );
  }
  if (!original) {
    // Three sub-states:
    //   1. No record picked         → recordId is null
    //   2. Record picked, fetching  → originalLoading true
    //   3. Fetch returned nothing   → record gone or no read access
    return (
      <div>
        <PaneHead
          icon={BranchFork20Filled}
          title="Diff"
          sub="Pick a record to compare against."
          group="write"
        />
        {!recordId ? (
          <MessageBar layout="multiline" intent="info" style={{ maxWidth: 720 }}>
            <MessageBarBody>
              No record selected — switch to the <strong>Target</strong> pane and pick one. The diff
              compares your edits against the record&apos;s current persisted state.
            </MessageBarBody>
          </MessageBar>
        ) : originalLoading ? (
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 0' }}>
            <Spinner size="small" />
            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
              Loading the current record from Dataverse — this anchors the &ldquo;before&rdquo;
              column of the diff.
            </Caption1>
          </div>
        ) : (
          <MessageBar layout="multiline" intent="warning" style={{ maxWidth: 720 }}>
            <MessageBarBody>
              <MessageBarTitle>Couldn&apos;t load the record.</MessageBarTitle>
              GUID <code>{recordId}</code> didn&apos;t resolve in <code>{table}</code>. The record
              may have been deleted, or you might not have read access. You can still author the
              PATCH — the diff just won&apos;t show a &ldquo;before&rdquo; column.
            </MessageBarBody>
          </MessageBar>
        )}
      </div>
    );
  }

  const primaryName = String(original[tbl.primaryName] ?? '');

  return (
    <div>
      <PaneHead
        icon={BranchFork20Filled}
        title={primaryName || '(unnamed record)'}
        sub={
          changedCount > 0 ? (
            <>
              {changedCount} field{changedCount === 1 ? '' : 's'} changed · {unchangedCount}{' '}
              unchanged
            </>
          ) : (
            <>
              No changes yet — edit fields in the <strong>Field set</strong> pane to populate the
              diff.
            </>
          )
        }
        group="write"
      >
        <Tooltip
          content={showUnchanged ? 'Hiding unchanged rows' : 'Showing unchanged rows'}
          relationship="description"
        >
          <Button
            size="small"
            appearance={showUnchanged ? 'primary' : 'outline'}
            icon={showUnchanged ? <Eye20Regular /> : <EyeOff20Regular />}
            onClick={() => setShowUnchanged((v) => !v)}
          >
            {showUnchanged ? 'All fields' : 'Changes only'}
          </Button>
        </Tooltip>
        <Badge appearance="tint" color={changedCount > 0 ? 'brand' : 'subtle'}>
          {changedCount} change{changedCount === 1 ? '' : 's'}
        </Badge>
      </PaneHead>

      {changedCount === 0 && !showUnchanged && (
        <MessageBar layout="multiline" intent="info" style={{ marginBottom: 14 }}>
          <MessageBarBody>
            <MessageBarTitle>No diff to render.</MessageBarTitle>
            PATCH sends only the fields in the body — edit one or more columns in the{' '}
            <strong>Field set</strong> pane to see the before → after diff. You can also drill into
            any column for a single-property <code>PUT</code> instead.
          </MessageBarBody>
        </MessageBar>
      )}

      <div style={{ maxWidth: 1080, overflowX: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0 }}>
          <thead>
            <tr>
              <Th style={{ width: '22%' }}>Field</Th>
              <Th tone="before">Before</Th>
              <Th>—</Th>
              <Th tone="after">After</Th>
              <Th style={{ width: 88 }}>Drill</Th>
            </tr>
          </thead>
          <tbody>
            {displayRows.length === 0 && (
              <tr>
                <td
                  colSpan={5}
                  style={{
                    padding: '24px 8px',
                    textAlign: 'center',
                    color: tokens.colorNeutralForeground3,
                    fontStyle: 'italic',
                  }}
                >
                  {showUnchanged
                    ? 'No fields on this record (somehow).'
                    : 'No changes — toggle "All fields" to see the unchanged rows.'}
                </td>
              </tr>
            )}
            {displayRows.map((r) => (
              <DiffRowView
                key={r.col.logicalName}
                row={r}
                onDrill={() => onDrillColumn(r.col.logicalName)}
              />
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function DiffRowView({ row, onDrill }: { row: DiffRow; onDrill: () => void }) {
  const { col, before, after, changed } = row;
  const fadeStyle = !changed ? { opacity: 0.45 } : {};

  return (
    <tr style={fadeStyle}>
      <Td>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span
            style={{
              fontFamily: tokens.fontFamilyMonospace,
              fontSize: 12,
              color: tokens.colorNeutralForeground1,
              fontWeight: 500,
            }}
          >
            {col.logicalName}
          </span>
          <Caption1 style={{ color: tokens.colorNeutralForeground3, fontSize: 10 }}>
            {col.displayName} ·{' '}
            <span style={{ color: tokens.colorBrandForeground2 }}>{col.attributeType}</span>
          </Caption1>
        </div>
      </Td>
      <Td tone="before">
        <ValueCell col={col} value={before} faded={!changed} />
      </Td>
      <Td>
        <span
          style={{
            color: changed ? tokens.colorBrandForeground1 : tokens.colorNeutralForeground3,
            fontSize: 18,
            fontWeight: 700,
          }}
        >
          {changed ? '→' : '='}
        </span>
      </Td>
      <Td tone={changed ? 'after' : undefined}>
        <ValueCell col={col} value={after} faded={!changed} highlight={changed} />
      </Td>
      <Td>
        <Tooltip content={`Switch to PUT /${col.logicalName}`} relationship="description">
          <Button
            size="small"
            appearance="subtle"
            icon={<ChevronRight20Filled />}
            onClick={onDrill}
            aria-label={`Drill into ${col.logicalName}`}
          >
            PUT
          </Button>
        </Tooltip>
      </Td>
    </tr>
  );
}

// ──────────────────────────────────────────────────────────────
// Header/body cells + value formatting
// ──────────────────────────────────────────────────────────────
function Th({
  children,
  style,
  tone,
}: {
  children: React.ReactNode;
  style?: React.CSSProperties;
  tone?: 'before' | 'after';
}) {
  const accent =
    tone === 'before'
      ? { background: 'rgba(125,125,125,.06)' }
      : tone === 'after'
        ? { background: 'rgba(31,127,56,.08)', color: tokens.colorPaletteGreenForeground1 }
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

function Td({ children, tone }: { children: React.ReactNode; tone?: 'before' | 'after' }) {
  const accent =
    tone === 'before'
      ? { background: 'rgba(125,125,125,.025)' }
      : tone === 'after'
        ? {
            background: 'rgba(31,127,56,.04)',
            borderLeft: `2px solid ${tokens.colorPaletteGreenBorderActive}`,
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

function ValueCell({
  col,
  value,
  faded,
  highlight,
}: {
  col: ColumnMeta;
  value: unknown;
  faded?: boolean;
  highlight?: boolean;
}) {
  if (value == null || value === '') {
    return (
      <span style={{ color: tokens.colorNeutralForeground3, fontStyle: 'italic', fontSize: 11 }}>
        (empty)
      </span>
    );
  }
  return (
    <span
      style={{
        fontFamily:
          col.attributeType === 'Uniqueidentifier' || isLookupLike(col)
            ? tokens.fontFamilyMonospace
            : tokens.fontFamilyBase,
        fontSize: 12,
        color: highlight ? tokens.colorPaletteGreenForeground1 : tokens.colorNeutralForeground1,
        fontWeight: highlight ? 600 : 400,
        opacity: faded && !highlight ? 0.7 : 1,
        wordBreak: 'break-word',
      }}
    >
      {formatValue(col, value)}
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
  if (isLookupLike(col)) {
    if (typeof value === 'object' && value !== null && 'id' in value) {
      const lk = value as LookupFieldValue;
      return `${lk.targetEntity}(${lk.id.slice(0, 8)}…)`;
    }
    if (typeof value === 'string') return value;
  }
  const opts = columnOptions(col);
  if (opts) {
    const n = Number(value);
    if (Number.isFinite(n)) {
      const o = opts.find((x) => x.value === n);
      if (o) return `${o.label} · ${n}`;
    }
    if (col.attributeType === 'MultiSelectPicklist') {
      if (Array.isArray(value)) {
        return value.map((v) => opts.find((o) => o.value === v)?.label ?? String(v)).join(', ');
      }
      if (typeof value === 'string') {
        return value
          .split(',')
          .map((v) => opts.find((o) => String(o.value) === v.trim())?.label ?? v)
          .join(', ');
      }
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

// Convert a CreateFieldValue (in-memory shape) into a comparable wire value.
function toComparable(col: ColumnMeta, raw: CreateFieldValue): unknown {
  if (raw == null) return null;
  if (isLookupLike(col)) {
    const lk = raw as LookupFieldValue;
    return lk?.id ?? null;
  }
  if (col.attributeType === 'MultiSelectPicklist' && Array.isArray(raw)) {
    return raw.join(',');
  }
  return raw;
}

function sameAsBefore(col: ColumnMeta, before: unknown, after: CreateFieldValue): boolean {
  const cmp = toComparable(col, after);
  if (before == null && (cmp == null || cmp === '')) return true;
  if (cmp == null) return false;
  return String(before) === String(cmp);
}

// Suppress unused-import warning
void ArrowSwap20Regular;
