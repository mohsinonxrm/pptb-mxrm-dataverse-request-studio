// CollectionSubgrid — renders a 1:N / N:N expanded array inside a parent
// RecordDetailCard's accordion panel.
//
// Why a custom mini-table instead of reusing ResultsGrid:
//   • ResultsGrid is a full DataGrid with virtualization, multi-select,
//     resize, infinite-scroll, search, density toggle — none of which we
//     want inside an accordion. The user is scanning, not querying.
//   • A simple Fluent v9 Table compositions onto the page naturally,
//     stays the same height as its content, and lets us put a per-row
//     "expand this row inline" toggle without fighting DataGrid's row
//     virtualizer.
//
// Row expand behaviour:
//   • If a row is all scalars (no nested objects / arrays in the response),
//     we just render it as a plain table row.
//   • If a row has nested expands (e.g. `accounts → opportunities → owninguser`),
//     the row gets a chevron toggle. Click it → the row is REPLACED by a
//     full RecordDetailCard for that row, recursing into the deeper levels.
//   • Multiple rows can be expanded at once.

import { useMemo, useState } from 'react';
import {
  Table,
  TableHeader,
  TableHeaderCell,
  TableBody,
  TableRow,
  TableCell,
  Button,
  Caption1,
  tokens,
} from '@fluentui/react-components';
import { ChevronRight16Regular, ChevronDown16Regular } from '@fluentui/react-icons';
import { findTable } from '../../mock/metadata';
import { partitionRecord, prettifyKey } from './detailFieldPartitioner';
import { RecordDetailCard } from './RecordDetailCard';

export interface CollectionSubgridProps {
  rows: Record<string, unknown>[];
  /** Target entity logical name — drives nav-property awareness in
   *  expanded sub-cards. Empty string OK; we degrade gracefully. */
  entityLogical: string;
  /** Depth inside the recursive RecordDetailCard tree. Passed through
   *  to children for default-open / default-closed heuristics. */
  level: number;
}

export function CollectionSubgrid({ rows, entityLogical, level }: CollectionSubgridProps) {
  const tbl = findTable(entityLogical);

  // Compute the column set ONCE from the first row's scalar partition.
  // Per-row partitioning still happens later (for the row-detail toggle);
  // this just decides which columns to show in the table.
  //
  // We union scalar keys across the first few rows so that an option-set
  // column that's `null` in row 0 but populated in row 1 still gets a
  // header. Capped at 5 rows for the union to stay cheap.
  const columns = useMemo(() => {
    const seen = new Set<string>();
    const keys: string[] = [];
    for (const r of rows.slice(0, 5)) {
      const p = partitionRecord(r, tbl);
      for (const s of p.scalars) {
        // Skip the etag — it's noise at the column level.
        if (s.key === '@odata.etag') continue;
        if (!seen.has(s.key)) {
          seen.add(s.key);
          keys.push(s.key);
        }
      }
    }
    return keys;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rows, tbl?.logicalName]);

  // Per-row partitions (memoized). We use this to decide whether each
  // row has any nested expands — gates the row-expand toggle.
  const rowPartitions = useMemo(() => rows.map((r) => partitionRecord(r, tbl)), [rows, tbl]);

  const [expandedRowIdx, setExpandedRowIdx] = useState<Set<number>>(new Set());
  const toggleRow = (i: number) => {
    setExpandedRowIdx((prev) => {
      const next = new Set(prev);
      if (next.has(i)) next.delete(i);
      else next.add(i);
      return next;
    });
  };

  if (columns.length === 0) {
    return (
      <Caption1 style={{ color: tokens.colorNeutralForeground3, fontStyle: 'italic' }}>
        Rows have no scalar fields to show in this grid.
      </Caption1>
    );
  }

  // Column min-width: GUID columns need ~280px to show the full value
  // without ellipsizing; numeric/code columns are fine at ~120px; long
  // text columns get a generous min so the user can read at least 200
  // chars before truncation. We don't try to MEASURE — pick by heuristic
  // on the column key + content.
  const columnMinWidth = (key: string): number => {
    if (/^_.*_value$/.test(key)) return 280; // lookup wire form (GUID)
    if (key === '@odata.etag') return 90;
    if (/id$/i.test(key)) return 280; // primary key GUIDs
    if (/^(state|status|.*code|.*type)$/i.test(key)) return 130;
    if (/(age|count|amount|value|qty|number)$/i.test(key)) return 130;
    return 200;
  };

  const hasToggleColGlobal = rowPartitions.some((p) => hasNested(p));

  return (
    <div style={{ overflow: 'auto', maxWidth: '100%' }}>
      <Table
        size="small"
        arial-label={`${entityLogical || 'related'} rows`}
        // Layout fixed isn't right — we want columns to size by min-width
        // and let the outer container scroll. `min-content` on the table
        // would shrink-wrap the columns; we use natural sizing + a
        // `min-width` per column instead.
        style={{ width: 'max-content', minWidth: '100%' }}
      >
        <TableHeader>
          <TableRow>
            {hasToggleColGlobal && <TableHeaderCell style={{ width: 32, minWidth: 32 }} />}
            {columns.map((k) => (
              <TableHeaderCell key={k} style={{ minWidth: columnMinWidth(k) }}>
                <span style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}>
                  {prettifyKey(k)}
                </span>
              </TableHeaderCell>
            ))}
          </TableRow>
        </TableHeader>
        <TableBody>
          {rows.map((r, i) => {
            const p = rowPartitions[i];
            const expanded = expandedRowIdx.has(i);
            const nested = hasNested(p);
            const hasToggleCol = rowPartitions.some((rp) => hasNested(rp));
            return (
              <RowOrSubcard
                key={i}
                row={r}
                rowPartition={p}
                columns={columns}
                entityLogical={entityLogical}
                level={level}
                expanded={expanded}
                nested={nested}
                hasToggleCol={hasToggleCol}
                onToggle={() => toggleRow(i)}
              />
            );
          })}
        </TableBody>
      </Table>
    </div>
  );
}

function hasNested(p: ReturnType<typeof partitionRecord>): boolean {
  return p.navObjects.length > 0 || p.navCollections.length > 0 || p.navNulls.length > 0;
}

function RowOrSubcard({
  row,
  rowPartition,
  columns,
  entityLogical,
  level,
  expanded,
  nested,
  hasToggleCol,
  onToggle,
}: {
  row: Record<string, unknown>;
  rowPartition: ReturnType<typeof partitionRecord>;
  columns: string[];
  entityLogical: string;
  level: number;
  expanded: boolean;
  nested: boolean;
  hasToggleCol: boolean;
  onToggle: () => void;
}) {
  if (expanded) {
    // Replace the row's normal cells with a single full-width cell that
    // hosts a full RecordDetailCard. The card recurses through its
    // own nested expands automatically.
    return (
      <TableRow>
        {hasToggleCol && (
          <TableCell style={{ verticalAlign: 'top' }}>
            <Button
              appearance="subtle"
              size="small"
              icon={<ChevronDown16Regular />}
              onClick={onToggle}
              aria-label="Collapse row"
            />
          </TableCell>
        )}
        <TableCell
          colSpan={columns.length}
          style={{
            background: tokens.colorNeutralBackground2,
            padding: 12,
          }}
        >
          <RecordDetailCard
            record={row}
            entityLogical={entityLogical}
            level={level + 1}
            showHeadline={false}
          />
        </TableCell>
      </TableRow>
    );
  }

  return (
    <TableRow>
      {hasToggleCol && (
        <TableCell style={{ verticalAlign: 'middle' }}>
          {nested ? (
            <Button
              appearance="subtle"
              size="small"
              icon={<ChevronRight16Regular />}
              onClick={onToggle}
              aria-label="Expand row to show nested fields"
            />
          ) : null}
        </TableCell>
      )}
      {columns.map((k) => {
        const scalar = rowPartition.scalars.find((s) => s.key === k);
        if (!scalar) {
          return (
            <TableCell key={k}>
              <span style={{ color: tokens.colorNeutralForeground4 }}>—</span>
            </TableCell>
          );
        }
        // Formatted-value preference, same as the main scalar grid.
        if (scalar.formattedValue != null) {
          return (
            <TableCell key={k}>
              <span>{scalar.formattedValue}</span>
              <span
                style={{
                  marginLeft: 6,
                  fontFamily: tokens.fontFamilyMonospace,
                  fontSize: 10,
                  color: tokens.colorNeutralForeground3,
                }}
              >
                ({String(scalar.value)})
              </span>
            </TableCell>
          );
        }
        if (scalar.value == null) {
          return (
            <TableCell key={k}>
              <span style={{ color: tokens.colorNeutralForeground4 }}>—</span>
            </TableCell>
          );
        }
        const v = scalar.value;
        const isStringy = typeof v === 'string';
        const looksLikeGuid =
          isStringy &&
          /^[{(]?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}[)}]?$/.test(
            v as string,
          );
        return (
          <TableCell
            key={k}
            style={{
              fontFamily: looksLikeGuid ? tokens.fontFamilyMonospace : undefined,
              fontSize: looksLikeGuid ? 11 : undefined,
              // No maxWidth — outer container scrolls horizontally. We
              // still ellipsize within a generous cap so a single
              // wall-of-text field doesn't blow the column out to 1000+px.
              maxWidth: 480,
              overflow: 'hidden',
              textOverflow: 'ellipsis',
              whiteSpace: 'nowrap',
            }}
            title={String(v)}
          >
            {String(v)}
          </TableCell>
        );
      })}
    </TableRow>
  );
}
