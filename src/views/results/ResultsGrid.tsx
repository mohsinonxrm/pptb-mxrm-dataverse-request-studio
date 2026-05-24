// Virtualized DataGrid for Retrieve-Multiple results.
//
// Architecture:
//   • @fluentui-contrib/react-data-grid-react-window for 1-D virtualization
//     (handles 100k+ rows smoothly without paging tricks).
//   • ResizeObserver measures the container + header so the virtualized
//     body can size itself precisely (no scroll inside scroll inside scroll).
//   • Per-column sizing options computed from BOTH header display name
//     width and the column's attribute-type-based minimum (GUIDs need 260,
//     dates need 160, etc.) — prevents the misaligned-header overflow you
//     see when every column defaults to the same narrow width.
//   • Sort indicators rendered in `renderHeaderCell` from the active
//     `$orderby` state (we don't use Fluent's internal sortDirection because
//     we manage multi-sort externally).
//   • `getCellRenderer` dispatches per AttributeTypeCode → rich cells
//     (Switch for Boolean, Badge for Picklist/Status, bold-name +
//     target-type caption for Lookup, locale-formatted numbers + dates).
//   • Selection: `getRowId` returns the entity's primary GUID. Multi-select
//     via the DataGrid's built-in selection.
//   • Infinite scroll: `onItemsRendered` fires `onLoadMore` when the user
//     scrolls within 10 rows of the bottom and there are more pages.
//
// Notes:
//   - Sort state comes from `OrderbySpec[]` (the studio's multi-sort model).
//   - Column display info comes from the live `TableMeta` registry via
//     `resolveColumnDisplay`.
//   - Nested $expand objects are flattened to dotted keys via `flattenRows`
//     so the grid sees a flat column list — same model as model-driven
//     advanced-find's link-entity flattening.

import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  createTableColumn,
  makeStyles,
  tokens,
  TableCellLayout,
  Skeleton,
  SkeletonItem,
  useScrollbarWidth,
  useFluent,
  Spinner,
} from '@fluentui/react-components';
import {
  DataGrid,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridBody,
  DataGridRow,
  DataGridCell,
} from '@fluentui-contrib/react-data-grid-react-window';
import type { RowRenderer } from '@fluentui-contrib/react-data-grid-react-window';
import { ArrowSortUp16Regular, ArrowSortDown16Regular } from '@fluentui/react-icons';

import { findTable } from '../../mock/metadata';
import type { ExpandSpec } from '../../editors/ExpandEditor';
import type { OrderbySpec } from '../../editors/OrderbyEditor';
import { usePersistedSettings } from '../../host/usePersistedSettings';
import { StatusPill } from '../../primitives/StatusPill';

import {
  filterDisplayableColumns,
  getAssociatedNavProperty,
  getFormattedValue,
  getLookupTargetEntity,
} from './FormattedValueUtils';
import { resolveColumnDisplay, type ColumnDisplayInfo } from './columnDisplayMap';
import { flattenRows, ROW_KEY } from './flattenExpansions';
import { getCellRenderer } from './DataGridCellRenderers';

const ROW_HEIGHT_COZY = 44;
const ROW_HEIGHT_COMPACT = 32;

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    minHeight: 0,
    backgroundColor: tokens.colorNeutralBackground1,
    position: 'relative',
  },
  gridWrapper: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    overflow: 'hidden',
    backgroundColor: tokens.colorNeutralBackground1,
    position: 'relative',
    minHeight: 0,
  },
  gridContent: {
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    overflow: 'hidden', // virtualized DataGrid handles its own scroll
    width: '100%',
    minHeight: 0,
  },
  body: { scrollbarGutter: 'stable' },
  row: { boxSizing: 'border-box' },
  headerCellContent: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    width: '100%',
    minWidth: 0,
  },
  headerCellLabel: {
    fontWeight: 600,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    minWidth: 0,
  },
  sortIndicator: {
    display: 'flex',
    alignItems: 'center',
    marginLeft: 'auto',
    flexShrink: 0,
  },
  // Hide Fluent's default sort arrow — we render our own in renderHeaderCell
  // so we can drive it off the request's $orderby state (incl. multi-sort).
  headerCell: {
    '& > button > span:last-child': { display: 'none' },
  },
  infoBar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: tokens.spacingHorizontalM,
    padding: '6px 12px',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    minHeight: '30px',
    zIndex: 1,
  },
  emptyState: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    height: '100%',
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase300,
  },
  loadingMore: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    color: tokens.colorBrandForeground1,
    marginLeft: tokens.spacingHorizontalS,
  },
});

export interface ResultsGridProps {
  /** Raw OData envelope. `body.value[]` are the rows; `body['@odata.nextLink']` (if any) drives infinite scroll. */
  body: unknown;
  /** Root entity logical name. */
  table?: string;
  /** Active root $select — column order in the grid mirrors this user-chosen order. */
  select?: string[];
  /** Active $expand spec — for nested-column header labels and inner $select ordering. */
  expand?: ExpandSpec[];
  /** Active $orderby spec — drives sort-indicator arrows. */
  orderby?: OrderbySpec[];
  /** Optional: when set, the grid will show a "loading more" indicator and call onLoadMore near bottom. */
  hasMore?: boolean;
  isLoadingMore?: boolean;
  onLoadMore?: () => void;
  /** Optional: selected GUIDs reported back to parent. */
  onSelectionChange?: (recordIds: string[]) => void;
  /** Client-side row search query (filters across raw + formatted values). */
  searchQuery?: string;
  /** Row density. */
  density?: 'cozy' | 'compact';
  /** Status info rendered in the grid's bottom info bar (replaces the old top strip). */
  status?: { code: number; ms: number; bytes: number };
  /** Server-emitted @odata.count, when $count=true was set. */
  serverTotal?: number;
  /**
   * Optional column-order hint used when `select` is empty. Modes that
   * source their column projection from somewhere other than $select —
   * Predefined Query uses the saved view's layoutxml — pass the attribute
   * names in display order. The grid reorders the leftover-backfill step
   * to honor this list before falling back to JSON-key order.
   *
   * Has no effect when `select` is non-empty; explicit $select always wins.
   * Each entry is matched against both bare and `_<x>_value` forms so
   * lookup columns slot into the right position regardless of which form
   * the server returned.
   */
  preferredColumnOrder?: string[];
}

export function ResultsGrid({
  body, table, select, expand, orderby,
  hasMore, isLoadingMore, onLoadMore, onSelectionChange,
  searchQuery = '', density = 'cozy',
  status, serverTotal, preferredColumnOrder,
}: ResultsGridProps) {
  const styles = useStyles();
  const [settings] = usePersistedSettings();
  const useLogicalNames = settings.useLogicalNames;
  const valueMode = settings.valueDisplayMode;

  const { targetDocument } = useFluent();
  const scrollbarWidth = useScrollbarWidth({ targetDocument });

  // ── Container/header measurements — precise sizing for the virtualizer ──
  const containerRef = useRef<HTMLDivElement>(null);
  const headerRef = useRef<HTMLDivElement>(null);
  const [gridDimensions, setGridDimensions] = useState({ width: 0, height: 0 });
  const [headerHeight, setHeaderHeight] = useState(0);
  const [selectedItems, setSelectedItems] = useState<Set<string | number>>(new Set());

  const rowHeight = density === 'compact' ? ROW_HEIGHT_COMPACT : ROW_HEIGHT_COZY;

  // ── Flatten nested $expand objects to dotted-key columns ──
  const allRows = useMemo(() => {
    const raw = (body as { value?: Record<string, unknown>[] } | null)?.value ?? [];
    return flattenRows(raw);
  }, [body]);

  // ── Client-side search (raw + formatted values) ──
  const rows = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    if (!q) return allRows;
    return allRows.filter(r =>
      Object.keys(r).some(k => {
        if (k.includes('@')) return false; // skip annotation keys for the column scan
        const f = getFormattedValue(r, k);
        if (f != null && String(f).toLowerCase().includes(q)) return true;
        return String(r[k] ?? '').toLowerCase().includes(q);
      }),
    );
  }, [allRows, searchQuery]);

  // ── Container ResizeObserver ──
  useEffect(() => {
    const node = containerRef.current;
    if (!node) return;
    const obs = new ResizeObserver(entries => {
      for (const entry of entries) {
        const { width, height } = entry.contentRect;
        setGridDimensions({ width, height });
      }
    });
    obs.observe(node);
    return () => obs.disconnect();
  }, [rows.length === 0]);

  // ── Header ResizeObserver ──
  useEffect(() => {
    const node = headerRef.current;
    if (!node) return;
    const obs = new ResizeObserver(() => setHeaderHeight(node.offsetHeight));
    obs.observe(node);
    return () => obs.disconnect();
  }, [rows.length === 0]);

  // ── Clear selection when results change ──
  useEffect(() => {
    setSelectedItems(new Set());
  }, [body]);

  // ── Column key list (root keys after annotation filtering + flatten) ──
  //
  // The user's mental model: "I picked these columns in THIS order, show
  // them in THIS order — and ONLY these." Dataverse sneaks extra columns
  // into the response when they're referenced by `$orderby` (the engine
  // has to evaluate them to sort the rows), and N:1 expands without an
  // inner `$select` return every column on the related entity. Without
  // a filtering rule the grid would surface those silent extras.
  //
  // Order priority:
  //
  //   1. Root $select — in user order. For Lookup/Customer/Owner we
  //      register both bare and `_<x>_value` forms so whichever the
  //      server returned lands at the right slot.
  //   2. Per-expand inner $select — each expand's columns, dot-prefixed
  //      by the expand's nav name. Nested expands chain the prefix.
  //   3. For each expand WITHOUT an inner $select, pick up all dotted
  //      keys under that nav (the server returned everything; user
  //      explicitly wanted those, just without an inner column list).
  //   4. Leftover backfill — ONLY when no root $select was specified.
  //      In that case the user asked for "everything"; we show every
  //      key the server returned. With a root $select, server-injected
  //      orderby columns / etag siblings are intentionally hidden.
  //
  // Annotation keys (`@…`) and the synthetic `__rowKey` are filtered.
  const hasExplicitRootSelect = !!select && select.length > 0;
  const columnKeys = useMemo(() => {
    if (rows.length === 0) return [];
    const allPresent = new Set<string>();
    for (const r of rows) for (const k of Object.keys(r)) allPresent.add(k);

    const ordered: string[] = [];
    const seen = new Set<string>();
    const take = (k: string) => {
      if (!k || seen.has(k)) return;
      seen.add(k);
      // Only surface keys that the server actually returned; user might
      // $select an attribute that the server omits (e.g. null lookup).
      if (allPresent.has(k)) ordered.push(k);
    };

    // 1) Root $select — user order. Try both bare and `_<x>_value` form.
    if (hasExplicitRootSelect) {
      for (const s of select!) {
        take(s);
        if (!s.startsWith('_')) take(`_${s}_value`);
      }
    }

    // 2 + 3) Expand inner $select — recursive. When an expand has its
    //   own $select, use that user-ordered list. When it doesn't, fan in
    //   every dotted key under the nav prefix so users can see everything
    //   the server returned for that branch.
    const visitExpand = (items: ExpandSpec[] | undefined, prefix: string) => {
      if (!items) return;
      for (const e of items) {
        const navPrefix = prefix ? `${prefix}.${e.nav}` : e.nav;
        if (e.select && e.select.length > 0) {
          for (const c of e.select) {
            take(`${navPrefix}.${c}`);
            if (!c.startsWith('_')) take(`${navPrefix}._${c}_value`);
          }
        } else {
          // No inner $select declared — include all dotted keys under
          // this nav. Sort them so the output is stable across renders.
          const subKeys: string[] = [];
          const navDotPrefix = `${navPrefix}.`;
          for (const k of allPresent) {
            if (k.startsWith(navDotPrefix)) subKeys.push(k);
          }
          subKeys.sort();
          for (const k of subKeys) take(k);
        }
        if (e.nestedExpand && e.nestedExpand.length) {
          visitExpand(e.nestedExpand, navPrefix);
        }
      }
    };
    visitExpand(expand, '');

    // 4) Preferred column order hint — for modes where $select isn't on
    //    the wire but the user still expects a known projection order
    //    (Predefined Query reads this from the saved view's layoutxml so
    //    the grid columns match the column order in Dynamics 365 / Sales
    //    Hub). Applied BEFORE the catch-all backfill so the named columns
    //    land first; backfill then fills in any extras the server sent.
    //
    //    Each entry tries both bare and `_<x>_value` forms (same as
    //    explicit-select handling above) so lookup columns slot in
    //    regardless of which key the server returned.
    if (!hasExplicitRootSelect && preferredColumnOrder && preferredColumnOrder.length > 0) {
      for (const c of preferredColumnOrder) {
        take(c);
        if (!c.startsWith('_')) take(`_${c}_value`);
      }
    }

    // 5) Leftover backfill — only when root $select is empty. With an
    //    explicit $select, we deliberately hide server-injected extras
    //    (orderby cols, address1_composite when filtering on a sub-field,
    //    etc.) so the grid faithfully reflects the user's chosen projection.
    if (!hasExplicitRootSelect) {
      for (const k of allPresent) take(k);
    }

    return filterDisplayableColumns(ordered);
  }, [rows, select, expand, hasExplicitRootSelect, preferredColumnOrder]);

  // ── Column display info (parent table, leaf column meta, header label) ──
  const rootTable = table ? findTable(table) : undefined;
  const columnDisplay = useMemo(() => {
    const m = new Map<string, ColumnDisplayInfo>();
    for (const key of columnKeys) {
      m.set(key, resolveColumnDisplay(key, rootTable, expand, useLogicalNames));
    }
    return m;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [columnKeys, rootTable?.logicalName, expand, useLogicalNames]);

  // ── Sort state map (from $orderby) ──
  const sortMap = useMemo(() => {
    const m = new Map<string, 'ascending' | 'descending'>();
    for (const it of orderby ?? []) {
      const dir = it.dir === 'desc' ? 'descending' : 'ascending';
      m.set(it.col, dir);
      m.set(`_${it.col}_value`, dir); // OData lookup form
    }
    return m;
  }, [orderby]);

  // ── Build TableColumnDefinition[] ──
  const columns = useMemo(() => {
    return columnKeys.flatMap((key) => {
      const info = columnDisplay.get(key)!;
      const sortDir = sortMap.get(key);

      const renderHeaderCell = (suffix?: string) => () => (
        <span className={styles.headerCellContent}>
          <span className={styles.headerCellLabel}>
            {useLogicalNames ? info.logicalName : info.displayName}
            {suffix ? ` ${suffix}` : ''}
          </span>
          {sortDir === 'ascending' && (
            <span className={styles.sortIndicator}><ArrowSortUp16Regular /></span>
          )}
          {sortDir === 'descending' && (
            <span className={styles.sortIndicator}><ArrowSortDown16Regular /></span>
          )}
        </span>
      );

      const renderFormatted = (r: Record<string, unknown>) => {
        const raw = r[key];
        const formatted = getFormattedValue(r, key);

        let targetDisplay: string | undefined;
        if (info.column?.attributeType === 'Lookup'
          || info.column?.attributeType === 'Customer'
          || info.column?.attributeType === 'Owner') {
          // Polymorphic-aware: pull per-row target from the annotation.
          const targetLogical =
            getLookupTargetEntity(r, key) ??
            getAssociatedNavProperty(r, key) ??
            info.column?.targets?.[0];
          targetDisplay = targetLogical
            ? (findTable(targetLogical)?.displayName ?? targetLogical)
            : undefined;
        }

        // Collection-valued expand columns surface a sibling
        // `<key>@odata.nextLink` when Dataverse paginated the inner result.
        // Propagate the flag so the CollectionCellRenderer can show a `+`
        // marker after the item count.
        const collectionHasMore = Array.isArray(raw)
          ? r[`${key}@odata.nextLink`] != null
          : undefined;

        return (
          <TableCellLayout truncate>
            {getCellRenderer({
              col: info.column,
              rawValue: raw,
              formattedValue: formatted,
              targetEntityDisplay: targetDisplay,
              collectionHasMore,
            })}
          </TableCellLayout>
        );
      };

      const renderRaw = (r: Record<string, unknown>) => {
        const raw = r[key];
        return (
          <TableCellLayout truncate>
            {raw == null ? (
              <span style={{ color: tokens.colorNeutralForeground4 }}>—</span>
            ) : (
              <span style={{
                display: 'block',
                overflow: 'hidden',
                textOverflow: 'ellipsis',
                whiteSpace: 'nowrap',
                fontFamily: tokens.fontFamilyMonospace,
                fontSize: tokens.fontSizeBase200,
              }}>{String(raw)}</span>
            )}
          </TableCellLayout>
        );
      };

      const compareFn = (a: Record<string, unknown>, b: Record<string, unknown>) => {
        const av = String(getFormattedValue(a, key) ?? a[key] ?? '');
        const bv = String(getFormattedValue(b, key) ?? b[key] ?? '');
        return av.localeCompare(bv);
      };

      if (valueMode === 'raw') {
        return [createTableColumn<Record<string, unknown>>({
          columnId: key,
          compare: compareFn,
          renderHeaderCell: renderHeaderCell(),
          renderCell: renderRaw,
        })];
      }

      if (valueMode === 'both') {
        const anyDiffers = rows.some(r => {
          const f = getFormattedValue(r, key);
          return f != null && f !== r[key];
        });
        if (anyDiffers) {
          return [
            createTableColumn<Record<string, unknown>>({
              columnId: key,
              compare: compareFn,
              renderHeaderCell: renderHeaderCell(),
              renderCell: renderFormatted,
            }),
            createTableColumn<Record<string, unknown>>({
              columnId: `${key}__raw`,
              compare: compareFn,
              renderHeaderCell: renderHeaderCell('(Raw)'),
              renderCell: renderRaw,
            }),
          ];
        }
      }

      return [createTableColumn<Record<string, unknown>>({
        columnId: key,
        compare: compareFn,
        renderHeaderCell: renderHeaderCell(),
        renderCell: renderFormatted,
      })];
    });
  }, [columnKeys, columnDisplay, sortMap, valueMode, useLogicalNames, rows, styles]);

  // ── Column sizing options ──
  // Width is computed from BOTH the display-name pixel estimate AND a
  // type-based minimum (GUIDs need 260px, dates 160, etc.) — that way
  // neither the header nor typical data values overflow.
  const columnSizingOptions = useMemo(() => {
    const out: Record<string, { minWidth: number; defaultWidth: number; idealWidth?: number }> = {};
    const typeMin = (attrType: string | undefined): number => {
      switch (attrType) {
        case 'Uniqueidentifier': return 260;
        case 'DateTime':         return 160;
        case 'Boolean':          return 110;
        case 'Lookup':
        case 'Customer':
        case 'Owner':            return 200;
        case 'Money':            return 130;
        case 'Picklist':
        case 'State':
        case 'Status':           return 140;
        case 'Memo':             return 240;
        case 'Integer':
        case 'BigInt':
        case 'Decimal':
        case 'Double':           return 110;
        default:                 return 140;
      }
    };
    for (const c of columns) {
      const cid = String(c.columnId);
      const baseKey = cid.endsWith('__raw') ? cid.slice(0, -5) : cid;
      const info = columnDisplay.get(baseKey);
      const label = useLogicalNames
        ? (info?.logicalName ?? cid)
        : (info?.displayName ?? cid);
      // ~7px per char + 56px padding (sort icon + cell padding)
      const headerEstimate = Math.ceil(label.length * 7) + 56;
      const dataMin = typeMin(info?.column?.attributeType);
      const ideal = Math.max(120, Math.min(Math.max(headerEstimate, dataMin), 360));
      out[cid] = {
        minWidth: Math.max(80, Math.min(dataMin, 200)),
        defaultWidth: ideal,
        idealWidth: ideal,
      };
    }
    return out;
  }, [columns, columnDisplay, useLogicalNames]);

  // ── Selection ──
  const handleSelectionChange = useCallback(
    (_e: unknown, data: { selectedItems: Set<string | number> }) => {
      setSelectedItems(data.selectedItems);
      onSelectionChange?.(Array.from(data.selectedItems).map(String));
    },
    [onSelectionChange],
  );

  // ── getRowId: prefer the synthetic `__rowKey` injected by the flattener
  //    (parent-PK alone is no longer unique once collection expands fan
  //    rows out). Falls back to the entity's primary key, then a JSON
  //    fingerprint. ──
  const getRowId = useCallback(
    (item: Record<string, unknown>) => {
      const synthetic = item[ROW_KEY];
      if (typeof synthetic === 'string') return synthetic;
      const pk = rootTable?.primaryKey;
      if (pk && typeof item[pk] === 'string') return item[pk] as string;
      // Final fallback — rare; only when both __rowKey and PK are missing.
      return JSON.stringify(item).slice(0, 64);
    },
    [rootTable?.primaryKey],
  );

  // ── Row renderer: skeleton while user is fast-scrolling ──
  const renderRow: RowRenderer<Record<string, unknown>> = useCallback(
    ({ item, rowId }, style, _index, isScrolling) => (
      <DataGridRow<Record<string, unknown>>
        key={rowId}
        style={{ ...style, height: `${rowHeight}px` }}
        className={styles.row}
      >
        {({ renderCell }) => (
          <DataGridCell focusMode="group">
            {isScrolling ? (
              <Skeleton style={{ width: '100%' }}>
                <SkeletonItem shape="rectangle" animation="pulse" appearance="translucent" />
              </Skeleton>
            ) : (
              renderCell(item)
            )}
          </DataGridCell>
        )}
      </DataGridRow>
    ),
    [styles.row, rowHeight],
  );

  // ── Infinite scroll: trigger onLoadMore when scrolled near the bottom ──
  const handleItemsRendered = useCallback(
    ({ visibleStopIndex }: { visibleStartIndex: number; visibleStopIndex: number }) => {
      if (!hasMore || !onLoadMore || isLoadingMore) return;
      const threshold = Math.min(10, Math.max(1, Math.floor(rows.length * 0.1)));
      if (visibleStopIndex >= rows.length - threshold) {
        onLoadMore();
      }
    },
    [hasMore, onLoadMore, isLoadingMore, rows.length],
  );

  // ── Empty / loading states ──
  // Friendly empty-state copy that distinguishes "filter matched nothing"
  // from "executed against an empty table". The successful 200 came back,
  // it just had `value: []`. Include the executed URL so the user can
  // sanity-check the criteria they sent (e.g. spotting a typo in a filter
  // value) without leaving the Results tab.
  if (rows.length === 0) {
    const responseStatus = status?.code;
    return (
      <div className={styles.emptyState} style={{ flexDirection: 'column', gap: 6, padding: '32px 16px' }}>
        <div style={{ fontSize: tokens.fontSizeBase400, fontWeight: tokens.fontWeightSemibold }}>
          No records matched
        </div>
        <div style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground2, textAlign: 'center', maxWidth: 520 }}>
          {responseStatus === 200
            ? 'The request succeeded but Dataverse returned zero rows. Check your $filter / $top — or try removing conditions to widen the result set.'
            : 'The request returned no rows.'}
        </div>
        {status && (
          <div style={{
            fontSize: tokens.fontSizeBase100,
            fontFamily: tokens.fontFamilyMonospace,
            color: tokens.colorNeutralForeground3,
            marginTop: 8,
          }}>
            HTTP {status.code} · {status.ms} ms · {(status.bytes / 1024).toFixed(1)} KB
          </div>
        )}
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <div className={styles.gridWrapper}>
        <div ref={containerRef} className={styles.gridContent}>
          <DataGrid
            aria-label={table ?? 'Results'}
            items={rows}
            columns={columns}
            // We render sort arrows ourselves (driven by $orderby state).
            sortable={false}
            resizableColumns
            resizableColumnsOptions={{ autoFitColumns: false }}
            columnSizingOptions={columnSizingOptions}
            selectionMode="multiselect"
            selectedItems={selectedItems}
            onSelectionChange={handleSelectionChange}
            getRowId={getRowId}
          >
            <DataGridHeader ref={headerRef as React.RefObject<HTMLDivElement>} style={{ paddingRight: scrollbarWidth }}>
              <DataGridRow
                style={{ minHeight: '40px', maxHeight: '40px' }}
                selectionCell={{ checkboxIndicator: { 'aria-label': 'Select all rows' } }}
              >
                {({ renderHeaderCell }) => (
                  <DataGridHeaderCell className={styles.headerCell}>
                    {renderHeaderCell()}
                  </DataGridHeaderCell>
                )}
              </DataGridRow>
            </DataGridHeader>
            <DataGridBody<Record<string, unknown>>
              itemSize={rowHeight}
              height={Math.max(0, gridDimensions.height - headerHeight)}
              listProps={{
                useIsScrolling: true,
                className: styles.body,
                onItemsRendered: handleItemsRendered,
              }}
            >
              {renderRow}
            </DataGridBody>
          </DataGrid>
        </div>

        {/* Info bar — status + size + row count + selection (moved from top strip per UX feedback) */}
        <div className={styles.infoBar}>
          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
            {status && (
              <StatusPill
                status={status.code >= 400 ? 'danger' : status.code >= 300 ? 'warning' : 'success'}
                code={status.code}
                ms={status.ms}
              />
            )}
            {status && (
              <span style={{ color: tokens.colorNeutralForeground3 }}>
                {(status.bytes / 1024).toFixed(1)} KB
              </span>
            )}
            <span>
              <strong>{rows.length.toLocaleString()}</strong> row{rows.length === 1 ? '' : 's'}
              {searchQuery && rows.length !== allRows.length && (
                <> of <strong>{allRows.length.toLocaleString()}</strong></>
              )}
              {serverTotal != null && serverTotal >= 0 && (
                <> · server total <strong>{serverTotal.toLocaleString()}</strong></>
              )}
              {hasMore && !isLoadingMore && ' · more available'}
              {isLoadingMore && (
                <span className={styles.loadingMore}>
                  <Spinner size="tiny" /> Loading more…
                </span>
              )}
            </span>
          </span>
          <span>
            {selectedItems.size > 0 && (
              <>
                <strong>{selectedItems.size}</strong> selected
              </>
            )}
          </span>
        </div>
      </div>
    </div>
  );
}
