import { useMemo, useState } from 'react';
import {
  Button, mergeClasses, tokens, MessageBar, MessageBarBody, Badge, Input,
  TableCellLayout, createTableColumn, type TableColumnDefinition,
  DataGrid, DataGridHeader, DataGridHeaderCell, DataGridBody, DataGridRow, DataGridCell,
} from '@fluentui/react-components';
// Plain (non-virtualized) Fluent v9 DataGrid. We tried the
// `@fluentui-contrib/react-data-grid-react-window` variant for wide-
// entity perf but the scroll-jumps + height-measurement fight wasn't
// worth the savings — most entities have ≤500 columns and the native
// browser scroll handles that fine. Revisit virtualization if/when a
// real customer reports perf issues on a wider table.
import {
  TextBulletList20Filled, ChevronRight20Regular,
  Checkmark20Filled, Search20Regular,
} from '@fluentui/react-icons';
import { useStudioStyles } from '../primitives/styles';
import { PaneHead } from './PaneHead';
import { ApplyOverridesBanner } from './ApplyOverridesBanner';
import { findTable, isCompanionLogicalReadOnly, type ColumnMeta } from '../mock/metadata';
import { groupColorVar } from '../theme/theme';
import type { RequestGroup } from '../registry/requestTypes';
import { SortableList, SortableItem } from '../primitives/Sortable';

export interface SelectEditorProps {
  table: string;
  selectedIds: string[];
  setSelectedIds: (ids: string[]) => void;
  group?: RequestGroup;
  /** When true, surface a banner — $apply replaces $select in the response shape. */
  applyActive?: boolean;
}

export function SelectEditor({ table, selectedIds, setSelectedIds, group = 'read', applyActive }: SelectEditorProps) {
  const s = useStudioStyles();
  const tbl = findTable(table);
  const [pickSel, setPickSel] = useState<string | null>(null);
  const [search, setSearch] = useState('');

  const { available, selected, orphanCount } = useMemo(() => {
    // Read picker filter:
    //   1. Honor IsValidForRead — a handful of platform-internal columns
    //      (entity-image side fields, `isprivate` etc.) are read-forbidden;
    //      the server 400s if they appear in $select.
    //   2. Hide companion logical read-only columns — the denormalized
    //      `*name` / `*yominame` text fields that Dataverse auto-populates
    //      from lookups/picklists. They're not real selectable columns in
    //      modern Web API; the resolved name comes through the
    //      `Prefer: odata.include-annotations="...FormattedValue"` hint
    //      we already surface in the Prefer pane.
    //   3. `IsValidODataAttribute` is intentionally NOT used — Microsoft
    //      reports it wrong on inherited entities (task, email, etc.).
    //      See pptbClient.getEntityAttributes for the full story.
    const all = (tbl?.columns ?? [])
      .filter(c => c.isValidForRead !== false)
      .filter(c => !isCompanionLogicalReadOnly(c));
    const filtered = (cs: ColumnMeta[]) => {
      if (!search) return cs;
      const q = search.toLowerCase();
      return cs.filter(c => c.displayName.toLowerCase().includes(q) || c.logicalName.toLowerCase().includes(q));
    };
    // Multi-select bulk-pick UX: the left "Columns" pane shows EVERY column
    // regardless of selection state. A Fluent v9 Checkbox per row reflects
    // membership in `selectedIds`; toggling it appends/removes immediately.
    // This lets users tick many columns rapidly without items disappearing
    // mid-flow (the previous "moved-to-right-pane" pattern was tedious for
    // bulk picks). Selected pane (right) still shows ordered + DnD list.
    const avail = filtered(all);
    // Build the displayed Selected list — preserve user-given order. If an
    // id doesn't match any known column on this table (e.g. a `_value` form
    // pasted via the OData parser for a column that doesn't actually exist
    // on the target entity), emit a synthetic "orphan" entry instead of
    // silently dropping it. Without this the UI would say "0 selected" and
    // "All columns (no $select emitted)" while the encoder still emits the
    // orphan id — a confusing UX mismatch.
    let orphans = 0;
    const sel2: ColumnMeta[] = selectedIds.map(id => {
      const hit = all.find(c => c.logicalName === id);
      if (hit) return hit;
      orphans++;
      // Synthetic ColumnMeta for display + click handlers. The
      // attributeType 'String' is just a placeholder so the type-badge
      // renders; the row gets a distinct visual style via
      // displayName-prefix-only-when-orphan.
      return {
        logicalName: id,
        displayName: `(unknown) ${id}`,
        attributeType: 'String',
        format: 'Text',
        maxLength: 0,
      } as ColumnMeta;
    });
    return { available: avail, selected: sel2, orphanCount: orphans };
  }, [tbl, selectedIds, search]);


  // ── DataGrid setup for the Columns pane ──
  //
  // Single source of truth: `selectedIds` (ordered). DataGrid's
  // selectedItems is derived from it as a Set. On selection change we
  // diff against the previous selectedIds to preserve insertion order
  // (DataGrid emits an unordered Set; we append additions, remove subtractions).
  //
  // The DataGrid auto-prepends a checkbox column when `selectionMode` is
  // set, so we only define the three data columns: Display Name, Logical
  // Name, Attribute Type. Each is sortable via the built-in compare fns.
  type GridItem = { logicalName: string; displayName: string; attributeType: string };
  const gridItems: GridItem[] = useMemo(
    () => available.map(c => ({
      logicalName: c.logicalName,
      displayName: c.displayName,
      attributeType: c.attributeType,
    })),
    [available],
  );

  const gridColumns: TableColumnDefinition<GridItem>[] = useMemo(() => [
    createTableColumn<GridItem>({
      columnId: 'displayName',
      compare: (a, b) => a.displayName.localeCompare(b.displayName),
      renderHeaderCell: () => 'Display name',
      renderCell: (item) => (
        <TableCellLayout truncate>
          {item.displayName}
        </TableCellLayout>
      ),
    }),
    createTableColumn<GridItem>({
      columnId: 'logicalName',
      compare: (a, b) => a.logicalName.localeCompare(b.logicalName),
      renderHeaderCell: () => 'Logical name',
      renderCell: (item) => (
        <TableCellLayout truncate>
          <code style={{
            fontFamily: tokens.fontFamilyMonospace,
            fontSize: tokens.fontSizeBase200,
            color: tokens.colorNeutralForeground2,
          }}>
            {item.logicalName}
          </code>
        </TableCellLayout>
      ),
    }),
    createTableColumn<GridItem>({
      columnId: 'attributeType',
      compare: (a, b) => a.attributeType.localeCompare(b.attributeType),
      renderHeaderCell: () => 'Type',
      renderCell: (item) => (
        <TableCellLayout>
          <Badge appearance="outline" size="small">{item.attributeType}</Badge>
        </TableCellLayout>
      ),
    }),
  ], []);

  const selectedSetForGrid = useMemo(() => new Set<string | number>(selectedIds), [selectedIds]);

  // Diff the incoming Set against the current ordered list. Anything new
  // gets appended; anything missing gets removed (preserving the order of
  // the survivors).
  const onGridSelectionChange = (_evt: unknown, data: { selectedItems: Set<string | number> }) => {
    const incoming = data.selectedItems;
    const survivors = selectedIds.filter(id => incoming.has(id));
    const additions: string[] = [];
    for (const id of incoming) {
      const s = String(id);
      if (!selectedIds.includes(s)) additions.push(s);
    }
    setSelectedIds([...survivors, ...additions]);
  };

  const moveLeft = () => {
    if (pickSel) {
      setSelectedIds(selectedIds.filter(id => id !== pickSel));
      setPickSel(null);
    }
  };
  const moveUp = () => {
    if (!pickSel) return;
    const i = selectedIds.indexOf(pickSel);
    if (i > 0) {
      const next = [...selectedIds];
      [next[i - 1], next[i]] = [next[i], next[i - 1]];
      setSelectedIds(next);
    }
  };
  const moveDown = () => {
    if (!pickSel) return;
    const i = selectedIds.indexOf(pickSel);
    if (i >= 0 && i < selectedIds.length - 1) {
      const next = [...selectedIds];
      [next[i], next[i + 1]] = [next[i + 1], next[i]];
      setSelectedIds(next);
    }
  };

  return (
    <div>
      <PaneHead
        icon={TextBulletList20Filled}
        title="$select"
        sub="Pick which columns the response should include. Empty = all (warned, costly)."
        group={group}
      >
        <Input
          contentBefore={<Search20Regular />}
          placeholder="Filter columns…"
          value={search}
          onChange={(_, d) => setSearch(d.value)}
          size="small"
          style={{ width: 220 }}
        />
        <Badge appearance="ghost">{selected.length} selected</Badge>
      </PaneHead>

      {applyActive && <ApplyOverridesBanner clause="$select" />}

      {/* "No $select" warning gates off `selectedIds.length` (the actual
          state), not `selected.length` (the displayed list). An orphan id
          on a parsed query — i.e. a column logical name that doesn't
          exist in the target table's metadata — counts as state for the
          encoder but doesn't render in the Selected list, so a length-of-
          displayed mismatch was previously firing this banner falsely. */}
      {selectedIds.length === 0 && (
        <MessageBar layout="multiline" intent="warning" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            No <code>$select</code> — Dataverse will return <strong>all columns</strong>. Pick a subset to keep responses small.
          </MessageBarBody>
        </MessageBar>
      )}

      {/* Surface parser-orphans separately so the user can spot and
          remove them. They render in the Selected list with a `(unknown)`
          prefix and the user can remove via the left-arrow / double-click. */}
      {orphanCount > 0 && (
        <MessageBar layout="multiline" intent="warning" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            <strong>{orphanCount}</strong> selected column{orphanCount === 1 ? '' : 's'} not found in this entity's metadata — likely an unknown <code>_&lt;x&gt;_value</code> form pasted from a URL. They're kept in state for round-trip faithfulness (the encoder still emits them), but DRS can't type-check them. Remove if unintended.
          </MessageBarBody>
        </MessageBar>
      )}

      <div className={s.selectGrid}>
        <div className={s.colList}>
          <div className={s.colListH}>
            Columns · {available.length}
            {selectedIds.length > 0 && (
              <span style={{ marginLeft: 'auto', fontSize: 10, color: tokens.colorNeutralForeground3 }}>
                {selectedIds.length} selected
              </span>
            )}
          </div>
          {/* Fluent v9 DataGrid for the column picker:
                • Built-in multi-select checkbox column (auto-prepended).
                • Sortable headers — click "Display name", "Logical name",
                  or "Type" to sort by that column.
                • `selectedItems` is derived from `selectedIds` (the
                  ordered state). `onSelectionChange` diffs against the
                  previous ordered list so insertion order survives. */}
          {/* Plain DataGrid wrapper — native browser scroll, no
              virtualization. Bounded height comes from `s.colList`'s
              `maxHeight: 60vh`. Reintroduce virtualization later if
              a customer reports perf on a wide entity. */}
          <div style={{ flex: 1, minHeight: 0, overflow: 'auto' }}>
            {available.length === 0 ? (
              <div style={{ padding: '20px 14px', textAlign: 'center', fontSize: 12, color: tokens.colorNeutralForeground3 }}>
                <em>No columns match "{search}"</em>
              </div>
            ) : (
              <DataGrid
                items={gridItems}
                columns={gridColumns}
                sortable
                selectionMode="multiselect"
                selectedItems={selectedSetForGrid}
                onSelectionChange={onGridSelectionChange}
                getRowId={(item) => (item as GridItem).logicalName}
                size="small"
                style={{ minWidth: 0 }}
              >
                <DataGridHeader>
                  <DataGridRow selectionCell={{ checkboxIndicator: { 'aria-label': 'Select all columns' } }}>
                    {({ renderHeaderCell }) => (
                      <DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>
                    )}
                  </DataGridRow>
                </DataGridHeader>
                <DataGridBody<GridItem>>
                  {({ item, rowId }) => (
                    <DataGridRow<GridItem>
                      key={rowId}
                      selectionCell={{ checkboxIndicator: { 'aria-label': `Select ${item.displayName}` } }}
                    >
                      {({ renderCell }) => (
                        <DataGridCell>{renderCell(item)}</DataGridCell>
                      )}
                    </DataGridRow>
                  )}
                </DataGridBody>
              </DataGrid>
            )}
          </div>
        </div>

        {/* Center column: keyboard-fallback arrow controls. The right-pane
            "add" arrow is gone (checkbox handles that). The left-arrow still
            removes from Selected for keyboard-only users; up/down reorders
            Selected for parity with DnD. */}
        <div className={s.colArrows}>
          <Button icon={<ChevronRight20Regular style={{ transform: 'rotate(180deg)' }} />} appearance="subtle" onClick={moveLeft} disabled={!pickSel} aria-label="Remove from selected" />
          <Button icon={<ChevronRight20Regular style={{ transform: 'rotate(-90deg)' }} />} appearance="subtle" onClick={moveUp}   disabled={!pickSel} aria-label="Move up" />
          <Button icon={<ChevronRight20Regular style={{ transform: 'rotate(90deg)' }} />}  appearance="subtle" onClick={moveDown} disabled={!pickSel} aria-label="Move down" />
        </div>

        <div className={s.colList}>
          <div className={s.colListH}>
            Selected · {selectedIds.length}
            <span style={{ marginLeft: 'auto', fontSize: 10, color: tokens.colorNeutralForeground3 }}>order matters</span>
          </div>
          <div className={s.colListBody}>
            {selectedIds.length === 0 && (
              <div style={{ padding: '20px 14px', textAlign: 'center', fontSize: 12, color: tokens.colorNeutralForeground3 }}>
                <em>All columns</em> (no <code>$select</code> emitted)
              </div>
            )}
            {/* The Selected pane is the authoritative order Dataverse will
                use for column projection. Drag-reorder via the grip on
                each row. Keyboard: Tab to grip → Space (pick up) →
                ArrowUp/Down → Space (drop). Up/Down arrow buttons in the
                colArrows toolbar still work as an alternative. */}
            <SortableList
              ids={selectedIds}
              onReorder={(from, to) => {
                const next = [...selectedIds];
                const [moved] = next.splice(from, 1);
                next.splice(to, 0, moved);
                setSelectedIds(next);
              }}
            >
              {selected.map(c => {
                const isOrphan = c.displayName.startsWith('(unknown) ');
                return (
                  <SortableItem key={c.logicalName} id={c.logicalName}>
                    {({ gripProps, Grip }) => (
                      <button
                        type="button"
                        className={mergeClasses(s.colRow, pickSel === c.logicalName && s.colRowSelected)}
                        onClick={() => setPickSel(c.logicalName)}
                        onDoubleClick={() => setSelectedIds(selectedIds.filter(id => id !== c.logicalName))}
                        style={isOrphan ? { fontStyle: 'italic', opacity: 0.85 } : undefined}
                      >
                        <span {...gripProps} aria-label={`Drag to reorder ${c.displayName}`}>
                          <Grip />
                        </span>
                        <Checkmark20Filled
                          style={{
                            width: 14, height: 14,
                            color: isOrphan ? tokens.colorPaletteRedForeground1 : groupColorVar(group),
                          }}
                        />
                        <span>{c.displayName}</span>
                        {isOrphan ? (
                          <Badge appearance="tint" color="warning" size="small">unknown</Badge>
                        ) : (
                          <span className={s.colTypeBadge}>{c.attributeType}</span>
                        )}
                        <span className={s.colMeta}>{c.logicalName}</span>
                      </button>
                    )}
                  </SortableItem>
                );
              })}
            </SortableList>
          </div>
        </div>
      </div>
    </div>
  );
}
