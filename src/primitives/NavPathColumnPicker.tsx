// NavPathColumnPicker — drill-down column picker that walks single-valued
// navigation properties (N:1) to reach columns on related entities.
//
// Why this exists:
//   Per /webapi/query/filter-rows, Dataverse OData supports filtering on
//   `<nav>/<col>` and even chains like `primarycontactid/createdby/fullname`.
//   Per /webapi/query/aggregate-data, $apply/groupby supports the same path
//   form: `groupby((primarycontactid/fullname))`. Both clauses use this
//   picker so the path-construction story is identical.
//
// Why N:1 only:
//   The same docs say "You can't include conditions on collection-valued
//   navigation properties that are nested in a lookup navigation property"
//   — e.g. `primarycontactid/new_contact_account/any(...)` is forbidden.
//   By restricting drill-in to N:1 navs we make the forbidden pattern
//   impossible to construct (collection-nav filtering happens via the
//   separate lambda flow in FilterEditor).
//
// Output shape:
//   The picker returns a string in the form
//     `<col>` | `<nav>/<col>` | `<nav>/<nav>/<col>` | …
//   Encoders downstream (filterTree.ruleToOData, applyToOData) consume
//   this directly — they rewrite the trailing segment to `_<col>_value`
//   for Lookup-family columns via the existing `attrRef` helper, so a
//   path like `primarycontactid/primarycontactid` for a chained lookup
//   becomes `primarycontactid/_primarycontactid_value` automatically.

import { useEffect, useMemo, useState } from 'react';
import {
  Popover, PopoverTrigger, PopoverSurface, Button, Input, tokens,
  Caption1, Body1, makeStyles, Spinner,
} from '@fluentui/react-components';
import {
  ArrowEnter20Regular, ChevronRight16Regular,
  Search20Regular, ArrowExit20Regular, Dismiss16Regular,
} from '@fluentui/react-icons';
import { findTable, __subscribeLiveTables, type ColumnMeta, type NavProperty, type TableMeta } from '../mock/metadata';
import { useLiveTable } from '../host/useLiveMetadata';

const useStyles = makeStyles({
  // Trigger button. Fluent v9's <Button> appearance variants paint their
  // surface tokens at a specificity that can win over plain makeStyles
  // class rules — empirically on 9.73.8 the chip rendered as a near-white
  // pill in dark mode. Setting bg/fg/border via Fluent neutral tokens with
  // `!important` resolves it: the !important defeats Fluent's atomic class,
  // and the tokens themselves still flip with the FluentProvider theme.
  triggerButton: {
    minWidth: 0,
    justifyContent: 'flex-start',
    fontWeight: 400,
    backgroundColor: `${tokens.colorNeutralBackground1} !important`,
    color: `${tokens.colorNeutralForeground1} !important`,
  },
  triggerPath: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    minWidth: 0,
  },
  triggerPathSegment: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '12px',
    // Color inherits from the trigger button.
  },
  triggerSeparator: {
    color: tokens.colorNeutralForeground3,
    flexShrink: 0,
  },
  surface: {
    padding: '8px',
    width: '380px',
    maxHeight: 'min(60vh, 480px)',
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  breadcrumb: {
    display: 'flex',
    alignItems: 'center',
    flexWrap: 'wrap',
    gap: '2px',
    padding: '4px 0',
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2,
  },
  crumbChip: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '11px',
    padding: '2px 8px',
    borderRadius: tokens.borderRadiusSmall,
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground2,
    cursor: 'pointer',
    border: 'none',
    ':hover': { backgroundColor: tokens.colorNeutralBackground3Hover },
  },
  crumbChipLast: {
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    fontWeight: 600,
    cursor: 'default',
  },
  list: {
    flexGrow: 1,
    overflowY: 'auto',
    display: 'flex',
    flexDirection: 'column',
  },
  optionRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '6px 8px',
    border: 'none',
    background: 'transparent',
    cursor: 'pointer',
    textAlign: 'left',
    width: '100%',
    fontFamily: 'inherit',
    fontSize: '13px',
    borderRadius: tokens.borderRadiusSmall,
    // Native <button> elements DON'T inherit text color — they have a UA
    // default of `ButtonText` (system-dark on most platforms). That makes
    // the option label look black-on-dark inside the popover when the page
    // is in dark mode. Pin to a theme-aware foreground token so children
    // (Body1 typography and any plain spans) inherit the correct color.
    color: tokens.colorNeutralForeground1,
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      color: tokens.colorNeutralForeground1Hover,
    },
  },
  optionLabel: {
    flexGrow: 1,
    minWidth: 0,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  optionType: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '10px',
    color: tokens.colorNeutralForeground3,
    backgroundColor: tokens.colorNeutralBackground3,
    padding: '0 5px',
    borderRadius: '3px',
    flexShrink: 0,
  },
  groupHeader: {
    fontSize: '10px',
    fontWeight: 600,
    textTransform: 'uppercase',
    letterSpacing: '0.06em',
    color: tokens.colorNeutralForeground3,
    padding: '8px 8px 4px',
  },
});

export interface NavPathColumnPickerProps {
  /** Root entity logical name — anchors the start of every path. */
  rootTable: string;
  /** Slash-delimited path: `name` | `primarycontactid/fullname` | `primarycontactid/createdby/fullname`. */
  value: string;
  onChange: (path: string) => void;
  /** Optional per-column leaf gate — e.g. exclude DateTime for groupby. */
  leafFilter?: (col: ColumnMeta) => boolean;
  /** Optional column-type allowlist — used by typed filter operators. */
  allowedColumnTypes?: ReadonlyArray<ColumnMeta['attributeType']>;
  /** Placeholder when value is empty. */
  placeholder?: string;
  /** Disabled state. */
  disabled?: boolean;
  /** Display logical names instead of display names (matches DisplaySettings). */
  useLogicalNames?: boolean;
  /** Optional button size. */
  size?: 'small' | 'medium';
}

/**
 * Walks a sequence of N:1 nav segments (NO leaf column) and returns the
 * entities visited. Used while the user is DRAFTING a path inside the
 * popover: `draftNavs` contains only nav names; the leaf column is picked
 * separately and triggers commit.
 *
 * Why separate from leaf resolution: in Dataverse, a Lookup attribute and
 * the nav property that traverses it share the same name (`primarycontactid`
 * is BOTH the lookup column on `account` AND the nav property to `contact`).
 * A single-segment string can't disambiguate "the user drilled in" vs "the
 * user picked the lookup column". Tracking nav-only segments internally
 * removes the ambiguity completely.
 */
function navWalk(
  rootTable: string,
  draftNavs: string[],
): {
  tables: TableMeta[];
  current: TableMeta | undefined;
  pendingTarget?: string;
} {
  const tables: TableMeta[] = [];
  const root = findTable(rootTable);
  if (!root) return { tables, current: undefined };
  tables.push(root);
  let cursor: TableMeta | undefined = root;
  for (const seg of draftNavs) {
    const nav = cursor?.navigationProperties.find(
      (n: NavProperty) => n.name === seg && n.cardinality === 'ManyToOne',
    );
    if (!nav) return { tables, current: undefined };
    const target = findTable(nav.targetEntity);
    if (!target) return { tables, current: undefined, pendingTarget: nav.targetEntity };
    tables.push(target);
    cursor = target;
  }
  return { tables, current: cursor };
}

/**
 * Resolves an externally-committed value (the path that went on the wire).
 * Convention: the LAST segment is the leaf column; everything before is
 * a sequence of N:1 nav segments. So we navWalk all but the last, then
 * look up the leaf on the deepest entity.
 */
function resolveExternalPath(
  rootTable: string,
  path: string,
): {
  tables: TableMeta[];
  current: TableMeta | undefined;
  leaf?: ColumnMeta;
  pendingTarget?: string;
} {
  if (!path) {
    const root = findTable(rootTable);
    return { tables: root ? [root] : [], current: root };
  }
  const segs = path.split('/');
  const navSegs = segs.slice(0, -1);
  const leafSeg = segs[segs.length - 1];
  const walked = navWalk(rootTable, navSegs);
  if (!walked.current) {
    return {
      tables: walked.tables,
      current: undefined,
      pendingTarget: walked.pendingTarget,
    };
  }
  const leaf = walked.current.columns.find(c => c.logicalName === leafSeg);
  return { tables: walked.tables, current: walked.current, leaf };
}

/**
 * Display label for a path segment, given the parent entity it sits on
 * and the desired naming mode.
 */
function labelFor(
  table: TableMeta | undefined,
  segment: string,
  useLogicalNames: boolean,
): string {
  if (!table) return segment;
  const col = table.columns.find(c => c.logicalName === segment);
  if (col) return useLogicalNames ? col.logicalName : col.displayName;
  const nav = table.navigationProperties.find(n => n.name === segment);
  if (nav) {
    // Nav name is functional; use the target table's display name as a
    // friendlier breadcrumb segment when in display-name mode.
    if (useLogicalNames) return nav.name;
    const target = findTable(nav.targetEntity);
    return target?.displayName ?? nav.name;
  }
  return segment;
}

export function NavPathColumnPicker({
  rootTable, value, onChange, leafFilter, allowedColumnTypes,
  placeholder = 'Pick a column…', disabled, useLogicalNames, size = 'small',
}: NavPathColumnPickerProps) {
  const styles = useStyles();
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState('');
  /**
   * `draftNavs` is the sequence of nav segments the user is BUILDING
   * inside the popover. Crucially it does NOT include a leaf column —
   * the leaf only exists AT commit time when the user clicks a column.
   * Storing only navs avoids the `primarycontactid`-is-both-a-column-and-
   * a-nav ambiguity that broke the previous string-based draft model.
   */
  const [draftNavs, setDraftNavs] = useState<string[]>([]);

  // Bump on every metadata-registry update. Without this, `navWalk` is
  // memoized against `[rootTable, draftNavs]` only — but `navWalk` calls
  // `findTable()` internally, which depends on the live registry. When
  // a lazily-loaded target (e.g. `contact`) lands AFTER the picker mounted,
  // the registry fires, the component re-renders, but the useMemo returns
  // its stale cached value (pendingTarget: 'contact' forever, spinner
  // stuck). Subscribing to the registry version and including it in the
  // memo deps forces a recomputation on registry change.
  const [registryVersion, setRegistryVersion] = useState(0);
  useEffect(() => __subscribeLiveTables(() => setRegistryVersion(v => v + 1)), []);

  // Walk the draft to find the current entity inside the popover.
  const draftWalked = useMemo(
    () => navWalk(rootTable, draftNavs),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [rootTable, draftNavs, registryVersion],
  );

  // Lazy-load the deepest target entity if it isn't in the registry yet.
  useLiveTable(draftWalked.pendingTarget ?? null);

  // When the popover opens, pre-fill `draftNavs` with the nav segments
  // already in the committed `value` (everything before the leaf). This
  // way re-opening a path like `primarycontactid/fullname` lands the user
  // at the contact level — ready to swap the leaf column without having
  // to re-drill from root. Picking the trailing leaf is a no-op replace.
  useEffect(() => {
    if (!open) return;
    const segs = value ? value.split('/') : [];
    setDraftNavs(segs.length > 1 ? segs.slice(0, -1) : []);
    setQuery('');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  // ── Options at the current draft-walked level ──
  const current = draftWalked.current;
  const { columnsAtLevel, navsAtLevel } = useMemo(() => {
    if (!current) return { columnsAtLevel: [] as ColumnMeta[], navsAtLevel: [] as NavProperty[] };
    let cols = current.columns;
    // Always hide File / Image — never meaningful in filter/groupby paths.
    cols = cols.filter(c => c.attributeType !== 'File' && c.attributeType !== 'Image');
    if (allowedColumnTypes) cols = cols.filter(c => allowedColumnTypes.includes(c.attributeType));
    if (leafFilter) cols = cols.filter(leafFilter);
    // Only N:1 navs are eligible drill-in steps — per spec, conditions on
    // collection-valued navs nested in a lookup are forbidden. This also
    // makes the forbidden lambda-in-lookup pattern impossible to construct.
    const navs = current.navigationProperties.filter(
      (n: NavProperty) => n.cardinality === 'ManyToOne',
    );
    return { columnsAtLevel: cols, navsAtLevel: navs };
  }, [current, allowedColumnTypes, leafFilter]);

  const filteredColumns = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return columnsAtLevel;
    return columnsAtLevel.filter(c =>
      c.displayName.toLowerCase().includes(q) || c.logicalName.toLowerCase().includes(q),
    );
  }, [columnsAtLevel, query]);

  const filteredNavs = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return navsAtLevel;
    return navsAtLevel.filter(n =>
      n.name.toLowerCase().includes(q) || n.targetEntity.toLowerCase().includes(q),
    );
  }, [navsAtLevel, query]);

  // ── Path manipulation ──
  // drillInto APPENDS a nav segment to the internal nav stack — no leaf
  // is committed, the popover stays open.
  const drillInto = (navName: string) => {
    setDraftNavs(prev => [...prev, navName]);
    setQuery('');
  };
  // pickColumn finalizes the path. The wire format is the nav chain plus
  // the leaf column, joined with `/`. Empty nav chain → just the leaf.
  const pickColumn = (logicalName: string) => {
    const finalPath = draftNavs.length > 0
      ? `${draftNavs.join('/')}/${logicalName}`
      : logicalName;
    onChange(finalPath);
    setOpen(false);
  };
  // backTo(idx) — idx -1 = root, idx 0 = stay at first nav, etc.
  const backTo = (segmentIndex: number) => {
    if (segmentIndex < 0) { setDraftNavs([]); setQuery(''); return; }
    setDraftNavs(prev => prev.slice(0, segmentIndex + 1));
    setQuery('');
  };
  const clearAndClose = () => {
    onChange('');
    setOpen(false);
  };

  // ── Trigger label ──
  // External `value` follows the "navs… + leaf" convention. Resolve via
  // resolveExternalPath so every segment is labeled in the right scope.
  const triggerLabel = useMemo(() => {
    if (!value) return placeholder;
    const segs = value.split('/');
    const rootMeta = findTable(rootTable);
    if (!rootMeta) return value;
    const labels: string[] = [];
    let cursor: TableMeta | undefined = rootMeta;
    for (let i = 0; i < segs.length; i++) {
      const seg = segs[i];
      labels.push(labelFor(cursor, seg, !!useLogicalNames));
      // For every segment except the last, walk via N:1 nav.
      if (i < segs.length - 1) {
        const nav: NavProperty | undefined = cursor?.navigationProperties.find(
          (n: NavProperty) => n.name === seg && n.cardinality === 'ManyToOne',
        );
        cursor = nav ? findTable(nav.targetEntity) : undefined;
      }
    }
    return labels.join(' › ');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value, rootTable, useLogicalNames, placeholder, registryVersion]);

  // ── Breadcrumb inside the popover (built from draftNavs) ──
  // Always shows the root chip + one chip per drilled-in nav. The LAST
  // chip is the "you are here" indicator and isn't clickable; the earlier
  // chips back the user up to that level.
  const draftCrumbs = useMemo(() => {
    const rootMeta = findTable(rootTable);
    const rootLabel = useLogicalNames
      ? (rootMeta?.logicalName ?? rootTable)
      : (rootMeta?.displayName ?? rootTable);
    const out: { label: string; clickable: boolean; navIndex: number }[] = [
      { label: rootLabel, clickable: draftNavs.length > 0, navIndex: -1 },
    ];
    let cursor: TableMeta | undefined = rootMeta;
    for (let i = 0; i < draftNavs.length; i++) {
      // Label with the TARGET entity's display name — the nav itself
      // identifies which lookup we drilled through, but the user thinks
      // about the destination, not the relationship name.
      const nav: NavProperty | undefined = cursor?.navigationProperties.find(
        (n: NavProperty) => n.name === draftNavs[i] && n.cardinality === 'ManyToOne',
      );
      const target = nav ? findTable(nav.targetEntity) : undefined;
      out.push({
        label: useLogicalNames
          ? draftNavs[i]
          : (target?.displayName ?? draftNavs[i]),
        clickable: i < draftNavs.length - 1,
        navIndex: i,
      });
      cursor = target;
    }
    return out;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [draftNavs, rootTable, useLogicalNames, registryVersion]);

  // `autoSize: 'height-always'` lets Fluent's popover machinery shrink
  // the surface to fit between the trigger and the viewport edge. With
  // the static `below-start` we had before, popovers near the bottom
  // of the viewport would render off-screen and get clipped.
  return (
    <Popover
      open={open}
      onOpenChange={(_, d) => setOpen(d.open)}
      positioning={{ position: 'below', align: 'start', autoSize: 'height-always' }}
      withArrow={false}
    >
      <PopoverTrigger disableButtonEnhancement>
        <Button
          size={size}
          appearance="outline"
          disabled={disabled}
          className={styles.triggerButton}
          style={{ maxWidth: 320 }}
        >
          <span className={styles.triggerPath}>
            {value ? (
              <span className={styles.triggerPathSegment}>{triggerLabel}</span>
            ) : (
              <span style={{ color: tokens.colorNeutralForeground3 }}>{placeholder}</span>
            )}
          </span>
        </Button>
      </PopoverTrigger>
      <PopoverSurface className={styles.surface}>
        {/* Breadcrumb */}
        <div className={styles.breadcrumb}>
          {draftCrumbs.map((c, i) => (
            <span key={i} style={{ display: 'inline-flex', alignItems: 'center', gap: 2 }}>
              {i > 0 && <ChevronRight16Regular className={styles.triggerSeparator} />}
              <button
                type="button"
                className={`${styles.crumbChip} ${!c.clickable ? styles.crumbChipLast : ''}`}
                disabled={!c.clickable}
                onClick={() => c.clickable && backTo(c.navIndex)}
              >
                {c.label}
              </button>
            </span>
          ))}
          <span style={{ flexGrow: 1 }} />
          {value && (
            <Button
              size="small"
              appearance="subtle"
              icon={<Dismiss16Regular />}
              onClick={clearAndClose}
              aria-label="Clear column"
              title="Clear column"
            />
          )}
        </div>

        {/* Search */}
        <Input
          size="small"
          appearance="filled-lighter"
          placeholder="Search columns or relationships…"
          contentBefore={<Search20Regular />}
          value={query}
          onChange={(_, d) => setQuery(d.value)}
          autoFocus
        />

        {/* List */}
        <div className={styles.list}>
          {draftWalked.pendingTarget && (
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: 12 }}>
              <Spinner size="extra-small" />
              <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                Loading metadata for <code>{draftWalked.pendingTarget}</code>…
              </Caption1>
            </div>
          )}

          {filteredColumns.length === 0 && filteredNavs.length === 0 && !draftWalked.pendingTarget && (
            <Caption1 style={{ color: tokens.colorNeutralForeground3, padding: 12, fontStyle: 'italic' }}>
              No matches{query ? ` for "${query}"` : ''}.
            </Caption1>
          )}

          {filteredColumns.length > 0 && (
            <>
              <div className={styles.groupHeader}>Columns · {filteredColumns.length}</div>
              {filteredColumns.map(c => (
                <button
                  key={c.logicalName}
                  type="button"
                  className={styles.optionRow}
                  onClick={() => pickColumn(c.logicalName)}
                >
                  <Body1 className={styles.optionLabel}>
                    {useLogicalNames ? c.logicalName : c.displayName}
                  </Body1>
                  <span className={styles.optionType}>{c.attributeType}</span>
                </button>
              ))}
            </>
          )}

          {filteredNavs.length > 0 && (
            <>
              <div className={styles.groupHeader}>
                Drill into related (single-valued) · {filteredNavs.length}
              </div>
              {filteredNavs.map(n => {
                const target = findTable(n.targetEntity);
                const targetLabel = useLogicalNames
                  ? n.targetEntity
                  : (target?.displayName ?? n.targetEntity);
                return (
                  <button
                    key={n.name}
                    type="button"
                    className={styles.optionRow}
                    onClick={() => drillInto(n.name)}
                  >
                    <ArrowEnter20Regular style={{ color: tokens.colorBrandForeground1, flexShrink: 0 }} />
                    <span className={styles.optionLabel}>
                      <span style={{ fontFamily: tokens.fontFamilyMonospace, fontWeight: 600 }}>
                        {n.name}
                      </span>
                      <Caption1 style={{ display: 'block', color: tokens.colorNeutralForeground3 }}>
                        → {targetLabel}
                      </Caption1>
                    </span>
                    <span className={styles.optionType}>N:1</span>
                  </button>
                );
              })}
            </>
          )}
        </div>

        {/* Footer actions */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderTop: `1px solid ${tokens.colorNeutralStroke2}`, paddingTop: 6 }}>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            Click a column to pick · click a relationship to drill in
          </Caption1>
          <Button
            size="small"
            appearance="subtle"
            icon={<ArrowExit20Regular />}
            onClick={() => setOpen(false)}
          >
            Close
          </Button>
        </div>
      </PopoverSurface>
    </Popover>
  );
}
