// FilterEditor — the heart of the studio.
//
// Renders the v2.2 visual vocabulary 1:1:
//   • Group card with combinator toggle (AND / OR), drag handle, condition count
//   • Plain `qb-rule` row (column · operator · value)
//   • `fn` block — Dataverse function (ƒ badge, Microsoft.Dynamics.CRM. prefix,
//     description, PropertyName/PropertyValue rows below)
//   • `lam` block — lambda over a collection nav (λ badge, nav-prop, any/all,
//     alias, nested predicate using the alias prefix)
//   • Generated OData card at the bottom of the page
//
// Each rendered node is fully recursive — groups inside lambdas inside groups…

import { useEffect, useMemo, useState } from 'react';
import {
  Button,
  Combobox,
  Option,
  Tooltip,
  Badge,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  Popover,
  PopoverTrigger,
  PopoverSurface,
  ToggleButton,
  Switch,
  tokens,
} from '@fluentui/react-components';
import { SegmentedToggle } from '../../primitives/SegmentedToggle';
import {
  Filter20Filled,
  BranchFork20Regular,
  Delete20Regular,
  Copy20Regular,
  Add20Regular,
  ReOrderDotsHorizontal20Regular,
} from '@fluentui/react-icons';
import { useStudioStyles } from '../../primitives/styles';
import { groupColorVar } from '../../theme/theme';
import { PaneHead } from '../PaneHead';
import {
  findColumn,
  findTable,
  isCompanionLogicalReadOnly,
  resolveNavPath,
  type ColumnMeta,
  type TableMeta,
} from '../../mock/metadata';
import { useLiveTable } from '../../host/useLiveMetadata';
import { ApplyOverridesBanner } from '../ApplyOverridesBanner';
import {
  type FilterGroup,
  type FilterRule,
  type FilterFunctionNode,
  type FilterLambdaNode,
  type FilterNode,
  type Combinator,
  countRules,
  groupToOData,
  addChild,
  patchNode,
  removeNode,
  reorderInGroup,
  validateTree,
  newId,
  defaultRule,
  defaultGroup,
  defaultFunction,
  defaultLambda,
} from './filterTree';
import { SortableList, SortableItem, type GripProps } from '../../primitives/Sortable';
import { findOperator, OPERATORS, operatorsFor, OP_CATEGORIES } from './operators';
import { FilterValueInput } from './FilterValueInput';
import { OperatorPicker } from './OperatorPicker';
import { NavPathColumnPicker } from '../../primitives/NavPathColumnPicker';
import { AntipatternIcon } from '../../primitives/AntipatternIcon';
import { detectColumnAntipatterns } from '../../engine/antipatterns';

/**
 * Resolves the leaf `ColumnMeta` for a slash-delimited nav-path PLUS the
 * `TableMeta` it lives on. The owner table matters for `useColumnDetail`
 * — when filtering on `primarycontactid/accountrolecode`, we need to
 * fetch the OptionSet against the CONTACT entity, not the root (account).
 *
 * Returns empty fields if any intermediate target hasn't been loaded yet
 * (the picker has its own loading affordance for that case).
 */
function resolvePathLeaf(
  rootTable: TableMeta,
  path: string,
): {
  col?: ColumnMeta;
  ownerTable?: TableMeta;
} {
  if (!path) return {};
  // Delegates to the canonical metadata-driven resolver (one source of truth
  // shared with the $filter encoder, picker, and pre-warmer) so leaf type
  // resolution is identical everywhere.
  const r = resolveNavPath(rootTable, path);
  return { col: r.leaf, ownerTable: r.ownerTable };
}

/** Convenience: just the leaf column. */
function resolvePathLeafColumn(rootTable: TableMeta, path: string): ColumnMeta | undefined {
  return resolvePathLeaf(rootTable, path).col;
}

/**
 * Walks the nav segments of a path and returns the FIRST target entity
 * that hasn't been loaded into the live registry yet. Returns null when
 * the whole chain is resolved. RuleRow uses this to fire `useLiveTable`
 * — once that load completes, the re-render walks further and may
 * surface the next pending target, and so on.
 */
function findPendingNavTarget(rootTable: TableMeta, path: string): string | null {
  if (!path || !path.includes('/')) return null;
  // Same canonical resolver — `pendingTarget` is the first nav hop whose
  // entity isn't in the registry yet. Drives the per-row `useLiveTable` so a
  // saved/pasted nav-path resolves level-by-level as targets load.
  return resolveNavPath(rootTable, path).pendingTarget ?? null;
}
import type { RequestGroup } from '../../registry/requestTypes';

export interface FilterEditorProps {
  table: string;
  tree: FilterGroup;
  setTree: (next: FilterGroup) => void;
  group?: RequestGroup;
  /** Estimated URL byte length (drives URL warnings) */
  urlBytes?: number;
  /** When true, show a compact variant (no banner, no generated card) — used inside $expand */
  compact?: boolean;
  /** When true, $apply is on at the root and this $filter is ignored — show a warning. */
  applyActive?: boolean;
}

export function FilterEditor({
  table,
  tree,
  setTree,
  group = 'read',
  urlBytes = 0,
  compact = false,
  applyActive = false,
}: FilterEditorProps) {
  const s = useStudioStyles();
  const tbl = findTable(table);
  const total = useMemo(() => countRules(tree), [tree]);
  const odata = useMemo(() => (tbl ? groupToOData(tree, tbl) : ''), [tree, tbl]);
  const v = validateTree(tree, urlBytes);

  // Note: we do NOT pre-warm collection-nav target tables here. On wide
  // entities like `incident` there can be 50+ relationships, and warming
  // every target up-front trips Dataverse's 100-concurrent-request cap.
  // Each LambdaBlock fetches its *own* target only when the user actually
  // picks that nav (see `useLiveTable(nav?.targetEntity)` inside LambdaBlock).
  if (!tbl) return <div>Pick a target table first.</div>;

  const update = (id: string, patch: Parameters<typeof patchNode>[2]) =>
    setTree(patchNode(tree, id, patch));
  const removeId = (id: string) => setTree(removeNode(tree, id));
  const onReorderChildren = (parentId: string, from: number, to: number) =>
    setTree(reorderInGroup(tree, parentId, from, to));
  // `levelTable` + `levelAlias` are pushed in by FilterGroupCard / LambdaBlock
  // so newly-added rules know which entity's columns to seed from and
  // whether to inject a lambda-alias prefix. Without this, a rule added
  // inside `incident_customer_accounts/any(i: …)` ended up with `col`
  // set to the OUTER table's first column with no prefix — Dataverse
  // then complained "<col> is property of entity 'accounts'. Conditions
  // on property other than current navigation property in any/all is not
  // supported." (the exact symptom the user hit).
  const addToGroup = (
    parentId: string,
    kind: 'rule' | 'group' | 'function' | 'lambda',
    extra?: string,
    levelTable: TableMeta = tbl,
    levelAlias?: string,
  ) => {
    const firstCol = levelTable.columns[0].logicalName;
    const prefix = levelAlias ? `${levelAlias}/` : '';
    let child: FilterNode;
    if (kind === 'rule') child = defaultRule(prefix + firstCol);
    else if (kind === 'group') child = defaultGroup(prefix + firstCol);
    // defaultFunction picks a column compatible with the chosen function's allowedTypes —
    // e.g. picking `LastXDays` lands on a DateTime column, not the primary key.
    // NOTE: Dataverse query functions inside lambdas use PropertyName *without*
    // the alias prefix (the lambda scope is implicit); we therefore don't prefix
    // function `col` here even when levelAlias is set.
    else if (kind === 'function') child = defaultFunction(levelTable, extra ?? 'LastXDays');
    else {
      // Nested lambda — Dataverse rejects this in practice, but if it
      // ever ships we'd anchor on the current level's collection navs.
      const collections = levelTable.navigationProperties.filter(
        (n) => n.cardinality !== 'ManyToOne',
      );
      const nav = extra ?? collections[0]?.name;
      if (!nav) return;
      const navMeta = levelTable.navigationProperties.find((n) => n.name === nav);
      const target = navMeta ? findTable(navMeta.targetEntity) : undefined;
      const targetFirstCol = target?.columns[0].logicalName ?? firstCol;
      child = defaultLambda(nav, targetFirstCol);
    }
    setTree(addChild(tree, parentId, child));
  };

  const header = !compact && (
    <PaneHead
      icon={Filter20Filled}
      title="$filter"
      sub="Combine rules with AND / OR · drop in Dataverse query functions · filter on related collections with λ any/all."
      group={group}
    >
      <Badge appearance="ghost">
        {total} rule{total === 1 ? '' : 's'}
      </Badge>
    </PaneHead>
  );

  const banner = !compact && total === 0 && (
    <MessageBar layout="multiline" intent="info" style={{ marginBottom: 12 }}>
      <MessageBarBody>
        <MessageBarTitle>No filter</MessageBarTitle>
        All rows will be returned (subject to <code>$top</code> / server caps). Use{' '}
        <strong>+ Add rule</strong> for a column comparison, <strong>Dataverse function</strong>
        &nbsp;(<code>ƒ</code>) for the built-in named conditions (<code>LastXDays</code>,{' '}
        <code>InFiscalYear</code>, <code>EqualUserId</code>, …), or <strong>Add lambda</strong>
        &nbsp;(<code>λ</code>) to filter on a related collection (e.g.{' '}
        <code>contact_customer_accounts/any(c: c/jobtitle eq 'CEO')</code>). <br />
        <br />
        <strong>Negation:</strong> toggle <code>not</code> on individual <code>contains</code> /
        <code>startswith</code> / <code>endswith</code> rules. For everything else, pick the
        explicit&nbsp;<code>Not*</code> sibling from the function list (<code>NotIn</code>,{' '}
        <code>NotBetween</code>, <code>NotEqualUserId</code>,<code>DoesNotContainValues</code>, …).
        Dataverse rejects <code>not</code> on groups and on <code>Microsoft.Dynamics.CRM.*</code>{' '}
        functions directly.
      </MessageBarBody>
    </MessageBar>
  );

  const warnings = !compact && (
    <>
      {v.ruleWarn && !v.ruleError && (
        <MessageBar layout="multiline" intent="warning" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            <strong>{v.ruleCount} conditions</strong> — Dataverse caps at 500 per query. Consider
            compressing OR groups with <code>In(...)</code>.
          </MessageBarBody>
        </MessageBar>
      )}
      {v.ruleError && (
        <MessageBar layout="multiline" intent="error" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            <strong>{v.ruleCount} conditions</strong> exceeds the{' '}
            <strong>500-condition hard cap</strong>.
          </MessageBarBody>
        </MessageBar>
      )}
      {v.urlWarn && !v.urlError && (
        <MessageBar layout="multiline" intent="warning" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            URL is <strong>{(urlBytes / 1024).toFixed(1)} KB</strong> — GET caps at 32 KB. Switch to{' '}
            <code>$batch</code> for larger requests.
          </MessageBarBody>
        </MessageBar>
      )}
    </>
  );

  return (
    <div>
      {header}
      {applyActive && !compact && <ApplyOverridesBanner clause="$filter" />}
      {banner}
      {warnings}

      <FilterGroupCard
        group={tree}
        depth={0}
        table={tbl}
        requestGroup={group}
        onUpdate={update}
        onAdd={addToGroup}
        onRemove={removeId}
        onReorderChildren={onReorderChildren}
      />

      {!compact && total > 0 && (
        <div className={s.generatedCard}>
          <div className={s.generatedLabel}>Generated</div>
          <div className={s.generatedCode}>
            <span className={s.generatedKey}>$filter</span>=<span>{odata}</span>
          </div>
          <div style={{ marginTop: 8, display: 'flex', justifyContent: 'flex-end' }}>
            <Tooltip content="Copy OData expression" relationship="label">
              <Button
                icon={<Copy20Regular />}
                appearance="subtle"
                size="small"
                onClick={() => navigator.clipboard?.writeText(odata)}
              >
                Copy
              </Button>
            </Tooltip>
          </div>
        </div>
      )}
    </div>
  );
}

// ============================================================
// Group card — recursive
// ============================================================
interface GroupCardProps {
  group: FilterGroup;
  depth: number;
  table: TableMeta;
  requestGroup: RequestGroup;
  /**
   * Lambda alias propagated from the enclosing `LambdaBlock`. When set,
   * newly added rules / nested groups inside this group must prefix
   * their column with `<alias>/` so the filter expression remains
   * scoped to the lambda variable (otherwise Dataverse rejects it with
   * "Conditions on property other than current navigation property in
   * any/all is not supported").
   */
  lambdaAlias?: string;
  onUpdate: (id: string, patch: Parameters<typeof patchNode>[2]) => void;
  onAdd: (
    parentId: string,
    kind: 'rule' | 'group' | 'function' | 'lambda',
    extra?: string,
    levelTable?: TableMeta,
    levelAlias?: string,
  ) => void;
  onRemove: (id: string) => void;
  /** Reorder children within THIS group's `rules` array. Threaded down
   *  from FilterEditor so the encoder mutates the tree by parent id. */
  onReorderChildren: (parentId: string, from: number, to: number) => void;
  /** When set, this group is rendered AS a sortable child of a parent
   *  group — apply gripProps to the group-card's own drag handle. */
  gripProps?: GripProps;
}

function FilterGroupCard({
  group,
  depth,
  table,
  requestGroup,
  lambdaAlias,
  onUpdate,
  onAdd,
  onRemove,
  onReorderChildren,
  gripProps,
}: GroupCardProps) {
  const s = useStudioStyles();
  const stripeGroups: RequestGroup[] = ['read', 'relate', 'write', 'execute', 'binary'];
  const stripe =
    depth === 0
      ? groupColorVar(requestGroup)
      : groupColorVar(stripeGroups[depth % stripeGroups.length]);
  const setComb = (c: Combinator) => onUpdate(group.id, { combinator: c });
  const toggleNegated = () => onUpdate(group.id, { negated: !group.negated });
  const conditionWord =
    group.rules.length === 1 ? '1 condition' : `${group.rules.length} conditions`;

  return (
    <div
      className={s.filterGroup}
      style={{ borderLeftColor: stripe, marginBottom: depth === 0 ? 0 : 6 }}
    >
      <div className={s.filterGroupHeader}>
        {depth > 0 && (
          <span {...(gripProps ?? {})} aria-label={gripProps ? 'Drag to reorder group' : undefined}>
            <ReOrderDotsHorizontal20Regular className={s.dragHandle} aria-hidden={!gripProps} />
          </span>
        )}
        <SegmentedToggle ariaLabel="Combinator">
          <ToggleButton
            checked={group.combinator === 'and'}
            onClick={() => setComb('and')}
            style={
              group.combinator === 'and' ? { backgroundColor: stripe, color: '#fff' } : undefined
            }
          >
            AND
          </ToggleButton>
          <ToggleButton
            checked={group.combinator === 'or'}
            onClick={() => setComb('or')}
            style={
              group.combinator === 'or' ? { backgroundColor: stripe, color: '#fff' } : undefined
            }
          >
            OR
          </ToggleButton>
        </SegmentedToggle>
        {/* Group-level NOT — empirically supported by Dataverse (tests
            G.1–G.7). When ON, the encoder wraps the group's body in
            `not (…)`. Caveat: `not (<group with a Microsoft.Dynamics.CRM.* fn>)`
            is rejected by Dataverse (test G.9); validateRequest surfaces
            a warning in that case. */}
        <Tooltip
          content={
            group.negated
              ? 'Group is negated — emits `not (…)`. Click to remove.'
              : 'Negate group — emits `not (…)`. Allowed by Dataverse on any group EXCEPT one containing `Microsoft.Dynamics.CRM.*` functions.'
          }
          relationship="description"
        >
          <ToggleButton
            checked={!!group.negated}
            onClick={toggleNegated}
            size="small"
            appearance={group.negated ? 'primary' : 'subtle'}
            style={{
              minWidth: 44,
              fontFamily: tokens.fontFamilyMonospace,
              fontWeight: 700,
              ...(group.negated
                ? { backgroundColor: tokens.colorPaletteRedBackground3, color: '#fff' }
                : {}),
            }}
            aria-label={group.negated ? 'Disable group NOT' : 'Negate group'}
          >
            NOT
          </ToggleButton>
        </Tooltip>
        <span style={{ fontSize: 12, color: tokens.colorNeutralForeground3 }}>{conditionWord}</span>
        <div style={{ flexGrow: 1 }} />
        {depth === 0 && (
          <Badge
            appearance="tint"
            size="small"
            style={{
              fontFamily: tokens.fontFamilyMonospace,
              fontWeight: 700,
            }}
          >
            λ supported
          </Badge>
        )}
        {depth > 0 && (
          <Tooltip content="Remove group" relationship="label">
            <Button
              icon={<Delete20Regular />}
              appearance="subtle"
              size="small"
              onClick={() => onRemove(group.id)}
              aria-label="Remove group"
            />
          </Tooltip>
        )}
      </div>

      <div className={s.filterGroupBody}>
        {group.rules.length === 0 && (
          <div
            style={{
              padding: '8px 4px',
              fontSize: 12,
              color: tokens.colorNeutralForeground3,
              fontStyle: 'italic',
            }}
          >
            Empty group — add a condition below.
          </div>
        )}
        {/* Drag-reorder children within THIS group only. Cross-group drag
            is intentionally not supported — would require complex move
            semantics (different lambda scopes, alias rewriting, etc.).
            Each child renders its existing static drag handle but with
            dnd-kit listeners spread on the wrapping span. */}
        <SortableList
          ids={group.rules.map((c) => c.id)}
          onReorder={(from, to) => onReorderChildren(group.id, from, to)}
        >
          {group.rules.map((child) => (
            <SortableItem key={child.id} id={child.id}>
              {({ gripProps: childGripProps }) => (
                <NodeRenderer
                  node={child}
                  parentTable={table}
                  depth={depth}
                  requestGroup={requestGroup}
                  lambdaAlias={lambdaAlias}
                  onUpdate={onUpdate}
                  onAdd={onAdd}
                  onRemove={onRemove}
                  onReorderChildren={onReorderChildren}
                  gripProps={childGripProps}
                />
              )}
            </SortableItem>
          ))}
        </SortableList>
      </div>

      <div className={s.filterGroupFooter}>
        <AddMenu group={group.id} table={table} lambdaAlias={lambdaAlias} onAdd={onAdd} />
      </div>
    </div>
  );
}

// ============================================================
// Add menu — split into 4 kinds
// ============================================================
function AddMenu({
  group,
  table,
  lambdaAlias,
  onAdd,
}: {
  group: string;
  table: TableMeta;
  lambdaAlias?: string;
  onAdd: (
    parentId: string,
    kind: 'rule' | 'group' | 'function' | 'lambda',
    extra?: string,
    levelTable?: TableMeta,
    levelAlias?: string,
  ) => void;
}) {
  const collections = table.navigationProperties.filter((n) => n.cardinality !== 'ManyToOne');
  // Every onAdd call below propagates the LEVEL's table + alias so that
  // FilterEditor.addToGroup seeds the new rule/group/function from the
  // correct entity's columns and prefixes lambda-scoped rules.
  return (
    <>
      <Button
        size="small"
        appearance="subtle"
        icon={<Add20Regular />}
        onClick={() => onAdd(group, 'rule', undefined, table, lambdaAlias)}
      >
        Add rule
      </Button>
      <Button
        size="small"
        appearance="subtle"
        icon={<BranchFork20Regular />}
        onClick={() => onAdd(group, 'group', undefined, table, lambdaAlias)}
      >
        Add group
      </Button>
      {/* Dataverse function picker — available at root AND inside lambdas.
          Empirical tests 6.5b and 13.15 confirm Dataverse query functions
          work inside `/any(c: …)` predicates as long as `PropertyName` is
          the BARE attribute logical name (no alias prefix, no `_value`
          form, no nav-path). The encoder enforces all three rules; the
          UI surfaces every category. The previous "root-only" gate was
          based on docs silence, not actual behavior. */}
      {/* `autoSize: 'height-always'` keeps the long DV-function list from
          overflowing the viewport and pushing the page layout down (the
          ~60 function options easily exceed available height on shorter
          screens). Floating-UI shrinks the popover to fit. */}
      <Menu positioning={{ position: 'below', align: 'start', autoSize: 'height-always' }}>
        <MenuTrigger disableButtonEnhancement>
          <Button
            size="small"
            appearance="subtle"
            style={{ color: '#8764b8' }}
            icon={
              <span
                style={{
                  fontFamily: tokens.fontFamilyMonospace,
                  fontWeight: 700,
                  fontSize: 14,
                  lineHeight: 1,
                  color: '#8764b8',
                }}
              >
                ƒ
              </span>
            }
          >
            Dataverse function
          </Button>
        </MenuTrigger>
        <MenuPopover style={{ maxHeight: 460, overflowY: 'auto' }}>
          <MenuList>
            {OP_CATEGORIES.filter((c) =>
              [
                'date-relative',
                'date-rolling',
                'fiscal',
                'range',
                'set',
                'choices',
                'hierarchy',
                'user-context',
              ].includes(c.id),
            ).map((cat) => (
              <span key={cat.id}>
                <span
                  style={{
                    display: 'block',
                    padding: '6px 10px',
                    fontSize: 10,
                    fontWeight: 600,
                    letterSpacing: '0.06em',
                    textTransform: 'uppercase',
                    color: tokens.colorNeutralForeground3,
                  }}
                >
                  {cat.label}
                </span>
                {OPERATORS.filter((o) => o.category === cat.id)
                  .slice(0, 12)
                  .map((op) => (
                    <MenuItem
                      key={op.id}
                      onClick={() => onAdd(group, 'function', op.id, table, lambdaAlias)}
                      secondaryContent={
                        <span
                          style={{
                            fontFamily: tokens.fontFamilyMonospace,
                            fontSize: 10,
                            color: tokens.colorNeutralForeground3,
                          }}
                        >
                          {op.odata.replace('Microsoft.Dynamics.CRM.', '')}
                        </span>
                      }
                    >
                      {op.label}
                    </MenuItem>
                  ))}
              </span>
            ))}
          </MenuList>
        </MenuPopover>
      </Menu>
      {collections.length > 0 && (
        <AddLambdaPicker group={group} table={table} lambdaAlias={lambdaAlias} onAdd={onAdd} />
      )}
    </>
  );
}

/**
 * Searchable Add-lambda picker. Replaces the previous MenuPopover (which
 * had no filtering and clipped off-viewport on entities with many
 * relationships). Uses Fluent v9 Popover + Combobox so the user can type
 * to narrow by nav name or target entity, and the surface auto-positions
 * inside the viewport.
 */
function AddLambdaPicker({
  group,
  table,
  lambdaAlias,
  onAdd,
}: {
  group: string;
  table: TableMeta;
  lambdaAlias?: string;
  onAdd: (
    parentId: string,
    kind: 'rule' | 'group' | 'function' | 'lambda',
    extra?: string,
    levelTable?: TableMeta,
    levelAlias?: string,
  ) => void;
}) {
  const collections = table.navigationProperties.filter((n) => n.cardinality !== 'ManyToOne');
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState('');

  const matches = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return collections;
    return collections.filter(
      (n) => n.name.toLowerCase().includes(q) || n.targetEntity.toLowerCase().includes(q),
    );
  }, [collections, query]);

  return (
    // Structured positioning with `autoSize: 'height-always'` — see the
    // OperatorPicker comment for the rationale. Without it, an entity with
    // many collection navs (account has ~50+) renders a popover that
    // overflows the viewport and pushes the page layout.
    <Popover
      open={open}
      onOpenChange={(_, d) => {
        setOpen(d.open);
        if (!d.open) setQuery('');
      }}
      positioning={{ position: 'below', align: 'start', autoSize: 'height-always' }}
      withArrow={false}
    >
      <PopoverTrigger disableButtonEnhancement>
        <Button
          size="small"
          appearance="subtle"
          style={{ color: tokens.colorBrandForeground1 }}
          icon={
            <span
              style={{
                fontFamily: tokens.fontFamilyMonospace,
                fontWeight: 700,
                fontSize: 14,
                lineHeight: 1,
                color: tokens.colorBrandForeground1,
              }}
            >
              λ
            </span>
          }
        >
          Add lambda
        </Button>
      </PopoverTrigger>
      <PopoverSurface
        style={{
          padding: 8,
          width: 360,
          maxHeight: 'min(60vh, 460px)',
          overflowY: 'auto',
        }}
      >
        <div
          style={{
            fontSize: 10,
            fontWeight: 600,
            letterSpacing: '0.06em',
            textTransform: 'uppercase',
            color: tokens.colorNeutralForeground3,
            padding: '4px 4px 8px',
          }}
        >
          Collection navigations · {collections.length}
        </div>
        <Combobox
          freeform
          size="small"
          placeholder="Search by nav or target entity…"
          value={query}
          onChange={(e) => setQuery((e.target as HTMLInputElement).value)}
          onOptionSelect={(_, d) => {
            if (d.optionValue) {
              onAdd(group, 'lambda', d.optionValue, table, lambdaAlias);
              setOpen(false);
              setQuery('');
            }
          }}
          style={{ width: '100%' }}
          listbox={{ style: { maxHeight: 360 } }}
        >
          {matches.map((n) => (
            <Option key={n.name} value={n.name} text={n.name}>
              <span style={{ display: 'flex', flexDirection: 'column', minWidth: 0 }}>
                <span
                  style={{
                    fontFamily: tokens.fontFamilyMonospace,
                    fontWeight: 600,
                    color: tokens.colorNeutralForeground1,
                  }}
                >
                  {n.name}
                </span>
                <span style={{ fontSize: 11, color: tokens.colorNeutralForeground3 }}>
                  → {n.targetEntity}
                  <span style={{ marginLeft: 8, fontFamily: tokens.fontFamilyMonospace }}>
                    {n.cardinality}
                  </span>
                </span>
              </span>
            </Option>
          ))}
          {matches.length === 0 && (
            <Option value="__none" text="" disabled>
              <span style={{ color: tokens.colorNeutralForeground3, fontSize: 11 }}>
                No matches for "{query}"
              </span>
            </Option>
          )}
        </Combobox>
      </PopoverSurface>
    </Popover>
  );
}

// ============================================================
// Node renderer — dispatches to row / fn / lam / group
// ============================================================
function NodeRenderer({
  node,
  parentTable,
  depth,
  requestGroup,
  lambdaAlias,
  onUpdate,
  onAdd,
  onRemove,
  onReorderChildren,
  gripProps,
}: {
  node: FilterNode;
  parentTable: TableMeta;
  depth: number;
  requestGroup: RequestGroup;
  /** Lambda alias inherited from the enclosing scope (if any). Nested
   *  groups / rules / functions stay in the same scope; a nested lambda
   *  starts a NEW scope and assigns its own alias. */
  lambdaAlias?: string;
  onUpdate: (id: string, patch: Parameters<typeof patchNode>[2]) => void;
  onAdd: (
    parentId: string,
    kind: 'rule' | 'group' | 'function' | 'lambda',
    extra?: string,
    levelTable?: TableMeta,
    levelAlias?: string,
  ) => void;
  onRemove: (id: string) => void;
  onReorderChildren: (parentId: string, from: number, to: number) => void;
  /** Drag-handle listeners coming from the enclosing SortableItem. Applied
   *  to each child's existing static drag-grip icon. */
  gripProps?: GripProps;
}) {
  if (node.type === 'group') {
    return (
      <FilterGroupCard
        group={node}
        depth={depth + 1}
        table={parentTable}
        requestGroup={requestGroup}
        lambdaAlias={lambdaAlias}
        onUpdate={onUpdate}
        onAdd={onAdd}
        onRemove={onRemove}
        onReorderChildren={onReorderChildren}
        gripProps={gripProps}
      />
    );
  }
  if (node.type === 'rule') {
    return (
      <RuleRow
        rule={node}
        table={parentTable}
        lambdaAlias={lambdaAlias}
        onUpdate={onUpdate}
        onRemove={onRemove}
        gripProps={gripProps}
      />
    );
  }
  if (node.type === 'function') {
    return (
      <FunctionBlock
        fn={node}
        table={parentTable}
        onUpdate={onUpdate}
        onRemove={onRemove}
        gripProps={gripProps}
      />
    );
  }
  return (
    <LambdaBlock
      node={node}
      parentTable={parentTable}
      depth={depth}
      requestGroup={requestGroup}
      onUpdate={onUpdate}
      onAdd={onAdd}
      onRemove={onRemove}
      onReorderChildren={onReorderChildren}
      gripProps={gripProps}
    />
  );
}

// ============================================================
// Plain rule row — qb-rule
// ============================================================
function RuleRow({
  rule,
  table,
  lambdaAlias,
  onUpdate,
  onRemove,
  gripProps,
}: {
  rule: FilterRule;
  table: TableMeta;
  /** Lambda alias inherited from enclosing scope (e.g. "c" inside
   *  `contact_customer_accounts/any(c:…)`). When set, the column picker
   *  stays in flat-list mode (drill-in via N:1 navs lives only at root
   *  scope today). When unset, we use NavPathColumnPicker so the user
   *  can walk `primarycontactid/fullname` style paths. */
  lambdaAlias?: string;
  /** Drag-handle listeners from the enclosing SortableItem. */
  gripProps?: GripProps;
  onUpdate: (id: string, patch: Parameters<typeof patchNode>[2]) => void;
  onRemove: (id: string) => void;
}) {
  const s = useStudioStyles();
  // Strip the lambda alias to get the "path within the lambda's target
  // entity". Inside the lambda, `table` IS the target entity, so paths
  // walk from there. Outside, `table` is the root and `pathOnLevel` ==
  // `rule.col`.
  const pathOnLevel = useMemo(() => {
    if (!lambdaAlias) return rule.col;
    return rule.col.startsWith(`${lambdaAlias}/`)
      ? rule.col.slice(lambdaAlias.length + 1)
      : rule.col;
  }, [rule.col, lambdaAlias]);

  // Resolve the leaf column + the entity it actually lives on. Single
  // helper works for both scopes now: walk `pathOnLevel` against `table`
  // (which is the lambda's target inside a lambda, or the root entity
  // outside one). The resolver naturally handles `col` (no nav-path),
  // `nav/col`, and `nav/nav/col` shapes.
  const lastSeg = pathOnLevel.split('/').pop() ?? pathOnLevel;
  // Eagerly load related entities along the nav-path so leaves on
  // related entities can resolve outside the picker (e.g. on initial
  // load of a saved request). Each useLiveTable resolves one level;
  // the re-render walks further.
  const pendingNav = findPendingNavTarget(table, pathOnLevel);
  useLiveTable(pendingNav);
  const { col, ownerTable } = (() => {
    const r = resolvePathLeaf(table, pathOnLevel);
    return r.col ? r : { col: findColumn(table, lastSeg), ownerTable: table };
  })();
  const op = findOperator(rule.op);
  const noVal = op?.arity === 0;
  // Always-hidden types (File / Image — not filterable on content)
  const visibleColumns = useMemo(
    () =>
      table.columns.filter(
        (c) =>
          c.attributeType !== 'File' &&
          c.attributeType !== 'Image' &&
          c.isValidForRead !== false &&
          !isCompanionLogicalReadOnly(c),
      ),
    [table],
  );
  // Symmetric inverse filter: when the user has picked a typed operator (e.g.
  // `contains`, which is String/Memo only), the column dropdown is narrowed to
  // columns whose AttributeType is in op.allowedTypes. `eq`/`ne` and null-check
  // allow every type so the picker shows the full list.
  const columnsForCurrentOp = useMemo(() => {
    const allowed = op?.allowedTypes;
    if (!allowed) return visibleColumns;
    return visibleColumns.filter((c) => allowed.includes(c.attributeType));
  }, [visibleColumns, op]);
  // Operators valid for plain rule rows: comparison, string, null-check
  const validOps = useMemo(() => {
    if (!col) return OPERATORS.filter((o) => ['comparison', 'string'].includes(o.category));
    return operatorsFor(col.attributeType, table.logicalName).filter(
      (o) => o.kind === 'comparison' || o.kind === 'odata-fn' || o.kind === 'null-check',
    );
  }, [col, table]);
  // NOT is only meaningful when the operator is one of the three OData string
  // functions — per Filter-Builder-Scenarios.html §2A, Dataverse does not
  // support negation on plain comparison rules or groups.
  const canNegate = op?.kind === 'odata-fn';

  const onOpChange = (opId: string) => {
    const next = OPERATORS.find((o) => o.id === opId);
    if (!next) return;
    // If the new op doesn't support column RHS, force literal mode
    const valKind: 'literal' | 'column' = next.supportsColumnRhs
      ? (rule.valKind ?? 'literal')
      : 'literal';
    const patch: Parameters<typeof patchNode>[2] = { op: opId, val: '', valKind };
    // Inverse-direction enforcement: if the current column isn't compatible
    // with the new op's allowedTypes, snap to the first compatible column.
    if (next.allowedTypes && col && !next.allowedTypes.includes(col.attributeType)) {
      const compatible = visibleColumns.find((c) => next.allowedTypes!.includes(c.attributeType));
      if (compatible) {
        const prefix = rule.col.includes('/')
          ? rule.col.split('/').slice(0, -1).join('/') + '/'
          : '';
        patch.col = prefix + compatible.logicalName;
      }
    }
    onUpdate(rule.id, patch);
  };

  // Column-vs-column compare is documented by Dataverse ONLY for the
  // "same row, same table" case:
  //   • Direct on root: `firstname eq lastname` (both cols on root)
  //   • Inside a lambda: `c/firstname eq c/lastname` (both cols on the
  //     lambda's target row — the encoder adds the alias prefix on the
  //     RHS automatically).
  // Nav-path column compare like `primarycontactid/fullname eq
  // primarycontactid/lastname` is NOT documented and Dataverse rejects
  // it in practice. The "cross table column comparisons" Limitations
  // section of /webapi/query/filter-rows is explicit: "OData supports
  // filtering on column values in the same row, but they must be in
  // the same table." Nav-paths address columns on a RELATED row, so
  // they fail the same-row test. Gate the toggle off for these rules.
  // `pathOnLevel` strips any lambda alias prefix, so `/` in the remainder
  // is unambiguously a nav-path. Inside a lambda, `c/jobtitle` strips to
  // `jobtitle` (no slash, not a nav-path); `c/primarycontactid/name`
  // strips to `primarycontactid/name` (slash, IS a nav-path).
  const isNavPath = pathOnLevel.includes('/');
  const canCompareColumn = op?.supportsColumnRhs && !isNavPath;
  const valKind = rule.valKind ?? 'literal';

  // Unified column picker for both scopes:
  //   • Root scope: paths walk from the root entity (= `table`).
  //   • Lambda scope: paths walk from the lambda's target (= `table`,
  //     since the inner FilterGroupCard receives the target as `table`).
  //     The lambda alias prefix is re-attached when storing rule.col.
  // Both modes use NavPathColumnPicker — the picker is alias-agnostic,
  // RuleRow does the prefixing. This lets users build nav-paths like
  // `c/primarycontactid/name` inside a lambda — OData supports this
  // and Dataverse evaluates it as "this iteration variable's lookup's
  // column", same scope as the rest of the lambda body.
  const onPathPicked = (pickedPath: string) => {
    // resolvePathLeaf walks within the level's entity (lambda target or
    // root). Operator-compat snap uses the leaf column's type so the
    // typed value-input renders correctly.
    const leaf = resolvePathLeafColumn(table, pickedPath);
    let nextOp = rule.op;
    if (leaf && op && op.allowedTypes && !op.allowedTypes.includes(leaf.attributeType)) {
      nextOp = validOps[0]?.id ?? 'eq';
    }
    // Re-prepend the lambda alias before storing so the encoder emits
    // `c/.../leaf` (vs `leaf` at root). Empty `pickedPath` is the cleared
    // state (user dismissed) — store empty too.
    const stored = pickedPath ? (lambdaAlias ? `${lambdaAlias}/${pickedPath}` : pickedPath) : '';
    onUpdate(rule.id, { col: stored, op: nextOp, val: '', valKind: 'literal' });
  };

  // Inline antipattern detection: when the resolved leaf column has
  // performance hazards (calculated/rollup/formula, logical, or
  // large-text + scan op), surface an amber warning icon next to the
  // column picker. Clicking it opens a popover with the message + a
  // Learn-more link. The same flags also aggregate into the URL Bar
  // drawer via `detectRetrieveMultipleAntipatterns`.
  const ruleAntipatterns = useMemo(() => {
    // For nav-path rules, the leaf's owner table matters — use the
    // already-resolved (col, ownerTable) from resolvePathLeaf if needed.
    // For now we use the row-level `col` lookup which works for both
    // bare and nav-paths against the level's table.
    if (!col) return [];
    return detectColumnAntipatterns(col, rule.op);
  }, [col, rule.op]);

  return (
    <div className={`${s.filterRow} ${noVal ? s.filterRowNoVal : ''}`}>
      <span {...(gripProps ?? {})} aria-label={gripProps ? 'Drag to reorder rule' : undefined}>
        <ReOrderDotsHorizontal20Regular className={s.dragHandle} aria-hidden={!gripProps} />
      </span>
      {/* The column picker + antipattern icon share the SAME grid cell
          via a nested flex container, so adding the icon doesn't bump
          the filterRow's 5-column grid out of alignment (which would
          shove the operator picker into the value slot, hide the value
          input, and stack the delete button onto a new row — the bug
          we just fixed). */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 4, minWidth: 0 }}>
        <div style={{ flex: 1, minWidth: 0 }}>
          <NavPathColumnPicker
            // `table` is the level's entity — root entity at root scope,
            // lambda target at lambda scope. The picker is alias-agnostic;
            // `onPathPicked` re-prepends the alias when storing.
            rootTable={table.logicalName}
            value={pathOnLevel}
            onChange={onPathPicked}
            allowedColumnTypes={op?.allowedTypes}
            placeholder="Pick column or drill via lookup…"
          />
        </div>
        <AntipatternIcon antipatterns={ruleAntipatterns} />
      </div>
      <OperatorPicker
        table={table.logicalName}
        col={col}
        value={rule.op}
        onChange={onOpChange}
        only={['comparison', 'string']}
      />
      {op ? (
        <span style={{ display: 'flex', gap: 6, alignItems: 'center', minWidth: 0 }}>
          {canNegate && (
            <NotToggle
              negated={!!rule.negated}
              onChange={(n) => onUpdate(rule.id, { negated: n })}
            />
          )}
          {canCompareColumn && (
            <ValueKindToggle
              kind={valKind}
              onChange={(k) => onUpdate(rule.id, { valKind: k, val: '' })}
            />
          )}
          <span style={{ flexGrow: 1, minWidth: 0 }}>
            <FilterValueInput
              rule={{
                id: rule.id,
                col: rule.col,
                op: rule.op,
                val: rule.val,
                valKind: rule.valKind,
              }}
              op={op}
              col={col}
              // ownerTable, not the root, so `useColumnDetail` fetches
              // OptionSet/Targets/MaxLength from the entity the leaf
              // actually lives on (e.g. contact when the path is
              // `primarycontactid/accountrolecode`).
              parentTable={ownerTable ?? table}
              onChange={(p) => onUpdate(rule.id, p)}
            />
          </span>
        </span>
      ) : (
        <span />
      )}
      <Tooltip content="Remove rule" relationship="label">
        <Button
          icon={<Delete20Regular />}
          appearance="subtle"
          size="small"
          onClick={() => onRemove(rule.id)}
          aria-label="Remove rule"
        />
      </Tooltip>
    </div>
  );
}

/**
 * Per-rule literal/column toggle — a v9 Switch.
 * When ON, the RHS becomes a same-type column picker (column-vs-column compare).
 * Only enabled for the six comparison operators (eq/ne/gt/ge/lt/le) that
 * support a bare property-name RHS per the filter-rows docs.
 */
function ValueKindToggle({
  kind,
  onChange,
}: {
  kind: 'literal' | 'column';
  onChange: (k: 'literal' | 'column') => void;
}) {
  const isColumn = kind === 'column';
  return (
    <Tooltip
      content={
        isColumn
          ? 'On — RHS is a same-row column reference (column-vs-column compare). Toggle off for a literal value.'
          : 'Off — RHS is a typed literal. Toggle on to compare against another column on this row.'
      }
      relationship="description"
    >
      <span style={{ display: 'inline-flex', alignItems: 'center', flexShrink: 0 }}>
        <Switch
          checked={isColumn}
          onChange={(_, d) => onChange(d.checked ? 'column' : 'literal')}
          label={
            <span
              style={{
                fontSize: 11,
                fontWeight: 600,
                color: isColumn ? tokens.colorBrandForeground1 : tokens.colorNeutralForeground3,
                fontFamily: tokens.fontFamilyMonospace,
                letterSpacing: '0.04em',
              }}
            >
              {isColumn ? 'col' : 'value'}
            </span>
          }
          labelPosition="after"
        />
      </span>
    </Tooltip>
  );
}

/**
 * Per-condition NOT toggle. Per Filter-Builder-Scenarios.html §2A:
 *   - Emits `not <expression>` wrapping the rule's OData function or
 *     Dataverse query function.
 *   - Only valid on contains / startswith / endswith and on
 *     Microsoft.Dynamics.CRM.* query functions — NEVER on plain comparisons
 *     and NEVER on whole groups (Dataverse rejects `not (group)`).
 */
function NotToggle({ negated, onChange }: { negated: boolean; onChange: (n: boolean) => void }) {
  return (
    <Tooltip
      content={
        negated
          ? 'Negation is on — the encoder emits `not <expression>`'
          : 'Negate this condition (wraps it in `not(...)`)'
      }
      relationship="description"
    >
      <ToggleButton
        size="small"
        shape="rounded"
        checked={negated}
        onClick={() => onChange(!negated)}
        aria-label="Negate condition"
        style={
          negated
            ? {
                backgroundColor: tokens.colorPaletteRedBackground2,
                color: tokens.colorPaletteRedForeground1,
              }
            : undefined
        }
      >
        <span style={{ fontFamily: tokens.fontFamilyMonospace, fontWeight: 700 }}>not</span>
      </ToggleButton>
    </Tooltip>
  );
}

// ============================================================
// Dataverse function block — .fn from v2.2
// ============================================================
function FunctionBlock({
  fn,
  table,
  onUpdate,
  onRemove,
  gripProps,
}: {
  fn: FilterFunctionNode;
  table: TableMeta;
  onUpdate: (id: string, patch: Parameters<typeof patchNode>[2]) => void;
  onRemove: (id: string) => void;
  gripProps?: GripProps;
}) {
  const s = useStudioStyles();
  const op = findOperator(fn.op);
  const lastSeg = fn.col.split('/').pop() ?? fn.col;
  const col = findColumn(table, lastSeg);

  // Function picker — when a column is selected, only show functions valid for that column's type.
  const fnOptions = useMemo(() => {
    if (!col) return OPERATORS.filter((o) => o.kind.startsWith('dv-fn'));
    return operatorsFor(col.attributeType, table.logicalName).filter((o) =>
      o.kind.startsWith('dv-fn'),
    );
  }, [col, table]);
  // Inverse direction: PropertyName picker is filtered to columns compatible with the chosen
  // function's allowedTypes — e.g. LastXDays only shows DateTime columns. File/Image always hidden.
  const compatibleColumns = useMemo(() => {
    const visible = table.columns.filter(
      (c) =>
        c.attributeType !== 'File' &&
        c.attributeType !== 'Image' &&
        c.isValidForRead !== false &&
        !isCompanionLogicalReadOnly(c),
    );
    const allowed = op?.allowedTypes;
    if (!allowed) return visible;
    return visible.filter((c) => allowed.includes(c.attributeType));
  }, [table, op]);

  const onColChange = (logicalName: string) =>
    onUpdate(fn.id, { col: logicalName, val: '', vals: undefined, values: undefined });
  const onFnChange = (opId: string) => {
    const next = findOperator(opId);
    if (!next) return;
    // Always clear `negated` when switching functions. Dataverse rejects
    // `not` on any Microsoft.Dynamics.CRM.* function (0x80060888), so we
    // never want a stale `true` carrying forward from a legacy request.
    const patch: Parameters<typeof patchNode>[2] = {
      op: opId,
      val: '',
      vals: undefined,
      values: undefined,
      negated: false,
    };
    // Two-argument shape: pick the right field for the wire format.
    //   • kind 'dv-fn-2' (InFiscalPeriodAndYear, …)         → fn.vals tuple
    //   • kind 'dv-fn-array' arity 2 (Between, NotBetween) → fn.values[0..1]
    // Mismatch causes the UI to render empty inputs even when state is
    // populated (the case we saw after parsing Between from a URL).
    if (next.arity === 2 && next.kind === 'dv-fn-2') patch.vals = ['', ''];
    if (next.arity === 2 && next.kind === 'dv-fn-array') patch.values = ['', ''];
    if (next.arity === 'n') patch.values = [];
    // If the current column isn't compatible with the new function, snap to the first compatible one.
    const allowed = next.allowedTypes;
    if (allowed && col && !allowed.includes(col.attributeType)) {
      const visible = table.columns.filter(
        (c) =>
          c.attributeType !== 'File' &&
          c.attributeType !== 'Image' &&
          c.isValidForRead !== false &&
          !isCompanionLogicalReadOnly(c),
      );
      const compatible = visible.find((c) => allowed.includes(c.attributeType));
      if (compatible) patch.col = compatible.logicalName;
    }
    onUpdate(fn.id, patch);
  };

  return (
    <div className={s.fnBlock}>
      <div className={s.fnHeader}>
        <span
          {...(gripProps ?? {})}
          aria-label={gripProps ? 'Drag to reorder function' : undefined}
        >
          <ReOrderDotsHorizontal20Regular className={s.dragHandle} aria-hidden={!gripProps} />
        </span>
        <span className={s.fnBadge}>ƒ</span>
        {/* NOTE: NotToggle is intentionally NOT rendered on FunctionBlock.
            Dataverse rejects `not Microsoft.Dynamics.CRM.<anything>` with:
              0x80060888 "Not operator along with the Custom Named Condition
              operators is not allowed".
            Each negatable function exposes an explicit `Not*` sibling
            instead — `NotBetween`, `NotIn`, `NotEqualUserId`,
            `DoesNotBeginWith`, `DoesNotContainValues`, etc. The user
            switches operator via the picker rather than toggling negation. */}
        <span className={s.fnPrefix}>Microsoft.Dynamics.CRM.</span>
        <Combobox
          size="small"
          value={op?.label ?? fn.op}
          selectedOptions={[fn.op]}
          onOptionSelect={(_, d) => d.optionValue && onFnChange(d.optionValue)}
          style={{ minWidth: 240 }}
        >
          {fnOptions.map((o) => (
            <Option key={o.id} value={o.id} text={o.label}>
              <span style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <span>{o.label}</span>
                <span
                  style={{
                    fontFamily: tokens.fontFamilyMonospace,
                    fontSize: 10,
                    color: tokens.colorNeutralForeground3,
                  }}
                >
                  {o.odata.replace('Microsoft.Dynamics.CRM.', '')}
                </span>
              </span>
            </Option>
          ))}
        </Combobox>
        {op?.hint && <span className={s.fnDescription}>{op.hint}</span>}
        <div style={{ flexGrow: 1 }} />
        <Tooltip content="Remove function" relationship="label">
          <Button
            icon={<Delete20Regular />}
            appearance="subtle"
            size="small"
            onClick={() => onRemove(fn.id)}
            aria-label="Remove function"
          />
        </Tooltip>
      </div>
      <div className={s.fnParams}>
        <div className={s.fnParamRow}>
          <span className={s.fnParamLabel}>PropertyName</span>
          <FilterableColumnCombo
            value={fn.col}
            currentDisplay={col?.displayName}
            columns={compatibleColumns}
            onChange={onColChange}
            placeholder={
              compatibleColumns.length === 0
                ? `No columns of type ${op?.allowedTypes?.join('/')} on this table`
                : 'Search columns…'
            }
            style={{ flexGrow: 1, maxWidth: 360 }}
          />
          {/* Inline antipattern indicator — same component as the rule
              row uses, surfacing calculated/logical/large-text hazards
              on the function's PropertyName column. */}
          {col && <AntipatternIcon antipatterns={detectColumnAntipatterns(col, fn.op)} />}
        </div>
        {op && renderFnValueRows(op, fn, col, table, (p) => onUpdate(fn.id, p), s)}
      </div>
    </div>
  );
}

// Per-function-id labels for the value rows. Falls back to the generic
// `PropertyValue` / `PropertyValue1` / `PropertyValue2` / `PropertyValues`
// when there's no specific label — but for the common rolling/fiscal
// functions we surface a much more readable "Days" / "Period" / "Year"
// so the user knows what kind of value to enter.
function valueLabelsFor(opId: string): { v?: string; v1?: string; v2?: string; vN?: string } {
  if (/^(LastX|NextX|OlderThanX)(Days)$/.test(opId)) return { v: 'Days' };
  if (/^(LastX|NextX|OlderThanX)(Hours)$/.test(opId)) return { v: 'Hours' };
  if (/^OlderThanXMinutes$/.test(opId)) return { v: 'Minutes' };
  if (/^(LastX|NextX|OlderThanX)(Weeks)$/.test(opId)) return { v: 'Weeks' };
  if (/^(LastX|NextX|OlderThanX)(Months)$/.test(opId)) return { v: 'Months' };
  if (/^(LastX|NextX|OlderThanX)(Years)$/.test(opId)) return { v: 'Years' };
  if (/^(LastX|NextX)FiscalPeriods$/.test(opId)) return { v: 'Fiscal periods' };
  if (/^(LastX|NextX)FiscalYears$/.test(opId)) return { v: 'Fiscal years' };
  if (opId === 'InFiscalPeriod') return { v: 'Fiscal period' };
  if (opId === 'InFiscalYear') return { v: 'Fiscal year' };
  if (
    /^(InFiscalPeriodAndYear|InOrAfterFiscalPeriodAndYear|InOrBeforeFiscalPeriodAndYear)$/.test(
      opId,
    )
  ) {
    return { v1: 'Fiscal period', v2: 'Fiscal year' };
  }
  if (opId === 'Between' || opId === 'NotBetween') return { vN: 'Range (low, high)' };
  if (opId === 'In' || opId === 'NotIn') return { vN: 'Values' };
  if (opId === 'ContainValues' || opId === 'DoesNotContainValues') return { vN: 'Options' };
  if (opId === 'On' || opId === 'OnOrAfter' || opId === 'OnOrBefore') return { v: 'Date' };
  if (/^(Above|AboveOrEqual|Under|UnderOrEqual|NotUnder)$/.test(opId))
    return { v: 'Reference record' };
  return {};
}

function renderFnValueRows(
  op: ReturnType<typeof findOperator>,
  fn: FilterFunctionNode,
  col: ColumnMeta | undefined,
  table: TableMeta,
  onChange: (patch: Parameters<typeof patchNode>[2]) => void,
  s: ReturnType<typeof useStudioStyles>,
): React.ReactNode {
  if (!op) return null;
  if (op.arity === 0) return null;
  const labels = valueLabelsFor(op.id);
  if (op.arity === 2) {
    // Two-argument shapes are split by kind:
    //   • dv-fn-2 (fiscal period+year family) — stores in fn.vals tuple,
    //     emits `PropertyValue1=…,PropertyValue2=…`
    //   • dv-fn-array arity 2 (Between / NotBetween) — stores in fn.values
    //     array, emits `PropertyValues=[low, high]`
    // We render TWO inputs in both cases (better UX than a generic array
    // editor for a 2-value range), but read/write the correct field.
    const usesArray = op.kind === 'dv-fn-array';
    const [a, b] = usesArray ? [fn.values?.[0] ?? '', fn.values?.[1] ?? ''] : (fn.vals ?? ['', '']);
    const writePair = (nextA: string, nextB: string) =>
      usesArray ? { values: [nextA, nextB] } : { vals: [nextA, nextB] as [string, string] };
    // Better default labels for Between — the existing labels object
    // already returns `vN: 'Range (low, high)'`. Repurpose into v1/v2.
    const v1Label =
      labels.v1 ?? (op.id === 'Between' || op.id === 'NotBetween' ? 'Low' : 'PropertyValue1');
    const v2Label =
      labels.v2 ?? (op.id === 'Between' || op.id === 'NotBetween' ? 'High' : 'PropertyValue2');
    return (
      <>
        <div className={s.fnParamRow}>
          <span className={s.fnParamLabel}>{v1Label}</span>
          <FilterValueInput
            rule={{ id: fn.id, col: fn.col, op: op.id, val: a }}
            op={{ ...op, arity: 1 }}
            col={col}
            parentTable={table}
            onChange={(p) => onChange(writePair((p.val ?? a) as string, b))}
          />
        </div>
        <div className={s.fnParamRow}>
          <span className={s.fnParamLabel}>{v2Label}</span>
          <FilterValueInput
            rule={{ id: fn.id, col: fn.col, op: op.id, val: b }}
            op={{ ...op, arity: 1 }}
            col={col}
            parentTable={table}
            onChange={(p) => onChange(writePair(a, (p.val ?? b) as string))}
          />
        </div>
      </>
    );
  }
  if (op.arity === 'n') {
    return (
      <div className={s.fnParamRow}>
        <span className={s.fnParamLabel}>{labels.vN ?? 'PropertyValues'}</span>
        <FilterValueInput
          rule={{ id: fn.id, col: fn.col, op: op.id, val: '', values: fn.values }}
          op={op}
          col={col}
          parentTable={table}
          onChange={(p) => onChange({ values: p.values })}
        />
      </div>
    );
  }
  // arity === 1
  return (
    <div className={s.fnParamRow}>
      <span className={s.fnParamLabel}>{labels.v ?? 'PropertyValue'}</span>
      <FilterValueInput
        rule={{ id: fn.id, col: fn.col, op: op.id, val: fn.val }}
        op={op}
        col={col}
        parentTable={table}
        onChange={(p) => onChange({ val: p.val })}
      />
    </div>
  );
}

// ============================================================
// Lambda block — .lam from v2.2
// ============================================================
function LambdaBlock({
  node,
  parentTable,
  depth,
  requestGroup,
  onUpdate,
  onAdd,
  onRemove,
  onReorderChildren,
  gripProps,
}: {
  node: FilterLambdaNode;
  parentTable: TableMeta;
  depth: number;
  requestGroup: RequestGroup;
  onUpdate: (id: string, patch: Parameters<typeof patchNode>[2]) => void;
  onAdd: (parentId: string, kind: 'rule' | 'group' | 'function' | 'lambda', extra?: string) => void;
  onRemove: (id: string) => void;
  onReorderChildren: (parentId: string, from: number, to: number) => void;
  gripProps?: GripProps;
}) {
  const s = useStudioStyles();
  const collections = parentTable.navigationProperties.filter((n) => n.cardinality !== 'ManyToOne');
  const nav = parentTable.navigationProperties.find((n) => n.name === node.nav);
  // Warm the selected nav's target entity. Re-renders this block (via
  // the live-table registry subscription) when the fetch resolves so
  // `findTable(nav.targetEntity)` flips from undefined → the loaded table.
  useLiveTable(nav?.targetEntity ?? null);
  const targetTbl = nav ? findTable(nav.targetEntity) : undefined;

  const onNavChange = (navName: string) => {
    const newNav = parentTable.navigationProperties.find((n) => n.name === navName);
    if (!newNav) return;
    const newTarget = findTable(newNav.targetEntity);
    const newAlias = navName[0] ?? node.alias;
    const newCol = newTarget?.columns[0].logicalName ?? 'name';
    onUpdate(node.id, {
      nav: navName,
      alias: newAlias,
      inner: {
        id: node.inner.id,
        type: 'group',
        combinator: node.inner.combinator,
        rules: [
          {
            id: newId('r'),
            type: 'rule',
            col: `${newAlias}/${newCol}`,
            op: 'eq',
            val: '',
          },
        ],
      },
    });
  };

  // Native Fluent v9 freeform + clearable Combobox — `value` IS the
  // displayed/typed string, synced to the picked nav on selection change.
  const [navQuery, setNavQuery] = useState<string>(node.nav);
  useEffect(() => {
    setNavQuery(node.nav);
  }, [node.nav]);
  const filteredCollections = useMemo(() => {
    const q = navQuery.trim().toLowerCase();
    if (!q || q === node.nav.toLowerCase()) return collections;
    return collections.filter(
      (n) => n.name.toLowerCase().includes(q) || n.targetEntity.toLowerCase().includes(q),
    );
  }, [collections, navQuery, node.nav]);

  return (
    <div className={s.lamBlock}>
      <div className={s.lamHeader}>
        <span {...(gripProps ?? {})} aria-label={gripProps ? 'Drag to reorder lambda' : undefined}>
          <ReOrderDotsHorizontal20Regular className={s.dragHandle} aria-hidden={!gripProps} />
        </span>
        <span className={s.lamBadge}>λ</span>
        {/* Lambda-level NOT — empirically supported (test G.8):
              `not contact_customer_accounts/any(c: c/firstname eq 'John')`
            returns 200. Useful for "no related row matches X". */}
        <Tooltip
          content={
            node.negated
              ? 'Lambda is negated — emits `not <nav>/<any|all>(…)`. Click to remove.'
              : 'Negate this lambda — emits `not <nav>/<any|all>(…)`. Useful for `no matching related row`.'
          }
          relationship="description"
        >
          <ToggleButton
            checked={!!node.negated}
            onClick={() => onUpdate(node.id, { negated: !node.negated })}
            size="small"
            appearance={node.negated ? 'primary' : 'subtle'}
            style={{
              minWidth: 44,
              fontFamily: tokens.fontFamilyMonospace,
              fontWeight: 700,
              ...(node.negated
                ? { backgroundColor: tokens.colorPaletteRedBackground3, color: '#fff' }
                : {}),
            }}
            aria-label={node.negated ? 'Disable lambda NOT' : 'Negate lambda'}
          >
            NOT
          </ToggleButton>
        </Tooltip>
        <Combobox
          freeform
          clearable
          size="small"
          value={navQuery}
          selectedOptions={[node.nav]}
          onChange={(e) => setNavQuery((e.target as HTMLInputElement).value)}
          onOptionSelect={(_, d) => {
            if (d.optionValue) {
              onNavChange(d.optionValue);
              setNavQuery(d.optionValue);
            } else {
              setNavQuery('');
            }
          }}
          placeholder="Search relationships…"
          style={{ minWidth: 220 }}
          listbox={{ style: { maxHeight: 320 } }}
        >
          {filteredCollections.length === 0 && (
            <Option value="__none" text="" disabled>
              <span style={{ color: tokens.colorNeutralForeground3, fontSize: 11 }}>
                No matches for "{navQuery}"
              </span>
            </Option>
          )}
          {filteredCollections.map((n) => (
            <Option key={n.name} value={n.name} text={n.name}>
              <span style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                {n.name}
                <span
                  style={{
                    fontFamily: tokens.fontFamilyMonospace,
                    fontSize: 10,
                    color: tokens.colorNeutralForeground3,
                  }}
                >
                  → {n.targetEntity}
                </span>
              </span>
            </Option>
          ))}
        </Combobox>
        <span
          style={{ color: tokens.colorNeutralForeground3, fontFamily: tokens.fontFamilyMonospace }}
        >
          /
        </span>
        <SegmentedToggle ariaLabel="Lambda kind">
          <ToggleButton
            checked={node.lambda === 'any'}
            onClick={() => onUpdate(node.id, { lambda: 'any' })}
            style={
              node.lambda === 'any'
                ? { backgroundColor: tokens.colorBrandForeground1, color: '#fff' }
                : undefined
            }
          >
            any
          </ToggleButton>
          <ToggleButton
            checked={node.lambda === 'all'}
            onClick={() => onUpdate(node.id, { lambda: 'all' })}
            style={
              node.lambda === 'all'
                ? { backgroundColor: tokens.colorBrandForeground1, color: '#fff' }
                : undefined
            }
          >
            all
          </ToggleButton>
        </SegmentedToggle>
        <span style={{ fontSize: 11, color: tokens.colorNeutralForeground3 }}>
          alias: <code style={{ fontFamily: tokens.fontFamilyMonospace }}>{node.alias}</code>
        </span>
        <div style={{ flexGrow: 1 }} />
        {targetTbl && (
          <span
            style={{
              fontSize: 10,
              color: tokens.colorNeutralForeground3,
              background: tokens.colorNeutralBackground3,
              padding: '2px 6px',
              borderRadius: 3,
            }}
          >
            → {targetTbl.displayName}
          </span>
        )}
        <Tooltip content="Remove lambda" relationship="label">
          <Button
            icon={<Delete20Regular />}
            appearance="subtle"
            size="small"
            onClick={() => onRemove(node.id)}
            aria-label="Remove lambda"
          />
        </Tooltip>
      </div>

      {targetTbl && (
        <div className={s.lamInner}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 8 }}>
            <span className={s.lamAliasPrefix}>{node.alias}/</span>
            <span style={{ fontSize: 11, color: tokens.colorNeutralForeground3 }}>
              every column reference is prefixed with the alias
            </span>
          </div>
          <FilterGroupCard
            group={node.inner}
            depth={depth + 1}
            table={targetTbl}
            requestGroup={requestGroup}
            // Inject the lambda's alias so any rule / nested group the
            // user adds inside this λ has its column auto-prefixed with
            // `<alias>/`. Without this, a second rule like
            // `_primarycontactid_value ne null` would arrive with no
            // prefix and Dataverse rejects it ("Conditions on property
            // other than current navigation property in any/all is not
            // supported").
            lambdaAlias={node.alias}
            onUpdate={onUpdate}
            onAdd={onAdd}
            onRemove={onRemove}
            onReorderChildren={onReorderChildren}
          />
        </div>
      )}
    </div>
  );
}

// ============================================================
// Reusable filterable column Combobox
// ============================================================
/**
 * Freeform + type-to-filter Combobox over a column list. Used by the
 * Dataverse function block's PropertyName picker and anywhere else we
 * need a long column dropdown to remain usable. Mirrors the Fluent v9
 * freeform pattern from
 * https://storybooks.fluentui.dev/react/llms/components-combobox.txt
 */
function FilterableColumnCombo({
  value,
  currentDisplay,
  columns,
  onChange,
  placeholder,
  style,
}: {
  value: string;
  currentDisplay?: string;
  columns: ColumnMeta[];
  onChange: (logicalName: string) => void;
  placeholder?: string;
  style?: React.CSSProperties;
}) {
  const [query, setQuery] = useState<string>(currentDisplay ?? value);
  useEffect(() => {
    setQuery(currentDisplay ?? value);
  }, [currentDisplay, value]);
  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q || q === (currentDisplay ?? '').toLowerCase()) return columns;
    return columns.filter(
      (c) => c.displayName.toLowerCase().includes(q) || c.logicalName.toLowerCase().includes(q),
    );
  }, [columns, query, currentDisplay]);

  return (
    <Combobox
      freeform
      clearable
      size="small"
      value={query}
      selectedOptions={value ? [value] : []}
      onChange={(e) => setQuery((e.target as HTMLInputElement).value)}
      onOptionSelect={(_, d) => {
        if (d.optionValue) {
          const picked = columns.find((c) => c.logicalName === d.optionValue);
          onChange(d.optionValue);
          setQuery(picked?.displayName ?? d.optionValue);
        } else {
          setQuery('');
        }
      }}
      placeholder={placeholder}
      style={style}
      listbox={{ style: { maxHeight: 320 } }}
    >
      {filtered.length === 0 && (
        <Option value="__none" text="" disabled>
          <span style={{ color: tokens.colorNeutralForeground3, fontSize: 11 }}>
            No matches for "{query}"
          </span>
        </Option>
      )}
      {filtered.map((c) => (
        <Option key={c.logicalName} value={c.logicalName} text={c.displayName}>
          <span style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
            {c.displayName}
            <span
              style={{
                fontFamily: tokens.fontFamilyMonospace,
                fontSize: 10,
                color: tokens.colorNeutralForeground3,
              }}
            >
              {c.logicalName}
            </span>
            <span
              style={{
                fontSize: 9,
                color: tokens.colorNeutralForeground3,
                backgroundColor: tokens.colorNeutralBackground3,
                padding: '0 4px',
                borderRadius: 3,
              }}
            >
              {c.attributeType}
            </span>
          </span>
        </Option>
      ))}
    </Combobox>
  );
}
