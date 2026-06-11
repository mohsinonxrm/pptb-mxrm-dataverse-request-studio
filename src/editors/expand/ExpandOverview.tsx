// $expand overview pane — the "right-pane editor" rendered when the active
// sidebar node is `expand` (root) or `expand/<id1>/.../expand` (any nested level).
//
// Shows:
//   - A breadcrumb of the current expand path so the user always knows which
//     entity's navigation properties they're picking from
//   - The list of navigation properties available at this level (metadata-driven)
//     — collection-valued vs single-valued visually distinguished
//   - The list of currently-expanded navs at this level (clickable to drill in)
//   - Depth / nav-count soft warnings (>3 / >10 per spec §10)

import { useMemo, useState } from 'react';
import {
  Caption1,
  MessageBar,
  MessageBarBody,
  Tooltip,
  Button,
  Badge,
  Subtitle2,
  mergeClasses,
  tokens,
  Combobox,
  Option,
  Field,
} from '@fluentui/react-components';
import {
  BranchFork20Filled,
  BranchFork20Regular,
  Add20Regular,
  Delete20Regular,
  ArrowEnter20Regular,
} from '@fluentui/react-icons';
import { useStudioStyles } from '../../primitives/styles';
import { PaneHead } from '../PaneHead';
import { findTable, type NavProperty } from '../../mock/metadata';
import {
  availableNavsAt,
  isCollectionExpand,
  totalExpandCount,
  maxExpandDepth,
  MAX_EXPANDS_PER_QUERY,
  hasAnyNestedExpand,
} from './expandTree';
import type { ExpandSpec } from '../ExpandEditor';
import type { RequestGroup } from '../../registry/requestTypes';

export interface ExpandOverviewProps {
  /** Parent-entity logical name at this level. */
  parentEntity: string;
  /** Expands that exist at this level (siblings the user is editing). */
  siblings: ExpandSpec[];
  /** The whole top-level expand list — used to count totals + depth across the tree. */
  rootExpand: ExpandSpec[];
  /** Path-segment trail for breadcrumb display (label per level). */
  breadcrumb: { label: string; path: string }[];
  /**
   * Cardinality of the IMMEDIATE parent $expand. `null` at root.
   * Drives the nav-availability filter and the parent-aware info bar.
   *   • null         → all cardinalities offered (root level)
   *   • 'ManyToOne'  → all cardinalities offered (single-valued parents
   *                    can host any nested expand per the docs)
   *   • 'OneToMany'  → only N:1 + 1:N offered (no N:N nested under a 1:N
   *                    in a query that already uses nested $expand)
   *   • 'ManyToMany' → nothing offered (Dataverse rejects all nested
   *                    expand inside an N:N)
   */
  parentCardinality: 'OneToMany' | 'ManyToOne' | 'ManyToMany' | null;
  onAdd: (navName: string) => void;
  onPickExisting: (expandId: string) => void;
  onRemove: (expandId: string) => void;
  group?: RequestGroup;
}

export function ExpandOverview({
  parentEntity,
  siblings,
  rootExpand,
  breadcrumb,
  parentCardinality,
  onAdd,
  onPickExisting,
  onRemove,
  group = 'read',
}: ExpandOverviewProps) {
  const s = useStudioStyles();
  const parentTbl = findTable(parentEntity);
  // isNested + parentIsNN are derived for the local conditionals below
  // (and the MessageBars). The actual nav filter uses parentCardinality
  // directly via availableNavsAt's new signature.
  const isNested = parentCardinality !== null;
  const parentIsNN = parentCardinality === 'ManyToMany';
  const parentIsCollection = parentCardinality === 'OneToMany';
  const available = availableNavsAt(parentEntity, siblings, { parentCardinality });
  const totalNavs = totalExpandCount(rootExpand);
  // Per /webapi/query/join-tables: when the query has NO nested $expand
  // anywhere, expanded collections come back un-paged — up to 5,000 rows
  // per collection. Worth flagging at the root overview.
  const queryHasNestedExpand = hasAnyNestedExpand(rootExpand);
  const hasFlatCollection =
    !queryHasNestedExpand && siblings.some((it) => isCollectionExpand(parentEntity, it));
  const depth = maxExpandDepth(rootExpand);
  const capExceeded = totalNavs >= MAX_EXPANDS_PER_QUERY;

  // Note: target-entity DisplayNames are NOT pre-fetched here. Showing
  // 50+ target entities would mean 50+ entity metadata fetches per
  // overview render, which trips Dataverse's 100-concurrent cap.
  // The labels fall back to the target logical name; once the user picks
  // a relationship, that single target loads via the per-expand fetch.

  return (
    <div>
      <PaneHead
        icon={BranchFork20Filled}
        title="$expand"
        sub={`${parentTbl?.displayName ?? parentEntity} · pick navigation properties to inline related rows.`}
        group={group}
      >
        <Badge appearance="ghost">{siblings.length} at this level</Badge>
      </PaneHead>

      {breadcrumb.length > 1 && (
        <div
          style={{
            marginBottom: 14,
            display: 'flex',
            alignItems: 'center',
            gap: 6,
            flexWrap: 'wrap',
          }}
        >
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Path:</Caption1>
          {breadcrumb.map((b, i) => (
            <span key={b.path} style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
              {i > 0 && <span style={{ color: tokens.colorNeutralForeground4 }}>→</span>}
              <Caption1
                style={{
                  fontFamily: tokens.fontFamilyMonospace,
                  color:
                    i === breadcrumb.length - 1
                      ? tokens.colorBrandForeground1
                      : tokens.colorNeutralForeground2,
                  fontWeight: i === breadcrumb.length - 1 ? 600 : 400,
                }}
              >
                {b.label}
              </Caption1>
            </span>
          ))}
        </div>
      )}

      {/* Hard caps + per-spec restrictions from the join-tables docs */}
      {capExceeded && (
        <MessageBar layout="multiline" intent="error" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            <strong>
              {totalNavs} of {MAX_EXPANDS_PER_QUERY} expand slots used.
            </strong>{' '}
            Dataverse rejects any query with more than {MAX_EXPANDS_PER_QUERY} <code>$expand</code>{' '}
            options total — remove some before adding more.
          </MessageBarBody>
        </MessageBar>
      )}
      {!capExceeded && totalNavs >= MAX_EXPANDS_PER_QUERY - 3 && (
        <MessageBar layout="multiline" intent="warning" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            <strong>
              {totalNavs} of {MAX_EXPANDS_PER_QUERY} expand slots used.
            </strong>{' '}
            Each expand is a server-side join — keep this list lean.
          </MessageBarBody>
        </MessageBar>
      )}
      {depth > 3 && (
        <MessageBar layout="multiline" intent="warning" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            Expand depth is <strong>{depth}</strong>. Every nested expand is another join — keep
            depth shallow for performance.
          </MessageBarBody>
        </MessageBar>
      )}
      {hasFlatCollection && !isNested && (
        <MessageBar layout="multiline" intent="info" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            Collections without a nested <code>$expand</code> aren't paginated — Dataverse can
            return up to <strong>5,000 related rows per parent record</strong>. Add a nested{' '}
            <code>$expand</code> or a top-level <code>Prefer: odata.maxpagesize</code> to control
            the size.
          </MessageBarBody>
        </MessageBar>
      )}
      {parentIsNN && (
        <MessageBar layout="multiline" intent="error" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            The enclosing expand is a <strong>many-to-many (N:N)</strong> relationship.{' '}
            <strong>Dataverse rejects any nested $expand inside an N:N</strong> with:{' '}
            <em>
              "The navigation property '…' cannot be expanded. Only many-to-one relationships are
              supported for nested expansion."
            </em>
          </MessageBarBody>
        </MessageBar>
      )}
      {parentIsCollection && (
        <MessageBar layout="multiline" intent="info" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            You're inside a <strong>collection (1:N)</strong> expand. You can nest either{' '}
            <strong>many-to-one (N:1)</strong> or <strong>one-to-many (1:N)</strong> relationships
            here. <strong>N:N navs are hidden</strong> — Dataverse rejects N:N expansion anywhere in
            a query that uses nested $expand. Also: <code>$top</code> and <code>$orderby</code>{' '}
            aren't supported inside this expand.
          </MessageBarBody>
        </MessageBar>
      )}
      {isNested && !parentIsNN && !parentIsCollection && (
        <MessageBar layout="multiline" intent="info" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            You're inside a <strong>single-valued (N:1)</strong> expand. All cardinalities can be
            nested here (e.g. the classic <code>tasks → contact → account → systemuser</code> chain
            in the docs).
          </MessageBarBody>
        </MessageBar>
      )}

      {/* Currently expanded at this level */}
      <section style={{ marginBottom: 18 }}>
        <Subtitle2 style={{ marginBottom: 8, display: 'block' }}>
          Currently expanded at this level
        </Subtitle2>
        {siblings.length === 0 && (
          <Caption1 style={{ color: tokens.colorNeutralForeground3, fontStyle: 'italic' }}>
            None — pick a navigation property below to add one.
          </Caption1>
        )}
        {siblings.length > 0 && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
            {siblings.map((it) => {
              const nav = parentTbl?.navigationProperties.find((n) => n.name === it.nav);
              const target = nav ? findTable(nav.targetEntity) : undefined;
              const isCol = isCollectionExpand(parentEntity, it);
              return (
                <div
                  key={it.id}
                  className={mergeClasses(s.inlineCard)}
                  style={{
                    padding: '10px 12px',
                    display: 'grid',
                    gridTemplateColumns: 'auto 1fr auto auto',
                    gap: 10,
                    alignItems: 'center',
                  }}
                >
                  <BranchFork20Regular style={{ color: tokens.colorBrandForeground1 }} />
                  <div style={{ minWidth: 0 }}>
                    <div style={{ fontFamily: tokens.fontFamilyMonospace, fontWeight: 600 }}>
                      {it.nav}
                    </div>
                    <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                      → {target?.displayName ?? nav?.targetEntity ?? '?'}
                      {' · '}
                      {nav?.cardinality}
                      {isCol ? ' (collection)' : ' (single)'}
                      {' · '}
                      {it.select.length ? `${it.select.length} cols` : 'all cols'}
                      {it.filter?.rules.length ? ` · ${it.filter.rules.length} filter` : ''}
                      {it.orderby.length ? ` · ${it.orderby.length} sort` : ''}
                      {it.top ? ` · top ${it.top}` : ''}
                      {it.nestedExpand?.length ? ` · ${it.nestedExpand.length} nested` : ''}
                    </Caption1>
                  </div>
                  <Button
                    size="small"
                    appearance="outline"
                    icon={<ArrowEnter20Regular />}
                    onClick={() => onPickExisting(it.id)}
                  >
                    Edit
                  </Button>
                  <Tooltip content="Remove this expand" relationship="label">
                    <Button
                      size="small"
                      appearance="subtle"
                      icon={<Delete20Regular />}
                      onClick={() => onRemove(it.id)}
                      aria-label="Remove expand"
                    />
                  </Tooltip>
                </div>
              );
            })}
          </div>
        )}
      </section>

      {/* Add navigation property — searchable Combobox + Add button */}
      <section>
        <Subtitle2 style={{ marginBottom: 8, display: 'block' }}>Add navigation property</Subtitle2>
        {available.length === 0 ? (
          <Caption1 style={{ color: tokens.colorNeutralForeground3, fontStyle: 'italic' }}>
            All eligible navigation properties at this level are already expanded.
          </Caption1>
        ) : (
          <NavPicker available={available} onAdd={onAdd} capExceeded={capExceeded} />
        )}
      </section>
    </div>
  );
}

/**
 * Searchable nav-property picker. v9 Combobox is typeable + filterable, which
 * scales to entities with hundreds of relationships — picking from a wall of
 * cards is impractical at that scale.
 */
function NavPicker({
  available,
  onAdd,
  capExceeded,
}: {
  available: NavProperty[];
  onAdd: (nav: string) => void;
  capExceeded: boolean;
}) {
  const [staged, setStaged] = useState<string | null>(null);
  const [filterText, setFilterText] = useState<string>('');

  // Manual filter so we can match on name, target entity, and cardinality.
  const filtered = useMemo(() => {
    if (!filterText) return available;
    const q = filterText.toLowerCase();
    return available.filter(
      (n) =>
        n.name.toLowerCase().includes(q) ||
        n.targetEntity.toLowerCase().includes(q) ||
        n.cardinality.toLowerCase().includes(q) ||
        (findTable(n.targetEntity)?.displayName.toLowerCase() ?? '').includes(q),
    );
  }, [available, filterText]);

  const stagedNav = staged ? available.find((n) => n.name === staged) : null;
  const stagedTarget = stagedNav ? findTable(stagedNav.targetEntity) : null;
  const stagedIsCol =
    stagedNav?.cardinality === 'OneToMany' || stagedNav?.cardinality === 'ManyToMany';

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 10, maxWidth: 720 }}>
      <Field
        label="Navigation property"
        hint={`${available.length} eligible relationship${available.length === 1 ? '' : 's'} on this entity. Type to search.`}
      >
        <div style={{ display: 'flex', gap: 8, alignItems: 'flex-end' }}>
          <Combobox
            size="medium"
            value={staged ?? filterText}
            selectedOptions={staged ? [staged] : []}
            placeholder="Search relationships…"
            onChange={(e) => {
              const text = (e.target as HTMLInputElement).value;
              setFilterText(text);
              // Typing clears any prior staged selection (the user is searching)
              if (staged && text !== staged) setStaged(null);
            }}
            onOptionSelect={(_, d) => {
              setStaged(d.optionValue ?? null);
              setFilterText('');
            }}
            style={{ flexGrow: 1 }}
          >
            {filtered.map((n) => {
              const tgt = findTable(n.targetEntity);
              const isCol = n.cardinality === 'OneToMany' || n.cardinality === 'ManyToMany';
              return (
                <Option key={n.name} value={n.name} text={n.name}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 2, minWidth: 0 }}>
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
                      → {tgt?.displayName ?? n.targetEntity}
                      <span style={{ marginLeft: 8, fontFamily: tokens.fontFamilyMonospace }}>
                        {n.cardinality}
                      </span>
                      <span style={{ marginLeft: 8 }}>
                        {n.cardinality === 'ManyToMany'
                          ? '(N:N — no nesting)'
                          : isCol
                            ? '(collection)'
                            : '(single)'}
                      </span>
                    </span>
                  </div>
                </Option>
              );
            })}
            {filtered.length === 0 && (
              <Option value="__none" text="" disabled>
                <span style={{ color: tokens.colorNeutralForeground3, fontSize: 12 }}>
                  No matches for "{filterText}"
                </span>
              </Option>
            )}
          </Combobox>
          <Button
            appearance="primary"
            icon={<Add20Regular />}
            disabled={!staged || capExceeded}
            onClick={() => {
              if (!staged) return;
              onAdd(staged);
              setStaged(null);
              setFilterText('');
            }}
          >
            Add
          </Button>
        </div>
      </Field>

      {stagedNav && (
        <div
          className={mergeClasses()}
          style={{
            padding: '10px 12px',
            borderRadius: tokens.borderRadiusMedium,
            backgroundColor: tokens.colorNeutralBackground2,
            border: `1px solid ${tokens.colorNeutralStroke2}`,
            display: 'flex',
            alignItems: 'flex-start',
            gap: 10,
          }}
        >
          <BranchFork20Regular style={{ color: tokens.colorBrandForeground1, marginTop: 2 }} />
          <div style={{ flexGrow: 1, minWidth: 0 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
              <span style={{ fontFamily: tokens.fontFamilyMonospace, fontWeight: 600 }}>
                {stagedNav.name}
              </span>
              <span style={{ color: tokens.colorNeutralForeground3, fontSize: 12 }}>
                → {stagedTarget?.displayName ?? stagedNav.targetEntity}
              </span>
              <Badge appearance="ghost" size="small">
                {stagedNav.cardinality}
              </Badge>
            </div>
            <div style={{ display: 'flex', gap: 4, marginTop: 6, flexWrap: 'wrap' }}>
              <Caption1 style={{ color: tokens.colorNeutralForeground3, marginRight: 6 }}>
                Inner clauses you'll be able to set:
              </Caption1>
              <Badge size="small" appearance="ghost">
                $select
              </Badge>
              <Badge size="small" appearance="ghost">
                $filter
              </Badge>
              {stagedIsCol && (
                <>
                  <Badge size="small" appearance="ghost">
                    $orderby
                  </Badge>
                  <Badge size="small" appearance="ghost">
                    $top
                  </Badge>
                </>
              )}
              {stagedNav.cardinality !== 'ManyToMany' && (
                <Badge size="small" appearance="ghost">
                  $expand (nested ManyToOne)
                </Badge>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
