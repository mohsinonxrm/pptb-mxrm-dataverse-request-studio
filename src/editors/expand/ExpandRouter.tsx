// Routes the active sidebar path (under $expand) to the right scoped editor.
//
// Path shape: 'expand' | 'expand/<id>' | 'expand/<id>/select' | 'expand/<id>/filter'
//           | 'expand/<id>/orderby' | 'expand/<id>/top' | 'expand/<id>/expand'
//           | 'expand/<id>/expand/<innerId>/...' (recursive — arbitrary depth)
//
// At every level the right pane reuses the same per-clause editors as the root
// level (SelectEditor / FilterEditor / OrderbyEditor / TopEditor) — but scoped
// to the related entity of that expand, and dispatching back through the
// nested `updateExpand` helper so deeply-nested edits don't leak.

import { Caption1, MessageBar, MessageBarBody, Spinner, tokens } from '@fluentui/react-components';
import { useStudioStyles } from '../../primitives/styles';
import { findTable } from '../../mock/metadata';
import { useLiveTable } from '../../host/useLiveMetadata';
import { type ExpandSpec } from '../ExpandEditor';
import {
  findExpandById, findExpandParentEntity, getExpandTarget, isCollectionExpand,
  updateExpand, removeExpand, addExpand, hasAnyNestedExpand,
} from './expandTree';
import { ExpandOverview } from './ExpandOverview';
import { SelectEditor } from '../SelectEditor';
import { FilterEditor } from '../filter/FilterEditor';
import { OrderbyEditor, type OrderbySpec } from '../OrderbyEditor';
import { TopEditor } from '../TopCountEditors';
import { emptyTree } from '../filter/filterTree';
import type { RequestGroup } from '../../registry/requestTypes';

export interface ExpandRouterProps {
  /** Slash-delimited path: 'expand', 'expand/<id>', 'expand/<id>/<clause>', etc. */
  path: string;
  /** Root entity logical name (the mode's target table). */
  rootEntity: string;
  /** Top-level expand list (the mode's state.expand). */
  rootExpand: ExpandSpec[];
  /** Setter for the top-level expand list. */
  setRootExpand: (items: ExpandSpec[]) => void;
  /** Notify parent that the active path should change (e.g. after Edit / Remove / Add). */
  setActivePath: (path: string) => void;
  group?: RequestGroup;
}

export function ExpandRouter({
  path, rootEntity, rootExpand, setRootExpand, setActivePath, group = 'read',
}: ExpandRouterProps) {
  const segments = path.split('/');
  // segments[0] === 'expand' guaranteed by caller

  // ── Find the expand spec the active path is pointing at ────
  // We walk segments from index 1, jumping over inner 'expand' boundaries.
  // Path patterns:
  //   ['expand']                                                   -> root overview
  //   ['expand', '<id>']                                            -> that expand's "summary" (auto-route to $select)
  //   ['expand', '<id>', '<clause>']                               -> scoped clause editor
  //   ['expand', '<id>', 'expand']                                 -> nested overview
  //   ['expand', '<id>', 'expand', '<innerId>', ...]               -> recurse
  //
  // We also build a breadcrumb trail with labels per segment.
  type Walk = {
    activeExpand: ExpandSpec | null;
    activeClause: string | null;       // 'select' | 'filter' | 'orderby' | 'top' | 'expand' | null
    parentEntity: string;              // entity at the active level (parent of `activeExpand` siblings)
    siblings: ExpandSpec[];            // expand list at the active level
    breadcrumb: { label: string; path: string }[];
    /**
     * Cardinality of the IMMEDIATE parent $expand (the one we're nested
     * UNDER at this level). `null` means we're at the root — no
     * enclosing expand. Drives `availableNavsAt`'s nesting rules and
     * the parent-aware info MessageBars in ExpandOverview.
     */
    parentCardinality: 'OneToMany' | 'ManyToOne' | 'ManyToMany' | null;
  };

  const walked: Walk = (() => {
    let entity = rootEntity;
    let siblings = rootExpand;
    let breadcrumb: { label: string; path: string }[] = [
      { label: findTable(rootEntity)?.displayName ?? rootEntity, path: 'expand' },
    ];
    // The cardinality of the most-recently-stepped-into expand. When we
    // pivot into another expand's `nestedExpand`, this becomes the
    // cardinality of THAT expand's nav — which is what governs what can
    // be nested beneath it. Starts at null (root level).
    let parentCardinality: Walk['parentCardinality'] = null;
    let cursor: ExpandSpec | null = null;
    let i = 1;
    let pathSoFar = 'expand';

    while (i < segments.length) {
      const idSeg = segments[i];
      const found = siblings.find(x => x.id === idSeg);
      if (!found) {
        return { activeExpand: null, activeClause: null, parentEntity: entity, siblings, breadcrumb, parentCardinality };
      }
      cursor = found;
      pathSoFar = `${pathSoFar}/${idSeg}`;
      const nav = findTable(entity)?.navigationProperties.find(n => n.name === found.nav);
      breadcrumb.push({ label: found.nav, path: pathSoFar });
      if (!nav) {
        return { activeExpand: cursor, activeClause: null, parentEntity: entity, siblings, breadcrumb, parentCardinality };
      }
      const next = segments[i + 1];
      if (!next) {
        return { activeExpand: cursor, activeClause: 'select', parentEntity: entity, siblings, breadcrumb, parentCardinality };
      }
      if (next === 'expand') {
        // Step into this expand's nestedExpand — the new level's parent
        // cardinality is THIS expand's nav cardinality.
        parentCardinality = nav.cardinality;
        entity = nav.targetEntity;
        siblings = cursor.nestedExpand ?? [];
        pathSoFar = `${pathSoFar}/expand`;
        breadcrumb.push({ label: 'expand', path: pathSoFar });
        i += 2;
        continue;
      }
      return { activeExpand: cursor, activeClause: next, parentEntity: entity, siblings, breadcrumb, parentCardinality };
    }
    return { activeExpand: null, activeClause: null, parentEntity: entity, siblings, breadcrumb, parentCardinality };
  })();

  // ── Dispatchers ────────────────────────────────────────────
  const dispatchPatch = (id: string, patch: Partial<ExpandSpec>) => {
    setRootExpand(updateExpand(rootExpand, id, patch));
  };
  const dispatchRemove = (id: string) => {
    setRootExpand(removeExpand(rootExpand, id));
    // Move focus back up one level
    const idx = path.indexOf(`/${id}`);
    const parent = idx > 0 ? path.slice(0, idx) : 'expand';
    setActivePath(parent);
  };
  const dispatchAdd = (parentId: string | null, navName: string) => {
    setRootExpand(addExpand(rootExpand, parentId, navName));
    // Focus the newly-added expand — it gets a fresh id; we don't have it here,
    // so simply stay on the overview (the user can click "Edit" on the new row).
  };

  // ── Hoisted target-entity lookup (must run unconditionally) ─
  // React's Rules of Hooks: every hook must run on every render in the
  // same order. We compute the active-expand's target-entity logical
  // name up here (may be null when there's no active expand or no
  // matching nav) and pass it into `useLiveTable` — null is a no-op
  // for the loader so the hook is safe to call eagerly.
  const activeExpandForHook = walked.activeExpand;
  const parentEntityForExpand = activeExpandForHook
    ? (findExpandParentEntity(rootExpand, activeExpandForHook.id, rootEntity) ?? walked.parentEntity)
    : walked.parentEntity;
  const parentTblForHook = findTable(parentEntityForExpand);
  const navMetaForHook = activeExpandForHook
    ? parentTblForHook?.navigationProperties.find(n => n.name === activeExpandForHook.nav)
    : undefined;
  const { loading: targetLoading } = useLiveTable(navMetaForHook?.targetEntity ?? null);

  // No bulk warming here — the active expand's target is already warmed
  // above via `useLiveTable(navMetaForHook?.targetEntity)`. Sibling cards
  // on the overview fall back to the raw target logical name; the user
  // clicking into one of them triggers its load.

  // ── Render ─────────────────────────────────────────────────
  if (walked.activeExpand === null) {
    // Overview at the active level (root or nested 'expand')
    return (
      <ExpandOverview
        parentEntity={walked.parentEntity}
        siblings={walked.siblings}
        rootExpand={rootExpand}
        breadcrumb={walked.breadcrumb}
        parentCardinality={walked.parentCardinality}
        onAdd={(navName) => {
          const enclosingIdx = segments.lastIndexOf('expand');
          const parentId = enclosingIdx > 0 ? segments[enclosingIdx - 1] : null;
          dispatchAdd(parentId, navName);
        }}
        onPickExisting={(expandId) => {
          setActivePath(`${path}/${expandId}/select`);
        }}
        onRemove={(id) => dispatchRemove(id)}
        group={group}
      />
    );
  }

  // We have an active expand and a clause — reuse the values computed
  // up at the hook hoist (parentEntityForExpand / navMetaForHook) so the
  // synchronous lookups stay consistent with what `useLiveTable` warmed.
  const activeExpand = walked.activeExpand;
  const navMeta = navMetaForHook;
  const target = getExpandTarget(parentEntityForExpand, activeExpand);
  const isCol = isCollectionExpand(parentEntityForExpand, activeExpand);

  if (!target) {
    // Still loading the target table — show a spinner instead of an
    // error. If the nav definitely points at an unknown entity (no
    // targetEntity name in the relationship) we keep the error message.
    if (navMeta?.targetEntity && targetLoading) {
      return (
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: 12 }}>
          <Spinner size="extra-small" />
          <span style={{ color: tokens.colorNeutralForeground3, fontSize: 12 }}>
            Loading metadata for <code>{navMeta.targetEntity}</code>…
          </span>
        </div>
      );
    }
    return (
      <MessageBar layout="multiline" intent="error">
        <MessageBarBody>
          Couldn't resolve the target entity for navigation property <code>{activeExpand.nav}</code>.
          {navMeta?.targetEntity && (
            <> Tried <code>{navMeta.targetEntity}</code>.</>
          )}
        </MessageBarBody>
      </MessageBar>
    );
  }

  const clause = walked.activeClause;
  // $top / $orderby are NOT supported in any expand once the query contains
  // a nested $expand anywhere — Dataverse returns:
  //   "Only $select and $filter clause can be provided while doing $expand on
  //    many-to-one relationship or nested one-to-many relationship."
  const queryHasNestedExpand = hasAnyNestedExpand(rootExpand);
  return (
    <ScopedClause
      clause={clause ?? 'select'}
      targetEntity={target.logicalName}
      expand={activeExpand}
      isCollection={isCol}
      breadcrumb={walked.breadcrumb}
      queryHasNestedExpand={queryHasNestedExpand}
      onPatch={(patch) => dispatchPatch(activeExpand.id, patch)}
      group={group}
    />
  );
}

// ────────────────────────────────────────────────────────────
// Scoped per-clause renderer
// ────────────────────────────────────────────────────────────
function ScopedClause({
  clause, targetEntity, expand, isCollection, breadcrumb, queryHasNestedExpand, onPatch, group,
}: {
  clause: string;
  targetEntity: string;
  expand: ExpandSpec;
  isCollection: boolean;
  breadcrumb: { label: string; path: string }[];
  /** True when the overall query has any nested $expand — disables $top/$orderby everywhere */
  queryHasNestedExpand: boolean;
  onPatch: (patch: Partial<ExpandSpec>) => void;
  group: RequestGroup;
}) {
  const s = useStudioStyles();
  // Breadcrumb header common to every nested clause
  const crumb = (
    <div style={{ marginBottom: 14, display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
      <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Editing inside:</Caption1>
      {breadcrumb.map((b, i) => (
        <span key={b.path} style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
          {i > 0 && <span style={{ color: tokens.colorNeutralForeground4 }}>→</span>}
          <Caption1 style={{
            fontFamily: tokens.fontFamilyMonospace,
            color: i === breadcrumb.length - 1 ? tokens.colorBrandForeground1 : tokens.colorNeutralForeground2,
            fontWeight: i === breadcrumb.length - 1 ? 600 : 400,
          }}>{b.label}</Caption1>
        </span>
      ))}
    </div>
  );
  void s;

  // Single-valued navs (ManyToOne / lookups):
  //   Dataverse error reads: "Only $select and $filter clause can be provided
  //   while doing $expand on many-to-one relationship".
  //   So: $select ✓, $filter ✓, $orderby ✗, $top ✗, nested $expand ✓ (ManyToOne only).
  if (!isCollection && (clause === 'orderby' || clause === 'top')) {
    return (
      <div>
        {crumb}
        <MessageBar layout="multiline" intent="info">
          <MessageBarBody>
            <code>${clause}</code> isn't allowed on a single-valued navigation property (<code>{expand.nav}</code>).
            Per the Dataverse docs, an <code>$expand</code> on a many-to-one relationship only accepts
            <code> $select</code>, <code>$filter</code>, and a nested <code>$expand</code>.
          </MessageBarBody>
        </MessageBar>
      </div>
    );
  }

  // $top / $orderby are blocked anywhere when the query contains any nested
  // expand — Dataverse rejects with:
  //   "Only $select and $filter clause can be provided while doing $expand
  //    on many-to-one relationship or nested one-to-many relationship."
  const blockedByNested = queryHasNestedExpand && (clause === 'top' || clause === 'orderby');

  switch (clause) {
    case 'select':
      return (
        <div>
          {crumb}
          <SelectEditor
            table={targetEntity}
            selectedIds={expand.select}
            setSelectedIds={(ids) => onPatch({ select: ids })}
            group={group}
          />
        </div>
      );
    case 'filter':
      return (
        <div>
          {crumb}
          <FilterEditor
            table={targetEntity}
            tree={expand.filter ?? emptyTree()}
            setTree={(t) => onPatch({ filter: t })}
            group={group}
          />
        </div>
      );
    case 'orderby':
      return (
        <div>
          {crumb}
          {blockedByNested && <NestedBlockedBanner clause="$orderby" />}
          <OrderbyEditor
            table={targetEntity}
            items={expand.orderby}
            setItems={(items: OrderbySpec[]) => onPatch({ orderby: items })}
            group={group}
          />
        </div>
      );
    case 'top':
      return (
        <div>
          {crumb}
          {blockedByNested && <NestedBlockedBanner clause="$top" />}
          <TopEditor
            top={expand.top ?? null}
            setTop={(n) => onPatch({ top: n })}
            maxPageSize={null}
            group={group}
          />
        </div>
      );
    default:
      return (
        <div>
          {crumb}
          <MessageBar layout="multiline" intent="info">
            <MessageBarBody>
              Unknown clause <code>{clause}</code> for expand <code>{expand.nav}</code>.
            </MessageBarBody>
          </MessageBar>
        </div>
      );
  }
}

function NestedBlockedBanner({ clause }: { clause: string }) {
  return (
    <MessageBar layout="multiline" intent="error" style={{ marginBottom: 12 }}>
      <MessageBarBody>
        The query contains a nested <code>$expand</code>, so <code>{clause}</code> is no longer accepted by Dataverse on any expand —
        the server returns: <em>"Only $select and $filter clause can be provided while doing $expand on many-to-one relationship or nested one-to-many relationship."</em>
        Remove all nested expands first, or accept that this value will be ignored.
      </MessageBarBody>
    </MessageBar>
  );
}

void findExpandById; // re-export for tests / future use
