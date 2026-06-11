// navPath — request-tree reference discovery for metadata pre-warming.
//
// The $filter encoder (and the picker, antipattern scanner, etc.) resolve a
// nav-path leaf's type against the RELATED entity's metadata (see
// `resolveNavPath` in mock/metadata.ts). That only works if the related
// entity has been fetched into the live registry. Today metadata is loaded
// lazily — only when the user manually drills the column picker — so a request
// that arrives fully-formed (saved-request reload, pasted OData URL, or a
// direct Execute) can hit the encoder before the related entity is present,
// and the leaf falls back to string-quoting (issue #33's failure mode).
//
// `collectReferencedEntities` walks a request's clause trees ($filter,
// $expand, $apply, $orderby) and returns every RELATED entity logical name
// referenced via a nav-path, lambda, or expand. The `useWarmReferencedTables`
// hook feeds that list to `metadata.getTable` so the registry is warm before
// the encoder runs — independent of whether any editor was ever opened.
//
// Resolution is lazy-aware and self-healing: a hop whose target isn't loaded
// yet is still ADDED to the set (so it gets warmed), but the walk stops there
// for this pass. Warming triggers a registry update; the hook re-runs and the
// walk now reaches one level deeper. Repeats until the whole graph is warm.

import { findTable, type TableMeta } from '../mock/metadata';
import type { FilterGroup } from '../editors/filter/filterTree';
import type { ExpandSpec } from '../editors/ExpandEditor';
import type { OrderbySpec } from '../editors/OrderbyEditor';
import type { ApplySpec } from '../editors/ApplyEditor';

export interface ReferencedTreeInput {
  filter?: FilterGroup;
  expand?: ExpandSpec[];
  apply?: ApplySpec;
  orderby?: OrderbySpec[];
}

/**
 * Add every related entity along a nav-path to `out`, walking N:1 hops as far
 * as the registry currently resolves. `alias` strips an enclosing lambda
 * alias. A bare column (no nav hop) contributes nothing.
 */
function addPathEntities(
  table: TableMeta | undefined,
  path: string,
  out: Set<string>,
  alias?: string,
): void {
  if (!table || !path) return;
  let p = path;
  if (alias && p.startsWith(alias + '/')) p = p.slice(alias.length + 1);
  const segs = p.split('/').filter(Boolean);
  let current: TableMeta | undefined = table;
  for (let i = 0; i < segs.length - 1; i++) {
    const nav = current?.navigationProperties.find(
      (n) => n.name === segs[i] && n.cardinality === 'ManyToOne',
    );
    if (!nav) return;
    out.add(nav.targetEntity);
    current = findTable(nav.targetEntity); // undefined → warm it, resolve deeper next pass
    if (!current) return;
  }
}

/** Walk a $filter (or $apply pre-filter) tree, collecting nav + lambda targets. */
function walkFilter(
  group: FilterGroup,
  table: TableMeta | undefined,
  out: Set<string>,
  alias?: string,
): void {
  if (!group) return;
  for (const node of group.rules) {
    if (node.type === 'rule' || node.type === 'function') {
      if (node.col) addPathEntities(table, node.col, out, alias);
    } else if (node.type === 'group') {
      walkFilter(node, table, out, alias);
    } else if (node.type === 'lambda') {
      const nav = table?.navigationProperties.find((n) => n.name === node.nav);
      if (!nav) continue;
      out.add(nav.targetEntity);
      // Inner predicate is scoped to the lambda's target entity.
      walkFilter(node.inner, findTable(nav.targetEntity), out, node.alias);
    }
  }
}

/** Walk a (possibly nested) $expand tree, collecting each target + inner clauses. */
function walkExpand(
  items: ExpandSpec[],
  parentTable: TableMeta | undefined,
  out: Set<string>,
): void {
  for (const it of items) {
    const nav = parentTable?.navigationProperties.find((n) => n.name === it.nav);
    if (!nav) continue;
    out.add(nav.targetEntity);
    const target = findTable(nav.targetEntity);
    if (it.filter) walkFilter(it.filter, target, out);
    if (it.orderby) for (const o of it.orderby) addPathEntities(target, o.col, out);
    if (it.nestedExpand) walkExpand(it.nestedExpand, target, out);
  }
}

/**
 * Return the related entity logical names referenced by a request's clause
 * trees, resolvable against the current registry. The root table itself is
 * excluded (it's loaded by the mode's own `useLiveTable`).
 */
export function collectReferencedEntities(
  rootTableLogical: string,
  input: ReferencedTreeInput,
): string[] {
  const root = findTable(rootTableLogical);
  if (!root) return [];
  const out = new Set<string>();

  if (input.filter) walkFilter(input.filter, root, out);
  if (input.apply?.prefilter) walkFilter(input.apply.prefilter, root, out);
  if (input.apply?.groupby) for (const g of input.apply.groupby) addPathEntities(root, g, out);
  if (input.apply?.aggregates)
    for (const a of input.apply.aggregates) addPathEntities(root, a.col, out);
  if (input.orderby) for (const o of input.orderby) addPathEntities(root, o.col, out);
  if (input.expand) walkExpand(input.expand, root, out);

  out.delete(rootTableLogical);
  return [...out];
}
