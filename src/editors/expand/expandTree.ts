// Tree helpers for the nested $expand model.
//
// Each `ExpandSpec` carries an `id` (UUID, via filterTree.newId) and an optional
// `nestedExpand: ExpandSpec[]`. These helpers walk the tree by id so updates
// to deeply-nested expands don't accidentally touch siblings.

import { type ExpandSpec, makeExpandSpec } from '../ExpandEditor';
import { findTable, type NavProperty, type TableMeta } from '../../mock/metadata';

// ────────────────────────────────────────────────────────────
// Lookups
// ────────────────────────────────────────────────────────────
export function findExpandById(items: ExpandSpec[], id: string): ExpandSpec | undefined {
  for (const it of items) {
    if (it.id === id) return it;
    if (it.nestedExpand) {
      const inner = findExpandById(it.nestedExpand, id);
      if (inner) return inner;
    }
  }
  return undefined;
}

/**
 * Resolves the parent-entity logical name for an expand item at any depth.
 * For a root-level expand the parent is the request's target entity; for nested
 * expands it's the target of the enclosing expand's nav property.
 */
export function findExpandParentEntity(
  items: ExpandSpec[],
  targetExpandId: string,
  rootEntity: string,
): string | undefined {
  function walk(level: ExpandSpec[], parentEntity: string): string | undefined {
    for (const it of level) {
      if (it.id === targetExpandId) return parentEntity;
      if (it.nestedExpand && it.nestedExpand.length > 0) {
        const tbl = findTable(parentEntity);
        const nav = tbl?.navigationProperties.find((n) => n.name === it.nav);
        if (nav) {
          const found = walk(it.nestedExpand, nav.targetEntity);
          if (found) return found;
        }
      }
    }
    return undefined;
  }
  return walk(items, rootEntity);
}

/** Returns the related-table metadata for the target of an expand item. */
export function getExpandTarget(
  parentTableLogical: string,
  expand: ExpandSpec,
): TableMeta | undefined {
  const parent = findTable(parentTableLogical);
  const nav = parent?.navigationProperties.find((n) => n.name === expand.nav);
  return nav ? findTable(nav.targetEntity) : undefined;
}

/** True when the expand's nav is collection-valued (1:N or N:N). */
export function isCollectionExpand(parentTableLogical: string, expand: ExpandSpec): boolean {
  const parent = findTable(parentTableLogical);
  const nav = parent?.navigationProperties.find((n) => n.name === expand.nav);
  return nav?.cardinality === 'OneToMany' || nav?.cardinality === 'ManyToMany';
}

// ────────────────────────────────────────────────────────────
// Mutation helpers (immutable)
// ────────────────────────────────────────────────────────────
export function updateExpand(
  items: ExpandSpec[],
  id: string,
  patch: Partial<ExpandSpec>,
): ExpandSpec[] {
  return items.map((it) => {
    if (it.id === id) return { ...it, ...patch };
    if (it.nestedExpand) return { ...it, nestedExpand: updateExpand(it.nestedExpand, id, patch) };
    return it;
  });
}

export function removeExpand(items: ExpandSpec[], id: string): ExpandSpec[] {
  return items
    .filter((it) => it.id !== id)
    .map((it) =>
      it.nestedExpand ? { ...it, nestedExpand: removeExpand(it.nestedExpand, id) } : it,
    );
}

/**
 * Add a new expand under the given parentId (or at root if parentId is null).
 * Pass the chosen nav property name.
 */
export function addExpand(
  items: ExpandSpec[],
  parentId: string | null,
  navName: string,
): ExpandSpec[] {
  const child = makeExpandSpec(navName);
  if (parentId === null) return [...items, child];
  return items.map((it) => {
    if (it.id === parentId) {
      return { ...it, nestedExpand: [...(it.nestedExpand ?? []), child] };
    }
    if (it.nestedExpand)
      return { ...it, nestedExpand: addExpand(it.nestedExpand, parentId, navName) };
    return it;
  });
}

// ────────────────────────────────────────────────────────────
// Validation / nav availability (sourced from
// learn.microsoft.com/.../webapi/query/join-tables — see audit notes
// at the call sites for the exact rules)
// ────────────────────────────────────────────────────────────

/** Dataverse hard cap — "You can include up to 15 $expand options in a query." */
export const MAX_EXPANDS_PER_QUERY = 15;

export interface NavAvailabilityOpts {
  /**
   * Cardinality of the IMMEDIATE parent $expand (the one this level is
   * nested under). `null` means we're at the top level — no restrictions.
   *
   * This replaces the older `isNested` flag, which collapsed all nested
   * levels into a single "allow only ManyToOne" rule. That was too strict
   * — per the docs, an N:1 expand can host any cardinality below it
   * (the classic `tasks → contact → account → systemuser` chain in the
   * Microsoft join-tables docs is precisely that), and a 1:N expand can
   * host N:1 AND further 1:N. Only N:N parents truly forbid all nesting.
   */
  parentCardinality: NavProperty['cardinality'] | null;
}

/**
 * Returns the navigation properties that the user can add to a given level.
 *
 * Rules (per /webapi/query/join-tables):
 *
 *   - The same nav can't appear twice within one $expand=...,...
 *   - Parent N:N → NOTHING nestable. Dataverse error:
 *       "The navigation property '<NAME>' cannot be expanded.
 *        Only many-to-one relationships are supported for nested expansion."
 *     (The wording is misleading — the error fires whenever the OUTER is
 *      N:N, regardless of the inner cardinality. We just hide everything.)
 *   - Parent 1:N → allow N:1 and 1:N (NOT N:N). Allowed by the docs'
 *     `accounts → contact_customer_accounts → owninguser` example, but
 *     N:N targets fail the "no N:N anywhere when nested $expand is in the
 *     query" rule, so we keep them out at this level.
 *   - Parent N:1 → allow any cardinality (N:1, 1:N, N:N). Classic deep-
 *     chain expand is N:1 → N:1 → N:1, but the docs don't restrict the
 *     inner cardinality when the outer is single-valued.
 *   - Root (parentCardinality == null) → allow any.
 */
export function availableNavsAt(
  parentTableLogical: string,
  existingAtLevel: ExpandSpec[],
  opts: NavAvailabilityOpts,
): NavProperty[] {
  if (opts.parentCardinality === 'ManyToMany') return [];
  const parent = findTable(parentTableLogical);
  if (!parent) return [];
  const used = new Set(existingAtLevel.map((e) => e.nav));
  return parent.navigationProperties.filter((n) => {
    if (used.has(n.name)) return false;
    // Inside a 1:N collection, N:N navs are not safely nestable in
    // Dataverse (the docs' "N:N when nested $expand exists in the query"
    // restriction). Filter them out at this level.
    if (opts.parentCardinality === 'OneToMany' && n.cardinality === 'ManyToMany') {
      return false;
    }
    return true;
  });
}

/** Total nav count anywhere in the tree — used against the 15 hard cap. */
export function totalExpandCount(items: ExpandSpec[]): number {
  let n = 0;
  for (const it of items) {
    n += 1;
    if (it.nestedExpand) n += totalExpandCount(it.nestedExpand);
  }
  return n;
}

/** Max depth of the tree (used for the depth-warning hint). */
export function maxExpandDepth(items: ExpandSpec[]): number {
  let max = 0;
  for (const it of items) {
    const inner =
      it.nestedExpand && it.nestedExpand.length > 0 ? 1 + maxExpandDepth(it.nestedExpand) : 1;
    if (inner > max) max = inner;
  }
  return max;
}

/**
 * True if the tree contains any nested $expand (an expand whose nestedExpand
 * has at least one entry). Used to gate $top/$orderby — Dataverse rejects
 * those when ANY nested $expand exists in the query.
 *
 * Error returned by Dataverse otherwise:
 *   "Only $select and $filter clause can be provided while doing $expand on
 *    many-to-one relationship or nested one-to-many relationship."
 */
export function hasAnyNestedExpand(items: ExpandSpec[]): boolean {
  for (const it of items) {
    if (it.nestedExpand && it.nestedExpand.length > 0) return true;
    if (it.nestedExpand && hasAnyNestedExpand(it.nestedExpand)) return true;
  }
  return false;
}

/**
 * True iff the tree contains nesting AND at least one expand in the tree
 * targets a collection-valued navigation property (1:N or N:N).
 *
 * This is the EXACT trigger for Dataverse's runtime restriction on top-
 * level `$top` and `$orderby`. The error message is:
 *
 *   "Only $select, $filter and $orderby clauses can be provided at top
 *    level while doing $expand on nested one-to-many relationship."
 *
 * (yes, the error contradicts itself by listing $orderby and then
 * forbidding it via "$top OR $orderby" — empirically, $top fails first
 * and $orderby is also blocked when present.)
 *
 * Empirical confirmation across our 7-test paste battery against
 * crm412218:
 *
 *   Test 3 — N:1 → N:1 → N:1 chain with $top → WORKED (no 1:N anywhere)
 *   Test 7 — single-level N:N, no nesting → WORKED (no nesting)
 *   All others (any 1:N involved in nested chain + $top) → blocked
 *
 * So: `hasAnyNestedExpand` alone is too aggressive (Test 3 disproves it),
 * and "no nesting at all" misses Tests 1/2/4/5/6. The combined check is
 * what matches Dataverse's actual behavior.
 *
 * `rootEntity` is required because expand items only carry the nav name;
 * we resolve cardinality by walking the metadata registry as we descend.
 */
export function hasCollectionInvolvedNestedExpand(
  items: ExpandSpec[],
  rootEntity: string,
): boolean {
  if (!hasAnyNestedExpand(items)) return false;
  return containsCollectionExpand(items, rootEntity);
}

function containsCollectionExpand(items: ExpandSpec[], parentEntity: string): boolean {
  const parent = findTable(parentEntity);
  if (!parent) return false;
  for (const it of items) {
    const nav = parent.navigationProperties.find((n) => n.name === it.nav);
    if (!nav) continue;
    if (nav.cardinality === 'OneToMany' || nav.cardinality === 'ManyToMany') return true;
    if (it.nestedExpand && containsCollectionExpand(it.nestedExpand, nav.targetEntity)) return true;
  }
  return false;
}
