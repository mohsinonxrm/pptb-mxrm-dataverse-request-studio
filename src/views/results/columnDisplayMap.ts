// Build a column-key → ColumnDisplayInfo map for the results grid.
//
// Each row column in a Dataverse OData response is one of:
//   1. A root-entity attribute              e.g. `title`
//   2. A root-entity lookup VALUE column    e.g. `_customerid_value`
//   3. A dotted path into a $expand'd nav   e.g. `customerid_account.name`
//                                            or  `customerid_account.primarycontactid.firstname`
//
// Given the request's `state.select` + `state.expand` and the root
// `TableMeta`, we resolve every result column to:
//   • the parent entity's TableMeta (root or a nested target)
//   • the leaf attribute's ColumnMeta on that table
//   • a friendly display name for the grid header (waterfall:
//     metadata.displayName → cleaned column key)
//
// Resolution is lazy w.r.t. metadata availability — when a target table
// isn't in the live registry yet we still surface the leaf attribute
// name so the header at least reads like a column, not "undefined".

import { findTable, type ColumnMeta, type NavProperty, type TableMeta } from '../../mock/metadata';
import type { ExpandSpec } from '../../editors/ExpandEditor';

export interface ColumnDisplayInfo {
  /** Where this attribute physically lives. */
  parentTable?: TableMeta;
  /** The leaf attribute's metadata (may be undefined while loading). */
  column?: ColumnMeta;
  /** Header display string the grid should render. */
  displayName: string;
  /** Logical-name string for `useLogicalNames` display mode. */
  logicalName: string;
  /** True if the result column is a lookup VALUE (`_<attr>_value`). */
  isLookupValue: boolean;
  /** Nav-property segments walked to reach this attribute (for nested expands). */
  navPath: string[];
}

/** Walks an `ExpandSpec` tree to find the spec whose nav path matches `path`. */
function findExpandByPath(items: ExpandSpec[], path: string[]): ExpandSpec | undefined {
  if (path.length === 0) return undefined;
  const [head, ...rest] = path;
  const match = items.find(it => it.nav === head);
  if (!match) return undefined;
  if (rest.length === 0) return match;
  return findExpandByPath(match.nestedExpand ?? [], rest);
}

/**
 * Walk a `TableMeta` + an array of nav names, returning the deepest
 * `TableMeta` we could resolve. Falls back to the last successfully
 * resolved table when a nav target isn't in the live registry yet.
 */
function walkToTargetTable(root: TableMeta, navPath: string[]): {
  table: TableMeta | undefined;
  lastNav: NavProperty | undefined;
} {
  let current: TableMeta | undefined = root;
  let lastNav: NavProperty | undefined;
  for (const name of navPath) {
    const nav = current?.navigationProperties.find(n => n.name === name);
    if (!nav) return { table: current, lastNav };
    lastNav = nav;
    current = findTable(nav.targetEntity);
    if (!current) return { table: undefined, lastNav };
  }
  return { table: current, lastNav };
}

/**
 * Cleanup helper: `_foo_value` → `foo`. Otherwise returns input untouched.
 */
function stripLookupWrapper(name: string): string {
  if (name.startsWith('_') && name.endsWith('_value')) return name.slice(1, -6);
  return name;
}

/**
 * Resolve one (dotted) result-column key against the request's
 * `expand` + the root `TableMeta`.
 */
export function resolveColumnDisplay(
  columnKey: string,
  rootTable: TableMeta | undefined,
  rootExpand: ExpandSpec[] | undefined,
  useLogicalNames: boolean,
): ColumnDisplayInfo {
  // Split on the LAST `/.../` segment to peel off annotation suffixes —
  // by this point `flattenRow` should already have kept @-suffixes on the
  // key, but we still allow plain dotted names like `acct.foo`.
  const parts = columnKey.split('.');
  const leaf = parts[parts.length - 1];
  const navPath = parts.slice(0, -1);

  // Resolve the parent table by walking the nav path.
  const { table: parentTable } = rootTable
    ? walkToTargetTable(rootTable, navPath)
    : { table: undefined };
  void findExpandByPath; // reserved for richer display-name resolution later

  // Find the leaf attribute. Lookup VALUE columns (`_xxx_value`) carry the
  // value at that exact key, but the underlying ColumnMeta is keyed by the
  // bare attribute name.
  const isLookupValue = leaf.startsWith('_') && leaf.endsWith('_value');
  const lookupKey = isLookupValue ? stripLookupWrapper(leaf) : leaf;
  const column = parentTable?.columns.find(c => c.logicalName === lookupKey);

  // Header display name waterfall.
  const leafDisplay = useLogicalNames
    ? (column?.logicalName ?? lookupKey)
    : (column?.displayName ?? lookupKey);

  let displayName = leafDisplay;
  if (navPath.length > 0) {
    // For nested expansions, qualify the header with the immediate parent
    // entity's display name (or logical name in logical-names mode) so the
    // user can tell "Email" on Contact from "Email" on Account.
    const parentLabel = useLogicalNames
      ? (parentTable?.logicalName ?? navPath[navPath.length - 1])
      : (parentTable?.displayName ?? navPath[navPath.length - 1]);
    displayName = `${leafDisplay} (${parentLabel})`;
  }

  const logicalName =
    navPath.length > 0
      ? [...navPath, lookupKey].join('.')
      : lookupKey;

  return {
    parentTable,
    column,
    displayName,
    logicalName,
    isLookupValue,
    navPath,
  };
}
