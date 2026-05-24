// Query anti-pattern detector — emits Advisory[] from a read-mode state.
//
// Sourced verbatim from
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/query-antipatterns
//
// The MS Learn doc enumerates seven server-side anti-pattern identifiers that
// Dataverse returns in `PerformanceValidationIssuesCauseTimeout` errors. We
// detect the same anti-patterns *client-side* so the user gets pre-flight
// feedback before paying the round-trip cost (and the throttling penalty).
//
//   PerformanceLeadingWildCard       — handled inside filterTree (see stripLeadingWildcards)
//   PerformanceLargeColumnSearch     — filter on Memo or String maxLength>850
//   OrderOnEnumAttribute             — already in OrderbyEditor warnings
//   OrderOnPropertiesFromJoinedTables— already in OrderbyEditor warnings
//   LargeAmountOfAttributes          — >25 columns in $select
//   LargeAmountOfLogicalAttributes   — any logical column in $select
//   FilteringOnCalculatedColumns     — any sourceType column in $filter
//
// Each detector function returns Advisory[] so the caller can concat without
// caring about which file produced which item.

import type { ReactNode } from 'react';
import { findColumn, findTable, isLargeText, isComputedColumn, isLogicalColumn, sourceTypeLabel, type NavProperty } from '../mock/metadata';
import type { ColumnMeta, TableMeta } from '../mock/metadata';
import type { FilterGroup } from '../editors/filter/filterTree';
import { adv, type Advisory } from '../primitives/advisories';

const LEARN_ANTIPATTERNS = 'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/query-antipatterns';
const LEARN_WILDCARDS    = 'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/wildcard-characters';

// Comma-join an array of ReactNodes (e.g. <code>'s) into a single ReactNode.
function joinNodes(nodes: ReactNode[], sep: string = ', '): ReactNode {
  return nodes.reduce<ReactNode[]>((acc, el, i) => i === 0 ? [el] : [...acc, sep, el], []);
}

/**
 * Address composite columns — `address1_city`, `address1_country`,
 * `address1_composite`, `address2_*`, etc. — are marked `IsLogical=true`
 * in the AttributeMetadata because their values live in a joined physical
 * table (`customeraddress`). Per the strict reading of the
 * LargeAmountOfLogicalAttributes antipattern, filtering on them WOULD
 * trip the warning.
 *
 * In practice these are the documented way to query customer addresses on
 * account/contact/lead — the server-side join cost is the expected
 * behavior, not a misuse pattern. Surfacing the warning every time the
 * user filters on `address1_city` adds noise without actionable advice.
 *
 * We exempt the `address[N]_*` family from FILTER-side logical-column
 * advisories only. Filtering by `address1_city` is the standard way to
 * scope customer queries — the join cost is one-shot and acceptable.
 *
 * SELECT-side stays subject to the advisory — picking 8 address sub-fields
 * in $select multiplies the join cost by every returned row, which IS the
 * canonical anti-pattern. If you really do need all the address columns,
 * the advisory tells you what it costs; if you don't, prune the projection.
 */
function isAddressCompositeColumn(c: ColumnMeta): boolean {
  return /^address[1-9]\d*_/.test(c.logicalName);
}

// ── $select detectors ──────────────────────────────────────────────────────

/**
 * Per query-antipatterns "Minimize the number of selected columns".
 * Threshold of 25 is conservative — the doc doesn't name a hard limit but
 * notes "large number of columns" trigger LargeAmountOfAttributes timeouts.
 */
export function detectLargeSelect(select: string[]): Advisory[] {
  if (select.length > 25) {
    return [adv.warn(
      'select-too-many',
      'antipattern',
      `${select.length} columns in $select — performance risk`,
      <span>
        Selecting more than ~25 columns slows the response and can trigger the{' '}
        <code>LargeAmountOfAttributes</code> antipattern. Pick only what the
        client renders.
      </span>,
      'select',
      LEARN_ANTIPATTERNS,
    )];
  }
  return [];
}

/**
 * Per query-antipatterns "Minimize the number of selected logical columns".
 * Logical columns live in a different physical table and force a join.
 */
export function detectLogicalSelect(select: string[], tbl: TableMeta): Advisory[] {
  const logicals = select
    .map(name => findColumn(tbl, name))
    .filter((c): c is ColumnMeta => !!c && isLogicalColumn(c));
  if (logicals.length === 0) return [];
  return [adv.warn(
    'select-logical',
    'antipattern',
    `${logicals.length} logical column${logicals.length === 1 ? '' : 's'} in $select`,
    <span>
      {joinNodes(logicals.map(c => <code key={c.logicalName}>{c.logicalName}</code>))}
      {' '}— values live in separate physical tables; each one forces a join.
      Per the <code>LargeAmountOfLogicalAttributes</code> antipattern.
    </span>,
    'select',
    LEARN_ANTIPATTERNS,
  )];
}

// ── $filter detectors ──────────────────────────────────────────────────────

/**
 * Walk a filter tree resolving every rule/function's column against the
 * correct owner table — handles:
 *   • Bare cols on the root table: `name` → root.name
 *   • Nav-path cols: `primarycontactid/fullname` → contact.fullname
 *   • Lambda-scoped cols: `c/description` inside `contact_customer_accounts/any(c:…)`
 *     → contact.description
 *   • Nested lambdas with inner nav-paths
 *
 * Yields `{ col, op }` per resolvable rule so callers can apply
 * per-antipattern filtering (e.g. large-text only fires on scan ops).
 */
function collectFilterColumns(
  filter: FilterGroup,
  rootTbl: TableMeta,
): Array<{ col: ColumnMeta; op: string }> {
  const out: Array<{ col: ColumnMeta; op: string }> = [];

  const walk = (group: FilterGroup, ownerTable: TableMeta, lambdaAlias?: string): void => {
    for (const node of group.rules) {
      if (node.type === 'rule' || node.type === 'function') {
        if (!node.col) continue;
        // Strip lambda alias prefix
        let path = node.col;
        if (lambdaAlias && path.startsWith(lambdaAlias + '/')) {
          path = path.slice(lambdaAlias.length + 1);
        }
        // Walk nav-path segments
        const segs = path.split('/');
        let cursor: TableMeta | undefined = ownerTable;
        for (let i = 0; i < segs.length - 1; i++) {
          const navMatch: NavProperty | undefined = cursor?.navigationProperties.find((n: NavProperty) => n.name === segs[i]);
          cursor = navMatch ? findTable(navMatch.targetEntity) : undefined;
          if (!cursor) break;
        }
        if (!cursor) continue;
        const leaf = cursor.columns.find(c => c.logicalName === segs[segs.length - 1]);
        if (leaf) out.push({ col: leaf, op: node.op });
      } else if (node.type === 'group') {
        walk(node, ownerTable, lambdaAlias);
      } else if (node.type === 'lambda') {
        const nav = ownerTable.navigationProperties.find((n: NavProperty) => n.name === node.nav);
        const target = nav ? findTable(nav.targetEntity) : undefined;
        if (target) walk(node.inner, target, node.alias);
      }
    }
  };
  walk(filter, rootTbl);
  return out;
}

/**
 * Per query-antipatterns "Avoid using conditions on large text columns".
 * Memo columns or String columns with MaxLength > 850 trigger the
 * PerformanceLargeColumnSearch antipattern — BUT only for SCAN operators
 * (contains/startswith/endswith). Exact-match `eq` on a Memo is fine: it
 * compares the whole value, not a substring scan. We scope to scan ops
 * so the URL-bar advisory matches the inline icon (which is also scan-only).
 */
const STRING_SCAN_OPS = new Set(['contains', 'startswith', 'endswith']);

export function detectLargeTextFilters(filter: FilterGroup, tbl: TableMeta): Advisory[] {
  // Only scan operators trigger the antipattern. `eq 'X'` on a Memo is an
  // exact match — not a scan. The inline icon at the rule level uses the
  // same scoping; we mirror it here so the aggregated drawer entry doesn't
  // fire when no rule actually has the issue.
  const pairs = collectFilterColumns(filter, tbl)
    .filter(({ col, op }) => isLargeText(col) && STRING_SCAN_OPS.has(op));
  if (pairs.length === 0) return [];
  const uniq = Array.from(new Map(pairs.map(p => [p.col.logicalName, p.col])).values());
  return [adv.warn(
    'filter-large-text',
    'antipattern',
    `contains/startswith/endswith on large-text column${uniq.length === 1 ? '' : 's'}`,
    <span>
      {joinNodes(uniq.map(c => <code key={c.logicalName}>{c.logicalName}</code>))}
      {' '}can't be efficiently indexed (Memo, or String &gt; 850 chars). Consider{' '}
      <strong>Dataverse Search</strong> instead.
    </span>,
    'filter',
    LEARN_ANTIPATTERNS,
  )];
}

/**
 * Per query-antipatterns "Minimize the number of selected logical columns" —
 * filtering on a logical column forces a join with the underlying physical
 * table that owns the data (composite address sub-fields, derived
 * status-trace metadata, etc.). Slower than filtering on a non-logical
 * column on the parent table.
 */
export function detectLogicalFilters(filter: FilterGroup, tbl: TableMeta): Advisory[] {
  const cols = collectFilterColumns(filter, tbl)
    .map(p => p.col)
    .filter(isLogicalColumn)
    // Exempt the address[N]_* composite family — the join is the
    // intended behavior for those, not a misuse pattern.
    .filter(c => !isAddressCompositeColumn(c));
  if (cols.length === 0) return [];
  const uniq = Array.from(new Map(cols.map(c => [c.logicalName, c])).values());
  return [adv.warn(
    'filter-logical',
    'antipattern',
    `Filter on logical column${uniq.length === 1 ? '' : 's'}`,
    <span>
      {joinNodes(uniq.map(c => <code key={c.logicalName}>{c.logicalName}</code>))}
      {' '}— values live in a joined physical table. Filtering forces a
      server-side join before the predicate; slower than filtering on
      sub-fields directly (e.g. <code>address1_city</code> instead of
      the composite).
    </span>,
    'filter',
    LEARN_ANTIPATTERNS,
  )];
}

/**
 * Per query-antipatterns "Avoid using formula or calculated columns in filter
 * conditions". Any sourceType > 0 column trips FilteringOnCalculatedColumns.
 */
export function detectComputedColumnFilters(filter: FilterGroup, tbl: TableMeta): Advisory[] {
  const cols = collectFilterColumns(filter, tbl)
    .map(p => p.col)
    .filter(isComputedColumn);
  if (cols.length === 0) return [];
  const uniq = Array.from(new Map(cols.map(c => [c.logicalName, c])).values());
  return [adv.warn(
    'filter-computed',
    'antipattern',
    `Filter on ${uniq.length === 1 ? `${sourceTypeLabel(uniq[0].sourceType)} column` : 'computed columns'}`,
    <span>
      {joinNodes(uniq.map(c => (
        <code key={c.logicalName}>
          {c.logicalName} ({sourceTypeLabel(c.sourceType)})
        </code>
      )))}
      {' '}— Dataverse computes these at retrieval time. Filtering on them is
      throttled (<code>FilteringOnCalculatedColumns</code>).
    </span>,
    'filter',
    LEARN_ANTIPATTERNS,
  )];
}

/**
 * Wildcard advisory — fed by stripLeadingWildcards on each contains/starts/ends
 * rule. The filter editor accumulates a list of stripped values and passes them
 * here so the user sees one consolidated advisory rather than N inline warnings.
 */
export interface StrippedWildcardEntry {
  col: string;
  raw: string;
  cleaned: string;
  kinds: string[];
}

export function buildWildcardAdvisory(entries: StrippedWildcardEntry[]): Advisory[] {
  if (entries.length === 0) return [];
  return [adv.info(
    'wildcard-stripped',
    'wildcard',
    `Stripped leading wildcard${entries.length === 1 ? '' : 's'} from ${entries.length} filter${entries.length === 1 ? '' : 's'}`,
    <span>
      Dataverse rejects leading wildcards. The following inputs were rewritten:
      <ul style={{ margin: '6px 0 0 0', paddingLeft: 18 }}>
        {entries.map((r, i) => (
          <li key={i} style={{ marginBottom: 2 }}>
            <code>{r.col}</code>: <code>{r.raw}</code> → <code>{r.cleaned}</code>
            <span style={{ marginLeft: 4, opacity: 0.7 }}>
              ({r.kinds.join(', ')})
            </span>
          </li>
        ))}
      </ul>
      <span style={{ display: 'block', marginTop: 6 }}>
        To search for a literal <code>%</code>, enclose it in brackets: <code>[%]</code>.
      </span>
    </span>,
    'filter',
    LEARN_WILDCARDS,
  )];
}

// ── Per-column rule advisory ──────────────────────────────────────────────
//
// Lightweight per-rule check used by the FilterEditor's inline icon. The
// drawer aggregation above is the canonical "everything DRS wants you to
// know" view; this is the discoverability signal at the source — so users
// notice WHICH rule the warning is about without scrolling through the
// drawer first.

export interface ColumnAntipattern {
  kind: 'calculated' | 'logical' | 'largeText';
  /** Title-cased short headline for popover. */
  title: string;
  /** One-sentence body explaining the impact. */
  body: string;
  /** MS Learn deep link for the user to verify. */
  learnMoreUrl: string;
}

/**
 * Detect every antipattern that applies to a single column when used in a
 * filter rule. `op` lets us scope the large-text check to string fns only
 * (since `eq` on a long-text column is not the same anti-pattern as
 * `contains()` on it).
 */
export function detectColumnAntipatterns(
  c: ColumnMeta,
  op?: string,
): ColumnAntipattern[] {
  const out: ColumnAntipattern[] = [];

  if (isComputedColumn(c)) {
    const label = sourceTypeLabel(c.sourceType) ?? 'computed';
    out.push({
      kind: 'calculated',
      title: `Filtering on a ${label} column`,
      body: `${c.displayName} is a ${label} column. Dataverse re-evaluates the formula for every row before applying the predicate — throttled via the FilteringOnCalculatedColumns antipattern.`,
      learnMoreUrl: LEARN_ANTIPATTERNS,
    });
  }

  // Exempt address[N]_* composite family — see isAddressCompositeColumn
  // rationale at the top of this file. Filtering on `address1_city` is
  // the documented way to query customer addresses on account/contact;
  // the join cost is expected, not a misuse pattern.
  if (isLogicalColumn(c) && !isAddressCompositeColumn(c)) {
    out.push({
      kind: 'logical',
      title: 'Filtering on a logical column',
      body: `${c.displayName} lives in a different physical table (composite address, derived metadata, etc.). Filtering forces a server-side join before the predicate.`,
      learnMoreUrl: LEARN_ANTIPATTERNS,
    });
  }

  // Large-text check only applies when the OPERATOR is a string scan.
  // `eq 'x'` on a Memo is exact-match — no anti-pattern. `contains('x')`
  // on the same Memo IS — it scans every row.
  const stringFnOps = new Set(['contains', 'startswith', 'endswith']);
  if (isLargeText(c) && op && stringFnOps.has(op)) {
    out.push({
      kind: 'largeText',
      title: `${op}() on a large-text column`,
      body: `${c.displayName} is ${c.attributeType === 'Memo' ? 'a Memo (multi-line text)' : 'a long String (MaxLength > 850)'} column. ${op}() can't use an index — scans every row. Consider Dataverse Search for full-text needs.`,
      learnMoreUrl: LEARN_ANTIPATTERNS,
    });
  }

  return out;
}

// ── Top-level composer for Retrieve Multiple ──────────────────────────────

export interface RetrieveMultipleAntipatternInput {
  table: string;
  select: string[];
  filter: FilterGroup;
  strippedWildcards?: StrippedWildcardEntry[];
}

/**
 * Composer — invoked once per render from RetrieveMultipleMode. Returns the
 * Advisory[] that gets passed to <UrlBar advisories=...>. Order-of-evaluation
 * matters here: blockers should always sort first inside the drawer, which
 * AdvisoryDrawer does via bucketAdvisories().
 */
export function detectRetrieveMultipleAntipatterns(input: RetrieveMultipleAntipatternInput): Advisory[] {
  const tbl = findTable(input.table);
  if (!tbl) return [];
  return [
    ...detectLargeSelect(input.select),
    ...detectLogicalSelect(input.select, tbl),
    ...detectLargeTextFilters(input.filter, tbl),
    ...detectLogicalFilters(input.filter, tbl),
    ...detectComputedColumnFilters(input.filter, tbl),
    ...buildWildcardAdvisory(input.strippedWildcards ?? []),
  ];
}
