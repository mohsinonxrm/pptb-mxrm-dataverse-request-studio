// validateRequest — final-pass validation of the filter / expand / orderby
// / apply tree, producing Advisory[] for the AdvisoryDrawer + Execute
// button gating.
//
// Why this exists (Gap C from the audit):
//   The UI gates prevent users from CONSTRUCTING most forbidden patterns
//   (lambda nesting under N:N, `not` on a Dataverse function, groupby on
//   DateTime, etc.). But state can also arrive from outside the UI:
//   - Loaded from a saved request (the persisted shape might predate a gate)
//   - Pasted from a URL
//   - Edited by hand
//   - Carried over across entity changes (we reset most things on
//     onTableChange, but bugs happen)
//
// This module walks the request shape once on each render and emits
// Advisory entries for anything that doesn't match the official docs.
// Errors block execute (surfaced via `disabledReasonFromAdvisories`);
// warnings + infos appear in the drawer but don't block.
//
// What this is NOT:
//   - A pre-flight count-query to predict the 50,000-row aggregate cap
//     (the docs and the user explicitly said we let Dataverse return
//     the error instead).
//   - A replacement for the per-clause inline banners — those stay.
//     This is the consolidated "execute-time" view of the same rules.
//
// Doc anchors:
//   /webapi/query/overview      — unsupported $skip/$search/$format/etc.
//   /webapi/query/filter-rows   — operators, 500-cond cap, NOT rules
//   /webapi/query/join-tables   — 15-expand cap, nested rules
//   /webapi/query/aggregate-data— $apply restrictions

import type { TableMeta } from '../mock/metadata';
import { findTable } from '../mock/metadata';
import type { FilterGroup, FilterNode } from '../editors/filter/filterTree';
import { countRules } from '../editors/filter/filterTree';
import { findOperator } from '../editors/filter/operators';
import type { ExpandSpec } from '../editors/ExpandEditor';
import type { OrderbySpec } from '../editors/OrderbyEditor';
import type { ApplySpec } from '../editors/ApplyEditor';
import {
  totalExpandCount, MAX_EXPANDS_PER_QUERY,
} from '../editors/expand/expandTree';
import { adv, type Advisory } from '../primitives/advisories';

export interface ValidateRequestInput {
  /** Root entity logical name. */
  table?: string;
  /** Top-level $filter tree (when $apply is OFF). */
  filter?: FilterGroup;
  /** $orderby spec. */
  orderby?: OrderbySpec[];
  /** $expand tree. */
  expand?: ExpandSpec[];
  /** $apply (aggregate) spec — when `enabled`, takes precedence over $filter. */
  apply?: ApplySpec;
}

/**
 * Run every spec-grounded check against the request and return an
 * `Advisory[]`. Modes append this to their existing antipattern + wildcard
 * advisories before passing the combined list to `<AdvisoryDrawer>`.
 */
export function validateRequest(input: ValidateRequestInput): Advisory[] {
  const out: Advisory[] = [];
  if (!input.table) return out;
  const rootTable = findTable(input.table);
  if (!rootTable) return out; // metadata not loaded yet

  // ── Filter validation ────────────────────────────────────────
  // When $apply is enabled, the active filter is apply.prefilter;
  // otherwise it's the top-level $filter. Either way we walk the same
  // grammar.
  const activeFilter = input.apply?.enabled ? input.apply.prefilter : input.filter;
  if (activeFilter) {
    validateFilterGroup(activeFilter, rootTable, undefined, out);
    // 500-condition hard cap. Per /webapi/query/filter-rows: "You can
    // include up to 500 total conditions in a query."
    const ruleCount = countRules(activeFilter);
    if (ruleCount > 500) {
      out.push(adv.err(
        'filter-cond-cap', 'validation',
        `Filter has ${ruleCount} conditions, exceeds the 500 hard cap.`,
        'Dataverse rejects queries with more than 500 total filter conditions. Compress OR-equality groups with `In(...)` to bring the count down.',
        'filter',
        'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query/filter-rows#condition-limits',
      ));
    }
  }

  // ── Expand validation ────────────────────────────────────────
  if (input.expand && input.expand.length > 0) {
    const total = totalExpandCount(input.expand);
    if (total > MAX_EXPANDS_PER_QUERY) {
      // Per docs (15-expand cap) — kept as advisory; we did not empirically
      // verify because the 16-nav probe (test 13.9) was contaminated by a
      // non-existent nav. Treat as a soft warning rather than a hard
      // blocker so users can execute and see what Dataverse actually does.
      out.push(adv.warn(
        'expand-cap', 'validation',
        `${total} expand options exceeds the documented ${MAX_EXPANDS_PER_QUERY} cap.`,
        `Microsoft documents a ${MAX_EXPANDS_PER_QUERY}-expand limit per query. Execute to see if your org actually enforces it.`,
        'expand',
        'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query/join-tables',
      ));
    }
    // Root $top + collection $expand was previously flagged as a conflict
    // but test 8.13 shows it works fine. Inner $top/$orderby alongside
    // nested expand is similarly accepted in practice (tests 8.3-8.5,
    // 9.7). The blanket "no $top/$orderby when nested" advisory was
    // overly aggressive — drop it.
    // Flat collection $expand info banner — kept because the 5000-row
    // implicit cap genuinely catches people out.
    const hasFlatCollection = input.expand.some(it => {
      const nav = rootTable.navigationProperties.find(n => n.name === it.nav);
      return nav && nav.cardinality !== 'ManyToOne' && (!it.nestedExpand || it.nestedExpand.length === 0);
    });
    if (hasFlatCollection) {
      out.push(adv.info(
        'expand-flat-collection', 'validation',
        'Flat collection expand can return up to 5,000 rows per parent.',
        <>Collections without a nested <code>$expand</code> aren't paginated. Add a nested <code>$expand</code> or use <code>Prefer: odata.maxpagesize</code> to control the size.</>,
        'expand',
        'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query/join-tables#single-expand-on-collection-valued-navigation-properties',
      ));
    }
    // Depth advisory dropped — empirical test 13.17 confirms depth 4
    // works fine. Microsoft's docs don't actually publish a depth cap
    // for $expand. Let users explore.
  }

  // ── Apply validation ────────────────────────────────────────
  if (input.apply?.enabled) {
    const apply = input.apply;

    // $orderby is rejected on $apply output.
    if (input.orderby && input.orderby.length > 0) {
      out.push(adv.err(
        'apply-orderby-conflict', 'validation',
        '`$orderby` on aggregate output is not supported.',
        <>Dataverse rejects <code>$orderby</code> on the alias columns produced by <code>$apply</code> with <em>"The query node SingleValueOpenPropertyAccess is not supported."</em> Remove the <code>$orderby</code> entries, or disable <code>$apply</code>.</>,
        'orderby',
        'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query/aggregate-data',
      ));
    }

    // groupby on DateTime is unsupported (test 10.8 → 0x80060888).
    // Also: groupby on a bare lookup or `_<attr>_value` is rejected
    // (tests 10.6 / 10.6b). For both, the picker prevents construction;
    // these checks catch state loaded from outside the UI.
    for (const g of apply.groupby) {
      const segs = g.split('/');
      const leafSeg = segs[segs.length - 1];
      if (segs.length === 1) {
        const col = rootTable.columns.find(c => c.logicalName === leafSeg);
        if (col?.attributeType === 'DateTime') {
          out.push(adv.err(
            `apply-groupby-datetime-${g}`, 'validation',
            `\`groupby\` on DateTime column \`${g}\` is not supported.`,
            'Dataverse does not support grouping by date/time values via OData $apply.',
            'apply',
            'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query/aggregate-data',
          ));
        }
        if (col?.attributeType === 'Lookup' || col?.attributeType === 'Customer' || col?.attributeType === 'Owner') {
          out.push(adv.err(
            `apply-groupby-lookup-${g}`, 'validation',
            `\`groupby\` on bare lookup \`${g}\` is rejected.`,
            <>Dataverse rejects both <code>{g}</code> and <code>_{g}_value</code> as group-by columns. Drill into the lookup and pick a scalar (e.g. <code>{g}/&lt;related-attr&gt;</code>).</>,
            'apply',
            'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query/aggregate-data',
          ));
        }
        if (leafSeg.startsWith('_') && leafSeg.endsWith('_value')) {
          out.push(adv.err(
            `apply-groupby-value-${g}`, 'validation',
            `\`groupby\` on \`${g}\` (_value form) is rejected.`,
            'Use the nav-path form (e.g. `primarycontactid/fullname`) — the `_<attr>_value` form is rejected by Dataverse for groupby.',
            'apply',
          ));
        }
      }
    }

    // Aggregate alias collisions — Dataverse rejects two aliases with
    // the same name in the same query.
    const seenAliases = new Set<string>();
    for (const a of apply.aggregates) {
      if (!a.alias.trim()) {
        out.push(adv.err(
          `apply-agg-empty-alias-${Math.random().toString(36).slice(2)}`, 'validation',
          'Aggregate is missing an alias.',
          'Every aggregate needs a non-empty alias for the output column.',
          'apply',
        ));
        continue;
      }
      if (seenAliases.has(a.alias)) {
        out.push(adv.err(
          `apply-agg-dup-alias-${a.alias}`, 'validation',
          `Duplicate aggregate alias \`${a.alias}\`.`,
          'Two aggregates can\'t share the same alias — each `as <alias>` must be unique within the query.',
          'apply',
        ));
      }
      seenAliases.add(a.alias);
      // Non-$count aggregates need a source column.
      if (a.fn !== '$count' && !a.col.trim()) {
        out.push(adv.err(
          `apply-agg-no-col-${a.alias}`, 'validation',
          `\`${a.fn}\` aggregate \`${a.alias}\` has no source column.`,
          'Scalar aggregates (sum / average / min / max) need a source column. `$count` is the only column-less aggregate.',
          'apply',
        ));
      }
    }
  }

  return out;
}

// ── Internals ──────────────────────────────────────────────────────

function validateFilterGroup(
  group: FilterGroup,
  parentTable: TableMeta,
  lambdaAlias: string | undefined,
  out: Advisory[],
): void {
  // Group-level NOT is empirically supported (tests G.1–G.7) — EXCEPT when
  // the negated group contains a Dataverse query function anywhere in its
  // descendants. Dataverse rejects `not (<…dv-fn…>)` with 405 (test G.9)
  // even when wrapped in parens. Surface as an error: there's no way to
  // make Dataverse accept this; the user must restructure or use the
  // explicit `Not*` sibling function instead.
  if (group.negated && groupHasDvFunctionDescendant(group)) {
    out.push(adv.err(
      `group-not-with-dvfn-${group.id}`, 'validation',
      'Negated group contains a Dataverse function — will be rejected.',
      <>
        Dataverse rejects <code>not (…Microsoft.Dynamics.CRM.X(…)…)</code> with HTTP&nbsp;405 even when wrapped in parens (test G.9). Move the function out of the negated group, or use the explicit <code>Not*</code> sibling (e.g. <code>NotIn</code>, <code>NotBetween</code>, <code>NotEqualUserId</code>, <code>DoesNotContainValues</code>) and drop the group-NOT.
      </>,
      'filter',
    ));
  }
  for (const node of group.rules) {
    validateFilterNode(node, parentTable, lambdaAlias, out);
  }
}

function groupHasDvFunctionDescendant(g: FilterGroup): boolean {
  for (const node of g.rules) {
    if (node.type === 'function') return true;
    if (node.type === 'group'   && groupHasDvFunctionDescendant(node)) return true;
    if (node.type === 'lambda'  && groupHasDvFunctionDescendant(node.inner)) return true;
  }
  return false;
}

function validateFilterNode(
  node: FilterNode,
  parentTable: TableMeta,
  lambdaAlias: string | undefined,
  out: Advisory[],
): void {
  if (node.type === 'group') {
    validateFilterGroup(node, parentTable, lambdaAlias, out);
    return;
  }
  if (node.type === 'rule') {
    // `not` on plain comparisons / null-checks / string-fns was previously
    // gated to odata-fn only — empirical tests 3.3-3.7 prove Dataverse
    // accepts `not <any-bool-expr>` (e.g. `not (revenue gt 100)`). Drop
    // that advisory. The encoder still emits `not` correctly per-kind.
    return;
  }
  if (node.type === 'function') {
    // `not` on Dataverse functions is universally rejected (test 5.7 → 405).
    if (node.negated) {
      out.push(adv.err(
        `fn-negated-${node.op}-${node.id}`, 'validation',
        `\`not\` on \`Microsoft.Dynamics.CRM.${node.op}\` is rejected by Dataverse.`,
        'Use the explicit `Not*` sibling (NotIn / NotBetween / NotEqualUserId / DoesNotContainValues / …) from the function picker instead.',
        'filter',
      ));
    }
    // DV functions inside a lambda are FINE — empirical tests 6.5b and
    // 13.15 (`contact_customer_accounts/any(c: EqualUserId(PropertyName='ownerid'))`)
    // succeed. The encoder strips the alias prefix automatically.
    // The earlier "undocumented inside lambda" warning was based on docs
    // silence, not actual behavior — dropped.
    //
    // Nav-path inside `PropertyName` (e.g. `primarycontactid/ownerid`)
    // is rejected (tests 7.7 / 7.8). The column picker doesn't surface
    // nav-paths for function nodes, but if state arrives with one
    // (legacy / paste / hand-edit) we should warn.
    if (node.col && node.col.includes('/') && !node.col.startsWith((lambdaAlias ?? '') + '/')) {
      // Strip any pure alias prefix; what's left must be a nav-path.
      const stripped = lambdaAlias && node.col.startsWith(lambdaAlias + '/')
        ? node.col.slice(lambdaAlias.length + 1) : node.col;
      if (stripped.includes('/')) {
        out.push(adv.err(
          `fn-nav-path-${node.op}-${node.id}`, 'validation',
          `Nav-path \`${stripped}\` in \`PropertyName\` is rejected.`,
          'Dataverse query functions don\'t support drilling through related tables in `PropertyName`. Pick a column on the current entity, or remove the function and use a direct comparison.',
          'filter',
        ));
      }
    }
    // Hierarchy fns require the entity to be configured for hierarchy
    // (test 13.11 → 0x80047020 on incident). The encoder swaps to the
    // entity's PK; if the entity isn't hierarchical Dataverse will still
    // reject. Soft hint, not a blocker.
    const op = findOperator(node.op);
    if (op?.category === 'hierarchy') {
      out.push(adv.info(
        `fn-hierarchy-${node.op}-${node.id}`, 'validation',
        `\`${node.op}\` needs a hierarchical relationship on \`${parentTable.logicalName}\`.`,
        <>Hierarchy functions only work on entities configured with a hierarchical relationship (e.g. <code>account.parentaccountid</code>). If your entity isn\'t configured, execute will fail with <code>0x80047020 "doesn't have a hierarchical relationship"</code>.</>,
        'filter',
      ));
    }
    return;
  }
  if (node.type === 'lambda') {
    // The nav must exist on parentTable and be collection-valued.
    const nav = parentTable.navigationProperties.find(n => n.name === node.nav);
    if (!nav) {
      out.push(adv.err(
        `lambda-unknown-nav-${node.id}`, 'validation',
        `Lambda navigates through unknown property \`${node.nav}\`.`,
        `\`${node.nav}\` isn't a navigation property on \`${parentTable.logicalName}\`. The relationship may have been removed or renamed since this request was saved.`,
        'filter',
      ));
      return;
    }
    if (nav.cardinality === 'ManyToOne') {
      out.push(adv.err(
        `lambda-on-n1-${node.id}`, 'validation',
        `Lambda \`${node.nav}\` is a single-valued (N:1) navigation property.`,
        'Lambda operators `any`/`all` only apply to collection-valued (1:N or N:N) navigation properties. Use a direct column comparison or `<nav>/<col>` path filter instead.',
        'filter',
      ));
      return;
    }
    if (!node.alias || !node.alias.trim()) {
      out.push(adv.err(
        `lambda-no-alias-${node.id}`, 'validation',
        'Lambda is missing an alias.',
        'Every `any`/`all` lambda needs a non-empty alias to scope its inner predicate (e.g. `c` in `any(c: c/email ne null)`).',
        'filter',
      ));
      return;
    }
    const innerTable = findTable(nav.targetEntity);
    if (!innerTable) {
      // Target metadata not loaded yet — skip deep validation.
      return;
    }
    validateFilterGroup(node.inner, innerTable, node.alias, out);
    return;
  }
}

// walkExpandTree was used by the $top/$orderby-when-nested advisory which
// has been dropped (empirical tests 8.3-8.5, 8.13, 9.7 show those work).
// Kept as a comment marker in case future advisories need a tree walker.
