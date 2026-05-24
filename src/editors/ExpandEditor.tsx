// $expand data model + OData encoder.
//
// The interactive UI lives in `expand/ExpandOverview.tsx` (overview pane for
// root + each nested level) and in the path-based right-pane router inside
// the mode component (which reuses SelectEditor/FilterEditor/OrderbyEditor/
// TopEditor scoped to the active expand's target entity).
//
// Nested $expand is supported recursively per the docs:
//   ?$expand=primarycontactid($select=fullname;$expand=createdby($select=fullname))

import { findTable } from '../mock/metadata';
import { groupToOData, type FilterGroup, newId } from './filter/filterTree';
import { orderbyToOData, type OrderbySpec } from './OrderbyEditor';
import { attrRefByName } from '../engine/odataAttr';

export interface ExpandSpec {
  /**
   * Stable id for tree-path navigation (e.g. `expand/<id>/select`). Generated
   * via `newId('e')` so it never collides with FilterEditor ids.
   */
  id: string;
  /** Navigation property name (e.g. primarycontactid, contact_customer_accounts) */
  nav: string;
  /** Inner $select column logical names */
  select: string[];
  /** Inner $top (collection-valued only) */
  top?: number | null;
  /** Inner $orderby (collection-valued only) — multi-column same as root */
  orderby: OrderbySpec[];
  /** Inner $filter (collection-valued only) — a full FilterGroup tree */
  filter?: FilterGroup;
  /** Nested $expand (recursive) — Dataverse supports nested expansion per spec §10 */
  nestedExpand?: ExpandSpec[];
}

/** Factory — used when adding a new expand at any level. */
export function makeExpandSpec(navName: string): ExpandSpec {
  return {
    id: newId('e'),
    nav: navName,
    select: [],
    top: null,
    orderby: [],
    nestedExpand: [],
  };
}

/** Encode the items as the value of `$expand=`. Supports nested $expand recursively per spec §10. */
export function expandToOData(items: ExpandSpec[], parentTableLogical: string): string {
  return items.map(it => emitOne(it, parentTableLogical)).join(',');
}

function emitOne(it: ExpandSpec, parentTableLogical: string): string {
  const inner: string[] = [];
  const parentTbl = findTable(parentTableLogical);
  const nav = parentTbl?.navigationProperties.find(n => n.name === it.nav);
  // Resolve the target entity so we can encode the inner $select's
  // lookup columns as `_<logical>_value` instead of the nav name.
  const targetTbl = nav ? findTable(nav.targetEntity) : undefined;

  if (it.select.length) {
    const encoded = it.select.map(c => attrRefByName(targetTbl, c)).join(',');
    inner.push(`$select=${encoded}`);
  }
  if (it.filter && it.filter.rules.length > 0 && nav) {
    const targetTbl = findTable(nav.targetEntity);
    if (targetTbl) {
      const expr = groupToOData(it.filter, targetTbl);
      if (expr) inner.push(`$filter=${expr}`);
    }
  }
  if (it.orderby.length) inner.push(`$orderby=${orderbyToOData(it.orderby)}`);
  if (it.top != null && it.top > 0) inner.push(`$top=${it.top}`);
  if (it.nestedExpand && it.nestedExpand.length > 0 && nav) {
    const child = it.nestedExpand
      .map(c => emitOne(c, nav.targetEntity))
      .join(',');
    if (child) inner.push(`$expand=${child}`);
  }
  return inner.length ? `${it.nav}(${inner.join(';')})` : it.nav;
}
