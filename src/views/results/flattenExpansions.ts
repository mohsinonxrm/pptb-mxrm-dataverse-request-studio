// Flatten nested `$expand` results into a tabular shape.
//
// Two transformations happen in one pass:
//
//   1. SINGLE-VALUED expands (N:1, e.g. `primarycontactid`) are dotted in
//      place — the inner object's keys are prefixed with `<navName>.` and
//      merged into the parent row.
//
//      Input  : { name: 'Acme', primarycontactid: { fullname: 'Jane' } }
//      Output : { name: 'Acme', 'primarycontactid.fullname': 'Jane' }
//
//   2. COLLECTION-VALUED expands (1:N / N:N, e.g. `contact_customer_accounts`)
//      are EXPLODED — the parent row is duplicated once per child, with the
//      child's fields dot-prefixed by the nav name. This mirrors a SQL
//      LEFT JOIN and matches how Model-Driven advanced-find displays
//      link-entity results.
//
//      Input  : { name: 'Acme', contact_customer_accounts: [
//                  { fullname: 'Jane' },
//                  { fullname: 'John' },
//                ] }
//      Output : [
//                { name: 'Acme', 'contact_customer_accounts.fullname': 'Jane' },
//                { name: 'Acme', 'contact_customer_accounts.fullname': 'John' },
//              ]
//
//      Empty collection → parent row emitted once with no child columns
//      (left-join semantics; the missing child columns render as em-dash).
//
//      Multiple sibling collections on the same parent → Cartesian product.
//      A row with both `contacts: [c1, c2]` and `tasks: [t1, t2, t3]`
//      emits 2 × 3 = 6 rows. The CartesianAdvisory hook surfaces a warning
//      so users aren't surprised. (Caller is responsible for the advisory.)
//
//   3. Annotation keys (`<col>@<…>`) are kept attached to their owning
//      column under the same dotted prefix so getFormattedValue still
//      resolves at the dotted key.
//
//   4. `<col>@odata.nextLink` on collections is preserved as a sibling
//      key on each exploded row so the cell renderer can show a "more
//      available" marker.
//
// Each output row gets a synthetic `__rowKey` containing a stable string
// derived from the indices it was assembled from. ResultsGrid uses this
// for selection / virtualization keying since the parent's primaryKey is
// no longer row-unique after multiplication.

export const ROW_KEY = '__rowKey';

const isPlainObject = (v: unknown): v is Record<string, unknown> =>
  v != null && typeof v === 'object' && !Array.isArray(v);

const isObjectArray = (v: unknown): v is Array<Record<string, unknown>> =>
  Array.isArray(v) && v.every((x) => x != null && typeof x === 'object' && !Array.isArray(x));

/** Lift every key in `row` under a `<prefix>.` namespace. Annotation
 *  suffix `@…` is preserved on the column it belongs to. */
function prefixKeys(row: Record<string, unknown>, prefix: string): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(row)) {
    const at = k.indexOf('@');
    const flatKey = at >= 0 ? `${prefix}.${k.slice(0, at)}${k.slice(at)}` : `${prefix}.${k}`;
    out[flatKey] = v;
  }
  return out;
}

/** Dotted-flatten a single-valued (N:1) expand into `out` in place. */
function flattenSingleValuedInto(
  row: Record<string, unknown>,
  prefix: string,
  out: Record<string, unknown>,
): void {
  for (const [k, v] of Object.entries(row)) {
    const at = k.indexOf('@');
    const flatKey = at >= 0 ? `${prefix}.${k.slice(0, at)}${k.slice(at)}` : `${prefix}.${k}`;
    if (isPlainObject(v)) {
      // Further nested N:1 (depth ≥ 2) — dot-flatten again.
      flattenSingleValuedInto(v, flatKey, out);
    } else if (isObjectArray(v)) {
      // A nested 1:N inside an N:1 (rare but possible if the user expands
      // both levels). We don't try to fan it out here — flattening from
      // within a single-valued path would change row counts mid-way and
      // make the column grid harder to reason about. Keep the array on
      // the row; the CollectionCellRenderer renders a count badge.
      out[flatKey] = v;
    } else {
      out[flatKey] = v;
    }
  }
}

/**
 * Flatten one server row into one-or-more grid rows. Single-valued expands
 * fold into a single base; collection-valued expands explode into multiple
 * outputs (Cartesian when ≥ 2 siblings).
 */
export function flattenRow(row: Record<string, unknown>): Array<Record<string, unknown>> {
  const base: Record<string, unknown> = {};
  const collections: Array<{
    prefix: string;
    /** Each item already recursively flattened (so a child's own 1:N
     *  fans out here too) and key-prefixed by the parent nav name. */
    rows: Array<Record<string, unknown>>;
    nextLink?: unknown;
  }> = [];

  for (const [key, value] of Object.entries(row)) {
    // Annotation key — keep on base. `@odata.nextLink` on a collection is
    // routed separately below.
    if (key.includes('@')) {
      // Skip the nextLink markers — we attach those per-collection.
      if (key.endsWith('@odata.nextLink')) continue;
      base[key] = value;
      continue;
    }

    if (isObjectArray(value)) {
      // 1:N / N:N expand — recursively flatten each item (so nested
      // collections also fan out), then prefix every key by the nav name.
      const itemRows = value.flatMap((it) => flattenRow(it));
      const prefixed = itemRows.map((itr) => prefixKeys(itr, key));
      collections.push({
        prefix: key,
        rows: prefixed,
        nextLink: row[`${key}@odata.nextLink`],
      });
      continue;
    }

    if (isPlainObject(value)) {
      // N:1 expand — dot-flatten into base.
      flattenSingleValuedInto(value, key, base);
      continue;
    }

    base[key] = value;
  }

  if (collections.length === 0) {
    // No collection expands — single row.
    return [{ ...base, [ROW_KEY]: 'r' }];
  }

  // Cartesian product across all sibling collections × base.
  // Each step expands the current combined set by each collection's items.
  // Empty collections preserve the row (left-join) — we add the @nextLink
  // marker but no child fields.
  let combined: Array<Record<string, unknown>> = [base];
  for (const c of collections) {
    if (c.rows.length === 0) {
      // Preserve parent rows. Attach the nextLink marker if present, in case
      // Dataverse paginated and we got "no items returned yet, but more
      // exist" — the cell renderer can still show that.
      if (c.nextLink != null) {
        combined = combined.map((b) => ({ ...b, [`${c.prefix}@odata.nextLink`]: c.nextLink }));
      }
      continue;
    }
    combined = combined.flatMap((b) =>
      c.rows.map((child) => ({
        ...b,
        ...child,
        // Sibling nextLink marker for the child cell renderer.
        ...(c.nextLink != null ? { [`${c.prefix}@odata.nextLink`]: c.nextLink } : {}),
      })),
    );
  }

  // Stamp each emitted row with a synthetic __rowKey so the grid can
  // uniquely identify multiplied rows (parent PK alone is not unique).
  return combined.map((r, i) => ({ ...r, [ROW_KEY]: `r${i}` }));
}

/** Flatten every server row, concatenating the multi-row outputs. */
export function flattenRows(rows: Array<Record<string, unknown>>): Array<Record<string, unknown>> {
  // Re-stamp __rowKey globally so it's unique across the whole grid, not
  // just within one parent's expansion.
  const out: Array<Record<string, unknown>> = [];
  let counter = 0;
  for (const r of rows) {
    for (const flat of flattenRow(r)) {
      out.push({ ...flat, [ROW_KEY]: `r${counter++}` });
    }
  }
  return out;
}

/**
 * True if any sibling collection-valued expand appears on the row.
 * The mode can use this to surface an advisory (e.g. "this query
 * returned 12 rows from 2 parents because of a sibling collection expand").
 */
export function hasSiblingCollectionExpands(row: Record<string, unknown>): boolean {
  let count = 0;
  for (const [k, v] of Object.entries(row)) {
    if (k.includes('@')) continue;
    if (isObjectArray(v)) {
      count++;
      if (count >= 2) return true;
    }
  }
  return false;
}
