// Change tracking in Retrieve Multiple (#233).
//
// The first request sends `Prefer: odata.track-changes`; Dataverse answers
// with the rows and an `@odata.deltaLink`. Following the delta link returns
// the rows created or updated since, plus one `$deletedEntity` entry (just
// `id` and `reason`) per row deleted since, and the next delta link.
//
// Dataverse rejects `$filter`, `$orderby`, `$expand` and `$top` next to
// `odata.track-changes`; `$select` is fine. A request that has them is
// blocked with an advisory per option rather than sent, and nothing is
// dropped behind the user's back.
//
// https://learn.microsoft.com/en-us/power-apps/developer/data-platform/use-change-tracking-synchronize-data-external-systems

import type { TableMeta } from '../metadata/metadata';
import { adv, type Advisory } from '../primitives/advisories';

/** Change tracking docs: the query options it rules out, and the delta link. */
export const CHANGE_TRACKING_DOCS =
  'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/use-change-tracking-synchronize-data-external-systems';

/** Whether the request's table can be tracked (#233), and why not. */
export interface ChangeTrackingAvailability {
  available: boolean;
  /** Why the option is off, when it isn't available. */
  reason?: string;
}

/** The query options Dataverse rejects with `odata.track-changes`, and the node that sets each. */
export const CHANGE_TRACKING_UNSUPPORTED: readonly { key: string; focusNode: string }[] = [
  { key: '$filter', focusNode: 'filter' },
  { key: '$orderby', focusNode: 'orderby' },
  { key: '$expand', focusNode: 'expand' },
  { key: '$top', focusNode: 'top' },
];

/** Why Track changes can't be turned on for this request, if it can't (#233). */
export function changeTrackingAvailability(opts: {
  table: string;
  tableMeta: TableMeta | undefined;
  countOnly?: boolean;
}): ChangeTrackingAvailability {
  if (!opts.table) return { available: false, reason: 'Pick a target table first.' };
  if (opts.countOnly) {
    return {
      available: false,
      reason: "Count only doesn't track changes. Turn Count only off in $count to use it.",
    };
  }
  if (!opts.tableMeta) {
    return { available: false, reason: "Loading the table's definition…" };
  }
  if (opts.tableMeta.changeTrackingEnabled !== true) {
    return {
      available: false,
      reason:
        "Change tracking isn't enabled for this table. Turn on Track changes in the table's " +
        'properties in Power Apps (Advanced options) to use it.',
    };
  }
  return { available: true };
}

/**
 * The advisories for a request with Track changes on: an error when the
 * table isn't tracked, and one per query option Dataverse rejects with it.
 * `queryParts` is what the URL sends (so `$top` left out for a page size, or
 * `$filter` under `$apply`, doesn't count). Nothing when Track changes is off,
 * or with Count only, which doesn't send it.
 */
export function changeTrackingAdvisories(opts: {
  trackChanges: boolean | undefined;
  countOnly?: boolean;
  tableMeta: TableMeta | undefined;
  queryParts: readonly { key: string }[];
}): Advisory[] {
  if (!opts.trackChanges || opts.countOnly) return [];
  const out: Advisory[] = [];
  // Only once the table has loaded: an unknown table isn't a reason to block.
  if (opts.tableMeta && opts.tableMeta.changeTrackingEnabled !== true) {
    out.push(
      adv.err(
        'change-tracking-unavailable',
        'validation',
        "Change tracking isn't enabled for this table",
        <>
          Dataverse returns a delta link only for tables with <strong>Track changes</strong> on.
          Turn it on in the table&apos;s properties in Power Apps (Advanced options), or turn Track
          changes off in the Prefer pane.
        </>,
        'prefer',
        CHANGE_TRACKING_DOCS,
      ),
    );
  }
  const sent = new Set(opts.queryParts.map((p) => p.key));
  for (const { key, focusNode } of CHANGE_TRACKING_UNSUPPORTED) {
    if (!sent.has(key)) continue;
    out.push(
      adv.err(
        `change-tracking-${focusNode}`,
        'validation',
        `Change tracking can't be combined with ${key}`,
        <>
          Dataverse rejects <code>{key}</code> when the request sends{' '}
          <code>Prefer: odata.track-changes</code>. Clear <code>{key}</code>, or turn Track changes
          off in the Prefer pane. <code>$select</code> works with change tracking.
        </>,
        focusNode,
        CHANGE_TRACKING_DOCS,
      ),
    );
  }
  return out;
}

/** The response's `@odata.deltaLink`, if it has one. */
export function deltaLinkOf(body: unknown): string | undefined {
  if (!body || typeof body !== 'object') return undefined;
  const link = (body as Record<string, unknown>)['@odata.deltaLink'];
  return typeof link === 'string' && link ? link : undefined;
}

/**
 * True for a delta response's entry for a deleted row: its `@odata.context`
 * ends in `/$deletedEntity`, and it carries only `id` and `reason`.
 */
export function isDeletedEntry(row: unknown): row is { id?: unknown; reason?: unknown } {
  if (!row || typeof row !== 'object' || Array.isArray(row)) return false;
  const ctx = (row as Record<string, unknown>)['@odata.context'];
  return typeof ctx === 'string' && /\$deletedEntity$/.test(ctx);
}

/** How many entries of a delta response are deleted rows. */
export function deletedRowCount(body: unknown): number {
  const value = (body as { value?: unknown } | null)?.value;
  return Array.isArray(value) ? value.filter(isDeletedEntry).length : 0;
}

/** The grid's Change column (#233): set on deleted rows only. */
export const CHANGE_KEY = '__change';
/** The deleted row's id, which the Change column shows next to its badge. */
export const DELETED_ID_KEY = '__deletedId';

/**
 * A delta response's entry as a grid row. A deleted row's entry has only
 * `id` and `reason`, which would otherwise become two columns of their own,
 * blank on every other row; it becomes a row whose Change column says
 * "Deleted" with its id. Every other row is returned as it is.
 */
export function deltaGridRow(row: Record<string, unknown>): Record<string, unknown> {
  if (!isDeletedEntry(row)) return row;
  const id = typeof row.id === 'string' ? row.id : '';
  return { [CHANGE_KEY]: id ? `Deleted ${id}` : 'Deleted', [DELETED_ID_KEY]: id };
}
