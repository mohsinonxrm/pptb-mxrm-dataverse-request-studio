// Pure partitioner — walks a Dataverse response record and splits its
// fields into four buckets that the renderer handles differently:
//
//   • scalars        — primitive values (with optional FormattedValue +
//                      lookup-target-table annotations paired in)
//   • navObjects     — nested objects (N:1 / 1:1 expanded lookups)
//   • navCollections — arrays of objects (1:N / N:N expanded collections)
//   • annotations    — `@odata.*` / `@Microsoft.Dynamics.CRM.*` metadata
//
// Why this is its own module:
//   1. The detail renderer is a render tree; keeping the analysis pure
//      means we can unit-test the partition logic without DOM.
//   2. Annotations come in shapes the casual reader shouldn't care about
//      (`<field>@OData.Community.Display.V1.FormattedValue`,
//      `_<lookup>_value@Microsoft.Dynamics.CRM.lookuplogicalname`, etc.)
//      — pairing them with their owning field once, here, means the
//      renderer never has to think about strings-with-@-in-them.
//   3. The same partition shape works for any record-bag we get back
//      from the API: RetrieveSingle bodies, Create/Update return=
//      representation, Action/Function single-entity outputs.

import type { TableMeta, NavProperty } from '../../mock/metadata';

// ── Public types ─────────────────────────────────────────────────────

export interface PartitionedScalar {
  /** Raw key as it appears in the response. May be `_<attr>_value` for
   *  lookups; we DON'T rewrite the key because we want what the API
   *  actually returned to be visible. */
  key: string;
  value: unknown;
  /** Paired @OData.Community.Display.V1.FormattedValue, if any. */
  formattedValue?: string;
  /** Paired @Microsoft.Dynamics.CRM.lookuplogicalname (lookup target table
   *  for polymorphic lookups), if any. */
  lookupTargetLogicalName?: string;
}

export interface PartitionedNavObject {
  /** Navigation property name as it appears in the response. */
  key: string;
  /** The nested record object. Recursively rendered by RecordDetailCard. */
  value: Record<string, unknown>;
  /** Target entity logical name from metadata, if available. */
  targetEntityLogical?: string;
  /** Target entity display name from metadata, if available. */
  targetEntityDisplay?: string;
  /** Cardinality from metadata — used for the badge label. */
  cardinality?: NavProperty['cardinality'];
}

export interface PartitionedNavCollection {
  key: string;
  value: Record<string, unknown>[];
  targetEntityLogical?: string;
  targetEntityDisplay?: string;
  cardinality?: NavProperty['cardinality'];
}

export interface PartitionedNavNull {
  key: string;
  targetEntityLogical?: string;
  targetEntityDisplay?: string;
  cardinality?: NavProperty['cardinality'];
}

export interface PartitionedAnnotation {
  key: string;
  value: unknown;
}

export interface RecordPartition {
  scalars: PartitionedScalar[];
  navObjects: PartitionedNavObject[];
  navCollections: PartitionedNavCollection[];
  /** Null/empty expanded nav properties — surfaced so the user knows the
   *  field was queried but came back unset. */
  navNulls: PartitionedNavNull[];
  /** Top-level @odata.context, @odata.etag, @odata.editLink, etc.
   *  (field-level annotations that paired with a scalar are NOT here —
   *  they've been folded into PartitionedScalar.) */
  annotations: PartitionedAnnotation[];
}

// ── Implementation ───────────────────────────────────────────────────

// Annotation suffixes we know how to fold INTO a scalar. Anything else
// stays as a top-level annotation in the bottom panel.
const FORMATTED_VALUE_SUFFIX = '@OData.Community.Display.V1.FormattedValue';
const LOOKUP_LOGICAL_NAME_SUFFIX = '@Microsoft.Dynamics.CRM.lookuplogicalname';
const NAV_ASSOCIATED_NAVPROP_SUFFIX = '@Microsoft.Dynamics.CRM.associatednavigationproperty';

/** True if the key looks like a generic OData/CRM annotation (contains '@'). */
function isAnnotation(k: string): boolean {
  return k.includes('@');
}

/**
 * Partition the given record. `tbl` is optional — when absent (no metadata
 * loaded for this entity), the partitioner falls back to inference:
 *   • objects → nav object
 *   • arrays  → nav collection
 *   • null on a key that has paired @… annotations → nav null
 *   • everything else → scalar
 */
export function partitionRecord(record: Record<string, unknown>, tbl?: TableMeta): RecordPartition {
  const scalars: PartitionedScalar[] = [];
  const navObjects: PartitionedNavObject[] = [];
  const navCollections: PartitionedNavCollection[] = [];
  const navNulls: PartitionedNavNull[] = [];
  const annotations: PartitionedAnnotation[] = [];

  // First pass: index annotations by their owner key so we can fold them.
  // Annotations come in `<owner>@<suffix>` form; the owner is everything
  // before the first '@'.
  const annotationByOwner = new Map<string, Map<string, unknown>>();
  for (const k of Object.keys(record)) {
    const at = k.indexOf('@');
    if (at < 0) continue;
    const owner = k.slice(0, at);
    const suffix = k.slice(at); // includes the '@'
    if (!annotationByOwner.has(owner)) annotationByOwner.set(owner, new Map());
    annotationByOwner.get(owner)!.set(suffix, record[k]);
  }

  // Second pass: walk non-annotation keys and classify.
  for (const k of Object.keys(record)) {
    if (isAnnotation(k)) {
      // Owner-prefixed annotations are folded into their owner below if
      // the owner exists. Top-level metadata (@odata.context, @odata.etag)
      // has an empty owner — those go to `annotations`.
      const at = k.indexOf('@');
      const owner = k.slice(0, at);
      if (!owner || !Object.prototype.hasOwnProperty.call(record, owner)) {
        annotations.push({ key: k, value: record[k] });
      }
      continue;
    }

    const value = record[k];
    const nav = tbl?.navigationProperties.find((n) => n.name === k);
    const folded = annotationByOwner.get(k);

    // Pull nav metadata once. For polymorphic lookups in the wire-form
    // (`_primarycontactid_value`), the target-table annotation gives us
    // the actual table. For expanded nav properties (`primarycontactid`),
    // we use the registry.
    const targetEntityLogical = nav?.targetEntity;
    const targetEntityDisplay = nav?.targetEntity;
    const cardinality = nav?.cardinality;

    if (Array.isArray(value)) {
      // 1:N or N:N expanded collection (always an array of objects, even
      // when empty). Treat object-array uniformly; we don't expect scalar
      // arrays in Dataverse responses.
      navCollections.push({
        key: k,
        value: value as Record<string, unknown>[],
        targetEntityLogical,
        targetEntityDisplay,
        cardinality,
      });
      continue;
    }

    if (value !== null && typeof value === 'object') {
      // N:1 or 1:1 expanded lookup — nested record.
      navObjects.push({
        key: k,
        value: value as Record<string, unknown>,
        targetEntityLogical,
        targetEntityDisplay,
        cardinality,
      });
      continue;
    }

    // Null on a key that the metadata knows is a nav property → "queried,
    // but unset". Render it as a placeholder accordion so the user
    // doesn't think the field was missing entirely.
    if (value === null && nav) {
      navNulls.push({
        key: k,
        targetEntityLogical,
        targetEntityDisplay,
        cardinality,
      });
      continue;
    }

    // Scalar — fold any paired field-level annotations.
    const formattedValue = folded?.get(FORMATTED_VALUE_SUFFIX);
    const lookupTargetLogicalName = folded?.get(LOOKUP_LOGICAL_NAME_SUFFIX);
    scalars.push({
      key: k,
      value,
      formattedValue: formattedValue == null ? undefined : String(formattedValue),
      lookupTargetLogicalName:
        lookupTargetLogicalName == null ? undefined : String(lookupTargetLogicalName),
    });
  }

  return { scalars, navObjects, navCollections, navNulls, annotations };
}

/** Drop the leading underscore + trailing `_value` from `_xxx_value` lookup
 *  field names so the displayed key reads like the schema-name version.
 *  Returns the key unchanged for non-lookup-wire-form fields. */
export function prettifyKey(key: string): string {
  const m = key.match(/^_(.+)_value$/);
  return m ? m[1] : key;
}

/** Best-effort primary-name / primary-key resolution. Uses metadata when
 *  available, otherwise scans the record for common Dataverse patterns
 *  (`name`, `fullname`, `<entity>id`, `subject`, `title`). */
export function pickHeadlineFields(
  record: Record<string, unknown>,
  tbl?: TableMeta,
): { headline: string; subline?: string } {
  const headlineKey =
    tbl?.primaryName ||
    (Object.prototype.hasOwnProperty.call(record, 'name')
      ? 'name'
      : Object.prototype.hasOwnProperty.call(record, 'fullname')
        ? 'fullname'
        : Object.prototype.hasOwnProperty.call(record, 'subject')
          ? 'subject'
          : Object.prototype.hasOwnProperty.call(record, 'title')
            ? 'title'
            : null);
  const idKey = tbl?.primaryKey;
  const headline =
    headlineKey != null && record[headlineKey] != null
      ? String(record[headlineKey])
      : '(unnamed record)';
  const idValue = idKey != null ? record[idKey] : undefined;
  const subline = idValue != null ? String(idValue) : undefined;
  return { headline, subline };
}

// Re-export the suffix constants for any caller (e.g. the annotations
// panel) that wants to label a field-level annotation as such.
export const ANNOTATION_SUFFIXES = {
  FORMATTED_VALUE: FORMATTED_VALUE_SUFFIX,
  LOOKUP_LOGICAL_NAME: LOOKUP_LOGICAL_NAME_SUFFIX,
  NAV_ASSOCIATED_NAVPROP: NAV_ASSOCIATED_NAVPROP_SUFFIX,
};
