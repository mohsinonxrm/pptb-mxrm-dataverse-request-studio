// URL composition for each Read mode. Returns both the relative path and the
// rendered relative URL plus an estimated byte length for validation.

import { ENV } from '../env/environment';
import {
  altKeyColumnType,
  findTable,
  findTableBySet,
  findColumn,
  isLookupLike,
  type AlternateKeyDef,
  type AttributeTypeCode,
  type TableMeta,
} from '../metadata/metadata';
import { bindPropertyFor, encodeFieldBody, parseRawBody } from './fieldBody';
import { groupToOData } from '../editors/filter/filterTree';
import { attrRefByName } from './odataAttr';
import { readRowLookup } from './recordSnapshot';
import { orderbyToOData } from '../editors/OrderbyEditor';
import { expandToOData } from '../editors/ExpandEditor';
import { applyToOData } from '../editors/ApplyEditor';
import type {
  RetrieveMultipleState,
  RetrieveSingleState,
  RetrieveNextLinkState,
  PredefinedQueryState,
} from '../state/readState';
import {
  defaultBypassOptions,
  type CreateState,
  type UpdateState,
  type UpsertState,
  type DeleteState,
  type MergeState,
  type CreateFieldValue,
} from '../state/writeState';
import type { AssociateState, DisassociateState } from '../state/relateState';
import { isSingleValuedNav } from '../state/relateState';

// Re-export execute builders so callers don't have to know they live in a
// separate module. Same surface contract — keeps imports terse.
export {
  buildExecuteAction,
  buildExecuteActionBody,
  buildExecuteFunction,
  buildExecuteWorkflow,
  buildExecuteWorkflowBody,
} from './executeBuilders';

export {
  buildManageFile,
  manageFilePipeline,
  buildManageImage,
  manageImagePipeline,
  buildManageAttachment,
  manageAttachmentPipeline,
  attachmentTargetLabel,
  formatSize,
  type BinaryPipelineStep,
} from './binaryBuilders';

export interface BuiltRequest {
  /** path + query (no host), e.g. /api/data/v9.2/accounts?$select=name */
  relativeUrl: string;
  /** /accounts?... (without /api/data/v9.2 prefix) — used by code generators */
  relativeNoBase: string;
  /** estimated byte length for the GET URL (per §11/§16 caps) */
  bytes: number;
  /** Encoded query-string fragments in the order they would be emitted (for display) */
  queryParts: { key: string; value: string }[];
  /** entitySet name without query string (e.g. `accounts`) */
  entitySet: string;
  /** Entity logical name for Xrm.WebApi (e.g. `account`) — distinct from entitySet
   *  because pluralization isn't a blind s-trim (opportunity → opportunities). */
  entityLogical: string;
  /** Record id segment when present (Retrieve Single) — `<guid>` extracted from `(...)` */
  recordId?: string;
}

/**
 * Count only (#272) asks for this annotation too, so the result can say
 * whether Dataverse stopped counting at its limit.
 */
export const COUNT_ONLY_PREFER =
  'odata.include-annotations="Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded"';

// Keep characters that the Dataverse Web API docs render literally — slashes,
// commas, equals, dollar signs, parentheses, spaces, and single quotes — so the
// generated URL reads like the official examples on Microsoft Learn instead of
// a wall of `%XX`. Browsers / cURL / Power Automate all accept this form.
//
// The previous `%20 → +` substitution was form-encoding (application/x-www-form-urlencoded),
// not URL-query encoding. RFC 3986 query strings don't use `+` for space, and
// some Dataverse code paths mis-decode it. Spaces stay as spaces (or `%20`).
const enc = (s: string) =>
  encodeURIComponent(s)
    .replace(/%2F/g, '/')
    .replace(/%2C/g, ',')
    .replace(/%3D/g, '=')
    .replace(/%24/g, '$')
    .replace(/%28/g, '(')
    .replace(/%29/g, ')')
    .replace(/%20/g, ' ') // literal space — matches MS Learn URL examples
    .replace(/%27/g, "'") // literal apostrophe — OData string delimiter
    .replace(/%3B/g, ';'); // literal semicolon — separates inner $expand options

function joinQuery(parts: { key: string; value: string }[]): string {
  return parts
    .filter((p) => p.value !== '' && p.value != null)
    .map((p) => `${p.key}=${p.value}`)
    .join('&');
}

/**
 * Encode a $select array as an OData-friendly column list. Lookup columns
 * are rewritten to `_<logical>_value` because the OData schema doesn't
 * expose `<logical>` as an addressable property on the entity type. See
 * `engine/odataAttr.ts` for the DRB-style background.
 */
function selectToOData(table: TableMeta | undefined, cols: string[]): string {
  if (!table) return cols.join(',');
  return cols.map((c) => attrRefByName(table, c)).join(',');
}

/**
 * `keepTop`: emit `$top` even when a page size is set. Dataverse ignores
 * `$top` next to `Prefer: odata.maxpagesize`, so the Code tab drops it; the
 * executor keeps it when the host can't send that Prefer (#219).
 */
export function buildRetrieveMultiple(
  s: RetrieveMultipleState,
  opts: { keepTop?: boolean } = {},
): BuiltRequest {
  const tbl = findTable(s.table);
  if (!tbl) {
    return {
      relativeUrl: '',
      relativeNoBase: '',
      bytes: 0,
      queryParts: [],
      entitySet: '',
      entityLogical: '',
    };
  }
  const parts: { key: string; value: string }[] = [];
  // Count only (#272): one row of the id column plus `$count=true`, whose
  // `@odata.count` is the number of rows matching $filter. Not `<set>/$count`:
  // Dataverse rejects $filter on it ("Could not find a property named … on
  // type 'Edm.Int32'"), and the host can't read its plain-text answer.
  if (s.countOnly) {
    if (tbl.primaryKey) parts.push({ key: '$select', value: enc(tbl.primaryKey) });
    const f = s.filter.rules.length ? groupToOData(s.filter, tbl) : '';
    if (f) parts.push({ key: '$filter', value: enc(f) });
    parts.push({ key: '$top', value: '1' }, { key: '$count', value: 'true' });
    return finishRetrieveMultiple(tbl, parts);
  }
  // Encoder now takes the table so it can render the filter() stage's
  // predicate the same way the top-level $filter does (operator allowlist +
  // per-attribute literal encoding).
  const applyStr = applyToOData(s.apply, tbl);
  if (applyStr) {
    parts.push({ key: '$apply', value: enc(applyStr) });
    // Per §14: $select/$filter/$orderby behave differently after $apply (often ignored)
  } else {
    if (s.select.length) parts.push({ key: '$select', value: enc(selectToOData(tbl, s.select)) });
    if (s.filter.rules.length) {
      const f = groupToOData(s.filter, tbl);
      if (f) parts.push({ key: '$filter', value: enc(f) });
    }
    if (s.expand.length)
      parts.push({ key: '$expand', value: enc(expandToOData(s.expand, s.table)) });
    if (s.orderby.length)
      parts.push({ key: '$orderby', value: enc(orderbyToOData(s.orderby, tbl)) });
  }
  const pageSized = s.prefer.maxpagesize != null && s.prefer.maxpagesize > 0;
  if (s.top != null && s.top > 0 && (!pageSized || opts.keepTop)) {
    parts.push({ key: '$top', value: String(s.top) });
  }
  if (s.countOn) parts.push({ key: '$count', value: 'true' });
  return finishRetrieveMultiple(tbl, parts);
}

function finishRetrieveMultiple(
  tbl: TableMeta,
  parts: { key: string; value: string }[],
): BuiltRequest {
  const path = `${ENV.apiBase}/${tbl.entitySetName}`;
  const qs = joinQuery(parts);
  const relativeUrl = qs ? `${path}?${qs}` : path;
  const noBase = `/${tbl.entitySetName}${qs ? `?${qs}` : ''}`;
  return {
    relativeUrl,
    relativeNoBase: noBase,
    bytes: relativeUrl.length,
    queryParts: parts,
    entitySet: tbl.entitySetName,
    entityLogical: tbl.logicalName,
  };
}

export function buildRetrieveSingle(s: RetrieveSingleState): BuiltRequest {
  const tbl = findTable(s.table);
  if (!tbl)
    return {
      relativeUrl: '',
      relativeNoBase: '',
      bytes: 0,
      queryParts: [],
      entitySet: '',
      entityLogical: '',
    };
  const parts: { key: string; value: string }[] = [];
  if (s.select.length) parts.push({ key: '$select', value: enc(selectToOData(tbl, s.select)) });
  if (s.expand.length) parts.push({ key: '$expand', value: enc(expandToOData(s.expand, s.table)) });

  const idSegment = s.recordId ? `(${s.recordId})` : '(<id>)';
  const path = `${ENV.apiBase}/${tbl.entitySetName}${idSegment}`;
  const qs = joinQuery(parts);
  const relativeUrl = qs ? `${path}?${qs}` : path;
  const noBase = `/${tbl.entitySetName}${idSegment}${qs ? `?${qs}` : ''}`;
  return {
    relativeUrl,
    relativeNoBase: noBase,
    bytes: relativeUrl.length,
    queryParts: parts,
    entitySet: tbl.entitySetName,
    entityLogical: tbl.logicalName,
    recordId: s.recordId ?? undefined,
  };
}

export function buildRetrieveNextLink(
  s: RetrieveNextLinkState,
  /** Entity set → logical name from the environment's entity list. */
  resolveSet?: (entitySet: string) => string | undefined,
): BuiltRequest {
  const url = s.url || '';
  // The nextLink URL is opaque; we don't decompose it. Return as-is for display.
  // Best-effort: extract the entity-set segment for breadcrumbs.
  let entitySet = '';
  const m = url.match(/\/api\/data\/v\d+\.\d+\/([^?(/]+)/);
  if (m) entitySet = m[1];
  // For relativeNoBase, lop off the host and api base if present.
  let noBase = url;
  noBase = noBase.replace(/^https?:\/\/[^/]+/, '');
  noBase = noBase.replace(/^\/api\/data\/v\d+\.\d+/, '');
  // The table comes from the entity list (or the loaded metadata) by entity
  // set name. Plurals don't strip back to a logical name (activitypointers,
  // addresses, custom sets), so an unknown set stays '' rather than a guess (#224).
  const entityLogical =
    (entitySet && (resolveSet?.(entitySet) ?? findTableBySet(entitySet)?.logicalName)) || '';
  return {
    relativeUrl: url.replace(/^https?:\/\/[^/]+/, ''),
    relativeNoBase: noBase || '/',
    bytes: url.length,
    queryParts: [],
    entitySet,
    entityLogical,
  };
}

// ────────────────────────────────────────────────────────────
// Write — Create
// ────────────────────────────────────────────────────────────
//
// URL is just `POST /<entitySet>`. When Prefer: return=representation is set,
// $select can scope the echoed response — that's the only OData query option
// allowed on the URL for Create. Everything else lives in the JSON body.

export function buildCreate(s: CreateState): BuiltRequest {
  const tbl = findTable(s.table);
  if (!tbl)
    return {
      relativeUrl: '',
      relativeNoBase: '',
      bytes: 0,
      queryParts: [],
      entitySet: '',
      entityLogical: '',
    };
  const parts: { key: string; value: string }[] = [];
  // $select only meaningful with return=representation
  if (s.prefer.returnRepresentation && s.returnSelect.length) {
    parts.push({ key: '$select', value: enc(selectToOData(tbl, s.returnSelect)) });
  }
  const path = `${ENV.apiBase}/${tbl.entitySetName}`;
  const qs = joinQuery(parts);
  const relativeUrl = qs ? `${path}?${qs}` : path;
  const noBase = `/${tbl.entitySetName}${qs ? `?${qs}` : ''}`;
  return {
    relativeUrl,
    relativeNoBase: noBase,
    bytes: relativeUrl.length,
    queryParts: parts,
    entitySet: tbl.entitySetName,
    entityLogical: tbl.logicalName,
  };
}

/**
 * Serialize a CreateState into the JSON body Dataverse expects.
 *
 * In JSON mode (`rawBody` non-null) the typed text is the body: it is sent
 * as parsed, without re-encoding, so nulls, annotations and keys the form
 * doesn't know survive. An unparseable body yields `{}` (the modes block
 * Execute with the parse error). Otherwise the body is encoded from
 * fieldValues + nullFields — see engine/fieldBody.encodeFieldBody for the
 * per-type rules.
 */
export function buildCreateBody(s: CreateState): Record<string, unknown> {
  const tbl = findTable(s.table);
  if (!tbl) return {};
  if (s.rawBody != null) {
    const parsed = parseRawBody(s.rawBody);
    return parsed.ok ? parsed.body : {};
  }
  return encodeFieldBody(tbl, s.fieldValues, s.nullFields ?? []);
}

// ────────────────────────────────────────────────────────────
// Write — Update (PATCH /<entitySet>(<id>))
// ────────────────────────────────────────────────────────────
//
// Per https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/update-delete-entities-using-web-api:
//   • Only $select is allowed on the URL — and only when Prefer: return=representation
//     is set (otherwise it has no effect; $expand is ignored on update).
//   • The body shape is identical to Create — same `<col>@odata.bind` for lookups,
//     same comma-separated form for multi-select picklists. We reuse buildCreateBody.

export function buildUpdate(s: UpdateState): BuiltRequest {
  const tbl = findTable(s.table);
  if (!tbl)
    return {
      relativeUrl: '',
      relativeNoBase: '',
      bytes: 0,
      queryParts: [],
      entitySet: '',
      entityLogical: '',
    };

  const idSegment = s.recordId ? `(${s.recordId})` : '(<id>)';
  // PUT single-column drills into the property URL: /<entitySet>(<id>)/<col>
  const propSegment = s.method === 'PUT' && s.putColumn ? `/${s.putColumn}` : '';

  const parts: { key: string; value: string }[] = [];
  // $select only makes sense on PATCH with return=representation; PUT doesn't
  // accept query options on the property URL per docs.
  if (s.method === 'PATCH' && s.prefer.returnRepresentation && s.returnSelect.length) {
    parts.push({ key: '$select', value: enc(selectToOData(tbl, s.returnSelect)) });
  }
  const path = `${ENV.apiBase}/${tbl.entitySetName}${idSegment}${propSegment}`;
  const qs = joinQuery(parts);
  const relativeUrl = qs ? `${path}?${qs}` : path;
  const noBase = `/${tbl.entitySetName}${idSegment}${propSegment}${qs ? `?${qs}` : ''}`;
  return {
    relativeUrl,
    relativeNoBase: noBase,
    bytes: relativeUrl.length,
    queryParts: parts,
    entitySet: tbl.entitySetName,
    entityLogical: tbl.logicalName,
    recordId: s.recordId ?? undefined,
  };
}

/**
 * Body shape:
 *   • PATCH → same as Create — partial diff with `<col>@odata.bind` for lookups.
 *   • PUT (single column) → `{ "value": <scalar> }` per docs.
 */
export function buildUpdateBody(s: UpdateState): Record<string, unknown> {
  if (s.method === 'PUT') {
    if (!s.putColumn) return { value: null };
    const v = s.fieldValues[s.putColumn];
    if (v == null) return { value: null };
    // Lookups still emit @odata.bind on PUT — but the wrapping is `{value: ...}`.
    // For lookup PUT, the standard syntax is actually:
    //   PUT /accounts(<id>)/primarycontactid/$ref  with  { '@odata.id': '/contacts(<id>)' }
    // — that's the Associate-single mode, not a property PUT. So lookups
    // aren't valid here; surface the GUID directly and let the user know.
    return { value: v as unknown };
  }
  // PATCH path — delegate to the Create encoder (single source of truth).
  // Forward `nullFields` so explicit-null entries land in the body as
  // `<col>: null` or `<nav>@odata.bind: null` per docs.
  return buildCreateBody({
    table: s.table,
    fieldValues: s.fieldValues,
    nullFields: s.nullFields ?? [],
    rawBody: s.rawBody ?? null,
    prefer: s.prefer,
    headers: s.headers,
    returnSelect: s.returnSelect,
    duplicateDetection: false,
    bypass: defaultBypassOptions(),
    dirty: s.dirty,
  });
}

// ────────────────────────────────────────────────────────────
// Write — Upsert (PATCH /<entitySet>(<id>) | (<altKey>=<val>,…))
// ────────────────────────────────────────────────────────────
//
// Per docs (use-upsert-insert-update-record):
//   • GUID-addressed: PATCH /accounts(<guid>)
//   • Alt-key-addressed: PATCH /accounts(accountnumber='ACC-0001')
//   • Composite alt-key: PATCH /sample_things(sample_key1=1,sample_key2=1)
//
// Alternate-key values follow OData literal rules: strings get single-quoted,
// numbers / GUIDs / booleans / date-times are bare. We dispatch on the column
// metadata to pick the right literal form. The key segment is part of the
// URL path, so every value is percent-encoded (`#`, `?`, `%`, spaces, … would
// otherwise truncate or corrupt the URL when it's copied out) — MS Learn's own
// example is `new_skts(new_name='M%26M')`. A lookup key column is addressed by
// its lookup property, `_<name>_value` (MS Learn: "Use an alternate key to
// reference a record").

/** Placeholder for a key part the user hasn't filled in yet (preview only —
 *  Execute is gated on every part being set). Mirrors the `(<id>)` placeholder. */
const UNSET_KEY_PART = '<value>';

/**
 * Characters Dataverse rejects inside an alternate-key value even when
 * percent-encoded (MS Learn: "Alternate keys with special character values").
 * `#` is listed on the maker-portal page for PATCH.
 */
const ALT_KEY_RESERVED_CHARS = /[/<>*%&:\\?+#]/;

const quoteKeyString = (value: string) => `'${encodeURIComponent(value.replace(/'/g, "''"))}'`;

function odataLiteral(value: string, type: AttributeTypeCode | undefined): string {
  // Unknown type → quoted as a string (the Upsert UI flags it, #64).
  if (!type) return quoteKeyString(value);
  switch (type) {
    case 'Integer':
    case 'BigInt':
    case 'Decimal':
    case 'Double':
    case 'Money':
    case 'Picklist':
    case 'MultiSelectPicklist':
    case 'State':
    case 'Status':
    case 'EntityName':
    case 'Uniqueidentifier':
    case 'Lookup':
    case 'Customer':
    case 'Owner': {
      const v = value.trim();
      return v === '' ? UNSET_KEY_PART : encodeURIComponent(v);
    }
    case 'DateTime': {
      // Edm.DateTimeOffset / Edm.Date literals are unquoted:
      // `(new_effectivedate=2024-05-01T00:00:00Z)`. `:` is a legal path
      // character and part of the literal, so keep it readable; `+` in an
      // offset (`+05:30`) must be encoded or it can read as a space.
      const v = value.trim();
      return v === '' ? UNSET_KEY_PART : encodeURIComponent(v).replace(/%3A/gi, ':');
    }
    case 'Boolean':
      return value === 'true' || value === '1' ? 'true' : 'false';
    default:
      return quoteKeyString(value);
  }
}

function altKeySegment(
  table: string,
  def: AlternateKeyDef,
  keyValues: Record<string, string>,
): string {
  const tbl = findTable(table);
  const parts = def.columns.map((col) => {
    const type = altKeyColumnType(tbl, def, col);
    const raw = keyValues[col] ?? '';
    // A lookup key column missing from the loaded columns still needs the
    // `_<name>_value` property name.
    const loaded = tbl?.columns.some((c) => c.logicalName === col);
    const name =
      !loaded && (type === 'Lookup' || type === 'Customer' || type === 'Owner')
        ? `_${col}_value`
        : attrRefByName(tbl, col);
    return `${name}=${odataLiteral(raw, type)}`;
  });
  return `(${parts.join(',')})`;
}

/** Key columns whose AttributeType is unknown — quoted as strings (#64). */
export function altKeyUnknownTypeColumns(table: string, def: AlternateKeyDef): string[] {
  const tbl = findTable(table);
  return def.columns.filter((col) => !altKeyColumnType(tbl, def, col));
}

/**
 * Key columns whose string value contains a character Dataverse rejects in
 * an alternate-key URL even when encoded. Drives the Upsert advisory that
 * points at the `$filter` workaround. Only parts emitted as quoted strings
 * are checked (String / Memo, or unknown type, which is quoted as a string)
 * — the `:` in a date-time literal is syntax, not data.
 */
export function altKeyReservedCharColumns(
  table: string,
  def: AlternateKeyDef,
  keyValues: Record<string, string>,
): string[] {
  const tbl = findTable(table);
  return def.columns.filter((col) => {
    const type = altKeyColumnType(tbl, def, col);
    if (type && type !== 'String' && type !== 'Memo') return false;
    return ALT_KEY_RESERVED_CHARS.test(keyValues[col] ?? '');
  });
}

export function buildUpsert(s: UpsertState): BuiltRequest {
  const tbl = findTable(s.table);
  if (!tbl)
    return {
      relativeUrl: '',
      relativeNoBase: '',
      bytes: 0,
      queryParts: [],
      entitySet: '',
      entityLogical: '',
    };

  let idSegment = '(<id>)';
  let recordId: string | undefined;
  const key = s.key;
  if (key.kind === 'guid') {
    if (key.recordId) {
      idSegment = `(${key.recordId})`;
      recordId = key.recordId;
    }
  } else {
    const def = tbl.alternateKeys?.find((k) => k.name === key.keyName);
    if (def) {
      idSegment = altKeySegment(s.table, def, key.keyValues);
    }
  }

  const parts: { key: string; value: string }[] = [];
  if (s.prefer.returnRepresentation && s.returnSelect.length) {
    parts.push({ key: '$select', value: enc(selectToOData(tbl, s.returnSelect)) });
  }
  const path = `${ENV.apiBase}/${tbl.entitySetName}${idSegment}`;
  const qs = joinQuery(parts);
  const relativeUrl = qs ? `${path}?${qs}` : path;
  const noBase = `/${tbl.entitySetName}${idSegment}${qs ? `?${qs}` : ''}`;
  return {
    relativeUrl,
    relativeNoBase: noBase,
    bytes: relativeUrl.length,
    queryParts: parts,
    entitySet: tbl.entitySetName,
    entityLogical: tbl.logicalName,
    recordId,
  };
}

export function buildUpsertBody(s: UpsertState): Record<string, unknown> {
  // Body is identical to Create — same encoding rules. Per docs, alt-key
  // values are NOT included in the body when using alt-key URL addressing.
  return buildCreateBody({
    table: s.table,
    fieldValues: s.fieldValues,
    nullFields: s.nullFields ?? [],
    rawBody: s.rawBody ?? null,
    prefer: s.prefer,
    headers: s.headers,
    returnSelect: s.returnSelect,
    duplicateDetection: false,
    bypass: defaultBypassOptions(),
    dirty: s.dirty,
  });
}

// ────────────────────────────────────────────────────────────
// Write — Delete (DELETE /<entitySet>(<id>) [ /<property> ])
// ────────────────────────────────────────────────────────────

export function buildDelete(s: DeleteState): BuiltRequest {
  const tbl = findTable(s.table);
  if (!tbl)
    return {
      relativeUrl: '',
      relativeNoBase: '',
      bytes: 0,
      queryParts: [],
      entitySet: '',
      entityLogical: '',
    };
  const idSegment = s.recordId ? `(${s.recordId})` : '(<id>)';
  const propSegment =
    s.scope.kind === 'single-property' && s.scope.column ? `/${s.scope.column}` : '';
  const path = `${ENV.apiBase}/${tbl.entitySetName}${idSegment}${propSegment}`;
  const noBase = `/${tbl.entitySetName}${idSegment}${propSegment}`;
  return {
    relativeUrl: path,
    relativeNoBase: noBase,
    bytes: path.length,
    queryParts: [],
    entitySet: tbl.entitySetName,
    entityLogical: tbl.logicalName,
    recordId: s.recordId ?? undefined,
  };
}

// ────────────────────────────────────────────────────────────
// Write — Merge (POST /Merge)
// ────────────────────────────────────────────────────────────
//
// URL: /Merge (unbound action — no entity set prefix). Body shape carries
// Target, Subordinate, UpdateContent, PerformParentingChecks.

export function buildMerge(s: MergeState): BuiltRequest {
  const tbl = findTable(s.table);
  if (!tbl)
    return {
      relativeUrl: '',
      relativeNoBase: '',
      bytes: 0,
      queryParts: [],
      entitySet: '',
      entityLogical: '',
    };
  const path = `${ENV.apiBase}/Merge`;
  return {
    relativeUrl: path,
    relativeNoBase: '/Merge',
    bytes: path.length,
    queryParts: [],
    entitySet: tbl.entitySetName,
    entityLogical: tbl.logicalName,
  };
}

/**
 * Build the Merge body from per-field choices.
 *
 * UpdateContent is composed by walking the user's `fieldChoices`:
 *   • 'subordinate' → copy the Subordinate row's value for that field
 *   • 'custom'      → take customValues[field]
 *   • 'target'      → SKIP (server keeps Target's value)
 *
 * The composed body looks like the example in the docs:
 *   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/merge-entity-using-web-api
 *
 *   {
 *     Target:      { @odata.type, accountid },
 *     Subordinate: { @odata.type, accountid },
 *     UpdateContent: { @odata.type, telephone1, websiteurl, description },
 *     PerformParentingChecks: true,
 *   }
 *
 * Incident tables ignore UpdateContent per docs — we still build it so the
 * preview shows what would have been sent if it weren't suppressed.
 */
export function buildMergeBody(s: MergeState): Record<string, unknown> {
  const tbl = findTable(s.table);
  if (!tbl) return {};
  const odataType = `Microsoft.Dynamics.CRM.${tbl.logicalName}`;
  const refKey = tbl.primaryKey;

  // Resolve the Subordinate row so we can read its values for 'subordinate'
  // choices. Reads from the live snapshot in state (populated by MergeMode
  // after a one-shot fetch). May be undefined if the user hasn't picked a
  // subordinate yet OR the fetch is still in flight; in that case
  // 'subordinate' choices are skipped at body-build time, matching the
  // safe default.
  const subRow = s.subordinateSnapshot ?? undefined;

  // Compose UpdateContent — only fields the user explicitly overrode
  const update: Record<string, CreateFieldValue> = {};
  for (const [field, choice] of Object.entries(s.fieldChoices)) {
    if (choice === 'target') continue;
    const col = findColumn(tbl, field);
    if (!col) continue;
    if (choice === 'subordinate') {
      if (!subRow) continue;
      if (isLookupLike(col)) {
        // The row carries lookups only as `_<col>_value` (never under the
        // logical name), so this must run before the generic null check.
        // Bind to the entity the annotation names — for polymorphic
        // lookups (customer / owner / regarding) `targets[0]` may be the
        // wrong entity set (#96). Unknown target → skip rather than guess.
        const lk = readRowLookup(subRow, col);
        if (lk?.targetEntity) {
          update[field] = { id: lk.id, targetEntity: lk.targetEntity };
        }
        continue;
      }
      const raw = subRow[field];
      if (raw == null) continue;
      if (col.attributeType === 'MultiSelectPicklist') {
        if (typeof raw === 'string' && raw) {
          update[field] = raw
            .split(',')
            .map((n) => Number(n))
            .filter((n) => Number.isFinite(n));
        }
        continue;
      }
      update[field] = raw as CreateFieldValue;
      continue;
    }
    if (choice === 'custom') {
      const cv = s.customValues[field];
      if (cv == null) continue;
      update[field] = cv;
    }
  }

  const updateContent = buildCreateBody({
    table: s.table,
    fieldValues: update,
    nullFields: [],
    prefer: {} as never,
    headers: [],
    returnSelect: [],
    duplicateDetection: false,
    bypass: defaultBypassOptions(),
    dirty: new Set(),
  });

  return {
    Target: {
      '@odata.type': odataType,
      [refKey]: s.targetId ?? '<target-guid>',
    },
    Subordinate: {
      '@odata.type': odataType,
      [refKey]: s.subordinateId ?? '<subordinate-guid>',
    },
    UpdateContent: {
      '@odata.type': odataType,
      ...updateContent,
    },
    PerformParentingChecks: s.performParentingChecks,
  };
}

// ────────────────────────────────────────────────────────────
// Relate — Associate (POST/PUT /<set>(<id>)/<nav>/$ref)
// ────────────────────────────────────────────────────────────
//
// One Associate state can produce 1+ HTTP requests:
//   • Collection-valued: one POST per target row.
//   • Single-valued:     one PUT (single target only).
//
// We surface BOTH: `buildAssociate` returns the BuiltRequest for the FIRST
// target (drives the URL bar + code generators), `buildAssociateRequests`
// returns the full list for multi-target preview + sequential execution.

/**
 * Wire shape for a single Associate request.
 *
 * **Single-valued (N:1) — PATCH @odata.bind (docs-preferred)**
 *   Per https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/associate-disassociate-entities-using-web-api#associate-with-a-single-valued-navigation-property
 *     PATCH /<source-set>(<sourceId>)
 *     { "<nav>@odata.bind": "<target-set>(<targetId>)" }
 *
 * **Collection-valued (1:N / N:N) — POST $ref**
 *   Per the same doc, the only supported form for collection-valued is:
 *     POST /<source-set>(<sourceId>)/<nav>/$ref
 *     { "@odata.id": "<absolute-url>/<target-set>(<targetId>)" }
 *
 * The MS doc lists PUT $ref as an "Other method" for single-valued, which
 * works at the wire but isn't exposed by PPTB's `dataverseAPI` (the host's
 * `.associate()` always POSTs). We emit the docs-preferred PATCH so the
 * Code tab snippets match what Microsoft documents AND the live executor
 * can route through `dvHost.update(...)` which DOES accept @odata.bind
 * annotations in the body.
 */
export interface AssociateRequest {
  /** HTTP verb — `POST` for collection-valued $ref, `PATCH` for single-valued @odata.bind. */
  method: 'POST' | 'PATCH';
  /**
   * Relative URL:
   *   • POST  → `/<set>(<id>)/<nav>/$ref`
   *   • PATCH → `/<set>(<id>)` (no /nav/$ref suffix — the body owns the link)
   */
  relativeUrl: string;
  relativeNoBase: string;
  /**
   * Body shape:
   *   • POST  → `{ "@odata.id": "<absolute-url>" }`
   *   • PATCH → `{ "<nav>@odata.bind": "<target-set>(<id>)" }`
   *
   * Both are Record<string, unknown> because PATCH's key is dynamic
   * (depends on the disambiguated nav property name).
   */
  body: Record<string, unknown>;
  /** Target row id this request points at (for display). */
  targetId: string;
}

/**
 * A single-valued (N:1) link is set or cleared with a PATCH of the source
 * record. Dataverse treats a PATCH to a missing id as an upsert, so
 * `If-Match: *` makes a mistyped source id fail with 404 instead of creating
 * a near-empty record (#254). Collection-valued requests are unaffected.
 */
export const PATCH_EXISTING_HEADERS: Readonly<Record<string, string>> = { 'If-Match': '*' };

/**
 * `headers` for one Associate / Disassociate request: plus `If-Match: *` on
 * the PATCH form. An If-Match the user added (any case, e.g. an ETag) is
 * kept — it already prevents the create, and a second one would conflict.
 */
export function relateRequestHeaders(
  method: string,
  headers: Record<string, string>,
): Record<string, string> {
  if (method !== 'PATCH') return headers;
  if (Object.keys(headers).some((k) => k.trim().toLowerCase() === 'if-match')) return headers;
  return { ...headers, ...PATCH_EXISTING_HEADERS };
}

/** Absolute URL used inside the @odata.id body — required by the spec for $ref endpoints. */
function absoluteRefUrl(entitySet: string, id: string): string {
  return `https://${ENV.host}${ENV.apiBase}/${entitySet}(${id})`;
}

export function buildAssociateRequests(s: AssociateState): AssociateRequest[] {
  const sourceTbl = findTable(s.table);
  if (!sourceTbl || !s.sourceId || !s.navProperty) return [];
  const nav = sourceTbl.navigationProperties.find((n) => n.name === s.navProperty);
  if (!nav) return [];
  const targetTbl = findTable(nav.targetEntity);
  if (!targetTbl) return [];

  if (isSingleValuedNav(nav)) {
    // Single-valued (N:1) — docs-preferred PATCH @odata.bind. Wire shape:
    //   PATCH /<source-set>(<sourceId>)
    //   { "<nav>@odata.bind": "<target-set>(<targetId>)" }
    // Only the first target is honored — single-valued semantics enforce
    // "one lookup at a time".
    const targetId = s.targets.filter(Boolean)[0];
    if (!targetId) return [];
    // Resolve the disambiguated bind property name (e.g.
    // `customerid_account@odata.bind`, `regardingobjectid_task_account@odata.bind`)
    // via the same helper Create/Update use — keeps the wire form consistent
    // across modes.
    const bindKey = bindPropertyFor(sourceTbl, { logicalName: nav.name }, targetTbl.logicalName);
    const path = `${ENV.apiBase}/${sourceTbl.entitySetName}(${s.sourceId})`;
    const noBase = `/${sourceTbl.entitySetName}(${s.sourceId})`;
    return [
      {
        method: 'PATCH',
        relativeUrl: path,
        relativeNoBase: noBase,
        // PATCH bind values are RELATIVE URLs (no scheme/host) per docs —
        // unlike POST $ref's @odata.id which requires absolute. The doc
        // example uses `"accounts(<id>)"` directly.
        body: { [bindKey]: `${targetTbl.entitySetName}(${targetId})` },
        targetId,
      },
    ];
  }

  // Collection-valued (1:N / N:N) — POST $ref per target. Unchanged.
  return s.targets.filter(Boolean).map((targetId) => {
    const path = `${ENV.apiBase}/${sourceTbl.entitySetName}(${s.sourceId})/${nav.name}/$ref`;
    const noBase = `/${sourceTbl.entitySetName}(${s.sourceId})/${nav.name}/$ref`;
    return {
      method: 'POST' as const,
      relativeUrl: path,
      relativeNoBase: noBase,
      body: { '@odata.id': absoluteRefUrl(targetTbl.entitySetName, targetId) },
      targetId,
    };
  });
}

/**
 * Returns the BuiltRequest for the FIRST queued Associate request. Used by
 * the UrlBar / CodeView. The Builder pane preview surfaces the full list via
 * buildAssociateRequests.
 */
export function buildAssociate(s: AssociateState): BuiltRequest {
  const sourceTbl = findTable(s.table);
  if (!sourceTbl)
    return {
      relativeUrl: '',
      relativeNoBase: '',
      bytes: 0,
      queryParts: [],
      entitySet: '',
      entityLogical: '',
    };
  // Even before targets are picked, render the URL skeleton so the user sees
  // the shape of the request as they configure the nav property. URL shape
  // depends on cardinality:
  //   • Single-valued (N:1) PATCH: /<source>(<id>)             — body owns the link
  //   • Collection-valued       POST: /<source>(<id>)/<nav>/$ref
  const nav = s.navProperty
    ? sourceTbl.navigationProperties.find((n) => n.name === s.navProperty)
    : undefined;
  const singleValued = nav ? isSingleValuedNav(nav) : false;
  const idSegment = s.sourceId ? `(${s.sourceId})` : '(<source-id>)';
  const navSegment = singleValued ? '' : nav ? `/${nav.name}/$ref` : '/<nav-property>/$ref';
  const path = `${ENV.apiBase}/${sourceTbl.entitySetName}${idSegment}${navSegment}`;
  const noBase = `/${sourceTbl.entitySetName}${idSegment}${navSegment}`;
  return {
    relativeUrl: path,
    relativeNoBase: noBase,
    bytes: path.length,
    queryParts: [],
    entitySet: sourceTbl.entitySetName,
    entityLogical: sourceTbl.logicalName,
    recordId: s.sourceId ?? undefined,
  };
}

export function buildAssociateBody(s: AssociateState): Record<string, unknown> {
  const reqs = buildAssociateRequests(s);
  // For the Code tab / preview we surface the FIRST body shape — the
  // multi-request preview lives in the Builder pane.
  if (reqs.length === 0) {
    // Skeleton body before targets are picked. Shape depends on cardinality:
    //   • Single-valued PATCH → `{ "<nav>@odata.bind": "<target-set>(<id>)" }`
    //   • Collection-valued POST → `{ "@odata.id": "<absolute-url>" }`
    const sourceTbl = findTable(s.table);
    const nav =
      sourceTbl && s.navProperty
        ? sourceTbl.navigationProperties.find((n) => n.name === s.navProperty)
        : undefined;
    const targetTbl = nav ? findTable(nav.targetEntity) : undefined;
    const singleValued = nav ? isSingleValuedNav(nav) : false;
    if (singleValued && sourceTbl && nav && targetTbl) {
      const bindKey = bindPropertyFor(sourceTbl, { logicalName: nav.name }, targetTbl.logicalName);
      return { [bindKey]: `${targetTbl.entitySetName}(<target-id>)` };
    }
    return {
      '@odata.id': targetTbl
        ? absoluteRefUrl(targetTbl.entitySetName, '<target-id>')
        : `https://${ENV.host}${ENV.apiBase}/<target-set>(<target-id>)`,
    };
  }
  return reqs[0].body;
}

// ────────────────────────────────────────────────────────────
// Relate — Disassociate
// ────────────────────────────────────────────────────────────
//
// Docs-preferred shapes (https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/associate-disassociate-entities-using-web-api):
//
//   • Single-valued (N:1): PATCH /<source-set>(<id>) with body
//     `{ "<nav>@odata.bind": null }` — clears the lookup. The doc shows
//     two equivalent body forms (with @odata.bind suffix and without);
//     we use @odata.bind for symmetry with the Associate body shape.
//
//   • Collection-valued (1:N / N:N) per-target: DELETE /<source-set>(<id>)/<nav>(<targetId>)/$ref
//     — removes one link from the collection. One DELETE per target.
//
// PPTB constraint that motivated the PATCH switch for single-valued:
// `dvHost.disassociate` REQUIRES a targetId in its signature, so the
// no-target single-valued DELETE form was previously silently skipped.
// PATCH @odata.bind: null routes through `dvHost.update(...)` which
// accepts arbitrary body shapes including null values.

export function buildDisassociate(s: DisassociateState): BuiltRequest {
  const sourceTbl = findTable(s.table);
  if (!sourceTbl)
    return {
      relativeUrl: '',
      relativeNoBase: '',
      bytes: 0,
      queryParts: [],
      entitySet: '',
      entityLogical: '',
    };
  const nav = s.navProperty
    ? sourceTbl.navigationProperties.find((n) => n.name === s.navProperty)
    : undefined;
  const idSegment = s.sourceId ? `(${s.sourceId})` : '(<source-id>)';
  const singleValued = nav ? isSingleValuedNav(nav) : false;

  // Single-valued → PATCH on the source record (body owns the null). URL
  // is just `/<source-set>(<id>)` — no /nav segment.
  if (singleValued) {
    const path = `${ENV.apiBase}/${sourceTbl.entitySetName}${idSegment}`;
    const noBase = `/${sourceTbl.entitySetName}${idSegment}`;
    return {
      relativeUrl: path,
      relativeNoBase: noBase,
      bytes: path.length,
      queryParts: [],
      entitySet: sourceTbl.entitySetName,
      entityLogical: sourceTbl.logicalName,
      recordId: s.sourceId ?? undefined,
    };
  }

  // Collection-valued → DELETE per target with /<nav>(<targetId>)/$ref.
  const navSegment = nav ? `/${nav.name}` : '/<nav-property>';
  const firstTarget = s.targetIds[0];
  const targetSegment = firstTarget ? `(${firstTarget})` : '(<target-id>)';
  const path = `${ENV.apiBase}/${sourceTbl.entitySetName}${idSegment}${navSegment}${targetSegment}/$ref`;
  const noBase = `/${sourceTbl.entitySetName}${idSegment}${navSegment}${targetSegment}/$ref`;
  return {
    relativeUrl: path,
    relativeNoBase: noBase,
    bytes: path.length,
    queryParts: [],
    entitySet: sourceTbl.entitySetName,
    entityLogical: sourceTbl.logicalName,
    recordId: s.sourceId ?? undefined,
  };
}

/**
 * Per-request wire shape:
 *
 *   • Single-valued (N:1) — exactly one PATCH with `@odata.bind: null` body.
 *     `targetId` is null because the URL doesn't carry a target.
 *
 *   • Collection-valued — one DELETE per target id.
 */
export interface DisassociateRequest {
  method: 'DELETE' | 'PATCH';
  relativeUrl: string;
  relativeNoBase: string;
  /** Target id for the link being removed. `null` for single-valued PATCH (URL has no target). */
  targetId: string | null;
  /** Body — only populated for the PATCH variant. */
  body?: Record<string, unknown>;
}

export function buildDisassociateRequests(s: DisassociateState): DisassociateRequest[] {
  const sourceTbl = findTable(s.table);
  if (!sourceTbl || !s.sourceId || !s.navProperty) return [];
  const nav = sourceTbl.navigationProperties.find((n) => n.name === s.navProperty);
  if (!nav) return [];

  const idSegment = `(${s.sourceId})`;
  const navSegment = `/${nav.name}`;

  if (isSingleValuedNav(nav)) {
    // Single-valued → exactly one PATCH @odata.bind: null on the source record.
    // Per docs (link in the section header above), this is the canonical
    // "clear the lookup" shape. The disambiguated bind key resolves via the
    // first target listed on the column metadata (we don't know which target
    // the lookup currently points at — and it doesn't matter, the server
    // clears the column regardless).
    //
    // Resolve the target entity for the bind-key disambiguation. For single-
    // target Lookups this is just the nav's targetEntity; for polymorphic
    // lookups we use that as a best-effort. Disambiguated bind keys all
    // resolve to the same wire effect when the value is null.
    const targetTbl = findTable(nav.targetEntity);
    const bindKey = bindPropertyFor(
      sourceTbl,
      { logicalName: nav.name },
      targetTbl?.logicalName ?? nav.targetEntity,
    );
    return [
      {
        method: 'PATCH',
        relativeUrl: `${ENV.apiBase}/${sourceTbl.entitySetName}${idSegment}`,
        relativeNoBase: `/${sourceTbl.entitySetName}${idSegment}`,
        targetId: null,
        body: { [bindKey]: null },
      },
    ];
  }
  // Collection-valued: one DELETE per target id — unchanged
  return s.targetIds.filter(Boolean).map((targetId) => ({
    method: 'DELETE' as const,
    relativeUrl: `${ENV.apiBase}/${sourceTbl.entitySetName}${idSegment}${navSegment}(${targetId})/$ref`,
    relativeNoBase: `/${sourceTbl.entitySetName}${idSegment}${navSegment}(${targetId})/$ref`,
    targetId,
  }));
}

/**
 * Collection a predefined query runs against when it targets a related
 * collection (#5): `<sourceSet>(<id>)/<nav>`. Null for the entity-set form,
 * or while the source table / nav aren't picked or loaded yet.
 */
export function predefinedCollectionPath(s: PredefinedQueryState): string | null {
  const r = s.related;
  if (!r || !r.sourceTable || !r.navProperty) return null;
  const src = findTable(r.sourceTable);
  if (!src) return null;
  return `${src.entitySetName}(${r.recordId || '<record-id>'})/${r.navProperty}`;
}

export function buildPredefinedQuery(s: PredefinedQueryState): BuiltRequest {
  const tbl = findTable(s.table);
  if (!tbl)
    return {
      relativeUrl: '',
      relativeNoBase: '',
      bytes: 0,
      queryParts: [],
      entitySet: '',
      entityLogical: '',
    };
  const parts: { key: string; value: string }[] = [];
  parts.push({ key: s.queryType, value: s.queryId ?? '<query-guid>' });
  if (s.top != null && s.top > 0 && (s.prefer.maxpagesize == null || s.prefer.maxpagesize === 0)) {
    parts.push({ key: '$top', value: String(s.top) });
  }
  const collection = predefinedCollectionPath(s) ?? tbl.entitySetName;
  const path = `${ENV.apiBase}/${collection}`;
  const qs = joinQuery(parts);
  const relativeUrl = qs ? `${path}?${qs}` : path;
  const noBase = `/${collection}${qs ? `?${qs}` : ''}`;
  return {
    relativeUrl,
    relativeNoBase: noBase,
    bytes: relativeUrl.length,
    queryParts: parts,
    entitySet: tbl.entitySetName,
    entityLogical: tbl.logicalName,
  };
}
