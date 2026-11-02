// Field-set body encoding — shared by the write builders (urlBuilder) and the
// FieldSetEditor's JSON view.
//
// Two directions:
//
//   encodeFieldBody  — in-memory `fieldValues` + `nullFields` → the JSON body
//                      Dataverse expects (lookups as `<nav>@odata.bind`,
//                      multi-select as "1,2,3", explicit nulls as `null`).
//   decodeFieldBody  — a user-typed JSON body → `fieldValues` + `nullFields`,
//                      plus a list of keys the form can't represent (so the
//                      UI can report them instead of silently dropping them).
//
// While the field set is in JSON mode the typed text is the body (see
// `rawBody` on the write states). `parseRawBody` is the single parser for
// that text, so the URL bar / Code tab / Execute all agree on what is sent.

import {
  findTable,
  findColumn,
  isLookupLike,
  type ColumnMeta,
  type TableMeta,
} from '../metadata/metadata';
import type { CreateFieldValue, LookupFieldValue } from '../state/writeState';

/**
 * Resolve the `@odata.bind` property name for a lookup-like column.
 *
 * For SINGLE-target lookups (most plain Lookup columns), Dataverse exposes
 * one nav-property whose name == the attribute logical name. Both
 * `<col>@odata.bind` and `<navProp>@odata.bind` work because they're the
 * same string.
 *
 * For POLYMORPHIC lookups (Customer, multi-target Lookup like
 * `regardingobjectid` on activities), the bare attribute name is NOT a
 * declared nav-property. The server returns:
 *   "An undeclared property '<attr>' which only has property annotations
 *    in the payload but no property value was found in the payload."
 *
 * The right name is the target-disambiguated nav-property — e.g.,
 * `customerid_account`, `customerid_contact`, `regardingobjectid_account_task`.
 * We look that up by matching the ManyToOne nav whose `referencingAttribute`
 * equals our column AND whose `targetEntity` equals the user's picked target.
 *
 * For OWNER columns (`ownerid` on most tables), Dataverse special-cases
 * the bare attribute name — it accepts both `/systemusers(...)` and
 * `/teams(...)` via a single nav. Empirically `ownerid@odata.bind` works
 * regardless of target; using the disambiguated form ALSO works. We use
 * the disambiguated form when we can find it, and fall back to the bare
 * name otherwise.
 *
 * Falls back to `<col>@odata.bind` when metadata can't resolve the nav —
 * single-target lookups always work this way, and the server will produce
 * the same clear error as before if the lookup happens to be polymorphic.
 */
export function bindPropertyFor(
  parentTbl: TableMeta,
  col: { logicalName: string },
  targetEntity: string,
): string {
  const nav = parentTbl.navigationProperties.find(
    (n) =>
      n.cardinality === 'ManyToOne' &&
      n.referencingAttribute === col.logicalName &&
      n.targetEntity === targetEntity,
  );
  return `${nav?.name ?? col.logicalName}@odata.bind`;
}

/**
 * Serialize fieldValues + nullFields into the JSON shape Dataverse expects.
 *
 * Per https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/create-entity-web-api:
 *
 *   • Lookup-like columns (Lookup / Customer / Owner) → `<col>@odata.bind`
 *     with value `/<entitySetName>(<guid>)`. The Customer/Owner polymorphic
 *     case uses the user-picked target's entity set.
 *   • MultiSelectPicklist → comma-separated string of integer values.
 *   • Everything else → the literal as-is.
 *
 * Explicit-null entries (nullFields) are emitted first:
 *   • Regular columns → `"<col>": null`
 *   • Lookup columns  → `"<nav>@odata.bind": null` (clears the lookup) per
 *     https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/associate-disassociate-entities-using-web-api#disassociate-by-using-a-single-valued-navigation-property
 *   We don't know which target the user "would have" picked, so the bind
 *   property resolves via the first target — the server clears the column
 *   regardless of which disambiguated form we use.
 *
 * Unset fields (`null`, `undefined`, `''` for strings, empty arrays, empty
 * lookup with no id) are dropped — only what the user explicitly populated
 * goes on the wire.
 */
export function encodeFieldBody(
  tbl: TableMeta,
  fieldValues: Record<string, CreateFieldValue>,
  nullFields: string[] = [],
): Record<string, unknown> {
  const body: Record<string, unknown> = {};

  for (const field of nullFields) {
    const col = findColumn(tbl, field);
    if (!col) continue;
    if (isLookupLike(col)) {
      const target = col.targets?.[0] ?? '';
      const bindProp = target ? bindPropertyFor(tbl, col, target) : `${col.logicalName}@odata.bind`;
      body[bindProp] = null;
    } else {
      body[col.logicalName] = null;
    }
  }

  for (const [field, raw] of Object.entries(fieldValues)) {
    if (raw == null) continue;
    if (nullFields.includes(field)) continue; // null wins; UI prevents this
    const col = findColumn(tbl, field);
    if (!col) continue;

    if (isLookupLike(col)) {
      const lk = raw as LookupFieldValue;
      if (!lk?.id) continue;
      // The target's entity set comes from its metadata when loaded, else
      // from the set name remembered when the value was read out of a JSON
      // body (the target's metadata may never have been fetched).
      const entitySet = findTable(lk.targetEntity)?.entitySetName ?? lk.entitySet;
      if (!entitySet) continue;
      // Polymorphic-safe: resolves to e.g. `regardingobjectid_account_task`
      // for activity → account lookups, or `customerid_contact` for
      // contact-Customer lookups. Falls back to bare attribute name for
      // single-target Lookups (same string).
      const bindProp = bindPropertyFor(tbl, col, lk.targetEntity);
      body[bindProp] = `/${entitySet}(${lk.id})`;
      continue;
    }
    if (col.attributeType === 'MultiSelectPicklist') {
      const arr = raw as number[];
      if (!Array.isArray(arr) || arr.length === 0) continue;
      body[col.logicalName] = arr.join(',');
      continue;
    }
    if (typeof raw === 'string' && raw === '') continue;
    body[col.logicalName] = raw;
  }
  return body;
}

// ── Raw JSON body ──────────────────────────────────────────────────────

export type RawBodyParse =
  | { ok: true; body: Record<string, unknown> }
  | { ok: false; error: string };

/** Parse the JSON-mode body text. The body must be a JSON object. */
export function parseRawBody(text: string): RawBodyParse {
  if (!text.trim()) return { ok: false, error: 'The body is empty. Use {} for an empty body.' };
  let parsed: unknown;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : 'Invalid JSON.' };
  }
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    return { ok: false, error: 'The body must be a JSON object ({ … }).' };
  }
  return { ok: true, body: parsed as Record<string, unknown> };
}

// ── JSON → form ────────────────────────────────────────────────────────

/** A body key the form view can't represent, with the reason. */
export interface UnmappedBodyKey {
  key: string;
  reason: string;
}

export interface DecodedFieldBody {
  values: Record<string, CreateFieldValue>;
  nullFields: string[];
  /** Keys that would be lost if the body were rebuilt from the form. */
  unmapped: UnmappedBodyKey[];
}

export interface DecodeOptions {
  /** Columns the form shows. Anything else is reported, not mapped. */
  editable?: (logicalName: string) => boolean;
  /** False when the caller has no explicit-null support — nulls are reported. */
  nullSupported?: boolean;
}

const GUID_RE =
  /^[{]?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}[}]?$/;

/** `/contacts(<guid>)`, `contacts(<guid>)` or an absolute URL ending in one. */
function parseBindTarget(v: string): { entitySet: string; id: string } | null {
  const m = v.trim().match(/(?:^|\/)([A-Za-z_][\w]*)\(([^)]*)\)\/?$/);
  return m ? { entitySet: m[1], id: m[2] } : null;
}

/**
 * Pick the lookup target an entity set refers to. Prefers loaded metadata;
 * without it, a single-target column has only one answer, and for
 * polymorphic columns we accept the plural form of a target's logical name
 * (contacts → contact, opportunities → opportunity). Returns undefined when
 * the target can't be told apart.
 */
function resolveTarget(targets: string[], entitySet: string): string | undefined {
  const loaded = targets.find((t) => findTable(t)?.entitySetName === entitySet);
  if (loaded) return loaded;
  // A loaded target whose entity set differs isn't the one referenced.
  const unloaded = targets.filter((t) => !findTable(t));
  if (targets.length === 1) return unloaded[0];
  const lower = entitySet.toLowerCase();
  return unloaded.find(
    (t) =>
      `${t}s` === lower ||
      `${t}es` === lower ||
      (t.endsWith('y') && `${t.slice(0, -1)}ies` === lower),
  );
}

// The JSON types `decodeFieldBody` maps, per attribute type. The JSON view's
// schema (engine/bodySchema) uses the same sets, so its warnings agree.
export const NUMBER_TYPES = new Set(['Integer', 'BigInt', 'Decimal', 'Double', 'Money']);
export const INT_CHOICE_TYPES = new Set(['Picklist', 'State', 'Status']);
export const STRING_TYPES = new Set(['String', 'Memo', 'Uniqueidentifier', 'DateTime']);

const typeOfJson = (v: unknown): string =>
  Array.isArray(v) ? 'an array' : v === null ? 'null' : `a ${typeof v}`;

/** Find the lookup column + (optional) target a `<prop>@odata.bind` key names. */
function resolveBindProperty(
  tbl: TableMeta,
  propName: string,
): { column: string; targets: string[]; navTarget?: string } | undefined {
  const navs = tbl.navigationProperties.filter((n) => n.cardinality === 'ManyToOne');
  const nav =
    navs.find((n) => n.name === propName) ??
    navs.find((n) => n.name.toLowerCase() === propName.toLowerCase());
  if (nav?.referencingAttribute) {
    const col = findColumn(tbl, nav.referencingAttribute);
    return {
      column: nav.referencingAttribute,
      targets: col && isLookupLike(col) ? col.targets : [nav.targetEntity],
      navTarget: nav.targetEntity,
    };
  }
  const col = findColumn(tbl, propName) ?? findColumn(tbl, propName.toLowerCase());
  if (col && isLookupLike(col)) return { column: col.logicalName, targets: col.targets };
  return undefined;
}

/**
 * Map a user-typed JSON body back onto fieldValues + nullFields. Everything
 * that can't be shown in the form (annotations, unknown columns, values of
 * the wrong JSON type, lookups whose target can't be resolved, …) is
 * reported in `unmapped` rather than dropped silently.
 */
export function decodeFieldBody(
  tbl: TableMeta,
  body: Record<string, unknown>,
  opts: DecodeOptions = {},
): DecodedFieldBody {
  const { editable, nullSupported = true } = opts;
  const values: Record<string, CreateFieldValue> = {};
  const nullFields: string[] = [];
  const unmapped: UnmappedBodyKey[] = [];
  const skip = (key: string, reason: string) => unmapped.push({ key, reason });
  const notEditable = (logical: string) => editable && !editable(logical);
  const setNull = (key: string, logical: string) => {
    if (!nullSupported) return skip(key, 'explicit null isn’t supported here');
    if (!nullFields.includes(logical)) nullFields.push(logical);
    delete values[logical];
  };

  for (const [key, v] of Object.entries(body)) {
    if (key.endsWith('@odata.bind')) {
      const propName = key.slice(0, -'@odata.bind'.length);
      const ref = resolveBindProperty(tbl, propName);
      if (!ref) {
        skip(key, `no lookup named ${propName} on ${tbl.logicalName}`);
        continue;
      }
      if (notEditable(ref.column)) {
        skip(key, `${ref.column} isn’t settable here, so the form hides it`);
        continue;
      }
      if (v === null) {
        setNull(key, ref.column);
        continue;
      }
      const target = typeof v === 'string' ? parseBindTarget(v) : null;
      if (!target) {
        skip(key, 'expected "/<entitySet>(<guid>)"');
        continue;
      }
      if (!GUID_RE.test(target.id)) {
        skip(key, 'the form only picks records by GUID (alternate-key references stay in JSON)');
        continue;
      }
      const targetEntity = resolveTarget(
        ref.navTarget ? [ref.navTarget] : ref.targets,
        target.entitySet,
      );
      if (!targetEntity) {
        skip(
          key,
          `“${target.entitySet}” isn't a known entity set for ${ref.navTarget ?? ref.targets.join(' / ')} (or its metadata isn't loaded)`,
        );
        continue;
      }
      values[ref.column] = {
        id: target.id.replace(/[{}]/g, ''),
        targetEntity,
        entitySet: target.entitySet,
      };
      continue;
    }

    if (key.includes('@')) {
      skip(key, 'annotation — the form has no control for it');
      continue;
    }
    const col = findColumn(tbl, key);
    if (!col) {
      skip(key, `not a column on ${tbl.logicalName}`);
      continue;
    }
    if (notEditable(col.logicalName)) {
      skip(key, `${col.logicalName} isn’t settable here, so the form hides it`);
      continue;
    }
    if (v === null) {
      setNull(key, col.logicalName);
      continue;
    }
    if (isLookupLike(col)) {
      skip(key, `lookups are set with "<nav>@odata.bind", not "${key}"`);
      continue;
    }
    const t = col.attributeType;
    if (t === 'MultiSelectPicklist') {
      const parts =
        typeof v === 'string'
          ? v.split(',').filter((p) => p.trim() !== '')
          : Array.isArray(v)
            ? v
            : null;
      const nums = parts?.map((p) => Number(typeof p === 'string' ? p.trim() : p));
      if (!nums || nums.some((n) => !Number.isInteger(n))) {
        skip(key, 'expected a comma-separated string of option values, e.g. "1,2"');
        continue;
      }
      values[col.logicalName] = nums;
      continue;
    }
    const ok = NUMBER_TYPES.has(t)
      ? typeof v === 'number'
      : INT_CHOICE_TYPES.has(t)
        ? Number.isInteger(v)
        : t === 'Boolean'
          ? typeof v === 'boolean'
          : STRING_TYPES.has(t)
            ? typeof v === 'string'
            : t === 'EntityName'
              ? typeof v === 'string' || typeof v === 'number'
              : false;
    if (!ok) {
      const want = NUMBER_TYPES.has(t)
        ? 'a number'
        : INT_CHOICE_TYPES.has(t)
          ? 'an integer option value'
          : t === 'Boolean'
            ? 'true or false'
            : STRING_TYPES.has(t)
              ? 'a string'
              : 'a value the form supports';
      skip(key, `${t} column — expected ${want}, got ${typeOfJson(v)}`);
      continue;
    }
    if (nullFields.includes(col.logicalName)) {
      nullFields.splice(nullFields.indexOf(col.logicalName), 1);
    }
    values[col.logicalName] = v as CreateFieldValue;
  }
  return { values, nullFields, unmapped };
}

/**
 * Columns a JSON body gives a non-empty value — lookups via their
 * `@odata.bind` key, whether or not the target can be resolved. Used by the
 * modes' required-field check while the field set is in JSON mode.
 */
export function columnsSetByBody(tbl: TableMeta, body: Record<string, unknown>): Set<string> {
  const out = new Set<string>();
  for (const [key, v] of Object.entries(body)) {
    if (v === null || v === undefined || v === '') continue;
    if (key.endsWith('@odata.bind')) {
      const ref = resolveBindProperty(tbl, key.slice(0, -'@odata.bind'.length));
      if (ref) out.add(ref.column);
      continue;
    }
    if (findColumn(tbl, key)) out.add(key);
  }
  return out;
}

/**
 * Required-column check shared by the modes while the field set is in JSON
 * mode. Returns the set of columns the typed body sets, or null when the
 * body isn't in JSON mode (callers then check fieldValues as before). An
 * unparseable body sets nothing.
 */
export function rawBodyColumns(tbl: TableMeta, rawBody: string | null | undefined) {
  if (rawBody == null) return null;
  const p = parseRawBody(rawBody);
  return p.ok ? columnsSetByBody(tbl, p.body) : new Set<string>();
}

/**
 * The column values a write will actually send, for the Update Diff pane:
 * the form's fieldValues, or — in JSON mode, where fieldValues are stale —
 * the typed body decoded back to columns. Explicit nulls are folded in as
 * `null` values (null wins, as in encodeFieldBody). `unmapped` lists body
 * keys that couldn't be mapped to a column; `error` is set when the typed
 * body doesn't parse (values is then empty).
 */
export function sentFieldValues(
  tbl: TableMeta | undefined,
  rawBody: string | null | undefined,
  fieldValues: Record<string, CreateFieldValue>,
  nullFields: string[],
): {
  values: Record<string, CreateFieldValue>;
  fromJson: boolean;
  unmapped: UnmappedBodyKey[];
  error?: string;
} {
  const withNulls = (values: Record<string, CreateFieldValue>, nulls: string[]) => {
    const out = { ...values };
    for (const n of nulls) out[n] = null;
    return out;
  };
  if (rawBody == null) {
    return { values: withNulls(fieldValues, nullFields), fromJson: false, unmapped: [] };
  }
  const p = parseRawBody(rawBody);
  if (!p.ok) return { values: {}, fromJson: true, unmapped: [], error: p.error };
  if (!tbl) return { values: {}, fromJson: true, unmapped: [] };
  const d = decodeFieldBody(tbl, p.body);
  return { values: withNulls(d.values, d.nullFields), fromJson: true, unmapped: d.unmapped };
}

/** Execute-blocking message for an invalid JSON-mode body, or null. */
export function rawBodyBlocker(rawBody: string | null | undefined): string | null {
  if (rawBody == null) return null;
  const p = parseRawBody(rawBody);
  return p.ok ? null : `The JSON body is invalid: ${p.error}`;
}

/**
 * The required-column rule for a create (#124). One definition shared by the
 * Create / Upsert Execute gate and the field-set editor's warning and badge,
 * in Form and JSON mode alike:
 *   - Only ApplicationRequired columns valid for create must be supplied.
 *     SystemRequired columns (ownerid, statecode, …) are filled in by
 *     Dataverse when absent, so they never warn.
 *   - A column is supplied when what will be sent gives it a non-empty
 *     value. JSON mode: the typed body sets it (see columnsSetByBody — the
 *     body is sent as typed, so a key the form can't show still counts).
 *     Form mode: a value that isn't null, '' or an empty list, and a lookup
 *     needs a record id (a seeded empty lookup is dropped from the body).
 *     An explicit null never counts.
 */
/**
 * Whether the form treats a column as required. Columns Advanced Find hides
 * (`owneridtype` and other system columns) don't count: the attribute load
 * used to leave them out, and the server fills them in (#44).
 */
export const isRequiredColumn = (c: ColumnMeta): boolean =>
  !!c.required && c.isValidForAdvancedFind !== false;

export function mustSupplyOnCreate(c: ColumnMeta): boolean {
  return (
    c.requiredLevel === 'ApplicationRequired' &&
    c.isValidForAdvancedFind !== false &&
    c.isValidForCreate !== false &&
    // File / Image columns are uploaded separately and hidden from the field
    // set, so they can't block a create (#305).
    c.attributeType !== 'File' &&
    c.attributeType !== 'Image'
  );
}

function isFieldValueSet(c: ColumnMeta, v: CreateFieldValue | undefined): boolean {
  if (v == null) return false;
  if (typeof v === 'string') return v !== '';
  if (Array.isArray(v)) return v.length > 0;
  if (isLookupLike(c)) return !!(v as LookupFieldValue | null)?.id;
  return true;
}

/**
 * Whether what will be sent gives a column a non-empty value — see
 * mustSupplyOnCreate. `rawBody` is the JSON-mode text, or null in Form
 * mode. Explicit nulls (Form mode's `nullFields`) are the caller's check.
 */
export function suppliedColumnTest(
  tbl: TableMeta,
  rawBody: string | null | undefined,
  fieldValues: Record<string, CreateFieldValue>,
): (c: ColumnMeta) => boolean {
  const inJson = rawBodyColumns(tbl, rawBody);
  return (c) =>
    inJson ? inJson.has(c.logicalName) : isFieldValueSet(c, fieldValues[c.logicalName]);
}

/** The columns (of `columns`) a create must supply that the body leaves unset. */
export function missingRequiredColumns(
  tbl: TableMeta,
  columns: ColumnMeta[],
  rawBody: string | null | undefined,
  fieldValues: Record<string, CreateFieldValue>,
  nullFields: string[],
): ColumnMeta[] {
  const supplied = suppliedColumnTest(tbl, rawBody, fieldValues);
  return columns
    .filter(mustSupplyOnCreate)
    .filter((c) => !supplied(c) || nullFields.includes(c.logicalName));
}
