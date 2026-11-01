// URL + body construction for the Execute group.
//
// Three modes:
//   • Execute Action      — POST /{action} (unbound), POST /<set>(<id>)/Microsoft.Dynamics.CRM.<action>
//                           (entity-bound) or POST /<set>/Microsoft.Dynamics.CRM.<action> (collection-bound)
//   • Execute Function    — GET  /{function}(p1=@p1)?@p1=value (same three binding shapes)
//   • Execute Workflow    — POST /workflows(<wf-id>)/Microsoft.Dynamics.CRM.ExecuteWorkflow
//
// Per docs:
//   - Bound action/function: must include the Microsoft.Dynamics.CRM
//     namespace in the URL segment, otherwise "Status Code: 400 Request
//     message has unresolved parameters".
//   - Unbound: just the bare name (no namespace prefix on the URL).
//   - Function params: prefer parameter aliases — GET /F(p1=@p1)?@p1='x' —
//     to avoid URL-length issues + DateTimeOffset bugs.
//   - Action params with ambiguous entity type need @odata.type discriminator.
//
// References:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-web-api-actions
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-web-api-functions

import { ENV } from '../mock/environment';
import { findTable } from '../mock/metadata';
import { type CsdlAction, type ActionParam } from '../mock/actionsCsdl';
// LIVE-AWARE action lookup. The mock `findAction` only sees the seeded
// ~30 actions in `mock/actionsCsdl.ts` — in PPTB the real action set has
// 1000+ entries from the parsed `$metadata` document. We use the cached
// sync accessor from csdlProvider so the builders return correct URL +
// body for any live action, not just mock ones.
//
// Without this swap, every non-mock action (QualifyLead, every Custom
// API, etc.) would return an empty BuiltRequest → broken code-gen
// snippets and missing body previews.
import { findActionSync } from '../host/csdlProvider';
import { ISO_DATE_RE } from '../primitives/urlParts';
import { isRecordGuid } from '../primitives/guid';
import type { BuiltRequest } from './urlBuilder';
import type {
  ExecuteActionState,
  ExecuteFunctionState,
  ExecuteWorkflowState,
} from '../state/executeState';
import type { CreateFieldValue } from '../state/writeState';

// ──────────────────────────────────────────────────────────────
// Execute Action
// ──────────────────────────────────────────────────────────────

export function buildExecuteAction(s: ExecuteActionState): BuiltRequest {
  const action = s.actionName ? findActionSync(s.actionName) : undefined;
  if (!action || action.kind !== 'Action') {
    return {
      relativeUrl: '',
      relativeNoBase: '',
      bytes: 0,
      queryParts: [],
      entitySet: '',
      entityLogical: '',
    };
  }

  const { segment, entitySet, entityLogical } = buildActionUrlSegment(action, s.boundRecordId);
  const path = `${ENV.apiBase}${segment}`;
  return {
    relativeUrl: path,
    relativeNoBase: segment,
    bytes: path.length,
    queryParts: [],
    entitySet,
    entityLogical,
    recordId: s.boundRecordId ?? undefined,
  };
}

export function buildExecuteActionBody(s: ExecuteActionState): Record<string, unknown> {
  const action = s.actionName ? findActionSync(s.actionName) : undefined;
  if (!action) return {};
  const body: Record<string, unknown> = {};
  for (const param of action.parameters) {
    const v = s.paramValues[param.name];
    if (v === undefined || v === null) {
      if (param.required) {
        // Surface required-but-unset as a placeholder so the user sees the shape
        body[param.name] = paramPlaceholder(param);
      }
      continue;
    }
    body[param.name] = encodeParamForBody(param, v);
  }
  return body;
}

// ──────────────────────────────────────────────────────────────
// Execute Function
// ──────────────────────────────────────────────────────────────

export function buildExecuteFunction(s: ExecuteFunctionState): BuiltRequest {
  const fn = s.functionName ? findActionSync(s.functionName) : undefined;
  if (!fn || fn.kind !== 'Function') {
    return {
      relativeUrl: '',
      relativeNoBase: '',
      bytes: 0,
      queryParts: [],
      entitySet: '',
      entityLogical: '',
    };
  }

  const { segment, queryParts, entitySet, entityLogical } = composeFunctionRequest(
    s,
    fn,
    /*wire*/ false,
  );
  const qs = queryParts.map((p) => `${p.key}=${p.value}`).join('&');
  const path = `${ENV.apiBase}${segment}${qs ? `?${qs}` : ''}`;
  const noBase = `${segment}${qs ? `?${qs}` : ''}`;
  return {
    relativeUrl: path,
    relativeNoBase: noBase,
    bytes: path.length,
    queryParts,
    entitySet,
    entityLogical,
    recordId: s.boundRecordId ?? undefined,
  };
}

/**
 * The path the live executor sends for a function call, relative to the Web
 * API base and without a leading slash (the shape `dataverseAPI.queryData`
 * takes). Built by the SAME composer as `buildExecuteFunction`, so what runs
 * is what the URL bar, URL preview and Code tab show (#55) — including the
 * parameter-alias toggle and the real entity set name for bound calls. The
 * only difference is percent-encoding of characters that would otherwise
 * change the URL's meaning on the wire.
 *
 * Returns null when the function isn't in the CSDL cache or a bound
 * function's entity set isn't resolved yet (table metadata not loaded).
 */
export function buildExecuteFunctionWirePath(s: ExecuteFunctionState): string | null {
  const fn = s.functionName ? findActionSync(s.functionName) : undefined;
  if (!fn || fn.kind !== 'Function') return null;
  const { segment, queryParts, entitySet } = composeFunctionRequest(s, fn, /*wire*/ true);
  if (fn.binding.kind !== 'unbound' && !entitySet) return null;
  if (fn.binding.kind === 'entity' && !s.boundRecordId) return null;
  const qs = queryParts.map((p) => `${p.key}=${p.value}`).join('&');
  return `${segment.replace(/^\//, '')}${qs ? `?${qs}` : ''}`;
}

/**
 * Shared function-URL composer. `wire` only switches value escaping:
 *   • alias values (query string) → on the wire, encodeURIComponent. In the
 *     preview (URL bar / Copy / Code tab) only `%`, `&`, `#` and `+` are
 *     escaped — the characters that would split the query, start a fragment,
 *     read as a space or start an escape — so a copied URL still means the
 *     same thing while quotes, brackets, spaces and `:` stay readable like
 *     the MS Learn examples (#61).
 *   • inline literals (path) → escape `%`, `#`, `?` and `\` in both forms,
 *     which would otherwise end the path or be rewritten by the URL parser.
 *     Other characters stay literal (aliases are the documented answer for
 *     awkward values).
 */
function composeFunctionRequest(
  s: ExecuteFunctionState,
  fn: CsdlAction,
  wire: boolean,
): {
  segment: string;
  queryParts: { key: string; value: string }[];
  entitySet: string;
  entityLogical: string;
} {
  const {
    segment: baseSegment,
    entitySet,
    entityLogical,
  } = buildActionUrlSegment(fn, s.boundRecordId);

  const escAlias = (v: string) =>
    wire ? encodeURIComponent(v) : v.replace(/[%&#+]/g, (c) => encodeURIComponent(c));
  const escInline = (v: string) => v.replace(/[%#?\\]/g, (c) => encodeURIComponent(c));

  // Compose the parameter portion: either inline literals or alias placeholders
  const namedParams = fn.parameters;
  let segment = baseSegment;
  const queryParts: { key: string; value: string }[] = [];

  if (namedParams.length > 0) {
    if (s.useParamAliases) {
      // Inline aliases — @p1, @p2, etc.
      const aliasPairs: string[] = [];
      let aliasIndex = 1;
      for (const param of namedParams) {
        const v = s.paramValues[param.name];
        if (v === undefined || v === null) continue;
        const alias = `@p${aliasIndex++}`;
        aliasPairs.push(`${param.name}=${alias}`);
        queryParts.push({
          key: alias,
          value: escAlias(encodeParamForUrl(param, v, /*aliased*/ true)),
        });
      }
      if (aliasPairs.length > 0) {
        segment = `${baseSegment}(${aliasPairs.join(',')})`;
      } else {
        segment = `${baseSegment}()`;
      }
    } else {
      // Inline literals — `(p1='val',p2=42)`
      const inlinePairs: string[] = [];
      for (const param of namedParams) {
        const v = s.paramValues[param.name];
        if (v === undefined || v === null) continue;
        inlinePairs.push(
          `${param.name}=${escInline(encodeParamForUrl(param, v, /*aliased*/ false))}`,
        );
      }
      segment =
        inlinePairs.length > 0 ? `${baseSegment}(${inlinePairs.join(',')})` : `${baseSegment}()`;
    }
  } else {
    // Parameterless functions still need `()` on the wire (WhoAmI()).
    // Dataverse accepts the bare name too, but `()` is the canonical form.
    segment = `${baseSegment}()`;
  }

  return { segment, queryParts, entitySet, entityLogical };
}

// ──────────────────────────────────────────────────────────────
// Execute Workflow
// ──────────────────────────────────────────────────────────────

export function buildExecuteWorkflow(s: ExecuteWorkflowState): BuiltRequest {
  const wfId = s.workflowId ?? '<workflow-id>';
  const segment = `/workflows(${wfId})/Microsoft.Dynamics.CRM.ExecuteWorkflow`;
  const path = `${ENV.apiBase}${segment}`;
  return {
    relativeUrl: path,
    relativeNoBase: segment,
    bytes: path.length,
    queryParts: [],
    entitySet: 'workflows',
    entityLogical: 'workflow',
    recordId: s.workflowId ?? undefined,
  };
}

export function buildExecuteWorkflowBody(s: ExecuteWorkflowState): Record<string, unknown> {
  return { EntityId: s.entityId ?? '<target-record-guid>' };
}

// ──────────────────────────────────────────────────────────────
// URL segment helpers
// ──────────────────────────────────────────────────────────────

function buildActionUrlSegment(
  action: CsdlAction,
  boundRecordId: string | null,
): { segment: string; entitySet: string; entityLogical: string } {
  const fullName = `${action.namespace}.${action.name}`;
  switch (action.binding.kind) {
    case 'unbound': {
      // Unbound — no namespace prefix on the URL, just the bare name.
      return {
        segment: `/${action.name}`,
        entitySet: '',
        entityLogical: '',
      };
    }
    // Bound but the table's metadata isn't loaded yet: show a placeholder
    // entity set rather than the bare (unbound) name, which would be a
    // different — usually non-existent — operation (#55). `entitySet` stays
    // empty so the modes / executor can tell the URL isn't ready.
    case 'entity': {
      const tbl = findTable(action.binding.entityType);
      if (!tbl) {
        const id = boundRecordId ?? `<${action.binding.entityType}-id>`;
        return {
          segment: `/<${action.binding.entityType}-set>(${id})/${fullName}`,
          entitySet: '',
          entityLogical: action.binding.entityType,
        };
      }
      const id = boundRecordId ?? `<${tbl.logicalName}-id>`;
      return {
        segment: `/${tbl.entitySetName}(${id})/${fullName}`,
        entitySet: tbl.entitySetName,
        entityLogical: tbl.logicalName,
      };
    }
    case 'collection': {
      const tbl = findTable(action.binding.entityType);
      if (!tbl) {
        return {
          segment: `/<${action.binding.entityType}-set>/${fullName}`,
          entitySet: '',
          entityLogical: action.binding.entityType,
        };
      }
      return {
        segment: `/${tbl.entitySetName}/${fullName}`,
        entitySet: tbl.entitySetName,
        entityLogical: tbl.logicalName,
      };
    }
  }
}

// ──────────────────────────────────────────────────────────────
// Parameter encoding
// ──────────────────────────────────────────────────────────────

/**
 * Encode a parameter value for inclusion in a POST body. The shape varies by
 * EDM type:
 *
 *   • Edm.* primitives → JSON-native (string/number/bool)
 *   • EntityReference  → `{ @odata.type: 'Microsoft.Dynamics.CRM.<entity>', <pk>: '<guid>' }`
 *   • EntitySpecific   → same as EntityReference but @odata.type is optional
 *                        (we include it for clarity)
 *   • EntityCollection → array of EntityReferences
 *   • Collection(...)  → array, JSON-native
 *   • ComplexType      → nested object — recursively encoded
 *   • OptionSetValue   → integer
 */
// ── EntitySpecific: existing record vs new instance (#1) ──────────
//
// An EntitySpecific parameter (`mscrm.<entity>`) takes either a reference to
// an existing row or a new entity instance built inline; the CSDL doesn't say
// which. Instance mode stores a field-value map keyed by column logical names
// (never starting with '@'); reference mode stores `{ '@ref': '<guid>' }`, so
// the two shapes can't collide. On the wire a reference is the same
// `@odata.type` + primary-key object Dataverse documents for action entity
// parameters (AddToQueue, Merge), and `{'@odata.id':'<set>(<id>)'}` in a
// function URL.

export interface EntitySpecificRef {
  '@ref': string;
}

export function isEntitySpecificRef(v: unknown): v is EntitySpecificRef {
  return (
    !!v &&
    typeof v === 'object' &&
    !Array.isArray(v) &&
    typeof (v as Record<string, unknown>)['@ref'] === 'string'
  );
}

/** Parameter names that almost always point at an existing row. */
const REFERENCE_PARAM_RE =
  /^(target|source|subject|owner|assignee|subordinate|record|parent|queue|destination|principal|user|team)/i;

/** Default mode for an EntitySpecific param the user hasn't touched yet. */
export function defaultEntitySpecificMode(param: ActionParam): 'reference' | 'instance' {
  return REFERENCE_PARAM_RE.test(param.name) ? 'reference' : 'instance';
}

function entityRefBody(entityType: string, id: string): Record<string, unknown> {
  const tbl = findTable(entityType);
  const pk = tbl?.primaryKey ?? `${entityType}id`;
  return { '@odata.type': `Microsoft.Dynamics.CRM.${entityType}`, [pk]: id };
}

/** `/accounts(<guid>)`, `accounts(<guid>)` or an absolute URL ending in one → set + id. */
function readRowPath(s: string): { set: string; id: string } | null {
  const m = /(?:^|\/)([A-Za-z_][A-Za-z0-9_]*)\(\{?([0-9a-fA-F-]{36})\}?\)$/.exec(s.trim());
  return m && isRecordGuid(m[2]) ? { set: m[1], id: m[2].toLowerCase() } : null;
}

const typeOf = (v: Record<string, unknown>) =>
  typeof v['@odata.type'] === 'string'
    ? (v['@odata.type'] as string).replace(/^#?Microsoft\.Dynamics\.CRM\./i, '')
    : undefined;

/**
 * The parameter value the Execute Action form stores for a value as it goes
 * on the wire — the inverse of `encodeParamForBody`, for a pasted request
 * body (cURL import, #271). `resolveSet` maps an entity set name to its
 * table. Returns an error when the value can't be read as that type, rather
 * than letting the encoder turn it into null.
 */
export function decodeParamFromBody(
  param: ActionParam,
  wire: unknown,
  resolveSet: (set: string) => string | undefined,
): { value: unknown } | { error: string } {
  const bad = (what: string) => ({ error: `${param.name}: expected ${what}.` });
  if (wire === null) return { value: null };
  const obj =
    wire && typeof wire === 'object' && !Array.isArray(wire)
      ? (wire as Record<string, unknown>)
      : null;
  // A reference as `{@odata.type, <pk>}` or `{@odata.id: '<set>(<id>)'}`.
  const readRef = (v: unknown, fallbackType?: string) => {
    if (!v || typeof v !== 'object' || Array.isArray(v)) return null;
    const o = v as Record<string, unknown>;
    if (typeof o['@odata.id'] === 'string') {
      const row = readRowPath(o['@odata.id']);
      const entityType = row && resolveSet(row.set);
      return row && entityType ? { id: row.id, entityType } : null;
    }
    const entityType = typeOf(o) ?? fallbackType;
    if (!entityType) return null;
    const pk = findTable(entityType)?.primaryKey ?? `${entityType}id`;
    const id = o[pk];
    return typeof id === 'string' && isRecordGuid(id) ? { id: id.toLowerCase(), entityType } : null;
  };
  switch (param.type) {
    case 'Edm.String':
    case 'Edm.Guid':
    case 'Edm.DateTimeOffset':
    case 'Edm.Binary':
      return typeof wire === 'string' ? { value: wire } : bad('a string');
    case 'Edm.Int32':
    case 'Edm.Int64':
    case 'Edm.Decimal':
    case 'Edm.Double':
    case 'OptionSetValue':
      return typeof wire === 'number' ? { value: wire } : bad('a number');
    case 'Edm.Boolean':
      return typeof wire === 'boolean' ? { value: wire } : bad('true or false');
    case 'Collection(Edm.String)':
    case 'Collection(Edm.Int32)':
    case 'Collection(Edm.Guid)':
      return Array.isArray(wire) ? { value: wire } : bad('an array');
    case 'EntityReference': {
      const ref = readRef(wire, param.entityType);
      return ref ? { value: ref } : bad('a row reference ({"@odata.type", "<id column>"})');
    }
    case 'EntityCollection': {
      if (!Array.isArray(wire)) return bad('an array of row references');
      const refs = wire.map((x) => readRef(x, param.entityType));
      return refs.every(Boolean) ? { value: refs } : bad('an array of row references');
    }
    case 'EntitySpecific': {
      if (!obj || !param.entityType) return bad('an object');
      const tbl = findTable(param.entityType);
      const pk = tbl?.primaryKey ?? `${param.entityType}id`;
      const keys = Object.keys(obj).filter((k) => k !== '@odata.type');
      // Just the row's id (or an @odata.id) → a reference to an existing row.
      if (keys.length === 1 && (keys[0] === pk || keys[0] === '@odata.id')) {
        const ref = readRef(obj, param.entityType);
        return ref ? { value: { '@ref': ref.id } } : bad('a row reference');
      }
      // Otherwise a new row, as a field-value map the form can show.
      const fields: Record<string, unknown> = {};
      for (const key of keys) {
        const v = obj[key];
        const bind = /^(.+)@odata\.bind$/.exec(key);
        const col = tbl?.columns.find((c) => c.logicalName === (bind ? bind[1] : key));
        if (!col) return { error: `${param.name}: ${param.entityType} has no column "${key}".` };
        if (bind) {
          const row = typeof v === 'string' ? readRowPath(v) : null;
          const target = row && resolveSet(row.set);
          if (!row || !target) return bad(`"${key}" to be "/<set>(<id>)"`);
          fields[col.logicalName] = { id: row.id, targetEntity: target, entitySet: row.set };
        } else if (col.attributeType === 'MultiSelectPicklist' && typeof v === 'string') {
          fields[col.logicalName] = v.split(',').map((n) => Number(n.trim()));
        } else fields[col.logicalName] = v;
      }
      return { value: fields };
    }
    case 'ComplexType': {
      if (!obj) return bad('an object');
      if (!param.complexType) return { value: obj };
      const out: Record<string, unknown> = {};
      for (const [k, v] of Object.entries(obj)) {
        const field = param.complexType.fields.find((f) => f.name === k);
        if (!field)
          return { error: `${param.name}: ${param.complexType.name} has no field "${k}".` };
        const r = decodeParamFromBody(field, v, resolveSet);
        if ('error' in r) return r;
        out[k] = r.value;
      }
      return { value: out };
    }
  }
}

export function encodeParamForBody(param: ActionParam, value: unknown): unknown {
  if (value === null || value === undefined) return null;
  switch (param.type) {
    case 'Edm.String':
    case 'Edm.Guid':
      return String(value);
    case 'Edm.Int32':
    case 'Edm.Int64':
    case 'Edm.Decimal':
    case 'Edm.Double':
    case 'OptionSetValue':
      return Number(value);
    case 'Edm.Boolean':
      return value === true || value === 'true' || value === 1;
    case 'Edm.DateTimeOffset':
      return String(value); // ISO 8601 string passes through
    case 'Edm.Binary':
      return String(value); // Base64-encoded — caller's responsibility
    case 'Collection(Edm.String)':
    case 'Collection(Edm.Int32)':
    case 'Collection(Edm.Guid)':
      return Array.isArray(value) ? value : [];
    case 'EntityReference': {
      // EntityReference points to an EXISTING row — encoded as
      // `{ "@odata.type": "Microsoft.Dynamics.CRM.<type>", "<pk>": "<id>" }`.
      const ref = value as { id?: string; entityType?: string };
      if (!ref?.id) return null;
      const entityType = ref.entityType ?? param.entityType ?? '';
      const tbl = findTable(entityType);
      const pk = tbl?.primaryKey ?? `${entityType}id`;
      return {
        '@odata.type': `Microsoft.Dynamics.CRM.${entityType}`,
        [pk]: ref.id,
      };
    }
    case 'EntitySpecific': {
      // Reference mode (#1) — an existing row of the declared type.
      if (isEntitySpecificRef(value)) {
        const id = value['@ref'].trim();
        if (!id || !param.entityType) return null;
        return entityRefBody(param.entityType, id);
      }
      // EntitySpecific is a NEW entity instance constructed inline (e.g.
      // OpportunityClose passed to WinOpportunity). The user has filled in
      // a CreateState-like fieldValues map for the target entity; we encode
      // it using the same rules as buildCreateBody (lookups → @odata.bind,
      // multi-select → comma string, etc.) and prepend @odata.type.
      const entityType = param.entityType ?? '';
      if (!entityType) return null;
      const fieldValues = (value as Record<string, CreateFieldValue>) ?? {};
      const body: Record<string, unknown> = {
        '@odata.type': `Microsoft.Dynamics.CRM.${entityType}`,
      };
      const tbl = findTable(entityType);
      if (!tbl) return body;
      for (const [field, raw] of Object.entries(fieldValues)) {
        if (raw == null) continue;
        const col = tbl.columns.find((c) => c.logicalName === field);
        if (!col) continue;
        // Lookup-like → @odata.bind
        if (
          col.attributeType === 'Lookup' ||
          col.attributeType === 'Customer' ||
          col.attributeType === 'Owner'
        ) {
          const lk = raw as { id?: string; targetEntity?: string; entitySet?: string };
          if (!lk?.id) continue;
          const target = lk.targetEntity ?? col.targets[0];
          // A value read out of a typed body carries its entity set, so the
          // bind survives when the target's metadata isn't loaded (#271).
          const set = findTable(target)?.entitySetName ?? lk.entitySet;
          if (!set) continue;
          body[`${col.logicalName}@odata.bind`] = `/${set}(${lk.id})`;
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
    case 'EntityCollection': {
      const arr = value as Array<{ id: string; entityType?: string }>;
      if (!Array.isArray(arr)) return [];
      return arr.map((ref) => {
        const entityType = ref.entityType ?? param.entityType ?? '';
        const tbl = findTable(entityType);
        const pk = tbl?.primaryKey ?? `${entityType}id`;
        return {
          '@odata.type': `Microsoft.Dynamics.CRM.${entityType}`,
          [pk]: ref.id,
        };
      });
    }
    case 'ComplexType': {
      const obj = value as Record<string, unknown>;
      if (!obj || typeof obj !== 'object') return null;
      // Recursively encode nested fields based on the complex type definition
      const out: Record<string, unknown> = {};
      if (param.complexType) {
        for (const f of param.complexType.fields) {
          if (f.name in obj) out[f.name] = encodeParamForBody(f, obj[f.name]);
        }
      } else {
        Object.assign(out, obj);
      }
      return out;
    }
  }
}

/**
 * Encode a parameter value as an OData literal for a function URL (the same
 * literal works inline and as a parameter-alias value). Unescaped — the
 * composer percent-encodes for the URL position. Per MS Learn ("Use Web API
 * functions", "Web API functions and actions sample"):
 *
 *   string          → 'value'  (single quotes, inner ' doubled)
 *   DateTimeOffset  → 2024-01-01T00:00:00Z  (unquoted)
 *   guid            → bare guid
 *   int/decimal/double → bare number
 *   boolean         → true / false
 *   collection      → JSON array: ["account","contact"], [1,2]
 *   enum            → Microsoft.Dynamics.CRM.EntityFilters'Entity'
 *                     (flags: comma-joined members, 'Entity,Attributes')
 *
 *   • EntityReference in a function: not supported as a function URL param —
 *     use `Target=@tid&@tid={'@odata.id':'/entitySet(<guid>)'}` instead.
 *     We emit the @odata.id form when aliased.
 */
/**
 * Why a function parameter value can't go into the URL as typed, or null
 * (#221): GUIDs and record references must be GUIDs, dates ISO 8601. The
 * encoder quotes such values, so they can't change the call, but Dataverse
 * would reject them.
 */
export function invalidFunctionParamReason(
  params: readonly ActionParam[],
  values: Record<string, unknown>,
): string | null {
  for (const p of params) {
    const v = values[p.name];
    if (v == null || v === '') continue;
    const ids: string[] =
      p.type === 'Edm.Guid'
        ? [String(v).trim()]
        : p.type === 'Collection(Edm.Guid)' && Array.isArray(v)
          ? v.map((x) => String(x).trim())
          : p.type === 'EntityReference'
            ? [String((v as { id?: unknown }).id ?? '').trim()].filter(Boolean)
            : p.type === 'EntitySpecific' && isEntitySpecificRef(v)
              ? [v['@ref'].trim()].filter(Boolean)
              : [];
    const bad = ids.find((id) => !isRecordGuid(id));
    if (bad !== undefined) return `Parameter ${p.name}: "${bad}" isn't a GUID.`;
    if (p.type === 'Edm.DateTimeOffset' && !ISO_DATE_RE.test(String(v).trim())) {
      return `Parameter ${p.name}: "${String(v)}" isn't an ISO 8601 date or date-time.`;
    }
  }
  return null;
}

export function encodeParamForUrl(param: ActionParam, value: unknown, aliased: boolean): string {
  if (value === null || value === undefined) return 'null';
  // Unquoted literals only when the value has the type's shape; anything else
  // is quoted and escaped, so it can't add parameters to the call (#221).
  // invalidFunctionParamReason blocks Execute on those.
  const quoted = (s: string) => `'${s.replace(/'/g, "''")}'`;
  switch (param.type) {
    case 'Edm.String':
      return quoted(String(value));
    case 'Edm.DateTimeOffset': {
      const s = String(value).trim();
      return s === '' ? 'null' : ISO_DATE_RE.test(s) ? s : quoted(s);
    }
    case 'Edm.Guid': {
      const s = String(value).trim();
      return isRecordGuid(s) ? s : quoted(s);
    }
    case 'Edm.Int32':
    case 'Edm.Int64':
    case 'Edm.Decimal':
    case 'Edm.Double':
      return String(Number(value));
    case 'OptionSetValue':
      return enumLiteral(param, value);
    case 'Edm.Boolean':
      return value === true || value === 'true' || value === 1 ? 'true' : 'false';
    case 'Collection(Edm.String)':
    case 'Collection(Edm.Guid)': {
      const arr = Array.isArray(value) ? value : [];
      return JSON.stringify(arr.map((s) => String(s)));
    }
    case 'Collection(Edm.Int32)': {
      const arr = Array.isArray(value) ? value : [];
      return JSON.stringify(arr.map((n) => Number(n)));
    }
    case 'EntityReference': {
      // Functions accept entity-typed params via @odata.id record-ref syntax.
      // Only valid when aliased (else inline would conflict with OData parser).
      const ref = value as { id?: string; entityType?: string };
      if (!ref?.id || !ref?.entityType) return 'null';
      const tbl = findTable(ref.entityType);
      const set = tbl?.entitySetName ?? `${ref.entityType}s`;
      void aliased; // both forms use the same record-ref shape
      return `{'@odata.id':'${set}(${ref.id.replace(/'/g, "''")})'}`;
    }
    case 'EntitySpecific': {
      // Only the reference form (#1) has a URL encoding.
      if (!isEntitySpecificRef(value) || !value['@ref'].trim() || !param.entityType) return 'null';
      const tbl = findTable(param.entityType);
      const set = tbl?.entitySetName ?? `${param.entityType}s`;
      return `{'@odata.id':'${set}(${value['@ref'].trim().replace(/'/g, "''")})'}`;
    }
    default:
      return String(value);
  }
}

/**
 * OData enum literal `<Namespace.EnumType>'<Member>'` for an OptionSetValue
 * parameter, e.g. `Microsoft.Dynamics.CRM.EndpointAccessType'Default'` (MS
 * Learn: Web API functions and actions sample, RetrieveCurrentOrganization).
 * A flags enum value that isn't a single member is decomposed into its
 * members, comma-joined. Without the enum type name (mock / unresolved CSDL
 * type) the integer is emitted as before; a value that maps to no member is
 * passed as its number inside the quotes, which OData also accepts.
 */
function enumLiteral(param: ActionParam, value: unknown): string {
  const n = Number(value);
  if (!param.enumType || !Number.isFinite(n)) return String(n);
  const members = param.optionSet ?? [];
  let names = members.find((m) => m.value === n)?.label;
  if (!names && param.enumIsFlags && n > 0) {
    // Single-bit members in the value; first name wins when several share a
    // bit (EntityFilters declares both Entity and Default as 1).
    const bits = members.filter(
      (m, i) =>
        m.value > 0 &&
        (m.value & (m.value - 1)) === 0 &&
        (n & m.value) === m.value &&
        members.findIndex((o) => o.value === m.value) === i,
    );
    const covered = bits.reduce((acc, m) => acc | m.value, 0);
    if (bits.length > 0 && covered === n) names = bits.map((m) => m.label).join(',');
  }
  return `${param.enumType}'${names ?? String(n)}'`;
}

// Placeholder used in body preview when a required param is unset
function paramPlaceholder(param: ActionParam): unknown {
  switch (param.type) {
    case 'Edm.String':
      return '<string>';
    case 'Edm.Guid':
      return '<guid>';
    case 'Edm.Int32':
    case 'Edm.Int64':
    case 'Edm.Decimal':
    case 'Edm.Double':
    case 'OptionSetValue':
      return 0;
    case 'Edm.Boolean':
      return false;
    case 'Edm.DateTimeOffset':
      return '<iso-8601-datetime>';
    case 'Collection(Edm.String)':
      return [];
    case 'Collection(Edm.Int32)':
      return [];
    case 'Collection(Edm.Guid)':
      return [];
    case 'EntityReference':
    case 'EntitySpecific':
      return {
        '@odata.type': `Microsoft.Dynamics.CRM.${param.entityType ?? '<entity>'}`,
        [param.entityType ? `${param.entityType}id` : '<entity>id']: '<guid>',
      };
    case 'EntityCollection':
      return [];
    case 'ComplexType':
      return {};
    case 'Edm.Binary':
      return '<base64>';
  }
}
