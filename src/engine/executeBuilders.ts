// URL + body construction for the Execute group.
//
// Three modes:
//   • Execute Action      — POST /{action} (unbound) or POST /<set>(<id>)/Microsoft.Dynamics.CRM.<action> (bound)
//   • Execute Function    — GET  /{function}(p1=@p1)?@p1=value
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

  const {
    segment: baseSegment,
    entitySet,
    entityLogical,
  } = buildActionUrlSegment(fn, s.boundRecordId);

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
        queryParts.push({ key: alias, value: encodeParamForUrl(param, v, /*aliased*/ true) });
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
        inlinePairs.push(`${param.name}=${encodeParamForUrl(param, v, /*aliased*/ false)}`);
      }
      segment =
        inlinePairs.length > 0 ? `${baseSegment}(${inlinePairs.join(',')})` : `${baseSegment}()`;
    }
  }

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
    case 'entity': {
      const tbl = findTable(action.binding.entityType);
      if (!tbl) return { segment: `/${action.name}`, entitySet: '', entityLogical: '' };
      const id = boundRecordId ?? `<${tbl.logicalName}-id>`;
      return {
        segment: `/${tbl.entitySetName}(${id})/${fullName}`,
        entitySet: tbl.entitySetName,
        entityLogical: tbl.logicalName,
      };
    }
    case 'collection': {
      const tbl = findTable(action.binding.entityType);
      if (!tbl) return { segment: `/${action.name}`, entitySet: '', entityLogical: '' };
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
          const lk = raw as { id?: string; targetEntity?: string };
          if (!lk?.id) continue;
          const target = lk.targetEntity ?? col.targets[0];
          const targetTbl = findTable(target);
          if (!targetTbl) continue;
          body[`${col.logicalName}@odata.bind`] = `/${targetTbl.entitySetName}(${lk.id})`;
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
 * Encode a parameter value for inclusion in a function URL.
 *
 *   • OData literal rules (when inline / not aliased):
 *       string  → 'value'  (single quotes, inner ' doubled)
 *       guid    → bare guid
 *       int/decimal/double → bare number
 *       boolean → true / false
 *       collection → ['a','b'] (rare inline; aliases recommended)
 *
 *   • When aliased (preferred):
 *       string  → 'value'  (same — the alias substitution happens server-side)
 *       all others → same form
 *
 *   • EntityReference in a function: not supported as a function URL param —
 *     use `Target=@tid&@tid={'@odata.id':'/entitySet(<guid>)'}` instead.
 *     We emit the @odata.id form when aliased.
 */
export function encodeParamForUrl(param: ActionParam, value: unknown, aliased: boolean): string {
  if (value === null || value === undefined) return 'null';
  switch (param.type) {
    case 'Edm.String':
    case 'Edm.DateTimeOffset': {
      const s = String(value).replace(/'/g, "''");
      return `'${s}'`;
    }
    case 'Edm.Guid':
      return String(value);
    case 'Edm.Int32':
    case 'Edm.Int64':
    case 'Edm.Decimal':
    case 'Edm.Double':
    case 'OptionSetValue':
      return String(Number(value));
    case 'Edm.Boolean':
      return value === true || value === 'true' || value === 1 ? 'true' : 'false';
    case 'Collection(Edm.String)': {
      const arr = (value as string[]) ?? [];
      return `[${arr.map((s) => `'${String(s).replace(/'/g, "''")}'`).join(',')}]`;
    }
    case 'Collection(Edm.Int32)': {
      const arr = (value as number[]) ?? [];
      return `[${arr.map((n) => Number(n)).join(',')}]`;
    }
    case 'Collection(Edm.Guid)': {
      const arr = (value as string[]) ?? [];
      return `[${arr.join(',')}]`;
    }
    case 'EntityReference': {
      // Functions accept entity-typed params via @odata.id record-ref syntax.
      // Only valid when aliased (else inline would conflict with OData parser).
      const ref = value as { id?: string; entityType?: string };
      if (!ref?.id || !ref?.entityType) return 'null';
      const tbl = findTable(ref.entityType);
      const set = tbl?.entitySetName ?? `${ref.entityType}s`;
      void aliased; // both forms use the same record-ref shape
      return `{'@odata.id':'${set}(${ref.id})'}`;
    }
    default:
      return String(value);
  }
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
