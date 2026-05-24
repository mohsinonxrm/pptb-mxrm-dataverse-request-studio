// State shapes for the Execute group.
//
// Execute-Action / Execute-Custom-API / Execute-Custom-Action all share the
// same conceptual shape (a CSDL action + named params) — the only thing that
// changes per mode is the action-picker filter (oob / custom-api / custom-
// action). One `ExecuteActionState` covers all three.
//
// Execute-Function is structurally similar but uses GET with URL-encoded
// parameter aliases instead of a JSON body.
//
// Execute-Workflow has its own URL pattern: POST /workflows(<id>)/
// Microsoft.Dynamics.CRM.ExecuteWorkflow with body `{ "EntityId": "<guid>" }`.

import type { HeaderItem } from '../editors/HeadersEditor';

// ──────────────────────────────────────────────────────────────
// Param value — a typed value tied back to the CSDL param definition.
// ──────────────────────────────────────────────────────────────
//
// Stored values follow these rules (the wire encoder reads the param type to
// know which shape to expect):
//
//   • Edm.String                       → string
//   • Edm.Int32 / Edm.Int64            → number
//   • Edm.Decimal / Edm.Double         → number
//   • Edm.Boolean                      → boolean
//   • Edm.Guid                         → string (lowercase guid form)
//   • Edm.DateTimeOffset               → ISO 8601 string
//   • OptionSetValue                   → number (option int value)
//   • Collection(Edm.String)           → string[]
//   • Collection(Edm.Int32)            → number[]
//   • Collection(Edm.Guid)             → string[]
//   • EntityReference                  → { id, entityType } — `@odata.type` added at encode time
//   • EntitySpecific                   → { id, entityType } — type known from CSDL, optional discriminator
//   • EntityCollection                 → Array<{ id, entityType }>
//   • ComplexType                      → Record<string, unknown> — nested per the type's fields
//
// `unknown` keeps the state flexible — the encoder enforces the per-type
// invariants at body-build time.
export type ExecParamValue = unknown;

// ──────────────────────────────────────────────────────────────
// Execute Action / Custom API / Custom Action
// ──────────────────────────────────────────────────────────────
export interface ExecuteActionState {
  /** The action name from the CSDL (e.g. "WinOpportunity"). */
  actionName: string | null;
  /** When bound to an entity: the source row GUID. */
  boundRecordId: string | null;
  /**
   * Map of paramName -> typed value. Encoded into the POST body per the
   * param's CSDL type. Unset params are excluded from the body.
   */
  paramValues: Record<string, ExecParamValue>;
  headers: HeaderItem[];
  /**
   * Which mode owns this state — drives the action picker filter and the
   * registry id surfaced in recent runs.
   */
  category: 'oob' | 'custom-api' | 'custom-action';
  dirty: Set<string>;
}

// ──────────────────────────────────────────────────────────────
// Execute Function
// ──────────────────────────────────────────────────────────────
export interface ExecuteFunctionState {
  /** The function name from the CSDL (e.g. "WhoAmI"). */
  functionName: string | null;
  /** When bound: the source row GUID. */
  boundRecordId: string | null;
  /** Map of paramName -> typed value. Encoded into the URL. */
  paramValues: Record<string, ExecParamValue>;
  /**
   * Use parameter aliases (recommended by docs to avoid URL-length issues and
   * DateTimeOffset bugs):
   *
   *   GET /Func(p1=@p1,p2=@p2)?@p1='val'&@p2=42
   *
   * When false, params are inlined:
   *
   *   GET /Func(p1='val',p2=42)
   */
  useParamAliases: boolean;
  headers: HeaderItem[];
  /** Source category — oob or custom-api. (Custom action functions are rare.) */
  category: 'oob' | 'custom-api';
  dirty: Set<string>;
}

// ──────────────────────────────────────────────────────────────
// Execute Workflow
// ──────────────────────────────────────────────────────────────
//
// POST /workflows(<workflow-id>)/Microsoft.Dynamics.CRM.ExecuteWorkflow
//   body: { "EntityId": "<target-guid>" }
//   → 204 No Content
//
// The workflow's primaryEntity in metadata determines which entity set the
// EntityId belongs to — we surface a record picker scoped to that table.

export interface ExecuteWorkflowState {
  /** Workflow definition id (uuid). */
  workflowId: string | null;
  /** Target record GUID — the entity the workflow operates on. */
  entityId: string | null;
  headers: HeaderItem[];
  dirty: Set<string>;
}
