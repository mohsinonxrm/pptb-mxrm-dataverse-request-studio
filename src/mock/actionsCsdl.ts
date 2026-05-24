// CSDL action / function type model — Microsoft Dataverse Web API.
//
// PPTB-only: the action / function catalog is fetched from the live
// `$metadata` document via `csdlProvider.fetchCsdlActions()` and cached
// per session. This file used to ship a seeded mock catalog alongside the
// types; that was removed as part of the "live-only, zero mock" sweep so
// nothing in the studio ever falls back to fixture data.
//
// What lives here now: pure type declarations consumed by the parser
// (csdlProvider.ts), the editors (ActionParamForm, ActionPicker), and the
// mode files. Anyone needing the live catalog calls
// `csdlProvider.actions.loadAll()` (async) or `findActionSync(name)` (sync
// cache lookup populated after the first load).
//
// References:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-web-api-actions
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-web-api-functions
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/actions
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/functions

// ──────────────────────────────────────────────────────────────
// Parameter type model — covers the EDM types Dataverse uses
// ──────────────────────────────────────────────────────────────
export type EdmType =
  | 'Edm.String'
  | 'Edm.Int32'
  | 'Edm.Int64'
  | 'Edm.Boolean'
  | 'Edm.DateTimeOffset'
  | 'Edm.Guid'
  | 'Edm.Decimal'
  | 'Edm.Double'
  | 'Edm.Binary'
  // EntityReference — needs @odata.type on the wire (parent type is mscrm.crmbaseentity)
  | 'EntityReference'
  // EntitySpecific — type is fixed (e.g. mscrm.opportunityclose); no @odata.type required but recommended
  | 'EntitySpecific'
  // Collection(mscrm.<entity>) — for batch params like Members in AddMembersTeam
  | 'EntityCollection'
  // Primitive collections — e.g. EntityNames in RetrieveTotalRecordCount
  | 'Collection(Edm.String)'
  | 'Collection(Edm.Int32)'
  | 'Collection(Edm.Guid)'
  // Option set int — e.g. Status params that reference an EnumType in CSDL
  | 'OptionSetValue'
  // Complex type — nested object (e.g. PrincipalAccess in GrantAccess)
  | 'ComplexType';

export interface ActionParam {
  name: string;
  type: EdmType;
  required: boolean;
  /** When type is EntityReference/EntitySpecific/EntityCollection: the target entity logical name. */
  entityType?: string;
  /** When type is OptionSetValue: the option set definition. */
  optionSet?: { value: number; label: string }[];
  /** When type is ComplexType: the complex type's display name + nested fields. */
  complexType?: { name: string; fields: ActionParam[] };
  /** Inline description from the CSDL (when annotations are loaded). */
  description?: string;
}

export interface ActionReturnType {
  /** void = 204, primitive = 200 with scalar, complex = 200 with object, entity/collection per OData. */
  kind: 'void' | 'primitive' | 'complex' | 'entity' | 'collection';
  /** Display name of the return type — e.g. "WhoAmIResponse" or "Edm.Guid". */
  typeName: string;
}

export type ActionBinding =
  | { kind: 'unbound' }
  | { kind: 'entity'; entityType: string }       // bound to a single row
  | { kind: 'collection'; entityType: string };   // bound to the entire entity collection

export interface CsdlAction {
  /** Wire name — what goes in the URL (e.g. "WinOpportunity", "new_AddNoteToContact"). */
  name: string;
  /** Display label — usually same as name but may be friendlier for custom APIs. */
  displayName?: string;
  /** Action verb — Action = POST + side effects, Function = GET + no side effects. */
  kind: 'Action' | 'Function';
  /** OData namespace — always Microsoft.Dynamics.CRM for OOB; custom messages may use a publisher prefix. */
  namespace: 'Microsoft.Dynamics.CRM' | string;
  binding: ActionBinding;
  parameters: ActionParam[];
  returnType: ActionReturnType;
  /** Composable functions can take $select/$filter modifiers on the URL. */
  isComposable?: boolean;
  /** Tag the source so the registry can route exec-action / exec-customapi / exec-customaction correctly. */
  source: 'oob' | 'custom-api' | 'custom-action';
  /** Inline description from the CSDL (annotated builds only). */
  description?: string;
  /**
   * True when the function is one of Dataverse's "query functions" — operators
   * used inside `$filter` (e.g. Last7Days, Between, EqualUserId, Contains).
   * They're declared as `<Function>` in CSDL but aren't meant to be invoked
   * standalone via Execute Function; they only make sense inline in a
   * filter expression. Detected via a structural signature in csdlProvider
   * (IsBound + PropertyName Edm.String + return Edm.Boolean).
   *
   * Reference:
   *   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/queryfunctions
   */
  isQueryFunction?: boolean;
}
