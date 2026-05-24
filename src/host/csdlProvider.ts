// CSDL provider — fetches the Dataverse `$metadata` document and parses it
// into the studio's `CsdlAction[]` shape (same type as `mock/actionsCsdl.ts`).
//
// Why this exists:
//   The studio's Execute group (ExecuteAction, ExecuteFunction, ExecuteCustomApi,
//   ExecuteCustomAction modes) is metadata-driven — the action picker and
//   the parameter form both read from `CsdlAction.parameters[]` to render
//   typed inputs (Edm.Int32 → SpinButton, EntityReference → record picker,
//   OptionSetValue → Combobox, etc).
//
//   In standalone mode we ship a curated mock in `src/mock/actionsCsdl.ts`.
//   In PPTB the source of truth is the live `$metadata` EDMX document —
//   accessible via `dvHost.metadata.getCSDLDocument()` (typed as
//   `getCSDLDocument(): Promise<string>` per @pptb/types).
//
// The doc is 1-5 MB of XML — we fetch once per session, cache, and parse
// using the browser's native `DOMParser` (no XML library needed). Parse
// runs ~50ms on a typical 2 MB document.
//
// Parser scope:
//   ✅ <Action> + <Function> elements at namespace level
//   ✅ <Parameter> with Type, Nullable, optional UnicodeEnabled
//   ✅ <ReturnType> with Type (void / primitive / entity / collection)
//   ✅ Binding detection — first Parameter named "bindingParameter" with
//      Type=mscrm.<entity> → bound-to-entity; Type=Collection(mscrm.<entity>)
//      → bound-to-collection; otherwise unbound
//   ✅ Composable functions (IsComposable="true")
//   ✅ Custom API / Custom Action source classification (heuristic — custom
//      messages have a publisher prefix like `new_` or `sample_`; OOB ones
//      live under `Microsoft.Dynamics.CRM`)
//
// Things we intentionally DON'T parse here:
//   ❌ EntityType / ComplexType / EnumType — those come from the regular
//      metadata endpoints (richer per-attribute info than the CSDL exposes)
//   ❌ NavigationProperty — same reason; relationships have their own typed
//      metadata endpoint with cascade configuration

import { dvHost } from './pptbBridge';
import type {
  CsdlAction, ActionParam, ActionBinding, ActionReturnType, EdmType,
} from '../mock/actionsCsdl';

// Tiny inline memoizer — CSDL is fetched at most twice per session (raw XML
// + parsed actions). A 4-line cache is enough; we don't need the full
// metadataCache machinery here.
const csdlCache = new Map<string, unknown>();
const csdlInflight = new Map<string, Promise<unknown>>();
async function memoize<T>(key: string, fetcher: () => Promise<T>): Promise<T> {
  if (csdlCache.has(key)) return csdlCache.get(key) as T;
  const pending = csdlInflight.get(key);
  if (pending) return pending as Promise<T>;
  const promise = (async () => {
    try {
      const value = await fetcher();
      csdlCache.set(key, value);
      return value;
    } finally {
      csdlInflight.delete(key);
    }
  })();
  csdlInflight.set(key, promise);
  return promise;
}

const K_CSDL_RAW = 'csdl:raw';
const K_CSDL_ACTIONS = 'csdl:actions';

// ── Public surface ──────────────────────────────────────────────────────

/**
 * Fetch + parse the live CSDL into `CsdlAction[]`. Cached for the session.
 * The raw XML is cached separately so re-parsing (e.g. after a manual cache
 * clear) doesn't re-fetch.
 */
export async function fetchCsdlActions(): Promise<CsdlAction[]> {
  return memoize(K_CSDL_ACTIONS, async () => {
    // fetchCsdlDocument is itself memoized so a second-load skips the network.
    const xml = await fetchCsdlDocument();
    return parseCsdl(xml);
  });
}

// ── Optional: provider object matching the metadata provider's shape ────
// Exposed for code that wants a single `actions` namespace to consume,
// without importing the parsing fn directly. The pattern mirrors
// `metadataProvider.metadata` so consumers can branch on `isEmbedded()`
// the same way.
export interface ActionsProvider {
  loadAll(): Promise<CsdlAction[]>;
  find(name: string): Promise<CsdlAction | undefined>;
  byCategory(cat: 'oob' | 'custom-api' | 'custom-action'): Promise<CsdlAction[]>;
  functionsByCategory(cat: 'oob' | 'custom-api'): Promise<CsdlAction[]>;
}

// Live-only actions provider. The studio is PPTB-targeted; the catalog
// always comes from the live `$metadata` document via `fetchCsdlActions()`.
// When not embedded, the underlying `dvHost.metadata.getCSDLDocument()`
// throws `HostNotAvailableError` and consumers handle the failure
// appropriately (picker shows an error message).
export const actions: ActionsProvider = {
  loadAll() { return fetchCsdlActions(); },
  async find(name) { return (await fetchCsdlActions()).find(a => a.name === name); },
  async byCategory(cat) {
    return (await fetchCsdlActions()).filter(a => a.source === cat && a.kind === 'Action');
  },
  async functionsByCategory(cat) {
    return (await fetchCsdlActions()).filter(a => a.source === cat && a.kind === 'Function');
  },
};

/**
 * SYNCHRONOUS action lookup — reads from the parsed-CSDL cache populated
 * by `fetchCsdlActions()`. Critical for the URL + body builders in
 * `executeBuilders.ts` which are called from useMemo on every render and
 * can't await an async fetch.
 *
 * The flow: when the user picks an action in the picker, the picker has
 * already kicked off `actions.loadAll()` which populates the `csdlCache`
 * entry under `K_CSDL_ACTIONS`. By the time `buildExecuteAction` runs,
 * the cache IS populated and this sync lookup hits.
 *
 * Returns undefined when the cache is empty (cold session, not yet loaded)
 * or when the named action doesn't exist. Callers must handle undefined
 * gracefully — typically by returning an empty BuiltRequest skeleton.
 */
export function findActionSync(name: string): CsdlAction | undefined {
  if (!name) return undefined;
  const cached = csdlCache.get(K_CSDL_ACTIONS) as CsdlAction[] | undefined;
  return cached?.find(a => a.name === name);
}

/** Raw `$metadata` XML — exposed for debug/inspection. Rarely needed by app
 *  code; the typical caller is `fetchCsdlActions()` above. Cached separately
 *  from the parsed result so a manual cache clear of the parsed actions
 *  doesn't trigger a re-download. */
export async function fetchCsdlDocument(): Promise<string> {
  return memoize(K_CSDL_RAW, () => dvHost.metadata.getCSDLDocument());
}

// ── XML parser ──────────────────────────────────────────────────────────

/**
 * Parse a CSDL/EDMX XML string into the studio's CsdlAction[].
 *
 * The EDMX shape we care about:
 *
 *   <edmx:Edmx>
 *     <edmx:DataServices>
 *       <Schema Namespace="Microsoft.Dynamics.CRM" ...>
 *         <Action Name="WinOpportunity">
 *           <Parameter Name="OpportunityClose" Type="mscrm.opportunityclose" Nullable="false" />
 *           <Parameter Name="Status" Type="Edm.Int32" Nullable="false" />
 *         </Action>
 *         <Function Name="WhoAmI" IsComposable="false">
 *           <ReturnType Type="mscrm.WhoAmIResponse" Nullable="false" />
 *         </Function>
 *         ...
 *       </Schema>
 *     </edmx:DataServices>
 *   </edmx:Edmx>
 */
// ── Type-resolution maps ──
//
// Two-pass parse so we can accurately classify parameter types:
//
//   Pass 1: walk every <EnumType> + <ComplexType> and collect them into
//   namespace-qualified maps (both `mscrm.<name>` and `Microsoft.Dynamics.
//   CRM.<name>` forms — both appear in real CSDL).
//
//   Pass 2: walk every <Action> + <Function>. When a parameter's raw Type
//   matches an EnumType, classify it as 'OptionSetValue' and embed the
//   enum members. When it matches a ComplexType, classify as 'ComplexType'
//   and embed the recursively-resolved fields. Falls back to the
//   csdlTypeToEdm heuristic for everything else.
//
// We export the type maps via the provider so editors can resolve
// nested-type references at render time (ComplexType properties can
// themselves reference other ComplexTypes / EnumTypes).

interface EnumTypeDef {
  /** Local name (e.g. "opportunity_statuscode"). */
  name: string;
  /** Underlying type — almost always Edm.Int32 in Dataverse. */
  underlyingType: string;
  members: Array<{ name: string; value: number }>;
}

interface ComplexTypeDef {
  name: string;
  /** Optional base type — chain-resolved to merge inherited properties. */
  baseType: string | null;
  properties: Array<{ name: string; type: string; nullable: boolean }>;
}

// ── Parser ──

export function parseCsdl(xml: string): CsdlAction[] {
  if (!xml || typeof DOMParser === 'undefined') return [];
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const parserError = doc.querySelector('parsererror');
  if (parserError) {
    console.error('[csdlProvider] failed to parse CSDL XML:', parserError.textContent);
    return [];
  }

  // Pass 1: collect EnumTypes + ComplexTypes
  const enumMap = new Map<string, EnumTypeDef>();
  const complexMap = new Map<string, ComplexTypeDef>();
  const schemas = doc.getElementsByTagName('Schema');
  for (const schema of Array.from(schemas)) {
    const ns = schema.getAttribute('Namespace') || 'Microsoft.Dynamics.CRM';
    for (const el of Array.from(schema.getElementsByTagName('EnumType'))) {
      const def = parseEnumType(el);
      if (def) registerType(enumMap, ns, def.name, def);
    }
    for (const el of Array.from(schema.getElementsByTagName('ComplexType'))) {
      const def = parseComplexType(el);
      if (def) registerType(complexMap, ns, def.name, def);
    }
  }
  // Cache the maps so editors can do lookups for nested types.
  __enumTypeMap = enumMap;
  __complexTypeMap = complexMap;

  // Pass 2: parse Actions + Functions with full type resolution
  const out: CsdlAction[] = [];
  for (const schema of Array.from(schemas)) {
    const ns = schema.getAttribute('Namespace') || 'Microsoft.Dynamics.CRM';
    for (const el of Array.from(schema.getElementsByTagName('Action'))) {
      const action = parseActionOrFunction(el, ns, 'Action');
      if (action) out.push(action);
    }
    for (const el of Array.from(schema.getElementsByTagName('Function'))) {
      const action = parseActionOrFunction(el, ns, 'Function');
      if (action) out.push(action);
    }
  }
  return out;
}

/**
 * Register a type under BOTH its `<ns>.<name>` and `mscrm.<name>` keys.
 * Real CSDL is inconsistent — the schema declares EnumType under e.g.
 * "Microsoft.Dynamics.CRM" namespace, but parameter Types reference it as
 * `mscrm.<name>` (legacy short form). Indexing both forms lets the lookup
 * succeed regardless of which form the consumer uses.
 */
function registerType<T>(map: Map<string, T>, ns: string, name: string, def: T): void {
  map.set(`${ns}.${name}`, def);
  map.set(`mscrm.${name}`, def);
}

// Module-level caches populated by parseCsdl. Module scope is fine: the
// provider is a singleton (one CSDL per session), and we always re-populate
// these on every parseCsdl call.
let __enumTypeMap: Map<string, EnumTypeDef> = new Map();
let __complexTypeMap: Map<string, ComplexTypeDef> = new Map();

/** Lookup an EnumType definition by its fully-qualified name. Exported for
 *  consumers (e.g. ActionParamForm) that need to render nested-type inputs. */
export function lookupEnumType(name: string): EnumTypeDef | undefined {
  return __enumTypeMap.get(name);
}
/** Lookup a ComplexType definition by its fully-qualified name. */
export function lookupComplexType(name: string): ComplexTypeDef | undefined {
  return __complexTypeMap.get(name);
}

function parseEnumType(el: Element): EnumTypeDef | null {
  const name = el.getAttribute('Name');
  if (!name) return null;
  const underlyingType = el.getAttribute('UnderlyingType') ?? 'Edm.Int32';
  const members = Array.from(el.getElementsByTagName('Member'))
    .map(m => ({
      name: m.getAttribute('Name') ?? '',
      // Some declarations omit Value (auto-increment from 0). When present
      // it's always a string; coerce to number.
      value: Number(m.getAttribute('Value') ?? '0'),
    }))
    .filter(m => m.name);
  return { name, underlyingType, members };
}

function parseComplexType(el: Element): ComplexTypeDef | null {
  const name = el.getAttribute('Name');
  if (!name) return null;
  const baseType = el.getAttribute('BaseType');

  // Walk BOTH <Property> AND <NavigationProperty> elements. Real Dataverse
  // ComplexTypes (e.g. PrincipalAccess) declare entity-typed slots as
  // NavigationProperty — we'd silently lose the Principal field otherwise.
  // The downstream type classification (resolveComplexFields) handles
  // entity-typed `mscrm.<name>` and the polymorphic `mscrm.crmbaseentity`
  // case via csdlTypeToEdm.
  //
  // Note: we use direct-child filtering (not getElementsByTagName) to avoid
  // picking up Property elements nested inside Annotation / NavigationProperty
  // children. ComplexType properties are always direct descendants.
  const properties: ComplexTypeDef['properties'] = [];
  for (const child of Array.from(el.children)) {
    const tag = child.localName;
    if (tag === 'Property' || tag === 'NavigationProperty') {
      const pName = child.getAttribute('Name');
      if (!pName) continue;
      properties.push({
        name: pName,
        type: child.getAttribute('Type') ?? 'Edm.String',
        nullable: child.getAttribute('Nullable') !== 'false',
      });
    }
  }
  return { name, baseType, properties };
}

function parseActionOrFunction(
  el: Element, namespace: string, kind: 'Action' | 'Function',
): CsdlAction | null {
  const name = el.getAttribute('Name');
  if (!name) return null;

  // BUG FIX (binding detection): per the OData CSDL spec, when IsBound="true"
  // the FIRST parameter is the binding parameter regardless of its name.
  // Real Dataverse uses `entity` (e.g. AddToQueue, RetrieveUserPrivileges)
  // and other names — not the literal `bindingParameter` token that some
  // doc examples use. We MUST consult IsBound, not the name.
  const isBound = el.getAttribute('IsBound') === 'true';

  const params: ActionParam[] = [];
  let binding: ActionBinding = { kind: 'unbound' };

  // First parameter under an IsBound action/function is the binding param.
  // The rest become regular parameters surfaced in the form.
  const paramEls = Array.from(el.getElementsByTagName('Parameter'));
  for (let i = 0; i < paramEls.length; i++) {
    const p = paramEls[i];
    const pName = p.getAttribute('Name') ?? '';
    const rawType = p.getAttribute('Type') ?? 'Edm.String';
    // BUG FIX (Nullable default): per OData CSDL spec, missing Nullable
    // defaults to "true" (= nullable = optional). Old code treated missing
    // Nullable as required, which red-badged every param without an explicit
    // attribute (e.g. WinOpportunity.Caller).
    const required = p.getAttribute('Nullable') === 'false';

    if (i === 0 && isBound) {
      // Binding: `mscrm.<entity>` for bound-to-entity,
      // `Collection(mscrm.<entity>)` for bound-to-collection.
      const isCollection = rawType.startsWith('Collection(');
      const stripped = isCollection
        ? rawType.replace(/^Collection\((.+)\)$/, '$1')
        : rawType;
      const entityType = stripped.replace(/^mscrm\./, '');
      binding = isCollection
        ? { kind: 'collection', entityType }
        : { kind: 'entity', entityType };
      continue;
    }

    // Two-pass classification:
    //   (1) Check if the raw type matches a parsed EnumType → OptionSetValue
    //       with members populated inline (used by the OptionSetInput).
    //   (2) Check if it matches a parsed ComplexType → ComplexType with
    //       fields populated recursively (each field classified the same
    //       way).
    //   (3) Otherwise fall back to the csdlTypeToEdm heuristic for entity
    //       refs / primitives / collections.
    const enumDef = lookupEnumType(rawType);
    const complexDef = !enumDef ? lookupComplexType(rawType) : undefined;

    if (enumDef) {
      params.push({
        name: pName,
        type: 'OptionSetValue',
        required,
        optionSet: enumDef.members.map(m => ({ value: m.value, label: m.name })),
      });
    } else if (complexDef) {
      params.push({
        name: pName,
        type: 'ComplexType',
        required,
        complexType: {
          name: complexDef.name,
          fields: resolveComplexFields(complexDef),
        },
      });
    } else {
      params.push({
        name: pName,
        type: csdlTypeToEdm(rawType),
        required,
        // Per-type metadata (entityType for refs) — for EntityReference /
        // EntitySpecific / EntityCollection only.
        ...(typeMetadataFromRaw(rawType)),
      });
    }
  }

  const returnType = parseReturnType(el);
  const isComposable =
    kind === 'Function' && el.getAttribute('IsComposable') === 'true';

  // Detect Dataverse query functions — they share a recognizable CSDL
  // signature: IsBound="true", first parameter named "PropertyName" of
  // type Edm.String, return type Edm.Boolean. These ~70 functions
  // (Last7Days, Between, EqualUserId, Contains, etc.) are designed to
  // be inlined inside `$filter` expressions, NOT invoked standalone via
  // Execute Function. Tagged so ExecuteFunctionMode's picker can filter
  // them out. Reference:
  //   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/queryfunctions
  let isQueryFunction = false;
  if (kind === 'Function' && isBound && returnType.typeName === 'Edm.Boolean') {
    const first = paramEls[0];
    if (first
      && first.getAttribute('Name') === 'PropertyName'
      && first.getAttribute('Type') === 'Edm.String'
    ) {
      isQueryFunction = true;
    }
  }

  return {
    name,
    namespace,
    kind,
    binding,
    parameters: params,
    returnType,
    isComposable: isComposable || undefined,
    source: classifySource(namespace, name),
    isQueryFunction: isQueryFunction || undefined,
  };
}

/**
 * Recursively resolve a ComplexType's full property set, walking the BaseType
 * chain to merge inherited properties, and classifying each property the
 * same way action parameters are classified (EnumType ref → OptionSetValue
 * with members; ComplexType ref → nested ComplexType with its own fields;
 * Edm primitives → primitive Edm types; etc.).
 *
 * Cycle protection: a `seen` set prevents infinite recursion on
 * self-referencing types. Dataverse CSDL doesn't currently emit cycles but
 * defensive guard is cheap.
 */
function resolveComplexFields(
  def: ComplexTypeDef,
  seen: Set<string> = new Set(),
): ActionParam[] {
  if (seen.has(def.name)) return [];
  const nextSeen = new Set(seen);
  nextSeen.add(def.name);

  // Walk base chain first so derived properties (later) overlay inherited
  // ones with the same name — matches OData's "derived type wins" semantics.
  const inherited: ActionParam[] = [];
  if (def.baseType) {
    const base = lookupComplexType(def.baseType);
    if (base) inherited.push(...resolveComplexFields(base, nextSeen));
  }

  const own: ActionParam[] = def.properties.map(p => {
    const enumDef = lookupEnumType(p.type);
    if (enumDef) {
      return {
        name: p.name,
        type: 'OptionSetValue' as const,
        required: !p.nullable,
        optionSet: enumDef.members.map(m => ({ value: m.value, label: m.name })),
      };
    }
    const childComplex = lookupComplexType(p.type);
    if (childComplex) {
      return {
        name: p.name,
        type: 'ComplexType' as const,
        required: !p.nullable,
        complexType: {
          name: childComplex.name,
          fields: resolveComplexFields(childComplex, nextSeen),
        },
      };
    }
    return {
      name: p.name,
      type: csdlTypeToEdm(p.type),
      required: !p.nullable,
      ...(typeMetadataFromRaw(p.type)),
    };
  });

  // De-dupe by name — derived overrides base.
  const merged = new Map<string, ActionParam>();
  for (const p of inherited) merged.set(p.name, p);
  for (const p of own) merged.set(p.name, p);
  return Array.from(merged.values());
}

function parseReturnType(el: Element): ActionReturnType {
  const rt = el.getElementsByTagName('ReturnType')[0];
  if (!rt) return { kind: 'void', typeName: 'void' };
  const raw = rt.getAttribute('Type') ?? '';

  if (raw.startsWith('Collection(')) {
    return { kind: 'collection', typeName: raw };
  }
  if (raw.startsWith('Edm.')) {
    return { kind: 'primitive', typeName: raw };
  }
  if (raw.startsWith('mscrm.')) {
    // Heuristic — entity-like names start with a lowercase letter (account),
    // complex types are PascalCase (WhoAmIResponse, RetrieveLicenseInfoResponse).
    const after = raw.slice('mscrm.'.length);
    const first = after.charAt(0);
    return first === first.toUpperCase()
      ? { kind: 'complex', typeName: after }
      : { kind: 'entity', typeName: after };
  }
  return { kind: 'complex', typeName: raw };
}

// ── Type translation ────────────────────────────────────────────────────

const PRIMITIVE_EDM: Record<string, EdmType> = {
  'Edm.String':          'Edm.String',
  'Edm.Int32':           'Edm.Int32',
  'Edm.Int64':           'Edm.Int64',
  'Edm.Boolean':         'Edm.Boolean',
  'Edm.DateTimeOffset':  'Edm.DateTimeOffset',
  'Edm.Guid':            'Edm.Guid',
  'Edm.Decimal':         'Edm.Decimal',
  'Edm.Double':          'Edm.Double',
  'Edm.Binary':          'Edm.Binary',
  'Collection(Edm.String)': 'Collection(Edm.String)',
  'Collection(Edm.Int32)':  'Collection(Edm.Int32)',
  'Collection(Edm.Guid)':   'Collection(Edm.Guid)',
};

function csdlTypeToEdm(raw: string): EdmType {
  if (PRIMITIVE_EDM[raw]) return PRIMITIVE_EDM[raw];

  // EntityReference family — `mscrm.crmbaseentity` is the canonical
  // polymorphic reference; everything else with `mscrm.` prefix that's NOT
  // PascalCase is an entity-specific lookup type.
  if (raw === 'mscrm.crmbaseentity') return 'EntityReference';
  if (raw.startsWith('Collection(mscrm.')) return 'EntityCollection';
  if (raw.startsWith('mscrm.')) {
    const after = raw.slice('mscrm.'.length);
    // PascalCase → complex type
    if (after.charAt(0) === after.charAt(0).toUpperCase() && !after.includes('.')) {
      return 'ComplexType';
    }
    return 'EntitySpecific';
  }
  // Microsoft.Dynamics.CRM.<enum-name> for option-set parameters
  if (raw.includes('Microsoft.Dynamics.CRM.') && raw.match(/[A-Z]/)) {
    return 'OptionSetValue';
  }
  // Default — string is the safest fallback for the inputs we don't know about
  return 'Edm.String';
}

function typeMetadataFromRaw(raw: string): Partial<ActionParam> {
  // EntityReference / EntitySpecific / EntityCollection — capture target entity logical name
  if (raw.startsWith('Collection(mscrm.')) {
    return { entityType: raw.replace(/^Collection\(mscrm\.(.+)\)$/, '$1') };
  }
  if (raw.startsWith('mscrm.') && raw !== 'mscrm.crmbaseentity') {
    const after = raw.slice('mscrm.'.length);
    // Skip complex-type names (PascalCase) — those aren't entity types
    if (after.charAt(0) === after.charAt(0).toLowerCase()) {
      return { entityType: after };
    }
  }
  return {};
}

// ── Source classification ──────────────────────────────────────────────
// We classify actions into custom-api / custom-action / oob buckets so the
// Execute picker can scope by source. Heuristic:
//
//   - Namespace "Microsoft.Dynamics.CRM" + lowercase-publisher-prefixed name
//     (e.g. `new_AddNoteToContact`, `sample_CustomApiSendNotification`)
//     → custom (further split: custom-action if it's an Action *or* uses
//       PascalCase verb, custom-api if it's a Function or has Output*
//       response shape)
//   - Else → oob

const CUSTOM_PREFIX_RE = /^[a-z][a-z0-9]+_[A-Z]/;

function classifySource(_namespace: string, name: string): CsdlAction['source'] {
  if (!CUSTOM_PREFIX_RE.test(name)) return 'oob';
  // Custom messages — we can't tell custom-api from custom-action purely
  // from name. Default to custom-api since that's the modern recommendation;
  // the editor picker still surfaces all custom messages either way.
  return 'custom-api';
}
