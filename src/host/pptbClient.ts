// PPTB Dataverse client wrapper — typed surface over `window.dataverseAPI`.
//
// Why the metadata $select strings here look "weird":
//
// The Dataverse metadata endpoint is finicky about which properties exist
// on which type-cast subtype. The base `AttributeMetadata` type doesn't
// have `Targets`, `MaxLength`, `Format`, etc. — those live on the type-cast
// subtypes (`LookupAttributeMetadata`, `StringAttributeMetadata`,
// `IntegerAttributeMetadata`, etc). Adding the wrong field to $select
// returns 0x80060888 and the whole fetch fails. Specific quirks like
// `IsValidForAdvancedFind/Value eq true` — the `/Value` part is required
// because the property is a `BooleanManagedProperty`, not a plain bool.
//
// Anything stored / cached by callers should be the raw types exposed
// here. Studio-shape transformations (TableMeta, ColumnMeta) happen one
// layer up.

// @pptb/types provides the global `window.dataverseAPI` typing.
/// <reference types="@pptb/types" />

// ── Raw Dataverse types ─────────────────────────────────────────────────

export interface EntityMetadata {
  LogicalName: string;
  SchemaName: string;
  DisplayName?: { UserLocalizedLabel?: { Label?: string } };
  EntitySetName: string;
  PrimaryIdAttribute: string;
  PrimaryNameAttribute: string;
  IsValidForAdvancedFind?: boolean;
  MetadataId: string;
  ObjectTypeCode: number;
}

export interface AttributeMetadata {
  LogicalName: string;
  SchemaName: string;
  DisplayName?: { UserLocalizedLabel?: { Label?: string } };
  AttributeType: string;
  AttributeTypeName?: { Value?: string };
  MetadataId: string;
  IsValidForAdvancedFind?: { Value: boolean };
  /**
   * True when the attribute lives in a different physical table (composite
   * address sub-fields, derived metadata, etc.). Filtering on a logical
   * attribute forces a server-side join — triggers the
   * `LargeAmountOfLogicalAttributes` antipattern. Available on the base
   * AttributeMetadata projection (no type-cast required).
   */
  IsLogical?: boolean;
  /**
   * Source kind for the attribute:
   *   0 = simple (regular column)
   *   1 = calculated (legacy formula)
   *   2 = rollup    (aggregate over related records)
   *   3 = formula   (Power Fx)
   *   4 = prompt    (AI-generated)
   *
   * Non-zero values trigger the `FilteringOnCalculatedColumns` antipattern
   * when used in `$filter`. Available on the base AttributeMetadata projection.
   */
  SourceType?: number;
  /** Only populated by `getAttributeDetailedMetadata` for String/Memo. */
  MaxLength?: number;
  /** Only populated by `getAttributeDetailedMetadata` for Decimal/Double/Money. */
  Precision?: number;
  /** Only populated by `getAttributeDetailedMetadata` for Integer/BigInt/Decimal/Double/Money. */
  MinValue?: number;
  MaxValue?: number;
  /** Only populated by `getAttributeDetailedMetadata` for DateTime + Integer. */
  Format?: string;
  DateTimeBehavior?: { Value?: string };
  MinSupportedValue?: string;
  MaxSupportedValue?: string;
  /** Only populated by `getAttributeWithOptionSet`. */
  OptionSet?: {
    Options?: Array<{
      Value: number;
      Label: { UserLocalizedLabel?: { Label?: string } };
    }>;
    TrueOption?: { Value: number; Label: { UserLocalizedLabel?: { Label?: string } } };
    FalseOption?: { Value: number; Label: { UserLocalizedLabel?: { Label?: string } } };
  };
  /** Only populated by getAttributeDetailedMetadata for Lookup/Customer/Owner. */
  Targets?: string[];

  /**
   * Whether this attribute can appear in the body of a POST (Create).
   * False for system-managed columns (createdon, modifiedon, etc.),
   * formula/rollup outputs, and read-only customizations. Fetched in the
   * base AttributeMetadata projection — no drill-in needed.
   * Reference: AttributeMetadata.IsValidForCreate.
   */
  IsValidForCreate?: boolean;
  /** Whether this attribute can appear in the body of a PATCH (Update).
   *  Subset of IsValidForCreate — some columns are write-once-on-create. */
  IsValidForUpdate?: boolean;
  /** Whether this attribute can be read at all. Always true on standard
   *  attributes; false on a few audit-internal columns. We project it for
   *  defensive filtering on $select in the Read modes. */
  IsValidForRead?: boolean;
  /**
   * Whether this attribute is exposed via the OData / Web API. False for:
   *   - Legacy `*_base` partner columns (every Money column has one)
   *   - Some inheritance-derived placeholders
   *   - A handful of internal platform columns
   *
   * Hard wall — if false, the attribute can't appear in $select / $filter
   * / $orderby / POST body. The server 400s with
   * `0x80060888 NavigationProperty 'foo' is not valid`. We filter server-
   * side via `$filter=IsValidODataAttribute eq true` AND defensively at
   * mapAttribute time, so non-OData columns never enter the live column
   * registry.
   */
  IsValidODataAttribute?: boolean;
  /**
   * Business-rule required level. Drives the required-field guard in
   * Create mode.
   *   - "None"                  no requirement
   *   - "Recommended"           UI hint only, not enforced server-side
   *   - "ApplicationRequired"   business-marked required ("blue asterisk")
   *   - "SystemRequired"        platform-enforced, can't be removed
   *
   * The maker shows this as the "Required" / "Business Required" /
   * "Business Recommended" picker on a column.
   */
  RequiredLevel?: {
    Value?: 'None' | 'Recommended' | 'ApplicationRequired' | 'SystemRequired';
    CanBeChanged?: boolean;
    ManagedPropertyLogicalName?: string;
  };
}

export interface RelationshipMetadata {
  SchemaName: string;
  RelationshipType: 'OneToManyRelationship' | 'ManyToOneRelationship' | 'ManyToManyRelationship';
  ReferencedEntity?: string;
  ReferencedAttribute?: string;
  ReferencingEntity?: string;
  ReferencingAttribute?: string;
  MetadataId: string;
  IsValidForAdvancedFind?: boolean;
  IsCustomRelationship?: boolean;
  /**
   * The OData navigation-property name on the REFERENCING entity. This is
   * the string you put in `$expand` or use in lambda paths (e.g.
   * `primarycontactid` for incident → contact, or `customerid_account` for
   * a polymorphic Customer pointing at account). Critically, for Customer
   * / Owner / polymorphic lookups the schema does NOT expose the bare
   * attribute name — only the target-disambiguated `<attr>_<target>` form.
   * Source of truth for `$expand`. See
   * https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-web-api-functions#address-lookup-and-customer-fields-in-functions
   */
  ReferencingEntityNavigationPropertyName?: string;
  /** Counterpart on the REFERENCED entity — collection-valued (1:N). */
  ReferencedEntityNavigationPropertyName?: string;
  // N:N-specific
  Entity1LogicalName?: string;
  Entity2LogicalName?: string;
  Entity1IntersectAttribute?: string;
  Entity2IntersectAttribute?: string;
  IntersectEntityName?: string;
  /** N:N nav property name on Entity1 / Entity2 — what `$expand` accepts. */
  Entity1NavigationPropertyName?: string;
  Entity2NavigationPropertyName?: string;
}

export interface WhoAmIResponse {
  UserId: string;
  BusinessUnitId: string;
  OrganizationId: string;
}

export interface PrivilegeCheckResponse {
  RolePrivileges: Array<{
    Depth: 'Basic' | 'Local' | 'Deep' | 'Global';
    PrivilegeId: string;
    BusinessUnitId: string;
    PrivilegeName: string;
  }>;
}

export interface Publisher {
  publisherid: string;
  friendlyname: string;
  uniquename: string;
  customizationprefix: string;
}

export interface Solution {
  solutionid: string;
  friendlyname: string;
  uniquename: string;
  solutionpackageversion?: string;
  version?: string;
  isvisible?: boolean;
  ismanaged?: boolean;
  _publisherid_value?: string;
}

export interface PublisherWithSolutions {
  publisher: Publisher;
  solutions: Solution[];
}

export interface SolutionComponent {
  msdyn_name: string;
  msdyn_displayname?: string;
  msdyn_logicalcollectionname?: string;
  msdyn_solutionid: string;
  msdyn_componenttype: number;
}

export interface AccessSummary {
  userId: string;
  canReadPublisher: boolean;
  canReadSolution: boolean;
  canReadCustomization: boolean;
  /** Has all three privileges → can use publisher → solution → entity cascade. */
  fullFilterMode: boolean;
  /** Has customization + solution but not publisher → solution-only cascade. */
  solutionsOnlyMode: boolean;
  /** Has customization + publisher but not solution → publishers-only mode. */
  publishersOnlyMode: boolean;
  /** Customization only → no scope filtering available, only "all entities". */
  metadataOnlyMode: boolean;
  /** No customization privilege → no metadata access at all. */
  noAccessMode: boolean;
}

// ── Availability check ─────────────────────────────────────────────────

export function isDataverseAvailable(): boolean {
  return typeof window !== 'undefined' && !!window.dataverseAPI;
}

// ── Metadata-call concurrency throttle ─────────────────────────────────
//
// Dataverse caps a connection at ~100 concurrent server requests
// (0x80072326). Our metadata loaders are normally well under that, but
// React 18's concurrent rendering + cascading editor mounts can fire a
// bursty wave before any of them complete. This semaphore guarantees no
// more than `METADATA_CONCURRENCY` metadata `queryData` calls are in
// flight at once — additional callers queue, in FIFO order.
//
// Only metadata helpers go through this. User-initiated record queries
// (RetrieveMultiple execute, RecordPicker typeahead, etc.) bypass it
// because they typically run one at a time.
const METADATA_CONCURRENCY = 6;
let __inFlightMetadata = 0;
const __metadataQueue: Array<() => void> = [];

async function throttledMetadataQuery<T>(query: string): Promise<T> {
  if (__inFlightMetadata >= METADATA_CONCURRENCY) {
    await new Promise<void>(resolve => { __metadataQueue.push(resolve); });
  }
  __inFlightMetadata++;
  try {
    return (await window.dataverseAPI!.queryData(query)) as T;
  } finally {
    __inFlightMetadata--;
    const next = __metadataQueue.shift();
    if (next) next();
  }
}

// ── Entity metadata ────────────────────────────────────────────────────

/** Fetch every entity in the org. Sorted client-side by display name. */
export async function getAllEntities(advancedFindOnly: boolean = true): Promise<EntityMetadata[]> {
  if (!isDataverseAvailable()) throw new Error('PPTB Dataverse API not available');

  let query =
    'EntityDefinitions?$select=LogicalName,SchemaName,DisplayName,EntitySetName,' +
    'PrimaryIdAttribute,PrimaryNameAttribute,IsValidForAdvancedFind,MetadataId,ObjectTypeCode';
  if (advancedFindOnly) {
    query += '&$filter=IsValidForAdvancedFind eq true';
  }

  const result = await throttledMetadataQuery<{ value?: unknown[] } & Record<string, unknown>>(query);
  const entities = result.value as unknown as EntityMetadata[];

  // Metadata endpoint doesn't support $orderby — sort client-side.
  entities.sort((a, b) => {
    const nameA = a.DisplayName?.UserLocalizedLabel?.Label || a.LogicalName;
    const nameB = b.DisplayName?.UserLocalizedLabel?.Label || b.LogicalName;
    return nameA.localeCompare(nameB);
  });

  return entities;
}

/** Convenience alias for `getAllEntities(true)`. */
export const getAllAdvancedFindEntities = (): Promise<EntityMetadata[]> => getAllEntities(true);

/** Single-entity metadata via the host's typed `getEntityMetadata` method. */
export async function getEntityMetadata(logicalName: string): Promise<EntityMetadata> {
  if (!isDataverseAvailable()) throw new Error('PPTB Dataverse API not available');
  return (await window.dataverseAPI!.getEntityMetadata(
    logicalName,
    true,
  )) as unknown as EntityMetadata;
}

/** Entity attributes — basic projection. Type-specific props via separate calls. */
export async function getEntityAttributes(
  logicalName: string,
  advancedFindOnly: boolean = true,
): Promise<AttributeMetadata[]> {
  if (!isDataverseAvailable()) throw new Error('PPTB Dataverse API not available');

  // `IsLogical` and `SourceType` are on the BASE AttributeMetadata projection
  // (no type-cast required) and we need them on every column so the filter
  // editor can surface the FilteringOnCalculatedColumns / LargeAmountOfLogicalAttributes
  // anti-patterns inline at the rule level — without forcing an eager
  // per-column detail fetch. Two extra fields in one round-trip.
  let query =
    `EntityDefinitions(LogicalName='${logicalName}')/Attributes` +
    `?$select=LogicalName,SchemaName,DisplayName,AttributeType,AttributeTypeName,` +
    `MetadataId,IsValidForAdvancedFind,IsLogical,SourceType,` +
    // Write-eligibility + required-level. All on the BASE projection (no
    // type-cast). Drives:
    //   • FieldSetEditor's column filter — Create only shows columns where
    //     IsValidForCreate is true; Update where IsValidForUpdate is true.
    //   • Create mode's required-field guard — gates Execute until every
    //     SystemRequired / ApplicationRequired column has a value.
    `IsValidForCreate,IsValidForUpdate,IsValidForRead,RequiredLevel,` +
    // IsValidODataAttribute — hard wall for Web API access. Filtered out
    // entirely below so non-OData columns (legacy `*_base` Money partners,
    // inheritance placeholders, internal platform columns) don't pollute
    // the registry.
    `IsValidODataAttribute`;
  // OData only honors ONE $filter per request — multiple `&$filter=` clauses
  // overwrite each other rather than AND-ing. Build the predicate as a
  // single boolean expression.
  //
  // We INTENTIONALLY do NOT filter on `IsValidODataAttribute eq true` here.
  // Microsoft's metadata reports that flag at the entity-type-root level,
  // not the runtime level. On inheritance-based entities (task, email,
  // appointment, letter, phonecall, fax — every activity type) ~90% of
  // attrs return `IsValidODataAttribute=false` even though they work fine
  // at runtime because Dataverse routes them through the parent
  // (activitypointer) automatically. Filtering on that flag would strip
  // `subject`, `description`, `ownerid`, `regardingobjectid`, `statecode`
  // and most other useful attrs from those entities.
  //
  // `IsValidForAdvancedFind` IS reliable across inheritance and is a
  // reasonable default-on filter (hides internal/system-only columns).
  const predicates: string[] = [];
  if (advancedFindOnly) {
    predicates.push('IsValidForAdvancedFind/Value eq true');
  }
  if (predicates.length > 0) {
    query += `&$filter=${predicates.join(' and ')}`;
  }

  const result = await throttledMetadataQuery<{ value?: unknown[] } & Record<string, unknown>>(query);
  const attrs = result.value as unknown as AttributeMetadata[];

  attrs.sort((a, b) => {
    const nameA = a.DisplayName?.UserLocalizedLabel?.Label || a.LogicalName;
    const nameB = b.DisplayName?.UserLocalizedLabel?.Label || b.LogicalName;
    return nameA.localeCompare(nameB);
  });

  return attrs;
}

/** Drill into a Boolean/State/Status/Picklist attribute to load its OptionSet. */
export async function getAttributeWithOptionSet(
  entityLogicalName: string,
  attributeLogicalName: string,
): Promise<AttributeMetadata> {
  if (!isDataverseAvailable()) throw new Error('PPTB Dataverse API not available');

  // Step 1: confirm AttributeType so we know which type-cast to use.
  const basicQuery =
    `EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(LogicalName='${attributeLogicalName}')` +
    `?$select=AttributeType,LogicalName,SchemaName,DisplayName,MetadataId`;

  const basicResult = await throttledMetadataQuery<{ value?: unknown[] } & Record<string, unknown>>(basicQuery);
  // Single-attribute reads return the record directly OR wrapped in `value[0]`.
  const basicAttr = (basicResult.value?.[0] || basicResult) as unknown as AttributeMetadata;

  if (!basicAttr || !basicAttr.LogicalName) {
    throw new Error(`Attribute ${attributeLogicalName} not found in entity ${entityLogicalName}`);
  }

  let typeCast = '';
  switch (basicAttr.AttributeType) {
    case 'Boolean':  typeCast = 'Microsoft.Dynamics.CRM.BooleanAttributeMetadata';  break;
    case 'State':    typeCast = 'Microsoft.Dynamics.CRM.StateAttributeMetadata';    break;
    case 'Status':   typeCast = 'Microsoft.Dynamics.CRM.StatusAttributeMetadata';   break;
    case 'Picklist': typeCast = 'Microsoft.Dynamics.CRM.PicklistAttributeMetadata'; break;
    case 'MultiSelectPicklist':
      typeCast = 'Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata';
      break;
    default:
      // No OptionSet on this type — return the basic shape.
      return basicAttr;
  }

  const fullQuery =
    `EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(LogicalName='${attributeLogicalName}')/${typeCast}` +
    `?$expand=OptionSet`;
  const fullResult = await throttledMetadataQuery<{ value?: unknown[] } & Record<string, unknown>>(fullQuery);
  const fullAttr = (fullResult.value?.[0] || fullResult) as unknown as AttributeMetadata;
  if (!fullAttr || !fullAttr.LogicalName) {
    throw new Error(`Failed to retrieve attribute ${attributeLogicalName} with OptionSet expansion`);
  }
  return fullAttr;
}

/** Drill into a typed attribute for MinValue/MaxValue/Precision/Format/Targets. */
export async function getAttributeDetailedMetadata(
  entityLogicalName: string,
  attributeLogicalName: string,
  attributeType: string,
): Promise<AttributeMetadata> {
  if (!isDataverseAvailable()) throw new Error('PPTB Dataverse API not available');

  let typeCast = '';
  let selectProps = '';

  switch (attributeType) {
    case 'Integer':
      typeCast = 'Microsoft.Dynamics.CRM.IntegerAttributeMetadata';
      selectProps = 'SchemaName,MaxValue,MinValue,Format';
      break;
    case 'BigInt':
      typeCast = 'Microsoft.Dynamics.CRM.BigIntAttributeMetadata';
      selectProps = 'SchemaName,MaxValue,MinValue';
      break;
    case 'Decimal':
      typeCast = 'Microsoft.Dynamics.CRM.DecimalAttributeMetadata';
      selectProps = 'SchemaName,MaxValue,MinValue,Precision';
      break;
    case 'Double':
    case 'Float':
      typeCast = 'Microsoft.Dynamics.CRM.DoubleAttributeMetadata';
      selectProps = 'SchemaName,MaxValue,MinValue,Precision';
      break;
    case 'Money':
      typeCast = 'Microsoft.Dynamics.CRM.MoneyAttributeMetadata';
      selectProps = 'SchemaName,MaxValue,MinValue,Precision';
      break;
    case 'String':
      typeCast = 'Microsoft.Dynamics.CRM.StringAttributeMetadata';
      selectProps = 'SchemaName,MaxLength,Format';
      break;
    case 'Memo':
      typeCast = 'Microsoft.Dynamics.CRM.MemoAttributeMetadata';
      selectProps = 'SchemaName,MaxLength,Format';
      break;
    case 'DateTime':
      typeCast = 'Microsoft.Dynamics.CRM.DateTimeAttributeMetadata';
      selectProps = 'SchemaName,Format,DateTimeBehavior';
      break;
    case 'Lookup':
    case 'Customer':
    case 'Owner':
      typeCast = 'Microsoft.Dynamics.CRM.LookupAttributeMetadata';
      selectProps = 'SchemaName,Targets';
      break;
    default: {
      const basicQuery =
        `EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(LogicalName='${attributeLogicalName}')` +
        `?$select=SchemaName,LogicalName,DisplayName,AttributeType,MetadataId`;
      const basicResult = await throttledMetadataQuery<{ value?: unknown[] } & Record<string, unknown>>(basicQuery);
      return (basicResult.value?.[0] || basicResult) as unknown as AttributeMetadata;
    }
  }

  const detailedQuery =
    `EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(LogicalName='${attributeLogicalName}')/${typeCast}` +
    `?$select=${selectProps}`;
  const result = await throttledMetadataQuery<{ value?: unknown[] } & Record<string, unknown>>(detailedQuery);
  const detailedAttr = (result.value?.[0] || result) as unknown as AttributeMetadata;

  if (!detailedAttr || !detailedAttr.SchemaName) {
    throw new Error(`Failed to retrieve detailed metadata for attribute ${attributeLogicalName}`);
  }

  // Type cast strips off the base properties — patch them back so callers
  // get a consistent shape.
  detailedAttr.LogicalName = attributeLogicalName;
  detailedAttr.AttributeType = attributeType;
  return detailedAttr;
}

/**
 * Bulk type-cast attribute fetch — returns every attribute of a given subtype
 * with the requested type-specific properties in ONE round trip.
 *
 * This is the supported pattern from
 * https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query-metadata-web-api#use-cast-segments-with-collections
 * and is what we use during entity load instead of N per-attribute fetches —
 * a `Promise.all` over per-attribute calls trips Dataverse's 100-concurrent
 * request cap (0x80072326) on entities with many lookups (e.g. account).
 *
 * Example: every lookup on `account` with its Targets in one request:
 *   EntityDefinitions(LogicalName='account')/Attributes/
 *     Microsoft.Dynamics.CRM.LookupAttributeMetadata?$select=LogicalName,Targets
 */
export async function getAttributesByTypeCast(
  entityLogicalName: string,
  typeCast: string,
  selectProps: string,
  expandProps?: string,
): Promise<AttributeMetadata[]> {
  if (!isDataverseAvailable()) throw new Error('PPTB Dataverse API not available');
  const expandPart = expandProps ? `&$expand=${expandProps}` : '';
  const query =
    `EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes/${typeCast}` +
    `?$select=${selectProps}${expandPart}`;
  const result = await throttledMetadataQuery<{ value?: unknown[] } & Record<string, unknown>>(query);
  return (result.value ?? []) as unknown as AttributeMetadata[];
}

/** Three parallel queries — 1:N + N:1 + N:N. Sorted client-side by SchemaName. */
export async function getEntityRelationships(
  logicalName: string,
  advancedFindOnly: boolean = true,
): Promise<{
  oneToMany: RelationshipMetadata[];
  manyToOne: RelationshipMetadata[];
  manyToMany: RelationshipMetadata[];
}> {
  if (!isDataverseAvailable()) throw new Error('PPTB Dataverse API not available');

  const buildQuery = (relType: string) => {
    let q =
      `EntityDefinitions(LogicalName='${logicalName}')/${relType}` +
      `?$select=SchemaName,RelationshipType,ReferencedEntity,ReferencedAttribute,` +
      `ReferencingEntity,ReferencingAttribute,` +
      // Critical for $expand on polymorphic lookups — Customer / Owner expose
      // their nav as `<attr>_<targetEntity>` (e.g. `customerid_account`),
      // not the bare attribute name. ReferencingEntityNavigationPropertyName
      // gives us the right value to put in $expand.
      `ReferencingEntityNavigationPropertyName,ReferencedEntityNavigationPropertyName,` +
      `MetadataId,IsValidForAdvancedFind,IsCustomRelationship`;
    if (advancedFindOnly) q += '&$filter=IsValidForAdvancedFind eq true';
    return q;
  };
  const buildM2MQuery = () => {
    let q =
      `EntityDefinitions(LogicalName='${logicalName}')/ManyToManyRelationships` +
      `?$select=SchemaName,RelationshipType,Entity1LogicalName,Entity2LogicalName,` +
      `Entity1IntersectAttribute,Entity2IntersectAttribute,IntersectEntityName,` +
      // N:N nav property names per side — what $expand accepts.
      `Entity1NavigationPropertyName,Entity2NavigationPropertyName,` +
      `MetadataId,IsValidForAdvancedFind,IsCustomRelationship`;
    if (advancedFindOnly) q += '&$filter=IsValidForAdvancedFind eq true';
    return q;
  };

  const [o2m, m2o, m2m] = await Promise.all([
    throttledMetadataQuery<{ value: unknown[] }>(buildQuery('OneToManyRelationships')),
    throttledMetadataQuery<{ value: unknown[] }>(buildQuery('ManyToOneRelationships')),
    throttledMetadataQuery<{ value: unknown[] }>(buildM2MQuery()),
  ]);

  const sort = (arr: unknown[]): RelationshipMetadata[] => {
    const rs = arr as RelationshipMetadata[];
    rs.sort((a, b) => a.SchemaName.localeCompare(b.SchemaName));
    return rs;
  };

  return {
    oneToMany: sort(o2m.value),
    manyToOne: sort(m2o.value),
    manyToMany: sort(m2m.value),
  };
}

// ── Alternate keys ─────────────────────────────────────────────────────
//
// Reference: EntityKeyMetadata — https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/entitykeymetadata
//
// Alternate keys let Upsert address a record by a business key (e.g.
// `accountnumber='ACC-9999'`) instead of the GUID primary key. They're
// optional per entity (most have none) and don't ship in the basic
// EntityDefinitions projection — separate `/Keys` child collection.
//
// Each key may compose multiple columns (composite keys). The order in
// `KeyAttributes` matters for the URL form
// `/<entityset>(col1='v1',col2='v2')`.

/** Raw shape returned by `/EntityDefinitions(...)/Keys` — the subset we project. */
export interface EntityKeyMetadata {
  /** Schema/logical name of the key definition (e.g. `mxrm_accountkey`). */
  LogicalName: string;
  SchemaName: string;
  /** Column logical names composing the key. Order matters. */
  KeyAttributes: string[];
  /** Display label — we use `UserLocalizedLabel.Label` for the UI. */
  DisplayName?: {
    UserLocalizedLabel?: { Label?: string };
    LocalizedLabels?: Array<{ Label?: string; LanguageCode?: number }>;
  };
  /** State of the index that backs this key — `Active` is the only one
   *  guaranteed to work at Upsert time. Pending / Failed mean the key is
   *  declared but the underlying database index isn't ready yet. */
  EntityKeyIndexStatus?: 'Pending' | 'InProgress' | 'Active' | 'Failed';
  MetadataId: string;
}

export async function getEntityKeys(logicalName: string): Promise<EntityKeyMetadata[]> {
  if (!isDataverseAvailable()) throw new Error('PPTB Dataverse API not available');
  // Only fetch the fields we actually project to UpsertMode's
  // AlternateKeyEditor — keeps the round-trip small. EntityKeyIndexStatus
  // is included so we can warn on non-Active keys (those won't work at
  // Upsert time even though they exist in metadata).
  const q =
    `EntityDefinitions(LogicalName='${logicalName}')/Keys` +
    `?$select=LogicalName,SchemaName,KeyAttributes,DisplayName,EntityKeyIndexStatus,MetadataId`;
  const r = await throttledMetadataQuery<{ value?: unknown[] } & Record<string, unknown>>(q);
  const arr = (r.value ?? []) as unknown as EntityKeyMetadata[];
  // Stable order by display name (fallback to logical name).
  arr.sort((a, b) => {
    const an = a.DisplayName?.UserLocalizedLabel?.Label ?? a.LogicalName;
    const bn = b.DisplayName?.UserLocalizedLabel?.Label ?? b.LogicalName;
    return an.localeCompare(bn);
  });
  return arr;
}

// ── Publishers + Solutions ─────────────────────────────────────────────

export async function getPublishers(): Promise<Publisher[]> {
  if (!isDataverseAvailable()) throw new Error('PPTB Dataverse API not available');
  const q =
    'publishers?$select=publisherid,friendlyname,uniquename,customizationprefix' +
    '&$filter=isreadonly eq false&$orderby=friendlyname asc';
  const r = await throttledMetadataQuery<{ value?: unknown[] } & Record<string, unknown>>(q);
  return r.value as unknown as Publisher[];
}

export async function getPublishersWithSolutions(): Promise<PublisherWithSolutions[]> {
  if (!isDataverseAvailable()) throw new Error('PPTB Dataverse API not available');

  const q =
    'publishers' +
    '?$select=publisherid,friendlyname,uniquename,customizationprefix' +
    '&$filter=isreadonly eq false and publisher_solution/any(s: s/isvisible eq true and s/solution_solutioncomponent/any(c: c/componenttype eq 1))' +
    '&$expand=publisher_solution(' +
      '$select=solutionid,friendlyname,isvisible,ismanaged,uniquename,version;' +
      '$filter=isvisible eq true and solution_solutioncomponent/any(c: c/componenttype eq 1))' +
    '&$orderby=friendlyname asc';
  const r = await throttledMetadataQuery<{ value?: unknown[] } & Record<string, unknown>>(q);
  const rows = r.value as unknown as Array<Publisher & { publisher_solution?: Solution[] }>;
  return rows.map(row => ({
    publisher: {
      publisherid: row.publisherid,
      friendlyname: row.friendlyname,
      uniquename: row.uniquename,
      customizationprefix: row.customizationprefix,
    },
    solutions: row.publisher_solution ?? [],
  }));
}

export async function getSolutionsByPublishers(publisherIds: string[]): Promise<Solution[]> {
  if (!isDataverseAvailable()) throw new Error('PPTB Dataverse API not available');
  if (publisherIds.length === 0) return [];
  const publisherFilter = publisherIds.map(id => `_publisherid_value eq ${id}`).join(' or ');
  const q =
    'solutions?$select=solutionid,friendlyname,uniquename,solutionpackageversion,' +
    '_publisherid_value,isvisible,ismanaged' +
    `&$filter=(${publisherFilter}) and isvisible eq true and ` +
    `solution_solutioncomponent/any(c: c/componenttype eq 1)` +
    '&$orderby=friendlyname asc';
  const r = await throttledMetadataQuery<{ value?: unknown[] } & Record<string, unknown>>(q);
  return r.value as unknown as Solution[];
}

export async function getAllSolutionsWithEntities(): Promise<Solution[]> {
  if (!isDataverseAvailable()) throw new Error('PPTB Dataverse API not available');
  const q =
    'solutions?$select=solutionid,friendlyname,uniquename,version,' +
    '_publisherid_value,isvisible,ismanaged' +
    '&$filter=isvisible eq true and solution_solutioncomponent/any(c: c/componenttype eq 1) ' +
    "and publisherid/isreadonly eq false" +
    '&$orderby=friendlyname asc';
  const r = await throttledMetadataQuery<{ value?: unknown[] } & Record<string, unknown>>(q);
  return r.value as unknown as Solution[];
}

/** Solution components (entities only — componenttype=1) for a set of solutions. */
export async function getSolutionComponents(solutionIds: string[]): Promise<SolutionComponent[]> {
  if (!isDataverseAvailable()) throw new Error('PPTB Dataverse API not available');
  if (solutionIds.length === 0) return [];
  // Issue one query per solution id and union the results — the
  // `msdyn_solutioncomponentsummaries` endpoint doesn't support a multi-id
  // filter in one request.
  const results = await Promise.all(
    solutionIds.map(id =>
      window.dataverseAPI!.queryData(
        `msdyn_solutioncomponentsummaries?` +
        `$select=msdyn_name,msdyn_displayname,msdyn_logicalcollectionname,msdyn_solutionid,msdyn_componenttype` +
        `&$filter=msdyn_componenttype eq 1 and msdyn_solutionid eq '${id}'`,
      )
    )
  );
  const merged: SolutionComponent[] = [];
  for (const r of results) {
    merged.push(...(r.value as unknown as SolutionComponent[]));
  }
  return merged;
}

/** Re-export the cache utility so hooks have a single import surface. */
import { metadataCache } from './cache';
export const filterCachedEntitiesByNames = (logicalNames: string[]): EntityMetadata[] =>
  metadataCache.filterCachedEntitiesByNames(logicalNames);

// ── WhoAmI + privileges + AccessSummary ────────────────────────────────

export async function whoAmI(): Promise<WhoAmIResponse | null> {
  if (!isDataverseAvailable()) return null;
  try {
    const result = await window.dataverseAPI!.execute({
      operationName: 'WhoAmI',
      operationType: 'function',
    });
    return result as unknown as WhoAmIResponse;
  } catch (e) {
    console.error('WhoAmI failed:', e);
    return null;
  }
}

export async function checkPrivilegeByName(userId: string, privilegeName: string): Promise<boolean> {
  if (!isDataverseAvailable()) return false;
  try {
    const q =
      `systemusers(${userId})/Microsoft.Dynamics.CRM.RetrieveUserPrivilegeByPrivilegeName(PrivilegeName='${privilegeName}')`;
    const result = await throttledMetadataQuery<{ value?: unknown[] } & Record<string, unknown>>(q);
    const response = result as unknown as PrivilegeCheckResponse;
    return !!response?.RolePrivileges?.length;
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    const isMissing = msg.includes('0x80040203') || msg.includes('does not exist');
    if (!isMissing) console.error(`checkPrivilegeByName(${privilegeName}) failed:`, e);
    return false;
  }
}

export async function getAccessSummary(): Promise<AccessSummary | null> {
  const user = await whoAmI();
  if (!user) return null;
  const [canReadPublisher, canReadSolution, canReadCustomization] = await Promise.all([
    checkPrivilegeByName(user.UserId, 'prvReadPublisher'),
    checkPrivilegeByName(user.UserId, 'prvReadSolution'),
    checkPrivilegeByName(user.UserId, 'prvReadCustomization'),
  ]);
  return {
    userId: user.UserId,
    canReadPublisher,
    canReadSolution,
    canReadCustomization,
    fullFilterMode:     canReadCustomization && canReadSolution && canReadPublisher,
    solutionsOnlyMode:  canReadCustomization && canReadSolution && !canReadPublisher,
    publishersOnlyMode: canReadCustomization && canReadPublisher && !canReadSolution,
    metadataOnlyMode:   canReadCustomization && !canReadSolution && !canReadPublisher,
    noAccessMode:      !canReadCustomization,
  };
}
