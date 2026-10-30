// ResponsePropertiesCard — surfaces the return-type contract for an action
// or function. Layout:
//
//   ┌────────────────────────────────────────────────────────────┐
//   │ Response properties              [from Custom API def]     │
//   ├──────────────┬────────────────┬────────────────────────────┤
//   │ Name         │ Type           │ Notes                       │
//   ├──────────────┼────────────────┼────────────────────────────┤
//   │ RenewalPrice │ Money          │ USD                         │
//   │ ValidUntil   │ DateTime       │ UTC                         │
//   └──────────────┴────────────────┴────────────────────────────┘
//
// For OOB actions/functions, the response is documented as a complex type
// (e.g. WhoAmIResponse). For custom APIs, response properties are declared
// per-row in the customapiresponseproperty table. We render the CSDL's
// returnType + any known complex-type fields.

import { tokens, mergeClasses, Badge, Caption1 } from '@fluentui/react-components';
import { Code20Filled } from '@fluentui/react-icons';
import { useStudioStyles } from '../primitives/styles';
import type { CsdlAction } from '../mock/actionsCsdl';
import { lookupComplexType } from '../host/csdlProvider';

// Known complex-type schemas — for actions/functions that return well-known
// response types. Custom APIs would source these from the customapiresponseproperty
// table in real Dataverse; we hardcode the common ones for the mock.
const KNOWN_RESPONSE_SCHEMAS: Record<
  string,
  Array<{ name: string; type: string; notes?: string }>
> = {
  WhoAmIResponse: [
    { name: 'UserId', type: 'Edm.Guid', notes: 'GUID of the calling user.' },
    { name: 'BusinessUnitId', type: 'Edm.Guid', notes: "User's business unit." },
    { name: 'OrganizationId', type: 'Edm.Guid', notes: 'Current org (tenant).' },
  ],
  RetrieveTotalRecordCountResponse: [
    {
      name: 'EntityRecordCountCollection',
      type: 'EntityRecordCountCollection',
      notes: 'Count keyed by entity logical name.',
    },
  ],
  GetTimeZoneCodeByLocalizedNameResponse: [
    { name: 'TimeZoneCode', type: 'Edm.Int32', notes: 'Resolved timezone code.' },
  ],
  AddToQueueResponse: [
    { name: 'QueueItemId', type: 'Edm.Guid', notes: 'GUID of the new queueitem row.' },
  ],
  RetrieveUserPrivilegesResponse: [
    {
      name: 'RolePrivileges',
      type: 'Collection(RolePrivilege)',
      notes: "Privileges granted by the user's security roles.",
    },
  ],
  RetrievePrincipalAccessResponse: [
    {
      name: 'AccessRights',
      type: 'AccessRights (bitmask)',
      notes: 'Effective access rights as a comma-separated string.',
    },
  ],
  ExportTranslationResponse: [
    {
      name: 'ExportTranslationFile',
      type: 'Edm.Binary',
      notes: 'Base64-encoded translation .zip.',
    },
  ],
  sample_SendNotificationResponse: [
    { name: 'NotificationId', type: 'Edm.Guid', notes: 'GUID of the notification row created.' },
    {
      name: 'DeliveredAt',
      type: 'Edm.DateTimeOffset',
      notes: 'Server time the notification was queued.',
    },
  ],
};

/** The element type of `Collection(<type>)`, or null for anything else (#317). */
export function collectionElementType(typeName: string): string | null {
  const m = /^Collection\((.+)\)$/.exec(typeName.trim());
  return m ? m[1].trim() : null;
}

/**
 * Properties of a return type from the loaded CSDL, base types first (a
 * derived property with the same name wins). Tried for `entity` too: the
 * parser calls any lowercase `mscrm.<name>` an entity, which includes
 * publisher-prefixed Custom API responses (`new_DoThingResponse`). For a
 * `Collection(<type>)`, the properties of one item (#317).
 */
export function csdlComplexProperties(
  kind: string,
  typeName: string,
): Array<{ name: string; type: string }> {
  if (kind === 'collection') {
    const element = collectionElementType(typeName);
    return element ? csdlComplexProperties('complex', element) : [];
  }
  if (kind !== 'complex' && kind !== 'entity') return [];
  const byName = new Map<string, { name: string; type: string }>();
  const seen = new Set<string>();
  const walk = (name: string) => {
    const def = lookupComplexType(name.includes('.') ? name : `mscrm.${name}`);
    if (!def || seen.has(def.name)) return;
    seen.add(def.name);
    if (def.baseType) walk(def.baseType);
    for (const p of def.properties) byName.set(p.name, { name: p.name, type: p.type });
  };
  walk(typeName);
  return [...byName.values()];
}

/**
 * The return type's kind to show. The parser calls a publisher-prefixed
 * Custom API response (`msdyn_ActivateProcessesResponse`) an entity; when
 * the CSDL has it as a complex type, say so (#317).
 */
export function returnKindLabel(rt: { kind: string; typeName: string }): string {
  return rt.kind === 'entity' && csdlComplexProperties('entity', rt.typeName).length > 0
    ? 'complex'
    : rt.kind;
}

export interface ResponsePropertiesCardProps {
  action: CsdlAction;
  /** Pill text — defaults to "from $metadata" / "from Custom API def" based on source. */
  pillText?: string;
}

export function ResponsePropertiesCard({ action, pillText }: ResponsePropertiesCardProps) {
  const s = useStudioStyles();
  const rt = action.returnType;
  // Hand-written notes first; otherwise the properties of the complex type
  // in the loaded CSDL (live $metadata inside PPTB) (#309).
  const schema: Array<{ name: string; type: string; notes?: string }> =
    KNOWN_RESPONSE_SCHEMAS[rt.typeName] ?? csdlComplexProperties(rt.kind, rt.typeName);
  const sourcePill =
    pillText ?? (action.source === 'custom-api' ? 'from Custom API def' : 'from $metadata');
  const element = rt.kind === 'collection' ? collectionElementType(rt.typeName) : null;
  const table = (
    <table style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0, fontSize: 11 }}>
      <thead>
        <Tr header>
          <Th>Name</Th>
          <Th>Type</Th>
          <Th>Notes</Th>
        </Tr>
      </thead>
      <tbody>
        {schema.map((prop) => (
          <Tr key={prop.name}>
            <Td mono>{prop.name}</Td>
            <Td>{prop.type}</Td>
            <Td>{prop.notes ?? ''}</Td>
          </Tr>
        ))}
      </tbody>
    </table>
  );

  return (
    <div className={mergeClasses(s.inlineCard)} style={{ padding: 12, maxWidth: 980 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
        <Code20Filled style={{ width: 16, height: 16, color: tokens.colorBrandForeground1 }} />
        <strong style={{ fontSize: 12 }}>Response properties</strong>
        <Badge appearance="ghost">{sourcePill}</Badge>
        <Badge appearance="tint" color="brand" style={{ marginLeft: 4 }}>
          {returnKindLabel(rt)}
        </Badge>
        <span style={{ flexGrow: 1 }} />
        <Caption1
          style={{
            color: tokens.colorNeutralForeground3,
            fontFamily: tokens.fontFamilyMonospace,
            fontSize: 10,
          }}
        >
          {rt.typeName}
        </Caption1>
      </div>

      {rt.kind === 'void' ? (
        <Caption1
          style={{
            color: tokens.colorNeutralForeground3,
            fontStyle: 'italic',
            display: 'block',
            padding: '6px 4px',
          }}
        >
          No response body — Dataverse returns <code>204 No Content</code>.
        </Caption1>
      ) : rt.kind === 'primitive' ? (
        <table
          style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0, fontSize: 11 }}
        >
          <thead>
            <Tr header>
              <Th>Name</Th>
              <Th>Type</Th>
              <Th>Notes</Th>
            </Tr>
          </thead>
          <tbody>
            <Tr>
              <Td mono>value</Td>
              <Td>{rt.typeName}</Td>
              <Td>
                Single primitive value — server wraps it as <code>{'{ value: ... }'}</code>.
              </Td>
            </Tr>
          </tbody>
        </table>
      ) : rt.kind === 'collection' ? (
        <>
          <Caption1
            style={{ color: tokens.colorNeutralForeground3, display: 'block', padding: '6px 4px' }}
          >
            Returns a collection — body envelope is{' '}
            <code>{'{ "@odata.context": "…", "value": [...] }'}</code>. Composable functions accept{' '}
            <code>$select</code> / <code>$filter</code> on the URL to scope columns.
            {schema.length > 0 && element && (
              <>
                {' '}
                Each item is a{' '}
                <code>{element.replace(/^(mscrm|Microsoft\.Dynamics\.CRM)\./, '')}</code>:
              </>
            )}
          </Caption1>
          {schema.length > 0 && table}
        </>
      ) : rt.kind === 'entity' && schema.length === 0 ? (
        // A publisher-prefixed Custom API response (`new_DoThingResponse`)
        // is parsed as an entity too; when the CSDL has it as a complex
        // type, the table below lists it (#317).
        <Caption1
          style={{ color: tokens.colorNeutralForeground3, display: 'block', padding: '6px 4px' }}
        >
          Returns an entity instance of type{' '}
          <code style={{ fontFamily: tokens.fontFamilyMonospace }}>{rt.typeName}</code>. Standard
          entity envelope with <code>@odata.etag</code>.
        </Caption1>
      ) : schema.length > 0 ? (
        table
      ) : (
        <Caption1
          style={{
            color: tokens.colorNeutralForeground3,
            display: 'block',
            padding: '6px 4px',
            fontStyle: 'italic',
          }}
        >
          Returns <code style={{ fontFamily: tokens.fontFamilyMonospace }}>{rt.typeName}</code>. Its
          properties aren&apos;t described in the loaded <code>$metadata</code>.
        </Caption1>
      )}
    </div>
  );
}

// Tiny inline table primitives so we don't bring in DataGrid for 3 rows.
function Tr({ children, header }: { children: React.ReactNode; header?: boolean }) {
  return (
    <tr style={{ background: header ? tokens.colorNeutralBackground3 : 'transparent' }}>
      {children}
    </tr>
  );
}
function Th({ children }: { children: React.ReactNode }) {
  return (
    <th
      style={{
        textAlign: 'left',
        padding: '6px 8px',
        fontSize: 10,
        fontWeight: 600,
        color: tokens.colorNeutralForeground2,
        borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
      }}
    >
      {children}
    </th>
  );
}
function Td({ children, mono }: { children: React.ReactNode; mono?: boolean }) {
  return (
    <td
      style={{
        padding: '6px 8px',
        borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
        fontFamily: mono ? tokens.fontFamilyMonospace : tokens.fontFamilyBase,
        verticalAlign: 'top',
      }}
    >
      {children}
    </td>
  );
}
