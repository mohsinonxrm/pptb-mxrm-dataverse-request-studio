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

// Known complex-type schemas — for actions/functions that return well-known
// response types. Custom APIs would source these from the customapiresponseproperty
// table in real Dataverse; we hardcode the common ones for the mock.
const KNOWN_RESPONSE_SCHEMAS: Record<string, Array<{ name: string; type: string; notes?: string }>> = {
  WhoAmIResponse: [
    { name: 'UserId',         type: 'Edm.Guid', notes: 'GUID of the calling user.' },
    { name: 'BusinessUnitId', type: 'Edm.Guid', notes: "User's business unit." },
    { name: 'OrganizationId', type: 'Edm.Guid', notes: 'Current org (tenant).' },
  ],
  RetrieveTotalRecordCountResponse: [
    { name: 'EntityRecordCountCollection', type: 'EntityRecordCountCollection', notes: 'Count keyed by entity logical name.' },
  ],
  GetTimeZoneCodeByLocalizedNameResponse: [
    { name: 'TimeZoneCode', type: 'Edm.Int32', notes: 'Resolved timezone code.' },
  ],
  AddToQueueResponse: [
    { name: 'QueueItemId', type: 'Edm.Guid', notes: 'GUID of the new queueitem row.' },
  ],
  RetrieveUserPrivilegesResponse: [
    { name: 'RolePrivileges', type: 'Collection(RolePrivilege)', notes: 'Privileges granted by the user\'s security roles.' },
  ],
  RetrievePrincipalAccessResponse: [
    { name: 'AccessRights', type: 'AccessRights (bitmask)', notes: 'Effective access rights as a comma-separated string.' },
  ],
  ExportTranslationResponse: [
    { name: 'ExportTranslationFile', type: 'Edm.Binary', notes: 'Base64-encoded translation .zip.' },
  ],
  sample_SendNotificationResponse: [
    { name: 'NotificationId',  type: 'Edm.Guid',   notes: 'GUID of the notification row created.' },
    { name: 'DeliveredAt',     type: 'Edm.DateTimeOffset', notes: 'Server time the notification was queued.' },
  ],
};

export interface ResponsePropertiesCardProps {
  action: CsdlAction;
  /** Pill text — defaults to "from $metadata" / "from Custom API def" based on source. */
  pillText?: string;
}

export function ResponsePropertiesCard({ action, pillText }: ResponsePropertiesCardProps) {
  const s = useStudioStyles();
  const rt = action.returnType;
  const schema = KNOWN_RESPONSE_SCHEMAS[rt.typeName] ?? [];
  const sourcePill = pillText ?? (action.source === 'custom-api' ? 'from Custom API def' : 'from $metadata');

  return (
    <div className={mergeClasses(s.inlineCard)} style={{ padding: 12, maxWidth: 980 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
        <Code20Filled style={{ width: 16, height: 16, color: tokens.colorBrandForeground1 }} />
        <strong style={{ fontSize: 12 }}>Response properties</strong>
        <Badge appearance="ghost">{sourcePill}</Badge>
        <Badge appearance="tint" color="brand" style={{ marginLeft: 4 }}>{rt.kind}</Badge>
        <span style={{ flexGrow: 1 }} />
        <Caption1 style={{ color: tokens.colorNeutralForeground3, fontFamily: tokens.fontFamilyMonospace, fontSize: 10 }}>
          {rt.typeName}
        </Caption1>
      </div>

      {rt.kind === 'void' ? (
        <Caption1 style={{ color: tokens.colorNeutralForeground3, fontStyle: 'italic', display: 'block', padding: '6px 4px' }}>
          No response body — Dataverse returns <code>204 No Content</code>.
        </Caption1>
      ) : rt.kind === 'primitive' ? (
        <table style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0, fontSize: 11 }}>
          <thead>
            <Tr header><Th>Name</Th><Th>Type</Th><Th>Notes</Th></Tr>
          </thead>
          <tbody>
            <Tr>
              <Td mono>value</Td>
              <Td>{rt.typeName}</Td>
              <Td>Single primitive value — server wraps it as <code>{'{ value: ... }'}</code>.</Td>
            </Tr>
          </tbody>
        </table>
      ) : rt.kind === 'collection' ? (
        <Caption1 style={{ color: tokens.colorNeutralForeground3, display: 'block', padding: '6px 4px' }}>
          Returns a collection — body envelope is <code>{'{ "@odata.context": "…", "value": [...] }'}</code>. Composable functions accept <code>$select</code> / <code>$filter</code> on the URL to scope columns.
        </Caption1>
      ) : rt.kind === 'entity' ? (
        <Caption1 style={{ color: tokens.colorNeutralForeground3, display: 'block', padding: '6px 4px' }}>
          Returns an entity instance of type <code style={{ fontFamily: tokens.fontFamilyMonospace }}>{rt.typeName}</code>. Standard entity envelope with <code>@odata.etag</code>.
        </Caption1>
      ) : schema.length > 0 ? (
        <table style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0, fontSize: 11 }}>
          <thead>
            <Tr header><Th>Name</Th><Th>Type</Th><Th>Notes</Th></Tr>
          </thead>
          <tbody>
            {schema.map(prop => (
              <Tr key={prop.name}>
                <Td mono>{prop.name}</Td>
                <Td>{prop.type}</Td>
                <Td>{prop.notes ?? ''}</Td>
              </Tr>
            ))}
          </tbody>
        </table>
      ) : (
        <Caption1 style={{ color: tokens.colorNeutralForeground3, display: 'block', padding: '6px 4px', fontStyle: 'italic' }}>
          Returns <code style={{ fontFamily: tokens.fontFamilyMonospace }}>{rt.typeName}</code>. No detailed schema in the mock CSDL — fetch the live <code>$metadata</code> for the full structure.
        </Caption1>
      )}
    </div>
  );
}

// Tiny inline table primitives so we don't bring in DataGrid for 3 rows.
function Tr({ children, header }: { children: React.ReactNode; header?: boolean }) {
  return <tr style={{ background: header ? tokens.colorNeutralBackground3 : 'transparent' }}>{children}</tr>;
}
function Th({ children }: { children: React.ReactNode }) {
  return (
    <th style={{
      textAlign: 'left',
      padding: '6px 8px',
      fontSize: 10,
      fontWeight: 600,
      color: tokens.colorNeutralForeground2,
      borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    }}>{children}</th>
  );
}
function Td({ children, mono }: { children: React.ReactNode; mono?: boolean }) {
  return (
    <td style={{
      padding: '6px 8px',
      borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
      fontFamily: mono ? tokens.fontFamilyMonospace : tokens.fontFamilyBase,
      verticalAlign: 'top',
    }}>{children}</td>
  );
}
