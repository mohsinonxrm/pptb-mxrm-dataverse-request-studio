import { describe, expect, it, vi } from 'vitest';
import type { CsdlAction } from '../mock/actionsCsdl';

const getCSDLDocument = vi.hoisted(() => vi.fn<() => Promise<string>>());
// Partial mock: keep the bridge's real exports (e.g. `subscribe`, which other
// host modules use at import time) and stub only the Dataverse calls.
vi.mock('./pptbBridge', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./pptbBridge')>()),
  dvHost: {
    metadata: { getCSDLDocument },
    queryData: vi.fn(async () => ({ value: [] })),
  },
}));

import {
  fetchCsdlActions,
  findActionSync,
  isCsdlSizeLimitError,
  lookupComplexType,
  lookupEnumType,
  parseCsdl,
} from './csdlProvider';

const CSDL = `<?xml version="1.0" encoding="utf-8"?>
<edmx:Edmx Version="4.0" xmlns:edmx="http://docs.oasis-open.org/odata/ns/edmx">
  <edmx:DataServices>
    <Schema Namespace="Microsoft.Dynamics.CRM" Alias="mscrm" xmlns="http://docs.oasis-open.org/odata/ns/edm">
      <EnumType Name="opportunity_statuscode">
        <Member Name="Won" Value="3" />
        <Member Name="Canceled" Value="4" />
      </EnumType>
      <ComplexType Name="PrincipalAccess">
        <Property Name="AccessMask" Type="mscrm.AccessRights" Nullable="false" />
        <NavigationProperty Name="Principal" Type="mscrm.principal" />
      </ComplexType>
      <EnumType Name="AccessRights" IsFlags="true">
        <Member Name="None" Value="0" />
        <Member Name="ReadAccess" Value="1" />
      </EnumType>
      <ComplexType Name="BaseThing">
        <Property Name="Id" Type="Edm.Guid" Nullable="false" />
        <Property Name="Name" Type="Edm.String" />
      </ComplexType>
      <ComplexType Name="DerivedThing" BaseType="mscrm.BaseThing">
        <Property Name="Name" Type="Edm.Int32" />
        <Property Name="Extra" Type="Edm.Boolean" />
      </ComplexType>
      <ComplexType Name="WhoAmIResponse">
        <Property Name="UserId" Type="Edm.Guid" Nullable="false" />
      </ComplexType>
      <Action Name="WinOpportunity">
        <Parameter Name="OpportunityClose" Type="mscrm.opportunityclose" Nullable="false" />
        <Parameter Name="Status" Type="Edm.Int32" Nullable="false" />
      </Action>
      <Action Name="AddToQueue" IsBound="true">
        <Parameter Name="entity" Type="mscrm.queue" Nullable="false" />
        <Parameter Name="Target" Type="mscrm.crmbaseentity" Nullable="false" />
        <Parameter Name="QueueItemProperties" Type="mscrm.queueitem" />
        <ReturnType Type="mscrm.AddToQueueResponse" Nullable="false" />
      </Action>
      <Action Name="BulkThing" IsBound="true">
        <Parameter Name="entityset" Type="Collection(mscrm.account)" Nullable="false" />
        <Parameter Name="Targets" Type="Collection(mscrm.contact)" />
        <Parameter Name="Names" Type="Collection(Edm.String)" />
        <ReturnType Type="Collection(mscrm.contact)" />
      </Action>
      <Action Name="SetStatus">
        <Parameter Name="Status" Type="mscrm.opportunity_statuscode" Nullable="false" />
      </Action>
      <Action Name="GrantAccess">
        <Parameter Name="Target" Type="mscrm.crmbaseentity" Nullable="false" />
        <Parameter Name="PrincipalAccess" Type="mscrm.PrincipalAccess" Nullable="false" />
      </Action>
      <Action Name="Echo">
        <Parameter Name="Thing" Type="mscrm.DerivedThing" />
      </Action>
      <Action Name="new_DoThing">
        <Parameter Name="Input" Type="Edm.String" />
        <ReturnType Type="Edm.String" />
      </Action>
      <Function Name="WhoAmI">
        <ReturnType Type="mscrm.WhoAmIResponse" Nullable="false" />
      </Function>
      <Function Name="RetrieveCount" IsComposable="true">
        <Parameter Name="EntityName" Type="Edm.String" Nullable="false" />
        <ReturnType Type="Edm.Int64" Nullable="false" />
      </Function>
      <Function Name="Last7Days" IsBound="true" IsComposable="true">
        <Parameter Name="PropertyName" Type="Edm.String" Nullable="false" />
        <ReturnType Type="Edm.Boolean" Nullable="false" />
      </Function>
    </Schema>
  </edmx:DataServices>
</edmx:Edmx>`;

function byName(actions: CsdlAction[], name: string): CsdlAction {
  const a = actions.find((x) => x.name === name);
  if (!a) throw new Error(`no ${name}`);
  return a;
}

describe('parseCsdl', () => {
  const actions = parseCsdl(CSDL);

  it('finds every Action and Function with its kind and namespace', () => {
    expect(actions.map((a) => [a.name, a.kind])).toEqual([
      ['WinOpportunity', 'Action'],
      ['AddToQueue', 'Action'],
      ['BulkThing', 'Action'],
      ['SetStatus', 'Action'],
      ['GrantAccess', 'Action'],
      ['Echo', 'Action'],
      ['new_DoThing', 'Action'],
      ['WhoAmI', 'Function'],
      ['RetrieveCount', 'Function'],
      ['Last7Days', 'Function'],
    ]);
    expect(new Set(actions.map((a) => a.namespace))).toEqual(new Set(['Microsoft.Dynamics.CRM']));
  });

  it('classifies primitive and entity-typed parameters; missing Nullable means optional', () => {
    const win = byName(actions, 'WinOpportunity');
    expect(win.binding).toEqual({ kind: 'unbound' });
    expect(win.parameters).toEqual([
      {
        name: 'OpportunityClose',
        type: 'EntitySpecific',
        required: true,
        entityType: 'opportunityclose',
      },
      { name: 'Status', type: 'Edm.Int32', required: true },
    ]);
    expect(win.returnType).toEqual({ kind: 'void', typeName: 'void' });
  });

  it('treats the first parameter of a bound operation as the binding, whatever its name', () => {
    const add = byName(actions, 'AddToQueue');
    expect(add.binding).toEqual({ kind: 'entity', entityType: 'queue' });
    expect(add.parameters).toEqual([
      { name: 'Target', type: 'EntityReference', required: true },
      {
        name: 'QueueItemProperties',
        type: 'EntitySpecific',
        required: false,
        entityType: 'queueitem',
      },
    ]);
    expect(add.returnType).toEqual({ kind: 'complex', typeName: 'AddToQueueResponse' });

    const bulk = byName(actions, 'BulkThing');
    expect(bulk.binding).toEqual({ kind: 'collection', entityType: 'account' });
    expect(bulk.parameters).toEqual([
      { name: 'Targets', type: 'EntityCollection', required: false, entityType: 'contact' },
      { name: 'Names', type: 'Collection(Edm.String)', required: false },
    ]);
    expect(bulk.returnType).toEqual({ kind: 'collection', typeName: 'Collection(mscrm.contact)' });
  });

  it('resolves EnumType parameters to option sets', () => {
    // toMatchObject: enum params may carry extra metadata (e.g. the qualified
    // enum name and IsFlags, used to emit Type'Member' literals).
    expect(byName(actions, 'SetStatus').parameters).toMatchObject([
      {
        name: 'Status',
        type: 'OptionSetValue',
        required: true,
        optionSet: [
          { value: 3, label: 'Won' },
          { value: 4, label: 'Canceled' },
        ],
      },
    ]);
  });

  it('resolves ComplexType parameters, including enum and navigation members', () => {
    const [, principalAccess] = byName(actions, 'GrantAccess').parameters;
    expect(principalAccess).toMatchObject({
      name: 'PrincipalAccess',
      type: 'ComplexType',
      required: true,
      complexType: {
        name: 'PrincipalAccess',
        fields: [
          {
            name: 'AccessMask',
            type: 'OptionSetValue',
            required: true,
            optionSet: [
              { value: 0, label: 'None' },
              { value: 1, label: 'ReadAccess' },
            ],
          },
          { name: 'Principal', type: 'EntitySpecific', required: false, entityType: 'principal' },
        ],
      },
    });
  });

  it('merges BaseType properties with the derived type winning', () => {
    const [thing] = byName(actions, 'Echo').parameters;
    expect(thing.complexType?.fields).toEqual([
      { name: 'Id', type: 'Edm.Guid', required: true },
      { name: 'Name', type: 'Edm.Int32', required: false },
      { name: 'Extra', type: 'Edm.Boolean', required: false },
    ]);
  });

  it('reads return types, composability and query-function shape', () => {
    expect(byName(actions, 'WhoAmI').returnType).toEqual({
      kind: 'complex',
      typeName: 'WhoAmIResponse',
    });
    const count = byName(actions, 'RetrieveCount');
    expect(count.isComposable).toBe(true);
    expect(count.returnType).toEqual({ kind: 'primitive', typeName: 'Edm.Int64' });
    expect(count.isQueryFunction).toBeUndefined();
    expect(byName(actions, 'Last7Days').isQueryFunction).toBe(true);
    expect(byName(actions, 'WinOpportunity').isComposable).toBeUndefined();
  });

  it('classifies publisher-prefixed messages as custom and the rest as OOB', () => {
    expect(byName(actions, 'new_DoThing').source).toBe('custom-api');
    expect(byName(actions, 'WinOpportunity').source).toBe('oob');
    expect(byName(actions, 'WhoAmI').source).toBe('oob');
  });

  it('exposes the parsed type maps under both qualified forms', () => {
    expect(lookupEnumType('mscrm.AccessRights')?.members).toHaveLength(2);
    expect(lookupEnumType('Microsoft.Dynamics.CRM.AccessRights')).toBe(
      lookupEnumType('mscrm.AccessRights'),
    );
    expect(lookupComplexType('mscrm.DerivedThing')?.baseType).toBe('mscrm.BaseThing');
  });

  it('returns no actions for empty or malformed XML', () => {
    const err = vi.spyOn(console, 'error').mockImplementation(() => {});
    expect(parseCsdl('')).toEqual([]);
    expect(parseCsdl('<edmx:Edmx><unclosed>')).toEqual([]);
    err.mockRestore();
  });
});

describe('fetchCsdlActions', () => {
  it('downloads the document once and serves later and concurrent calls from the cache', async () => {
    getCSDLDocument.mockResolvedValue(CSDL);
    expect(findActionSync('WhoAmI')).toBeUndefined();
    const [a, b] = await Promise.all([fetchCsdlActions(), fetchCsdlActions()]);
    await fetchCsdlActions();
    expect(getCSDLDocument).toHaveBeenCalledTimes(1);
    expect(a).toBe(b);
    expect(a.map((x) => x.name)).toContain('WinOpportunity');
    expect(findActionSync('WinOpportunity')?.kind).toBe('Action');
    expect(findActionSync('')).toBeUndefined();
  });
});

describe('isCsdlSizeLimitError (#79)', () => {
  it('recognises both forms of the PPTB 16 MB response cap', () => {
    expect(
      isCsdlSizeLimitError(
        'RangeError [ERR_OUT_OF_RANGE]: Cannot create a Buffer larger than 16777216 bytes',
      ),
    ).toBe(true);
    expect(isCsdlSizeLimitError('Metadata response exceeds the maximum size (16 MB)')).toBe(true);
  });

  it('does not match unrelated failures', () => {
    expect(isCsdlSizeLimitError('Network request failed')).toBe(false);
    expect(isCsdlSizeLimitError('401 Unauthorized')).toBe(false);
    expect(isCsdlSizeLimitError('')).toBe(false);
  });
});
