// #317: a `Collection(<complex>)` return lists one item's properties.
import { renderToStaticMarkup } from 'react-dom/server';
import { beforeAll, describe, expect, it } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import type { CsdlAction } from '../mock/actionsCsdl';

import {
  ResponsePropertiesCard,
  collectionElementType,
  csdlComplexProperties,
} from './ResponsePropertiesCard';
import { parseCsdl } from '../host/csdlProvider';

beforeAll(() => {
  parseCsdl(`<?xml version="1.0" encoding="utf-8"?>
<edmx:Edmx Version="4.0" xmlns:edmx="http://docs.oasis-open.org/odata/ns/edmx">
  <edmx:DataServices>
    <Schema Namespace="Microsoft.Dynamics.CRM" Alias="mscrm" xmlns="http://docs.oasis-open.org/odata/ns/edm">
      <ComplexType Name="BasePrivilege">
        <Property Name="Depth" Type="Edm.String" />
      </ComplexType>
      <ComplexType Name="RolePrivilegeItem" BaseType="mscrm.BasePrivilege">
        <Property Name="PrivilegeId" Type="Edm.Guid" />
        <Property Name="BusinessUnitId" Type="Edm.Guid" />
      </ComplexType>
    </Schema>
  </edmx:DataServices>
</edmx:Edmx>`);
});

const op = (kind: string, typeName: string, source: CsdlAction['source'] = 'oob') =>
  ({
    name: 'new_Op',
    kind: 'Function',
    namespace: 'Microsoft.Dynamics.CRM',
    binding: { kind: 'unbound' },
    parameters: [],
    returnType: { kind, typeName },
    source,
  }) as unknown as CsdlAction;

const render = (a: CsdlAction) =>
  renderToStaticMarkup(
    <FluentProvider theme={webLightTheme}>
      <ResponsePropertiesCard action={a} />
    </FluentProvider>,
  )
    .replace(/<[^>]+>/g, ' ')
    .replace(/&#x27;|&apos;/g, "'")
    .replace(/\s+/g, ' ');

describe('collectionElementType', () => {
  it('unwraps Collection(...)', () => {
    expect(collectionElementType('Collection(mscrm.RolePrivilegeItem)')).toBe(
      'mscrm.RolePrivilegeItem',
    );
    expect(collectionElementType(' Collection( Edm.String ) ')).toBe('Edm.String');
  });

  it('returns null for anything else', () => {
    expect(collectionElementType('mscrm.RolePrivilegeItem')).toBeNull();
    expect(collectionElementType('Collection(')).toBeNull();
  });
});

describe('csdlComplexProperties for a collection (#317)', () => {
  it("lists the element type's properties, base type first", () => {
    expect(csdlComplexProperties('collection', 'Collection(mscrm.RolePrivilegeItem)')).toEqual([
      { name: 'Depth', type: 'Edm.String' },
      { name: 'PrivilegeId', type: 'Edm.Guid' },
      { name: 'BusinessUnitId', type: 'Edm.Guid' },
    ]);
  });

  it('lists nothing for a collection of primitives, entities or unknown types', () => {
    expect(csdlComplexProperties('collection', 'Collection(Edm.String)')).toEqual([]);
    expect(csdlComplexProperties('collection', 'Collection(mscrm.account)')).toEqual([]);
    expect(csdlComplexProperties('collection', 'Collection(mscrm.Nope)')).toEqual([]);
  });
});

describe('ResponsePropertiesCard', () => {
  it("shows a known element type's properties as one item's", () => {
    const text = render(op('collection', 'Collection(mscrm.RolePrivilegeItem)'));
    expect(text).toContain('Each item is a RolePrivilegeItem :');
    expect(text).toMatch(/Depth Edm\.String .*PrivilegeId Edm\.Guid .*BusinessUnitId Edm\.Guid/);
  });

  it('keeps the generic collection text when the element type is unknown', () => {
    const text = render(op('collection', 'Collection(mscrm.account)'));
    expect(text).toContain('Returns a collection');
    expect(text).not.toContain('Each item is');
    expect(text).not.toContain('PrivilegeId');
  });

  it("names the loaded $metadata when a type's properties are unknown", () => {
    const text = render(op('complex', 'NoSuchResponse'));
    expect(text).toContain("properties aren't described in the loaded $metadata .");
    expect(text).toContain('from $metadata');
  });
});
