import { describe, expect, it } from 'vitest';
import { runDetail } from './Sidebar';
import { csdlComplexProperties } from '../editors/ResponsePropertiesCard';
import { parseCsdl } from '../host/csdlProvider';

// #309: Recent runs read "204 · 1 rows" for writes.
describe('runDetail', () => {
  it('counts rows for the read modes, with the right plural', () => {
    expect(runDetail({ modeId: 'retrieve-multiple', rowCount: 1 })).toBe(' · 1 row');
    expect(runDetail({ modeId: 'retrieve-single', rowCount: 0 })).toBe(' · 0 rows');
    expect(runDetail({ modeId: 'predefined-query', rowCount: 1234 })).toBe(
      ` · ${(1234).toLocaleString()} rows`,
    );
    expect(runDetail({ modeId: 'retrieve-nextlink' })).toBe(' · — rows');
  });

  it('shows the status alone for writes, functions and downloads without a summary', () => {
    expect(runDetail({ modeId: 'update', rowCount: 1 })).toBe('');
    expect(runDetail({ modeId: 'delete', rowCount: 0 })).toBe('');
    expect(runDetail({ modeId: 'execute-action', rowCount: 1 })).toBe('');
    expect(runDetail({ modeId: 'execute-function', rowCount: 1 })).toBe('');
    expect(runDetail({ modeId: 'manage-file', rowCount: 0 })).toBe('');
  });

  it("prefers the mode's summary", () => {
    expect(
      runDetail({ modeId: 'retrieve-multiple', rowCount: 6, summary: '6 rows · all pages' }),
    ).toBe(' · 6 rows · all pages');
    expect(runDetail({ modeId: 'associate', rowCount: 3, summary: '3 of 5 linked' })).toBe(
      ' · 3 of 5 linked',
    );
  });
});

// #309: response properties come from the loaded CSDL, including Custom API
// responses the parser classifies as entities and inherited properties.
describe('csdlComplexProperties', () => {
  parseCsdl(`<?xml version="1.0" encoding="utf-8"?>
<edmx:Edmx Version="4.0" xmlns:edmx="http://docs.oasis-open.org/odata/ns/edmx">
  <edmx:DataServices>
    <Schema Namespace="Microsoft.Dynamics.CRM" Alias="mscrm" xmlns="http://docs.oasis-open.org/odata/ns/edm">
      <ComplexType Name="CanBeReferencedResponse">
        <Property Name="CanBeReferenced" Type="Edm.Boolean" Nullable="false" />
      </ComplexType>
      <ComplexType Name="new_BaseResponse">
        <Property Name="Status" Type="Edm.String" />
      </ComplexType>
      <ComplexType Name="new_DoThingResponse" BaseType="mscrm.new_BaseResponse">
        <Property Name="Result" Type="Edm.Int32" />
      </ComplexType>
    </Schema>
  </edmx:DataServices>
</edmx:Edmx>`);

  it('lists a complex type', () => {
    expect(csdlComplexProperties('complex', 'CanBeReferencedResponse')).toEqual([
      { name: 'CanBeReferenced', type: 'Edm.Boolean' },
    ]);
  });

  it('finds a lowercase Custom API response (parsed as entity) with its base properties', () => {
    expect(csdlComplexProperties('entity', 'new_DoThingResponse')).toEqual([
      { name: 'Status', type: 'Edm.String' },
      { name: 'Result', type: 'Edm.Int32' },
    ]);
  });

  it('returns nothing for primitives or unknown types', () => {
    expect(csdlComplexProperties('primitive', 'Edm.String')).toEqual([]);
    expect(csdlComplexProperties('complex', 'NoSuchResponse')).toEqual([]);
  });
});
