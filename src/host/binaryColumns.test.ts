// #305: Image and File columns report `AttributeType: Virtual` and are not
// Advanced Find-eligible, so the attribute load dropped them and Manage
// Image / File Data found no columns.
import { afterEach, describe, expect, it, vi } from 'vitest';
import { getEntityAttributes, isBinaryColumn, type AttributeMetadata } from './pptbClient';
import { attributeTypeOf, mapAttribute } from './metadataProvider';

const attr = (
  LogicalName: string,
  AttributeType: string,
  typeName: string,
  af: boolean,
): AttributeMetadata => ({
  LogicalName,
  SchemaName: LogicalName,
  AttributeType,
  AttributeTypeName: { Value: typeName },
  MetadataId: LogicalName,
  IsValidForAdvancedFind: { Value: af },
});

describe('attributeTypeOf', () => {
  it('maps Virtual ImageType / FileType to Image / File', () => {
    expect(
      attributeTypeOf({ AttributeType: 'Virtual', AttributeTypeName: { Value: 'ImageType' } }),
    ).toBe('Image');
    expect(
      attributeTypeOf({ AttributeType: 'Virtual', AttributeTypeName: { Value: 'FileType' } }),
    ).toBe('File');
  });

  it('leaves other virtual columns unmapped', () => {
    expect(
      attributeTypeOf({ AttributeType: 'Virtual', AttributeTypeName: { Value: 'VirtualType' } }),
    ).toBeUndefined();
    expect(attributeTypeOf({ AttributeType: 'Virtual' })).toBeUndefined();
  });

  it('maps regular columns by AttributeType', () => {
    expect(
      attributeTypeOf({ AttributeType: 'String', AttributeTypeName: { Value: 'StringType' } }),
    ).toBe('String');
  });
});

describe('mapAttribute', () => {
  it('produces Image and File columns', () => {
    expect(mapAttribute(attr('entityimage', 'Virtual', 'ImageType', false))).toMatchObject({
      logicalName: 'entityimage',
      attributeType: 'Image',
    });
    expect(mapAttribute(attr('msdyn_blobfile', 'Virtual', 'FileType', false))).toMatchObject({
      logicalName: 'msdyn_blobfile',
      attributeType: 'File',
    });
  });

  it('drops label virtual columns', () => {
    expect(mapAttribute(attr('statecodename', 'Virtual', 'VirtualType', false))).toBeNull();
  });
});

describe('isBinaryColumn', () => {
  it('is true only for virtual Image / File columns', () => {
    expect(isBinaryColumn(attr('entityimage', 'Virtual', 'ImageType', false))).toBe(true);
    expect(isBinaryColumn(attr('f', 'Virtual', 'FileType', false))).toBe(true);
    expect(isBinaryColumn(attr('statecodename', 'Virtual', 'VirtualType', false))).toBe(false);
    expect(isBinaryColumn(attr('name', 'String', 'StringType', true))).toBe(false);
  });
});

describe('getEntityAttributes (Advanced Find only)', () => {
  afterEach(() => {
    delete (window as { dataverseAPI?: unknown }).dataverseAPI;
  });

  it('asks for virtual columns too and keeps only the binary ones', async () => {
    // What the server returns for the filter: Advanced Find columns plus
    // every virtual column.
    const queryData = vi.fn(async (_query: string) => ({
      value: [
        attr('name', 'String', 'StringType', true),
        attr('entityimage', 'Virtual', 'ImageType', false),
        attr('msdyn_blobfile', 'Virtual', 'FileType', false),
        attr('statecodename', 'Virtual', 'VirtualType', false),
      ],
    }));
    (window as unknown as { dataverseAPI: unknown }).dataverseAPI = { queryData };

    const attrs = await getEntityAttributes('account');

    const query = queryData.mock.calls[0][0];
    expect(query).toContain(
      "$filter=(IsValidForAdvancedFind/Value eq true or AttributeType eq Microsoft.Dynamics.CRM.AttributeTypeCode'Virtual')",
    );
    expect(attrs.map((a) => a.LogicalName).sort()).toEqual([
      'entityimage',
      'msdyn_blobfile',
      'name',
    ]);
  });

  it('returns everything unfiltered when Advanced Find only is off', async () => {
    const queryData = vi.fn(async (_query: string) => ({
      value: [
        attr('statecodename', 'Virtual', 'VirtualType', false),
        attr('internal', 'String', 'StringType', false),
      ],
    }));
    (window as unknown as { dataverseAPI: unknown }).dataverseAPI = { queryData };

    const attrs = await getEntityAttributes('account', false);

    expect(queryData.mock.calls[0][0]).not.toContain('$filter');
    expect(attrs).toHaveLength(2);
  });
});
