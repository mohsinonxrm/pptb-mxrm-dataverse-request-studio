// #305: Image and File columns report `AttributeType: Virtual` and are not
// Advanced Find-eligible, so the attribute load dropped them and Manage
// Image / File Data found no columns.
import { afterEach, describe, expect, it, vi } from 'vitest';
import { getEntityAttributes, isBinaryColumn, type AttributeMetadata } from './pptbClient';
import { attributeTypeOf, enrichBinaryColumns, mapAttribute } from './metadataProvider';
import { mustSupplyOnCreate } from '../engine/fieldBody';
import type { ColumnMeta } from '../mock/metadata';

const loadAttributesByTypeCast = vi.hoisted(() => vi.fn());
vi.mock('./dataverseMetadata', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./dataverseMetadata')>()),
  loadAttributesByTypeCast,
}));

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

describe('enrichBinaryColumns', () => {
  afterEach(() => loadAttributesByTypeCast.mockReset());

  const cols = () =>
    [
      mapAttribute(attr('name', 'String', 'StringType', true)),
      mapAttribute(attr('entityimage', 'Virtual', 'ImageType', false)),
      mapAttribute(attr('custom_photo', 'Virtual', 'ImageType', false)),
      mapAttribute(attr('custom_doc', 'Virtual', 'FileType', false)),
    ] as ColumnMeta[];

  it('fills size limits and full-image storage from the type casts', async () => {
    loadAttributesByTypeCast.mockImplementation(async (_t: string, cast: string) =>
      cast.endsWith('ImageAttributeMetadata')
        ? [
            {
              LogicalName: 'entityimage',
              MaxSizeInKB: 10240,
              CanStoreFullImage: false,
              IsPrimaryImage: true,
            },
            {
              LogicalName: 'custom_photo',
              MaxSizeInKB: 30720,
              CanStoreFullImage: true,
              IsPrimaryImage: false,
            },
          ]
        : [{ LogicalName: 'custom_doc', MaxSizeInKB: 131072 }],
    );
    const c = cols();
    await enrichBinaryColumns('account', c);
    expect(c[1]).toMatchObject({
      maxSizeInKB: 10240,
      canStoreFullImage: false,
      isPrimaryImage: true,
    });
    expect(c[2]).toMatchObject({ maxSizeInKB: 30720, canStoreFullImage: true });
    expect(c[2]).not.toHaveProperty('isPrimaryImage', true);
    expect(c[3]).toMatchObject({ maxSizeInKB: 131072 });
  });

  it('skips the requests when the table has no binary columns', async () => {
    await enrichBinaryColumns('x', [mapAttribute(attr('name', 'String', 'StringType', true))!]);
    expect(loadAttributesByTypeCast).not.toHaveBeenCalled();
  });

  it('keeps the defaults when a type cast fails', async () => {
    loadAttributesByTypeCast.mockRejectedValue(new Error('boom'));
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const c = cols();
    await enrichBinaryColumns('account', c);
    expect(c[1]).not.toHaveProperty('maxSizeInKB');
    warn.mockRestore();
  });
});

describe('mustSupplyOnCreate', () => {
  it("doesn't require File / Image columns, which can't be set in the field set", () => {
    const required = (c: ColumnMeta) =>
      ({ ...c, requiredLevel: 'ApplicationRequired' }) as ColumnMeta;
    expect(
      mustSupplyOnCreate(
        required(mapAttribute(attr('entityimage', 'Virtual', 'ImageType', false))!),
      ),
    ).toBe(false);
    expect(
      mustSupplyOnCreate(required(mapAttribute(attr('doc', 'Virtual', 'FileType', false))!)),
    ).toBe(false);
    expect(
      mustSupplyOnCreate(required(mapAttribute(attr('name', 'String', 'StringType', true))!)),
    ).toBe(true);
  });
});
