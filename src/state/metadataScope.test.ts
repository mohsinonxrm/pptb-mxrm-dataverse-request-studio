// #44: "Advanced Find only" scopes the pickers; it's off by default, and a
// value saved before it did anything is dropped once.
import { beforeAll, describe, expect, it, vi } from 'vitest';
import {
  isColumnInScope,
  isEntityInScope,
  isNavInScope,
  scopedColumns,
  scopedNavs,
} from './metadataScope';
import { SETTINGS_REVISION, defaultDisplaySettings, mergeWithDefaults } from './displaySettings';
import { mapAttribute } from '../host/metadataProvider';
import type { AttributeMetadata } from '../host/pptbClient';
import { __registerLiveTables, findTable, type TableMeta } from '../mock/metadata';
import { availableNavsAt } from '../editors/expand/expandTree';
import { relatedCollectionOptions } from '../editors/RelatedCollectionTarget';

const getEntityAttributes = vi.hoisted(() => vi.fn(async () => []));
const getEntityRelationships = vi.hoisted(() =>
  vi.fn(async () => ({ oneToMany: [], manyToOne: [], manyToMany: [] })),
);
vi.mock('../host/pptbClient', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/pptbClient')>()),
  getEntityAttributes,
  getEntityRelationships,
}));

describe('scope rules', () => {
  it('everything is in scope with the setting off', () => {
    expect(isEntityInScope({ isValidForAdvancedFind: false }, false)).toBe(true);
    expect(isColumnInScope({ isValidForAdvancedFind: false }, false)).toBe(true);
    expect(isNavInScope({ isValidForAdvancedFind: false }, false)).toBe(true);
  });

  it('with it on, only what Advanced Find shows; a missing flag counts as shown', () => {
    expect(isEntityInScope({ isValidForAdvancedFind: false }, true)).toBe(false);
    expect(isEntityInScope({}, true)).toBe(true);
    expect(isNavInScope({ isValidForAdvancedFind: false }, true)).toBe(false);
    expect(
      scopedColumns(
        [
          { logicalName: 'a', isValidForAdvancedFind: true },
          { logicalName: 'b', isValidForAdvancedFind: false },
        ],
        true,
      ).map((c) => c.logicalName),
    ).toEqual(['a']);
    expect(scopedNavs([{ isValidForAdvancedFind: false }], true)).toEqual([]);
  });

  it('Image and File columns are always offered (never Advanced Find-eligible)', () => {
    expect(isColumnInScope({ attributeType: 'Image', isValidForAdvancedFind: false }, true)).toBe(
      true,
    );
    expect(isColumnInScope({ attributeType: 'File', isValidForAdvancedFind: false }, true)).toBe(
      true,
    );
  });
});

describe('settings', () => {
  it('is off by default', () => {
    expect(defaultDisplaySettings.advancedFindOnly).toBe(false);
    expect(mergeWithDefaults(undefined).advancedFindOnly).toBe(false);
  });

  it('drops a value saved before it did anything, once', () => {
    const old = mergeWithDefaults({ advancedFindOnly: true, useLogicalNames: true });
    expect(old.advancedFindOnly).toBe(false);
    expect(old.useLogicalNames).toBe(true);
    expect(old.settingsRevision).toBe(SETTINGS_REVISION);
  });

  it('keeps a value saved since', () => {
    expect(
      mergeWithDefaults({ advancedFindOnly: true, settingsRevision: SETTINGS_REVISION })
        .advancedFindOnly,
    ).toBe(true);
  });
});

describe('metadata', () => {
  it('maps the column flag', () => {
    const raw = (af: boolean) =>
      ({
        LogicalName: 'x',
        SchemaName: 'x',
        AttributeType: 'String',
        MetadataId: 'x',
        IsValidForAdvancedFind: { Value: af },
      }) as AttributeMetadata;
    expect(mapAttribute(raw(false))?.isValidForAdvancedFind).toBe(false);
    expect(mapAttribute(raw(true))?.isValidForAdvancedFind).toBe(true);
  });

  it('loads columns and relationships unfiltered; the pickers scope them', async () => {
    const { loadEntityAttributes, loadEntityRelationships } =
      await import('../host/dataverseMetadata');
    await loadEntityAttributes('zz_scope_test');
    await loadEntityRelationships('zz_scope_test');
    expect(getEntityAttributes).toHaveBeenCalledWith('zz_scope_test', false);
    expect(getEntityRelationships).toHaveBeenCalledWith('zz_scope_test', false);
  });
});

describe('relationship pickers', () => {
  const table: TableMeta = {
    logicalName: 'zz_parent',
    entitySetName: 'zz_parents',
    displayName: 'Parent',
    primaryKey: 'zz_parentid',
    primaryName: 'zz_name',
    columns: [],
    navigationProperties: [
      {
        name: 'shown',
        relationshipName: 'shown_rel',
        cardinality: 'OneToMany',
        targetEntity: 'zz_child',
        isValidForAdvancedFind: true,
      },
      {
        name: 'hidden',
        relationshipName: 'hidden_rel',
        cardinality: 'OneToMany',
        targetEntity: 'zz_child',
        isValidForAdvancedFind: false,
      },
      {
        name: 'hidden_lookup',
        relationshipName: 'hidden_lookup_rel',
        cardinality: 'ManyToOne',
        targetEntity: 'zz_other',
        referencingAttribute: 'zz_otherid',
        isValidForAdvancedFind: false,
      },
    ],
  };
  beforeAll(() => {
    __registerLiveTables([table]);
  });

  it('$expand offers only Advanced Find relationships when on', () => {
    const names = (af: boolean) =>
      availableNavsAt('zz_parent', [], { parentCardinality: null, advancedFindOnly: af }).map(
        (n) => n.name,
      );
    expect(names(false)).toEqual(['shown', 'hidden', 'hidden_lookup']);
    expect(names(true)).toEqual(['shown']);
  });

  it('the related collection keeps the relationship already chosen', () => {
    const t = findTable('zz_parent')!;
    expect(relatedCollectionOptions(t, true)).toEqual([]);
    expect(
      relatedCollectionOptions(t, true, 'hidden_lookup_rel').map((o) => o.relationshipName),
    ).toEqual(['hidden_lookup_rel']);
    expect(relatedCollectionOptions(t, false)).toHaveLength(1);
  });
});
